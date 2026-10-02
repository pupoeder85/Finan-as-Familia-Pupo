"""
Painel financeiro da Família Pupo
---------------------------------
Lê a aba "Projecao_Mes_a_Mes" da planilha de planejamento e monta um painel
com resumo do mês, evolução, parcelas futuras, lançamentos e pontos de atenção.

Como rodar:
    pip install -r requirements.txt
    streamlit run app.py

Coloque a planilha na mesma pasta do app.py (ou informe o caminho na barra
lateral). Depois de salvar o Excel, clique em "Recarregar planilha".
"""

from __future__ import annotations

import datetime as dt
import html
import re
import unicodedata
from collections import defaultdict
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ----------------------------------------------------------------------------
# Configuração
# ----------------------------------------------------------------------------
ARQUIVO_PADRAO = "Planejamento_familiar_Pupo.xlsx"

MESES_ABREV = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
MESES_NOME = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho",
              "agosto", "setembro", "outubro", "novembro", "dezembro"]

COR = {"fixo": "#3D5A8A", "cartao": "#D5533F", "invest": "#C99419",
       "sobra": "#1E7A5F", "alerta": "#B4412F", "cinza": "#8A9691", "claro": "#DCE3E0"}
CORES_CAT = ["#3D5A8A", "#D5533F", "#1E7A5F", "#C99419", "#7A4FA0", "#2C8BA6", "#B5577E",
             "#6B7F3A", "#A8672E", "#4F6470", "#8E3B46", "#3F8F7F", "#9C8AC0", "#7B6B5D"]
COR_PAGTO = {"Visa": COR["cartao"], "Amex": COR["fixo"], "Mercado Pago": COR["invest"],
             "Elo": COR["sobra"], "Débito/boleto": COR["cinza"]}

# Regras de categoria: a primeira regra que casar com a descrição vence.
# Palavras com até 3 letras precisam casar como palavra inteira.
# Para ensinar uma loja nova, basta incluir a palavra na lista certa.
REGRAS_CATEGORIA = [
    ("Salário", ["salario"]),
    ("Investimento", ["investimento"]),
    ("Consórcio", ["consorcio"]),
    ("Transporte", ["carro", "gasolina", "uber", "veloe", "estar"]),
    ("Moradia", ["condominio", "casa", "copel", "sanepar", "internet", "sky", "diarista", "gas", "vivo"]),
    ("Pet", ["pelinho"]),
    ("Vestuário", ["calca", "c&a", "netshoes", "shein", "moda mundial", "jeans", "pernambucanas"]),
    ("Compras e casa", ["niver", "presente", "sofa", "madeira", "panela", "havan", "mercado", "shopee",
                        "conceito", "palie", "monitor", "pc"]),
    ("Lazer e viagens", ["playstation", "airbnb", "decolar", "latam", "ticketmaster", "acdc", "cinema",
                         "foco radical"]),
    ("Assinaturas", ["apple", "netflix", "play", "amazon", "microsoft", "coxa", "ea", "steam", "google"]),
    ("Saúde e beleza", ["academia", "farmacia", "cabelo", "barbearia", "embelezamento", "cremes", "fralda"]),
    ("Educação e igreja", ["pos", "escolinha", "igreja", "dizimo"]),
    ("Alimentação", ["burger", "coca", "pizza", "panificadora", "jacomar", "pao delicia", "frango",
                     "mini preco", "fratelli", "jantar", "boteco", "cafe", "vila bella", "refri"]),
    ("Consertos", ["conserto"]),
    ("Impostos e taxas", ["ir", "anuidade", "cartorio", "associacao"]),
]

TIPOS = {"cartao de credito": "Cartão", "gasto fixo": "Fixo",
         "salario": "Receita", "investimento": "Investimento"}


# ----------------------------------------------------------------------------
# Utilitários
# ----------------------------------------------------------------------------
def norm(s) -> str:
    """Minúsculas, sem acento e sem espaços nas pontas."""
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    s = unicodedata.normalize("NFKD", str(s))
    return "".join(c for c in s if not unicodedata.combining(c)).lower().strip()


def brl(v: float, casas: int = 0) -> str:
    txt = f"{abs(v):,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return ("-R$ " if v < 0 else "R$ ") + txt


def sem_latex(txt: str) -> str:
    """No st.markdown, texto entre dois "$" vira fórmula matemática. Escapa o "$" de "R$"."""
    return txt.replace("$", "\\$")


def rotulo(k: int) -> str:
    """202609 -> 'Set/26'"""
    return f"{MESES_ABREV[k % 100 - 1]}/{str(k // 100)[2:]}"


def nome_mes(k: int) -> str:
    return MESES_NOME[k % 100 - 1]


def somar_meses(k: int, n: int) -> int:
    y, m = divmod(k // 100 * 12 + (k % 100 - 1) + n, 12)
    return y * 100 + m + 1


def diff_meses(a: int, b: int) -> int:
    return (b // 100 - a // 100) * 12 + (b % 100 - a % 100)


def chave_mes(v) -> int | None:
    """Converte 'Ago/2026', datas do Excel (número serial) ou datetime em 202608."""
    if v is None or (isinstance(v, float) and pd.isna(v)) or v == "":
        return None
    if isinstance(v, (dt.datetime, dt.date, pd.Timestamp)):
        return v.year * 100 + v.month
    if isinstance(v, (int, float)):
        d = dt.datetime(1899, 12, 30) + dt.timedelta(days=round(float(v)))
        return d.year * 100 + d.month
    m = re.match(r"^([A-Za-zçÇ]{3})[a-zç]*/(\d{2,4})$", str(v).strip())
    if not m:
        return None
    abrev = norm(m.group(1))[:3]
    nomes = [x.lower() for x in MESES_ABREV]
    if abrev not in nomes:
        return None
    ano = int(m.group(2))
    ano += 2000 if ano < 100 else 0
    return ano * 100 + nomes.index(abrev) + 1


def categoria(desc: str) -> str:
    x = norm(desc)
    for cat, palavras in REGRAS_CATEGORIA:
        for p in palavras:
            if len(p) <= 3:
                if re.search(r"(^|[^a-z0-9])" + re.escape(p) + r"($|[^a-z0-9])", x):
                    return cat
            elif p in x:
                return cat
    return "Outros"


# ----------------------------------------------------------------------------
# Leitura da planilha
# ----------------------------------------------------------------------------
def ler_planilha(origem) -> pd.DataFrame:
    """Lê a aba de projeção e devolve um DataFrame limpo, um lançamento por linha."""
    xls = pd.ExcelFile(origem, engine="openpyxl")
    aba = next((n for n in xls.sheet_names if "projecao" in norm(n)), None)
    if aba is None:
        raise ValueError('Não encontrei a aba "Projecao_Mes_a_Mes" na planilha.')
    bruto = pd.read_excel(xls, sheet_name=aba, header=None, dtype=object)
    cab = [norm(c) for c in bruto.iloc[0].tolist()]

    def col(prefixos, padrao):
        for i, h in enumerate(cab):
            if any(h.startswith(p) for p in prefixos):
                return i
        return padrao

    C = dict(m=col(["mes projetado"], 0), d=col(["descricao"], 4), t=col(["tipo"], 5),
             cc=col(["cartao"], 6), p=col(["parcela"], 7), v=col(["valor"], 8),
             ate=col(["ate qual"], 9), pa=col(["compra parcelada"], 10))

    def celula(linha, nome):
        i = C[nome]
        v = linha[i] if i < len(linha) else None
        return None if (isinstance(v, float) and pd.isna(v)) else v

    linhas = []
    for linha in bruto.iloc[1:].itertuples(index=False):
        linha = list(linha)
        k = chave_mes(celula(linha, "m"))
        desc = str(celula(linha, "d") or "").strip()
        tipo = TIPOS.get(norm(celula(linha, "t")))
        if not k or not desc or not tipo:
            continue
        try:
            valor = round(float(celula(linha, "v") or 0), 2)
        except (TypeError, ValueError):
            valor = 0.0
        if valor == 0 and tipo != "Receita":
            continue
        cc = str(celula(linha, "cc") or "").strip() or ("Débito/boleto" if tipo == "Fixo" else "")
        linhas.append(dict(
            k=k, mes=rotulo(k), desc=desc,
            cat="Salário" if tipo == "Receita" else categoria(desc),
            tipo=tipo, pagto=cc, parcela=str(celula(linha, "p") or "").strip(),
            valor=valor, parcelada=norm(celula(linha, "pa")) == "sim",
            ate=str(celula(linha, "ate") or "").strip(),
        ))
    df = pd.DataFrame(linhas)
    if df.empty:
        raise ValueError("A aba de projeção não tem lançamentos válidos.")
    return corrigir_parcelas(df)


def corrigir_parcelas(df: pd.DataFrame) -> pd.DataFrame:
    """Corrige numerações como '7/4' usando uma parcela válida da mesma série."""
    padrao = df["parcela"].str.fullmatch(r"\d+/\d+")
    for _, g in df[padrao].groupby(["desc", "valor"]):
        partes = g["parcela"].str.split("/", expand=True).astype(int)
        validas = g[partes[0] <= partes[1]]
        if validas.empty:
            continue
        tot = partes.loc[validas.index, 1].max()
        ref = validas[partes.loc[validas.index, 1] == tot].sort_values("k").iloc[0]
        n0 = int(ref["parcela"].split("/")[0])
        for idx, r in g.iterrows():
            a, b = partes.loc[idx]
            if a > b or b != tot:
                n = n0 + diff_meses(ref["k"], r["k"])
                if 1 <= n <= tot:
                    df.at[idx, "parcela"] = f"{n}/{tot}"
    return df


# ----------------------------------------------------------------------------
# Cálculos
# ----------------------------------------------------------------------------
def eh_invest(df: pd.DataFrame) -> pd.Series:
    # Só conta como investimento o que está com Tipo = "Investimento" na planilha.
    return df["tipo"] == "Investimento"


def resumo_mes(df: pd.DataFrame, k: int) -> dict:
    m = df[df["k"] == k]
    inv = eh_invest(m)
    rec = m.loc[m["tipo"] == "Receita", "valor"].sum()
    fixo = -m.loc[(m["tipo"] == "Fixo") & ~inv, "valor"].sum()
    cart = -m.loc[(m["tipo"] == "Cartão") & ~inv, "valor"].sum()
    # Investimento sempre soma, esteja lançado como negativo (-426) ou positivo (426) na planilha.
    invest = m.loc[inv, "valor"].abs().sum()
    # Investimento fica à parte: é mostrado, mas não entra no cálculo da sobra.
    return dict(linhas=m, rec=rec, fixo=fixo, cart=cart, inv=invest,
                sobra=rec - fixo - cart)


def montar_modelo(df: pd.DataFrame) -> dict:
    hoje = dt.date.today()
    k_hoje = hoje.year * 100 + hoje.month
    meses = sorted(df.loc[(df["tipo"] == "Receita") & (df["valor"] > 0), "k"].unique().tolist())
    ultimo = meses[-1] if meses else k_hoje
    inicio = max(meses[0] if meses else k_hoje, somar_meses(k_hoje, 1))
    fim = somar_meses(inicio, 14)

    base = df.loc[(df["k"] == ultimo) & (df["tipo"] == "Receita"), "valor"].sum()
    sempre = df["ate"].map(norm) == "sempre"
    inv_df = eh_invest(df)
    rec_estimado = -df.loc[(df["k"] == ultimo) & (df["tipo"] != "Receita") & sempre & ~inv_df, "valor"].sum()

    futuro = []
    k = inicio
    while k <= fim:
        m = df[(df["k"] == k) & (df["tipo"] != "Receita") & ~inv_df]
        carro = -m.loc[m["desc"] == "Carro", "valor"].sum()
        parc = -m.loc[(m["desc"] != "Carro") & m["parcelada"], "valor"].sum()
        real = k in meses
        rec = -m.loc[(m["desc"] != "Carro") & ~m["parcelada"], "valor"].sum() if real else rec_estimado
        futuro.append(dict(k=k, mes=rotulo(k), carro=carro, parcelas=parc, fixos=rec, estimado=not real))
        k = somar_meses(k, 1)
    futuro = pd.DataFrame(futuro)

    p = df[df["parcelada"] & (df["desc"] != "Carro") & (df["k"] >= inicio) & (df["valor"] < 0)]
    parcelas = (p.groupby(["desc", "valor"])
                 .agg(fim=("k", "max"), restantes=("k", "count"), pagto=("pagto", "first"))
                 .reset_index())
    parcelas["mensal"] = -parcelas["valor"]
    parcelas["total_restante"] = parcelas["mensal"] * parcelas["restantes"]
    parcelas = parcelas.sort_values(["fim", "mensal"], ascending=[True, False])

    carro = df[df["desc"] == "Carro"].sort_values("k")
    passados = [m for m in meses if m <= k_hoje]
    dia_a_dia = [
        -df.loc[(df["k"] == m) & (df["tipo"] == "Cartão") & ~df["parcelada"] & ~sempre, "valor"].sum()
        for m in passados
    ]
    return dict(df=df, meses=meses, k_hoje=k_hoje, base=base, futuro=futuro,
                parcelas=parcelas, carro=carro,
                dia_a_dia=sum(dia_a_dia) / len(dia_a_dia) if dia_a_dia else 0)


def pontos_de_atencao(modelo: dict) -> list[str]:
    df, meses = modelo["df"], modelo["meses"]
    notas = []
    parc = df[df["parcelada"] & df["parcela"].str.fullmatch(r"\d+/\d+")]

    # várias parcelas da mesma compra no mesmo mês
    por_mes = defaultdict(list)
    for (k, d, v), g in parc.groupby(["k", "desc", "valor"]):
        if len(g) > 1 and g["parcela"].nunique() > 1:
            por_mes[k].append(f"{d} ({', '.join(g['parcela'])})")
    for k, itens in sorted(por_mes.items()):
        notas.append(f"Em {rotulo(k).lower()} há compras com várias parcelas lançadas no mesmo mês: "
                     f"{'; '.join(itens)}. Isso infla esse mês e esvazia os seguintes.")

    # mesmo valor e mesma parcela em descrições diferentes
    for (k, v, p), g in parc[parc["valor"] < 0].groupby(["k", "valor", "parcela"]):
        nomes = g["desc"].unique()
        if len(nomes) > 1:
            notas.append(f"{' e '.join(nomes)} aparecem com o mesmo valor ({brl(-v, 2)}) e parcela {p} "
                         f"em {rotulo(k).lower()}. Confira se não é a mesma compra duplicada.")

    # gastos recorrentes que sumiram em algum mês
    rec = df[(df["ate"].map(norm) == "sempre") & (df["tipo"] != "Receita")]
    recorrentes = [d for d, g in rec.groupby("desc") if g["k"].isin(meses).sum() >= len(meses) - 2]
    for k in meses:
        tem = set(df.loc[df["k"] == k, "desc"])
        faltam = [d for d in recorrentes if d not in tem]
        if 0 < len(faltam) < 6:
            notas.append(f"Em {rotulo(k).lower()} faltam gastos que se repetem nos outros meses: "
                         f"{', '.join(faltam)}. Se existiram, a sobra real é menor.")

    carro = modelo["carro"]
    if not carro.empty:
        ult = carro.iloc[-1]
        n = (carro["valor"] == ult["valor"]).sum()
        notas.append(f"O financiamento do carro ({brl(-ult['valor'])}/mês) vai até {ult['mes'].lower()}, "
                     f"cerca de {brl(-ult['valor'] * n)} no total contando os meses lançados.")
    return notas


# ----------------------------------------------------------------------------
# Gráficos
# ----------------------------------------------------------------------------
LAYOUT = dict(margin=dict(l=10, r=10, t=10, b=10), plot_bgcolor="rgba(0,0,0,0)",
              paper_bgcolor="rgba(0,0,0,0)", font=dict(size=13),
              legend=dict(orientation="h", y=-0.15, x=0.5, xanchor="center"),
              separators=",.")


def graf_categorias(r: dict, cores: dict) -> go.Figure:
    m = r["linhas"]
    m = m[m["tipo"].isin(["Fixo", "Cartão"]) & ~eh_invest(m)]
    s = (-m.groupby("cat")["valor"].sum()).loc[lambda x: x > 0.5].sort_values()
    fig = go.Figure(go.Bar(x=s.values, y=s.index, orientation="h",
                           marker_color=[cores.get(c, COR["cinza"]) for c in s.index],
                           hovertemplate="%{y}: R$ %{x:,.2f}<extra></extra>"))
    fig.update_layout(**LAYOUT, height=max(320, 28 * len(s) + 60), xaxis_tickprefix="R$ ",
                      xaxis_tickformat=",.0f")
    return fig


def graf_evolucao(modelo: dict) -> go.Figure:
    df, meses = modelo["df"], modelo["meses"]
    R = [resumo_mes(df, k) for k in meses]
    x = [nome_mes(k).capitalize() for k in meses]
    fig = go.Figure()
    for nome, chave, cor in [("Gastos fixos", "fixo", COR["fixo"]), ("Cartões", "cart", COR["cartao"])]:
        fig.add_bar(name=nome, x=x, y=[r[chave] for r in R], marker_color=cor,
                    hovertemplate=nome + ": R$ %{y:,.2f}<extra></extra>")
    fig.add_scatter(name="Receita", x=x, y=[r["rec"] for r in R], mode="lines+markers",
                    line=dict(color="#15201C", dash="dash"),
                    hovertemplate="Receita: R$ %{y:,.2f}<extra></extra>")
    fig.add_scatter(name="Sobra", x=x, y=[r["sobra"] for r in R], mode="lines+markers",
                    line=dict(color=COR["sobra"], width=3), marker=dict(size=9),
                    hovertemplate="Sobra: R$ %{y:,.2f}<extra></extra>")
    fig.add_scatter(name="Investimento acumulado", x=x,
                    y=[modelo["inv_acumulado"][k]["total"] for k in meses], mode="lines+markers",
                    line=dict(color=COR["invest"], width=2, dash="dot"), marker=dict(size=8),
                    hovertemplate="Investimento acumulado: R$ %{y:,.2f}<extra></extra>")
    fig.update_layout(**LAYOUT, barmode="stack", height=380, yaxis_tickprefix="R$ ",
                      yaxis_tickformat=",.0f", hovermode="x unified")
    return fig


def graf_futuro(modelo: dict) -> go.Figure:
    f, base = modelo["futuro"], modelo["base"]
    fig = go.Figure()
    fig.add_bar(name="Carro", x=f["mes"], y=f["carro"], marker_color=COR["fixo"])
    fig.add_bar(name="Parcelas de cartão", x=f["mes"], y=f["parcelas"], marker_color=COR["cartao"])
    fig.add_bar(name="Fixos e assinaturas (claro = estimado)", x=f["mes"], y=f["fixos"],
                marker_color=[COR["claro"] if e else COR["cinza"] for e in f["estimado"]])
    fig.add_scatter(name=f"Receita (base {brl(base)})", x=f["mes"], y=[base] * len(f), mode="lines",
                    line=dict(color="#15201C", dash="dash"))
    fig.update_traces(selector=dict(type="bar"), hovertemplate="%{fullData.name}: R$ %{y:,.2f}<extra></extra>")
    fig.update_layout(**LAYOUT, barmode="stack", height=420, yaxis_tickprefix="R$ ",
                      yaxis_tickformat=",.0f", hovermode="x unified")
    return fig


# ----------------------------------------------------------------------------
# Interface
# ----------------------------------------------------------------------------
st.set_page_config(page_title="Finanças da Família Pupo", page_icon="🏡", layout="wide")

# O Streamlit marca a página como inglês, e o Chrome/Edge oferece "traduzir",
# o que estraga os textos ("Set" vira "Conjunto", "Out" vira "Saída").
# Este trecho marca a página como português e desliga a tradução automática.
_SEM_TRADUCAO = """<script>
const d = window.parent.document;
d.documentElement.lang = "pt-BR";
d.documentElement.setAttribute("translate", "no");
d.documentElement.classList.add("notranslate");
if (!d.querySelector('meta[name="google"]')) {
  const m = d.createElement("meta"); m.name = "google"; m.content = "notranslate"; d.head.appendChild(m);
}
</script>"""
if hasattr(st, "iframe"):  # Streamlit mais novo
    st.iframe(_SEM_TRADUCAO, height=1)
else:  # versões antigas
    import streamlit.components.v1 as components
    components.html(_SEM_TRADUCAO, height=0)
st.markdown("""
<style>
.block-container{padding-top:2rem;max-width:1200px}
.hero{font-size:clamp(22px,3vw,34px);font-weight:800;line-height:1.2;letter-spacing:-.02em;margin:.2rem 0 1rem;max-width:30ch}
.ribbon{display:flex;height:48px;border-radius:12px;overflow:hidden;margin-bottom:1rem}
.ribbon div{display:flex;align-items:center;padding-left:12px;color:#fff;font-weight:700;font-size:14px;white-space:nowrap;overflow:hidden}
.tl{display:grid;grid-template-columns:70px 1fr auto;gap:12px;padding:10px 0;border-top:1px solid rgba(128,128,128,.25)}
.tl b.mo{font-size:16px}.tl .it{opacity:.75;font-size:14px}.tl .fr{font-weight:700;color:#1E7A5F;text-align:right;white-space:nowrap}
.tl .fr small{display:block;font-weight:400;opacity:.7}
</style>""", unsafe_allow_html=True)


@st.cache_data(show_spinner="Lendo a planilha…")
def carregar_do_disco(caminho: str, _mtime: float) -> pd.DataFrame:
    # _mtime entra na chave do cache: se o arquivo mudar, a leitura é refeita.
    return ler_planilha(caminho)


with st.sidebar:
    st.header("Planilha")
    caminho = st.text_input("Caminho do arquivo",
                            value=str(Path(__file__).with_name(ARQUIVO_PADRAO)))
    if st.button("Recarregar planilha", type="primary", width="stretch"):
        st.cache_data.clear()
        st.rerun()
    enviado = st.file_uploader("Ou envie um arquivo", type=["xlsx", "xlsm"])
    st.divider()
    saldo_inicial_inv = st.number_input(
        "Saldo já investido antes do primeiro mês (R$)", min_value=0.0, value=0.0, step=100.0,
        format="%.2f", help="Ponto de partida do investimento acumulado. Cada mês soma o aporte "
                            "do mês ao saldo anterior.")
    st.caption("Depois de salvar o Excel, clique em Recarregar. "
               "Se o arquivo mudou no disco, o painel também relê sozinho na próxima interação.")

try:
    if enviado is not None:
        df = ler_planilha(enviado)
        fonte = enviado.name
    else:
        p = Path(caminho)
        if not p.exists():
            st.info(f"Não encontrei **{p.name}**. Coloque a planilha na pasta do app, "
                    "informe o caminho na barra lateral ou envie o arquivo.")
            st.stop()
        df = carregar_do_disco(str(p), p.stat().st_mtime)
        fonte = f"{p.name} (salvo em {dt.datetime.fromtimestamp(p.stat().st_mtime):%d/%m/%Y %H:%M})"
except Exception as e:  # noqa: BLE001
    st.error(f"Não consegui ler a planilha: {e}")
    st.stop()

modelo = montar_modelo(df)

# Investimento acumulado: saldo anterior + aporte do mês
acumulado, saldo = {}, saldo_inicial_inv
for _m in modelo["meses"]:
    aporte = resumo_mes(df, _m)["inv"]
    acumulado[_m] = dict(anterior=saldo, aporte=aporte, total=saldo + aporte)
    saldo += aporte
modelo["inv_acumulado"] = acumulado
meses = modelo["meses"]
if not meses:
    st.warning("A planilha não tem nenhum mês com Salário lançado.")
    st.stop()

cats = sorted(df.loc[~df["tipo"].isin(["Receita"]) & ~eh_invest(df), "cat"].unique())
cores_cat = {c: CORES_CAT[i % len(CORES_CAT)] for i, c in enumerate(cats)}

# Cabeçalho
a, z = meses[0], meses[-1]
st.title("Finanças da Família Pupo")
st.caption(f"Receitas, gastos e parcelas de {nome_mes(a)} de {a // 100} a {nome_mes(z)} de {z // 100}, "
           f"com o que já está comprometido nos meses seguintes. Fonte: {fonte}.")

padrao = modelo["k_hoje"] if modelo["k_hoje"] in meses else ([m for m in meses if m < modelo["k_hoje"]] or meses)[-1]
varios_anos = len({m // 100 for m in meses}) > 1


def nome_aba(x: int) -> str:
    nome = nome_mes(x).capitalize()
    return f"{nome} {x // 100}" if varios_anos else nome


k = st.radio("Mês", meses, index=meses.index(padrao), horizontal=True,
             format_func=nome_aba, label_visibility="collapsed")

# ---------------- Resumo do mês ----------------
r = resumo_mes(df, k)
if r["sobra"] >= 0:
    frase = (f"Em {nome_mes(k)} entram {brl(r['rec'])} e sobram "
             f"<span style='color:{COR['sobra']}'>{brl(r['sobra'])}</span> depois de pagar tudo.")
else:
    frase = (f"Em {nome_mes(k)} entram {brl(r['rec'])} e faltam "
             f"<span style='color:{COR['alerta']}'>{brl(-r['sobra'])}</span> para fechar o mês.")
st.markdown(f"<div class='hero'>{frase}</div>", unsafe_allow_html=True)
if r["inv"]:
    ia_h = modelo["inv_acumulado"][k]
    st.markdown((
                f"<div>À parte, <b style='color:{COR['invest']}'>{brl(r['inv'])}</b> foram para investimentos. "
                f"Somando ao saldo anterior de {brl(ia_h['anterior'])}, o total investido chega a "
                f"<b style='color:{COR['invest']}'>{brl(ia_h['total'])}</b>.</div>"), unsafe_allow_html=True)

partes = [("Gastos fixos", r["fixo"], COR["fixo"]), ("Cartões", r["cart"], COR["cartao"]),
          ("Sobra", max(r["sobra"], 0), COR["sobra"])]
tot = max(r["rec"], r["fixo"] + r["cart"]) or 1
def _faixa(nome, v, cor):
    pct = f"{v / r['rec'] * 100:.0f}%" if r["rec"] and v / tot > .09 else ""
    return f"<div style='flex:0 0 {v / tot * 100:.2f}%;background:{cor}' title='{nome}: {brl(v, 2)}'>{pct}</div>"


faixas = "".join(_faixa(n, v, c) for n, v, c in partes)
st.markdown(f"<div class='ribbon'>{faixas}</div>", unsafe_allow_html=True)

cols = st.columns(5)
for col, (n, v, _) in zip(cols, partes):
    col.metric(n, brl(v))
ia = modelo["inv_acumulado"][k]
cols[3].metric("Investimento acumulado", brl(ia["total"]), delta=f"+{brl(ia['aporte'])} no mês",
               help="Saldo anterior + aporte do mês. Fica à parte e não entra no cálculo da sobra.")
cols[4].metric("Receita comprometida", f"{(r['fixo'] + r['cart']) / r['rec'] * 100:.0f}%" if r["rec"] else "—")

c1, c2 = st.columns([7, 5])
with c1:
    st.subheader("Para onde vai o dinheiro")
    st.caption("Gastos do mês por categoria")
    st.plotly_chart(graf_categorias(r, cores_cat), width="stretch", config={"displayModeBar": False})
with c2:
    st.subheader("Como foi pago")
    m = r["linhas"][r["linhas"]["tipo"].isin(["Fixo", "Cartão"])]
    pag = (-m.groupby("pagto")["valor"].sum()).loc[lambda x: x > 0.5].sort_values(ascending=False)
    fig = go.Figure(go.Bar(x=pag.values, y=pag.index, orientation="h", text=[brl(v, 2) for v in pag.values],
                           marker_color=[COR_PAGTO.get(i, COR["cinza"]) for i in pag.index],
                           hovertemplate="%{y}: R$ %{x:,.2f}<extra></extra>"))
    fig.update_layout(**LAYOUT, height=60 + 44 * len(pag), yaxis_autorange="reversed",
                      xaxis_visible=False)
    st.plotly_chart(fig, width="stretch", config={"displayModeBar": False})
    st.subheader("Maiores gastos do mês")
    top = r["linhas"][(r["linhas"]["valor"] < 0) & ~eh_invest(r["linhas"])].nsmallest(5, "valor")
    st.dataframe(pd.DataFrame({"Descrição": top["desc"], "Parcela": top["parcela"].replace("", "—"),
                               "Valor": -top["valor"]}),
                 hide_index=True, width="stretch",
                 column_config={"Valor": st.column_config.NumberColumn(format="R$ %.2f")})

# ---------------- Evolução ----------------
st.header("Mês a mês")
st.caption("Barras: gastos empilhados. Linha verde: o que sobra depois dos gastos. "
           "Linha pontilhada dourada: investimento acumulado (saldo anterior + aporte do mês), "
           "mostrado à parte e não descontado da sobra.")
st.plotly_chart(graf_evolucao(modelo), width="stretch", config={"displayModeBar": False})

st.subheader("Investimentos mês a mês")
st.caption("Cada mês soma o aporte ao saldo anterior. O saldo inicial pode ser ajustado na barra lateral.")
st.dataframe(
    pd.DataFrame([{"Mês": nome_aba(m), "Saldo anterior": v["anterior"], "Aporte do mês": v["aporte"],
                   "Saldo acumulado": v["total"]} for m, v in modelo["inv_acumulado"].items()]),
    hide_index=True, width="stretch",
    column_config={c: st.column_config.NumberColumn(format="R$ %.2f")
                   for c in ["Saldo anterior", "Aporte do mês", "Saldo acumulado"]})

# ---------------- Futuro ----------------
st.header("Parcelas e compromissos futuros")
f1, f2 = st.columns([7, 5])
with f1:
    st.subheader("Quanto já está comprometido")
    fut = modelo["futuro"]
    est = fut[fut["estimado"]]
    if not est.empty:
        st.caption(f"De {est.iloc[0]['mes'].lower()} em diante, os gastos fixos são estimados repetindo "
                   f"{rotulo(z).lower()}, porque a planilha ainda não os projeta.")
    st.plotly_chart(graf_futuro(modelo), width="stretch", config={"displayModeBar": False})
    if not est.empty:
        idx = sorted({0, len(est) // 3, 2 * len(est) // 3, len(est) - 1})
        cs = st.columns(len(idx))
        for c, i in zip(cs, idx):
            row = est.iloc[i]
            c.metric(f"Livre em {row['mes'].lower()}",
                     brl(modelo["base"] - row["carro"] - row["parcelas"] - row["fixos"]))
        if modelo["dia_a_dia"]:
            st.caption("Esse valor ainda precisa cobrir o dia a dia não parcelado (gasolina, mercado, "
                       f"restaurantes). Nos meses já fechados isso somou em média "
                       f"{brl(round(modelo['dia_a_dia'], -2))} por mês no cartão.")
with f2:
    st.subheader("Quando cada parcela acaba")
    st.caption("Valor que deixa de sair do orçamento no mês seguinte")
    acc, blocos = 0.0, []
    for fim, g in modelo["parcelas"].groupby("fim"):
        v = g["mensal"].sum()
        acc += v
        nomes = [f"{html.escape(re.sub(r'[*].*', '', d).strip())} ({brl(x)})" for d, x in zip(g["desc"], g["mensal"])]
        txt = ", ".join(nomes[:4]) + (f" e mais {len(nomes) - 4}" if len(nomes) > 4 else "")
        blocos.append(f"<div class='tl'><b class='mo'>{rotulo(fim)}</b><div class='it'>{txt}</div>"
                      f"<div class='fr'>+{brl(v)}/mês<small>{brl(acc)} liberados</small></div></div>")
    carro = modelo["carro"]
    if not carro.empty:
        u = carro.iloc[-1]
        blocos.append(f"<div class='tl'><b class='mo'>{u['mes']}</b><div class='it'>Última parcela do carro</div>"
                      f"<div class='fr'>+{brl(-u['valor'])}/mês</div></div>")
    st.markdown("".join(blocos) or "Nenhuma parcela futura encontrada.", unsafe_allow_html=True)

# ---------------- Lançamentos ----------------
st.header("Lançamentos")
m = df[df["k"] == k]
t1, t2, t3 = st.columns([3, 2, 2])
busca = t1.text_input("Buscar descrição", placeholder="ex.: gasolina")
f_cat = t2.selectbox("Categoria", ["Todas"] + sorted(m["cat"].unique()))
f_tipo = t3.selectbox("Tipo", ["Todos", "Fixo", "Cartão", "Investimento", "Receita"])
if busca:
    m = m[m["desc"].map(norm).str.contains(norm(busca), regex=False)]
if f_cat != "Todas":
    m = m[m["cat"] == f_cat]
if f_tipo != "Todos":
    m = m[m["tipo"] == f_tipo]
st.dataframe(
    m.sort_values("valor")[["desc", "cat", "tipo", "pagto", "parcela", "valor"]]
     .rename(columns={"desc": "Descrição", "cat": "Categoria", "tipo": "Tipo", "pagto": "Pagamento",
                      "parcela": "Parcela", "valor": "Valor"}),
    hide_index=True, width="stretch", height=min(560, 38 + 35 * len(m)),
    column_config={"Valor": st.column_config.NumberColumn(format="R$ %.2f")})
inv_sel = eh_invest(m)
txt_inv = f", investido à parte {brl(-m.loc[inv_sel, 'valor'].sum(), 2)}" if inv_sel.any() else ""
st.caption(f"{len(m)} lançamentos em {nome_mes(k)}, saldo da seleção {brl(m.loc[~inv_sel, 'valor'].sum(), 2)}{txt_inv}")

# ---------------- Pontos de atenção ----------------
st.header("Pontos de atenção na planilha")
notas = pontos_de_atencao(modelo)
st.markdown(sem_latex("\n".join(f"- {n}" for n in notas)) if notas else "Nada fora do comum nesta versão da planilha.")
