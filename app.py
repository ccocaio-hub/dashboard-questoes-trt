"""
Painel de Questões (TRT, FCC, Analista e Técnico) — Streamlit

Rodar:   streamlit run app.py
Precisa: questoes_completo.xlsx (aba "Questões") na mesma pasta deste arquivo.
Libs:    streamlit, pandas, openpyxl, plotly
"""
import html
import math
import re
import statistics
import unicodedata
import zlib
from collections import Counter, defaultdict
from io import BytesIO
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Painel de Questões", page_icon="📊", layout="wide")

# ───────────────────────────── configuração ─────────────────────────────
ARQUIVO = Path(__file__).parent / "questoes_completo.xlsx"
ABA = "Questões"

LIMIAR_OUTRAS = 3.0   # % mínimo do universo para uma matéria (ou assunto nível 2) ter fatia/barra própria
COR_OUTRAS = "#7f8c8d"
PAGINAS = ["🥧 Pizza", "📊 Barras", "⚖️ Comparar"]
PAG_COMPARAR = "⚖️ Comparar"
CMP_CORES = ["#4e79a7", "#f28e2b", "#59a14f", "#b07aa1"]            # uma cor por cenário
CMP_MODOS = ["Filtros próprios", "Total da base", "Resto (fora do cenário 1)"]
CMP_CAMPOS = [("escol", "escolaridade", "Escolaridade"), ("orgao", "orgao", "Órgão"),
              ("cargo", "Cargo", "Cargo"), ("ano", "ano", "Ano")]
HOVER = dict(bgcolor="white", bordercolor="#888888", font=dict(color="black", size=13))  # balão do mouse
COR_DEMAIS = "#8395a7"
PALETA = [
    "#4e79a7", "#f28e2b", "#e15759", "#76b7b2", "#59a14f", "#edc948", "#b07aa1", "#ff9da7",
    "#9c755f", "#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd", "#8c564b", "#e377c2",
    "#17becf", "#bcbd22", "#aec7e8", "#ffbb78", "#98df8a", "#ff9896", "#c5b0d5", "#86bcb6",
]
PALETA_DESTAQUE = [
    "#ff6b35", "#00b4d8", "#ffd166", "#06d6a0", "#c77dff", "#ef476f",
    "#8ac926", "#1982c4", "#ff9f1c", "#f15bb5", "#2ec4b6", "#9b5de5",
]

NIVEIS = ["materia", "nivel2", "nivel3", "nivel4", "nivel5"]
OBRIGATORIAS = ["id_questao", "anulada", "orgao", "Cargo", "ano", "escolaridade"] + NIVEIS
CHAVES = {"f_escol": "escolaridade", "f_orgao": "orgao", "f_cargo": "Cargo", "f_ano": "ano", "f_materia": "materia"}

CSS = """
<style>
.lg{display:flex;flex-wrap:wrap;gap:8px 24px;margin:4px 0 12px 0}
.lg-col{flex-direction:column;gap:7px;margin-top:8px}
.lg-i{display:flex;align-items:center;gap:8px}
.lg-c{width:14px;height:14px;border-radius:3px;flex:none}
.lg-v{opacity:.65;font-size:.9em}
.tt{font-size:.95rem;margin-top:4px}
.tt .cab,.tt summary,.tt .fol{display:flex;align-items:center;gap:10px;padding:4px 6px}
.tt .cab{font-weight:600;border-bottom:1px solid rgba(128,128,128,.45)}
.tt .nm{flex:1;min-width:0}
.tt .q{width:90px;text-align:right}
.tt .p{width:110px;text-align:right}
.tt summary{cursor:pointer;list-style:none}
.tt summary::-webkit-details-marker{display:none}
.tt summary::before{content:"▸";display:inline-block;width:1em;flex:none;opacity:.7}
.tt details[open]>summary::before{content:"▾"}
.tt summary:hover{background:rgba(128,128,128,.14)}
.tt .fol{padding-left:calc(6px + 1em + 10px)}
.tt .ch{margin-left:1.1rem;border-left:1px solid rgba(128,128,128,.3)}
.tt .raiz>summary,.tt .raiz>.fol{font-weight:600}
.tt .outras>summary .nm,.tt .fol.outras .nm{font-style:italic}
/* balão do mouse (hover) do Plotly: fundo branco e texto preto em qualquer tema/navegador */
.js-plotly-plot .hoverlayer .hovertext rect,
.js-plotly-plot .hoverlayer .hovertext path{fill:#ffffff !important;fill-opacity:1 !important;stroke:#888888 !important}
.js-plotly-plot .hoverlayer .hovertext text,
.js-plotly-plot .hoverlayer .hovertext tspan{fill:#000000 !important;stroke:none !important}
/* tabela de comparação */
.tc{overflow-x:auto}
.tc .cab,.tc summary,.tc .fol{min-width:640px}
.tt .c{width:122px;text-align:right;white-space:nowrap}
.tt .c small{opacity:.6}
.tt .d{width:78px;text-align:right;white-space:nowrap}
.tt .d.pos{color:#2e9e5b;font-weight:600}
.tt .d.neg{color:#e0554f;font-weight:600}
.tt .d.ruido{opacity:.5;font-weight:400}
.tt .d.nd,.tt .d.z{opacity:.6;font-weight:400}
.tt .cab .c{overflow:hidden;text-overflow:ellipsis}
</style>
"""


# ───────────────────────────── utilidades ─────────────────────────────
def fmt_int(n):
    return f"{int(n):,}".replace(",", ".")


def fmt_pct(x, casas=1):
    return f"{x:.{casas}f}".replace(".", ",") + "%"


def chave_orgao(o):
    m = re.search(r"\d+", str(o))
    return (int(m.group()) if m else 10**6, str(o))


def html_md(s):
    """Renderiza HTML (em uma linha só, para o markdown não interferir)."""
    st.markdown(s, unsafe_allow_html=True)


def mostrar_plotly(fig):
    try:
        st.plotly_chart(fig, width="stretch")
    except Exception:  # versões antigas do Streamlit
        st.plotly_chart(fig, use_container_width=True)


def mostrar_df(df, **kw):
    try:
        st.dataframe(df, width="stretch", **kw)
    except Exception:
        st.dataframe(df, use_container_width=True, **kw)


# ───────────────────────────── dados ─────────────────────────────
@st.cache_data(show_spinner="Carregando a base…")
def carregar():
    df = pd.read_excel(ARQUIVO, sheet_name=ABA)
    faltando = [c for c in OBRIGATORIAS if c not in df.columns]
    if faltando:
        raise ValueError(f"Colunas ausentes na aba '{ABA}': {faltando}")
    df["anulada"] = pd.to_numeric(df["anulada"], errors="coerce").fillna(0).astype(int)
    df["ano"] = df["ano"].astype(int)
    # assunto mais específico: nível 5, senão 4, 3, 2, senão a própria matéria
    df["assunto_detalhado"] = df[["nivel5", "nivel4", "nivel3", "nivel2", "materia"]].bfill(axis=1).iloc[:, 0]
    return df


def caminho(linha):
    """(materia, nivel2, ...) até o último nível preenchido."""
    nomes = []
    for x in linha:
        if not isinstance(x, str):
            break
        nomes.append(x)
    return tuple(nomes)


@st.cache_data
def nos_assunto():
    """{rótulo: caminho} de todos os assuntos (níveis 2 a 5) existentes na base."""
    df = carregar()
    vistos = {}
    for linha in df[NIVEIS].itertuples(index=False):
        c = caminho(linha)
        for k in range(2, len(c) + 1):
            vistos[" › ".join(c[:k])] = c[:k]
    return dict(sorted(vistos.items()))


# ───────────────────────────── abreviações dos assuntos ─────────────────────────────
MAX_CURTO = 18     # nomes (sem parênteses) até este tamanho ficam como estão
MAX_ABREV = 18     # se a versão abreviada couber aqui, usa ela inteira; senão corta nas 1ªs palavras + "…"
MIN_PALAVRAS = 2   # ao cortar, mantém pelo menos 2 palavras (ex.: "Dir. Adm…")

STOP = {"de", "da", "do", "das", "dos", "e", "em", "na", "no", "nas", "nos", "a", "o", "as", "os",
        "para", "por", "com", "ao", "à", "ou"}
ABREV = {
    "direito": "Dir", "direitos": "Dir", "administração": "Adm", "administrativo": "Adm",
    "administrativa": "Adm", "constitucional": "Const", "geral": "Ger.", "pública": "Púb.",
    "público": "Púb.", "públicos": "Púb.", "públicas": "Púb.", "legislação": "Legisl.",
    "específica": "Espec.", "específico": "Espec.", "tribunais": "Trib.", "ministérios": "Min.",
    "federais": "Fed.", "federal": "Fed.", "estadual": "Est.", "municipal": "Mun.",
    "processual": "Proc.", "processo": "Proc.", "trabalho": "Trab.", "previdenciário": "Prev.",
    "tributário": "Trib.", "financeiro": "Fin.", "financeira": "Fin.", "contabilidade": "Cont.",
    "contábil": "Cont.", "contábeis": "Cont.", "orçamentária": "Orç.", "orçamentário": "Orç.",
    "orçamento": "Orç.", "informática": "Inf.", "informação": "Inf.", "tecnologia": "Tec.",
    "língua": "Líng.", "portuguesa": "Port.", "raciocínio": "Rac.", "lógico": "Lóg.",
    "matemática": "Mat.", "recursos": "Rec.", "materiais": "Mat.", "humanos": "Hum.",
    "gestão": "Gest.", "serviço": "Serv.", "serviços": "Serv.", "segurança": "Seg.",
    "responsabilidade": "Resp.", "fiscal": "Fisc.", "introdução": "Intr.", "demonstrações": "Demonstr.",
    "noções": "Noç.", "princípios": "Princ.", "fundamentais": "Fund.", "organização": "Org.",
    "organizacional": "Org.", "empresarial": "Empr.", "internacional": "Intern.",
}
VOGAIS = set("aeiouáàâãéêíóôõú")


def abrev_palavra(p):
    ref = p.lower()
    if ref in ABREV:
        return ABREV[ref]
    if len(p) <= 5 or not p.isalpha():      # curtas, números e siglas ficam como estão
        return p
    if p.isupper():
        return p
    corte = 5 if p[3].lower() in VOGAIS else 4   # corta depois de uma consoante
    return p[:corte] + "."


def _partes(nome):
    sem = re.sub(r"\s*\(.*?\)", "", nome).strip()
    teve_par = sem != nome.strip()
    toks = [t.strip(",;:") for t in sem.split() if t not in {"-", "–", "—"}]
    conteudo = [t for t in toks if t and t.lower() not in STOP]
    return sem, teve_par, [abrev_palavra(t) for t in conteudo]


def abreviar_grupo(nomes):
    """{nome: nome curto} para irmãos (mesmo pai). Só abrevia nomes longos; garante nomes curtos distintos."""
    res, pend = {}, {}
    for n in nomes:
        sem, par, abr = _partes(n)
        if not abr:
            res[n] = n
        elif len(sem) <= MAX_CURTO:
            res[n] = sem + ("…" if par else "")
        else:
            completo = " ".join(abr)
            if len(completo) <= MAX_ABREV:
                res[n] = completo.rstrip(".") + "…" if par else completo
            else:
                pend[n] = (abr, par)
    usados = Counter(res.values())
    k = min(MIN_PALAVRAS, max((len(a) for a, _ in pend.values()), default=0))
    ativos = dict(pend)
    while ativos:
        cand = {}
        for n, (abr, par) in ativos.items():
            if k >= len(abr):
                cand[n] = " ".join(abr).rstrip(".") + "…" if par else " ".join(abr)
            else:
                cand[n] = " ".join(abr[:k]).rstrip(".") + "…"
        cont = Counter(cand.values())
        for n, c in list(cand.items()):
            if cont[c] == 1 and usados[c] == 0:
                res[n] = c; usados[c] += 1; del ativos[n]
        if ativos and all(k >= len(a) for a, _ in ativos.values()):
            break
        k += 1
    for n in ativos:           # sem como diferenciar: mantém o nome inteiro
        res[n] = n
    dup = {v for v, c in Counter(res.values()).items() if c > 1}
    for n, v in list(res.items()):
        if v in dup:
            res[n] = n
    return res


def rotulos_curtos(mapa):
    """{rótulo completo: rótulo curto}. Só o último nível fica sem abreviar."""
    prefixos = {cam[:k] for cam in mapa.values() for k in range(1, len(cam) + 1)}
    grupos = {}
    for p in prefixos:
        grupos.setdefault(p[:-1], []).append(p[-1])
    curto = {}
    for pai, nomes in grupos.items():
        for n, c in abreviar_grupo(nomes).items():
            curto[pai + (n,)] = c
    return {rot: " › ".join([curto[cam[:j]] for j in range(1, len(cam))] + [cam[-1]]) for rot, cam in mapa.items()}


@st.cache_data
def curtos():
    """{rótulo completo: rótulo abreviado} de todos os assuntos (níveis 2 a 5)."""
    return rotulos_curtos(nos_assunto())


# ───────────────────────────── filtros em cascata ─────────────────────────────
def universo(df, ignorar=None):
    """Questões que passam por todos os filtros (exceto, opcionalmente, o de chave `ignorar`)."""
    s = st.session_state
    m = pd.Series(True, index=df.index)
    for chave, col in CHAVES.items():
        if chave == ignorar:
            continue
        if chave == "f_materia":
            if s[chave] != "Todas":
                m &= df[col] == s[chave]
        elif s[chave]:
            m &= df[col].isin(s[chave])
    if not s["f_anuladas"]:
        m &= df["anulada"] == 0
    return df[m]


def limpar_filtros():
    for k in ("f_escol", "f_orgao", "f_cargo", "f_ano"):
        st.session_state[k] = []
    st.session_state["f_materia"] = "Todas"
    st.session_state["f_anuladas"] = True


def barra_lateral(df):
    s = st.session_state
    padroes = {"f_escol": [], "f_orgao": [], "f_cargo": [], "f_ano": [], "f_materia": "Todas",
               "f_anuladas": True, "destaques": [], "pagina": "🥧 Pizza"}
    for k, v in padroes.items():
        s.setdefault(k, list(v) if isinstance(v, list) else v)

    # 1) descarta seleções que deixaram de existir depois de outro filtro mudar.
    #    Os filtros mais à esquerda (Órgão → Cargo → Ano → Matéria) têm prioridade:
    #    percorremos da direita para a esquerda, então o que "sobra" inválido é o da direita.
    for _ in range(5):
        mudou = False
        for chave, col in reversed(list(CHAVES.items())):
            validos = set(universo(df, ignorar=chave)[col].unique())
            if chave == "f_materia":
                if s[chave] != "Todas" and s[chave] not in validos:
                    s[chave], mudou = "Todas", True
            else:
                novo = [v for v in s[chave] if v in validos]
                if novo != s[chave]:
                    s[chave], mudou = novo, True
        if not mudou:
            break

    sb = st.sidebar
    pagina = sb.radio("Gráfico", PAGINAS, key="pagina")
    if pagina == PAG_COMPARAR:
        sb.caption("Na página Comparar, cada cenário tem os seus próprios filtros.")
        return
    sb.markdown("---")
    sb.header("🎯 Filtros")

    def multi(rotulo, chave, col, ordem):
        cont = universo(df, ignorar=chave)[col].value_counts()
        sb.multiselect(rotulo, sorted(cont.index.tolist(), key=ordem), key=chave,
                       format_func=lambda v: f"{v} ({cont[v]})", placeholder="Todos")

    multi("Escolaridade", "f_escol", "escolaridade", str)
    multi("Órgão", "f_orgao", "orgao", chave_orgao)
    multi("Cargo", "f_cargo", "Cargo", str)
    multi("Ano", "f_ano", "ano", lambda a: -a)

    base_mat = universo(df, ignorar="f_materia")
    cont_mat = base_mat["materia"].value_counts()
    sb.selectbox(
        "Matéria", ["Todas"] + cont_mat.index.tolist(), key="f_materia",
        format_func=lambda v: f"Todas ({len(base_mat)})" if v == "Todas" else f"{v} ({cont_mat[v]})",
    )
    sb.caption("⚠️ Este filtro aceita apenas **uma** matéria por vez.")

    sb.checkbox("Incluir questões anuladas", key="f_anuladas")
    sb.button("Limpar filtros", on_click=limpar_filtros)


# ───────────────────────────── cálculos dos gráficos ─────────────────────────────
def agrupar_pizza(df_u, coluna):
    """Fatias: grupos com >= LIMIAR_OUTRAS% do universo; o resto vira 'Outras'."""
    cont = df_u[coluna].fillna("(sem assunto)").value_counts()
    pct = cont / len(df_u) * 100
    fortes = sorted(cont[pct >= LIMIAR_OUTRAS].items(), key=lambda kv: (-kv[1], kv[0]))
    linhas = [(nome, n, PALETA[i % len(PALETA)]) for i, (nome, n) in enumerate(fortes)]
    resto = cont[pct < LIMIAR_OUTRAS]
    if len(resto):
        linhas.append(("Outras", int(resto.sum()), COR_OUTRAS))
    out = pd.DataFrame(linhas, columns=["rotulo", "qtd", "cor"])
    out["pct"] = out["qtd"] / len(df_u) * 100
    return out


def nomes_destaques(rotulos):
    """{caminho: nome curto} na ordem escolhida (usa o caminho se houver nomes repetidos)."""
    mapa = nos_assunto()
    caminhos = [mapa[r] for r in rotulos if r in mapa]
    ultimos = [c[-1] for c in caminhos]
    repetidos = {u for u in ultimos if ultimos.count(u) > 1}
    return {c: (" › ".join(c[1:]) if c[-1] in repetidos else c[-1]) for c in caminhos}


def segmentar(df_u, destaques):
    """Segmento de cada questão: o destaque mais específico do seu caminho, ou 'Demais assuntos'."""
    if not destaques:
        return pd.Series("Demais assuntos", index=df_u.index)
    saida = []
    for linha in df_u[NIVEIS].itertuples(index=False):
        nomes = caminho(linha)
        seg = "Demais assuntos"
        for k in range(len(nomes), 1, -1):
            nome = destaques.get(nomes[:k])
            if nome:
                seg = nome
                break
        saida.append(seg)
    return pd.Series(saida, index=df_u.index)


def dados_barras(df_u, coluna, destaques):
    """Tabela grupo × segmento (contagens). Grupos < limiar viram 'Outras', exceto os que contêm destaque."""
    total = len(df_u)
    grupo = df_u[coluna].fillna("(sem assunto)")
    tab = pd.crosstab(grupo, segmentar(df_u, destaques))
    nomes_dest = [n for n in destaques.values() if n in tab.columns]
    cols = nomes_dest + (["Demais assuntos"] if "Demais assuntos" in tab.columns else [])
    tab = tab[cols]
    tot = tab.sum(axis=1)
    tem_dest = tab[nomes_dest].sum(axis=1) > 0 if nomes_dest else pd.Series(False, index=tab.index)
    pequenas = (tot / total * 100 < LIMIAR_OUTRAS) & ~tem_dest
    ordem = sorted(tab.index[~pequenas], key=lambda g: (-tot[g], g))
    final = tab.loc[ordem]
    n_outras = int(pequenas.sum())
    if n_outras:
        final.loc["Outras"] = tab[pequenas].sum()
    return final, n_outras, nomes_dest


def figura_pizza(pz):
    fig = go.Figure(go.Pie(
        labels=pz["rotulo"], values=pz["qtd"], marker=dict(colors=pz["cor"]),
        hole=0.4, sort=False, direction="clockwise",
        texttemplate="%{percent:.1%}", textposition="inside",
        hovertemplate="<b>%{label}</b><br>%{value} questões · %{percent:.1%}<extra></extra>",
    ))
    fig.update_layout(
        showlegend=False, height=520, margin=dict(l=10, r=10, t=10, b=10), separators=",.",
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        uniformtext_minsize=11, uniformtext_mode="hide", hoverlabel=HOVER,
    )
    fig.update_traces(hoverlabel=HOVER)
    return fig


def passo_eixo(maximo):
    """Passo "redondo" para o eixo (no máximo ~8 marcas)."""
    for p in (0.5, 1, 2, 5, 10, 20, 25, 50):
        if maximo / p <= 8:
            return p
    return 50


def figura_barras(final, total, n_outras, nomes_dest, unidade):
    cats = list(final.index)
    cor = {n: PALETA_DESTAQUE[i % len(PALETA_DESTAQUE)] for i, n in enumerate(nomes_dest)}
    cor["Demais assuntos"] = COR_DEMAIS
    fig = go.Figure()
    for seg in final.columns:
        qtd = final[seg].astype(int)
        fig.add_bar(
            y=cats, x=(qtd / total * 100).tolist(), name=seg, orientation="h",
            marker=dict(color=cor[seg]), customdata=qtd.tolist(),
            hovertemplate=("<b>%{y}</b><br>" + html.escape(seg) +
                           ": %{customdata} questões · %{x:.1f}% do universo<extra></extra>"),
        )
    tot = final.sum(axis=1)
    pct = tot / total * 100
    maximo = float(pct.max())
    passo = passo_eixo(maximo)
    marcas = [round(i * passo, 4) for i in range(int(maximo // passo) + 1)]
    # o valor de cada barra vai no rótulo do eixo vertical, para a maior barra poder tocar a borda direita
    rotulos = []
    for g in cats:
        extra = f" · {n_outras} {unidade}" if g == "Outras" else ""
        rotulos.append(f"{g}<br>{fmt_pct(pct[g])} · {int(tot[g])} q.{extra}")
    fig.update_layout(
        barmode="stack", bargap=0.18, showlegend=False, separators=",.",
        height=max(320, 64 * len(cats) + 90), margin=dict(l=10, r=12, t=10, b=40),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(range=[0, maximo], tickvals=marcas, ticksuffix="%",
                   tickformat=".1f" if passo < 1 else ".0f",
                   gridcolor="rgba(128,128,128,.25)", zeroline=False),
        yaxis=dict(autorange="reversed", automargin=True, tickmode="array", tickvals=cats, ticktext=rotulos),
        hoverlabel=HOVER,
    )
    fig.update_traces(hoverlabel=HOVER)
    return fig, cor


# ───────────────────────────── blocos de interface ─────────────────────────────
def legenda_html(itens, coluna=False):
    corpo = "".join(
        f'<div class="lg-i"><span class="lg-c" style="background:{c}"></span>'
        f'<span>{html.escape(n)}</span><span class="lg-v">{html.escape(d)}</span></div>'
        for c, n, d in itens
    )
    return f'<div class="lg{" lg-col" if coluna else ""}">{corpo}</div>'


def arvore_html(df_u, por_materia=True):
    """Menu expansível: matéria › nível 2 › … até o último nível, ordenado por quantidade.
    O que tem menos de LIMIAR_OUTRAS% do universo é reunido em "Outras" (igual ao gráfico), sempre no fim."""
    total = len(df_u)
    raiz = {}
    for linha in df_u[NIVEIS].itertuples(index=False):
        no = raiz
        for nome in caminho(linha):
            d = no.setdefault(nome, {"n": 0, "f": {}})
            d["n"] += 1
            no = d["f"]

    def com_outras(filhos):
        fortes = {k: v for k, v in filhos.items() if v["n"] / total * 100 >= LIMIAR_OUTRAS}
        pequenos = {k: v for k, v in filhos.items() if k not in fortes}
        if pequenos:
            fortes["Outras"] = {"n": sum(v["n"] for v in pequenos.values()), "f": pequenos, "outras": True}
        return fortes

    if por_materia:                      # gráfico por matéria: agrupa as matérias pequenas
        raiz = com_outras(raiz)
    else:                                # gráfico por nível 2 de uma matéria: agrupa os níveis 2 pequenos
        for d in raiz.values():
            d["f"] = com_outras(d["f"])

    def linha_html(nome, d):
        return (f'<span class="nm">{html.escape(nome)}</span>'
                f'<span class="q">{fmt_int(d["n"])}</span>'
                f'<span class="p">{fmt_pct(d["n"] / total * 100)}</span>')

    def nos(filhos, nivel):
        partes = []
        for nome, d in sorted(filhos.items(), key=lambda kv: (kv[1].get("outras", False), -kv[1]["n"], kv[0])):
            cl = ("raiz" if nivel == 0 else "no") + (" outras" if d.get("outras") else "")
            if d["f"]:
                abre = " open" if (nivel == 0 and len(filhos) == 1) else ""
                partes.append(f'<details class="{cl}"{abre}><summary>{linha_html(nome, d)}</summary>'
                              f'<div class="ch">{nos(d["f"], nivel + 1)}</div></details>')
            else:
                partes.append(f'<div class="fol {cl}">{linha_html(nome, d)}</div>')
        return "".join(partes)

    cab = '<div class="cab"><span class="nm">Matéria / assunto</span><span class="q">Questões</span><span class="p">% do universo</span></div>'
    return f'<div class="tt">{cab}{nos(raiz, 0)}</div>'


@st.cache_data
def exportar(df_u, formato):
    if formato == "csv":
        return df_u.to_csv(sep=";", index=False).encode("utf-8-sig")
    buf = BytesIO()
    df_u.to_excel(buf, index=False, sheet_name="Questões")
    return buf.getvalue()


def sem_acento(txt):
    return "".join(c for c in unicodedata.normalize("NFD", str(txt)) if unicodedata.category(c) != "Mn").lower()


TABELA_FILTROS = [  # (chave, coluna do df, rótulo, peso da coluna)
    ("tf_escol", "escolaridade", "escolaridade", 1.2),
    ("tf_orgao", "orgao", "orgao", 1.0),
    ("tf_ano", "ano", "ano", 0.9),
    ("tf_cargo", "Cargo", "cargo", 1.3),
    ("tf_materia", "materia", "matéria", 2.0),
]


def limpar_filtros_tabela():
    for chave, *_ in TABELA_FILTROS:
        st.session_state[chave] = []
    st.session_state["tf_id"] = ""
    st.session_state["tf_assunto"] = ""


def filtros_tabela(df_u):
    """Filtros logo acima do cabeçalho da tabela de questões (valem só para a tabela e para o download)."""
    s = st.session_state
    for chave, col, _, _ in TABELA_FILTROS:             # descarta seleções que saíram do universo
        s.setdefault(chave, [])
        validos = set(df_u[col].unique())
        s[chave] = [v for v in s[chave] if v in validos]
    s.setdefault("tf_id", "")
    s.setdefault("tf_assunto", "")

    larguras = [1.2] + [p for *_, p in TABELA_FILTROS] + [3.0]
    cols = st.columns(larguras)
    cols[0].text_input("id_questao", key="tf_id", placeholder="um ou mais IDs",
                       help="Digite um ou mais IDs (ou pedaços), separados por espaço ou vírgula.")
    for c, (chave, col, rotulo, _) in zip(cols[1:-1], TABELA_FILTROS):
        opcoes = sorted(df_u[col].unique().tolist(), key=chave_orgao if col == "orgao" else str)
        c.multiselect(rotulo, opcoes, key=chave, placeholder="Todos")
    cols[-1].text_input("assunto_detalhado", key="tf_assunto", placeholder="parte do nome de um assunto",
                        help="Procura em todos os níveis do assunto (matéria › nível 2 › … › último nível), "
                             "ignorando acentos e maiúsculas. Dica: cole aqui o nome de um item da tabela expansível.")

    out = df_u
    ids = [t for t in re.split(r"[\s,;]+", s["tf_id"].strip()) if t]
    if ids:
        col_id = out["id_questao"].astype(str)
        out = out[col_id.map(lambda x: any(t in x for t in ids))]
    for chave, col, _, _ in TABELA_FILTROS:
        if s[chave]:
            out = out[out[col].isin(s[chave])]
    termo = sem_acento(s["tf_assunto"].strip())
    if termo:
        caminhos = out[NIVEIS].apply(lambda r: sem_acento(" › ".join(caminho(r))), axis=1)
        out = out[caminhos.str.contains(termo, regex=False)]

    n1, n2 = st.columns([5, 1])
    n1.caption(f"Mostrando {fmt_int(len(out))} de {fmt_int(len(df_u))} questões do universo filtrado.")
    n2.button("Limpar filtros da tabela", on_click=limpar_filtros_tabela)
    return out


def bloco_tabelas(df_u):
    st.markdown("---")
    st.subheader("📚 Matérias e assuntos neste universo")
    st.caption("Clique na seta para abrir os níveis. Ordenado por quantidade; % em relação ao universo filtrado. "
               f"“Outras” reúne o que tem menos de {fmt_pct(LIMIAR_OUTRAS, 0)} (igual ao gráfico) e abre normalmente.")
    html_md(arvore_html(df_u, st.session_state["f_materia"] == "Todas"))

    st.markdown("---")
    st.subheader("📋 Questões")
    df_tab = filtros_tabela(df_u)
    mostra = pd.DataFrame({
        "id_questao": df_tab["id_questao"], "escolaridade": df_tab["escolaridade"], "orgao": df_tab["orgao"], "ano": df_tab["ano"],
        "cargo": df_tab["Cargo"], "matéria": df_tab["materia"], "assunto_detalhado": df_tab["assunto_detalhado"],
    })
    mostrar_df(mostra, hide_index=True, column_config={
        "id_questao": st.column_config.NumberColumn("id_questao", format="%d"),
        "ano": st.column_config.NumberColumn("ano", format="%d"),
    })
    c1, c2, _ = st.columns([1, 1, 4])
    c1.download_button("⬇️ Baixar CSV (todas as colunas)", exportar(df_tab, "csv"),
                       file_name="questoes_filtradas.csv", mime="text/csv")
    c2.download_button("⬇️ Baixar XLSX (todas as colunas)", exportar(df_tab, "xlsx"),
                       file_name="questoes_filtradas.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def metricas(df_u, df):
    c1, c2, c3 = st.columns(3)
    c1.metric("Questões filtradas", fmt_int(len(df_u)))
    c2.metric("Total da base", fmt_int(len(df)))
    c3.metric("% da base selecionada", fmt_pct(len(df_u) / len(df) * 100))


def pagina_pizza(df_u):
    por_materia = st.session_state["f_materia"] == "Todas"
    coluna = "materia" if por_materia else "nivel2"
    st.subheader("🥧 Distribuição por matéria" if por_materia
                 else f"🥧 Distribuição por assunto (nível 2) — {st.session_state['f_materia']}")
    pz = agrupar_pizza(df_u, coluna)
    esq, dir_ = st.columns([3, 2])
    with esq:
        mostrar_plotly(figura_pizza(pz))
    with dir_:
        html_md(legenda_html([(r.cor, r.rotulo, f"{fmt_pct(r.pct)} · {fmt_int(r.qtd)}")
                              for r in pz.itertuples()], coluna=True))


def seletor_destaques():
    """Multiselect em largura total. O widget só existe nesta página, então a seleção fica
    guardada em st.session_state["destaques"] para sobreviver à troca de página."""
    s = st.session_state
    if "w_destaques" not in s:
        s["w_destaques"] = list(s["destaques"])
    rot = curtos()
    s["w_destaques"] = [r for r in s["w_destaques"] if r in rot]
    st.multiselect(
        "Destacar assuntos (apenas para gráfico de barras)", list(rot.keys()), key="w_destaques",
        format_func=lambda r: rot[r], placeholder="Digite para buscar um assunto…",
        help="Caminho: Matéria › nível 2 › … › último nível. Nomes longos aparecem abreviados, "
             "exceto o último nível. Escolher um nível intermediário destaca todas as questões abaixo dele.",
    )
    s["destaques"] = list(s["w_destaques"])


def pagina_barras(df_u):
    por_materia = st.session_state["f_materia"] == "Todas"
    coluna = "materia" if por_materia else "nivel2"
    unidade = "matérias" if por_materia else "assuntos"
    st.subheader("📊 Matérias" if por_materia
                 else f"📊 Assuntos (nível 2) — {st.session_state['f_materia']}")
    seletor_destaques()
    destaques = nomes_destaques(st.session_state["destaques"])
    final, n_outras, nomes_dest = dados_barras(df_u, coluna, destaques)
    fig, cor = figura_barras(final, len(df_u), n_outras, nomes_dest, unidade)
    mostrar_plotly(fig)

    itens = [(cor[n], n, f"{fmt_int(final[n].sum())} questões") for n in nomes_dest]
    if "Demais assuntos" in final.columns:
        itens.append((COR_DEMAIS, "Demais assuntos", f"{fmt_int(final['Demais assuntos'].sum())} questões"))
    html_md(legenda_html(itens))
    ausentes = [n for n in destaques.values() if n not in nomes_dest]
    if ausentes:
        st.caption("Sem questões neste universo: " + "; ".join(ausentes))
    if not destaques:
        st.caption("Escolha assuntos em “Destacar assuntos” (barra lateral) para segmentá-los nas barras.")


# ───────────────────────────── comparação de cenários ─────────────────────────────
def ck(i, campo):
    return f"cmp_{i}_{campo}"


def chaves_persistentes():
    """Chaves cujo valor deve sobreviver quando o widget some (troca de página)."""
    ks = list(CHAVES) + ["f_anuladas", "w_destaques", "tf_id", "tf_assunto"] + [c for c, *_ in TABELA_FILTROS]
    ks += ["cmp_n", "cmp_anuladas", "cmp_materia", "cmp_k"]
    for i in range(4):
        ks += [ck(i, campo) for campo, *_ in CMP_CAMPOS] + [f"cmp_{i}_modo", f"cmp_{i}_nome"]
    return ks


def cmp_inicializar():
    s = st.session_state
    for k, v in {"cmp_n": 2, "cmp_anuladas": True, "cmp_materia": "Todas", "cmp_k": 1}.items():
        s.setdefault(k, v)
    for i in range(4):
        for campo, *_ in CMP_CAMPOS:
            s.setdefault(ck(i, campo), [])
        s.setdefault(f"cmp_{i}_modo", CMP_MODOS[0])
        s.setdefault(f"cmp_{i}_nome", "")


def cmp_base(df):
    return df if st.session_state["cmp_anuladas"] else df[df["anulada"] == 0]


def cmp_modo(i):
    return st.session_state[f"cmp_{i}_modo"] if i > 0 else CMP_MODOS[0]


def cmp_filtrar(base, i, ignorar=None):
    s = st.session_state
    m = pd.Series(True, index=base.index)
    for campo, col, _ in CMP_CAMPOS:
        if campo != ignorar and s[ck(i, campo)]:
            m &= base[col].isin(s[ck(i, campo)])
    return base[m]


def cmp_limpar_selecoes(df):
    """Descarta seleções que deixaram de existir (da direita para a esquerda), em cada cenário."""
    s = st.session_state
    base = cmp_base(df)
    for i in range(s["cmp_n"]):
        for _ in range(4):
            mudou = False
            for campo, col, _ in reversed(CMP_CAMPOS):
                validos = set(cmp_filtrar(base, i, ignorar=campo)[col].unique())
                novo = [v for v in s[ck(i, campo)] if v in validos]
                if novo != s[ck(i, campo)]:
                    s[ck(i, campo)], mudou = novo, True
            if not mudou:
                break


def cmp_copiar(i):
    """Copia os filtros da barra lateral (páginas Pizza/Barras) para o cenário i."""
    s = st.session_state
    for campo, chave in (("escol", "f_escol"), ("orgao", "f_orgao"), ("cargo", "f_cargo"), ("ano", "f_ano")):
        s[ck(i, campo)] = list(s.get(chave, []))


def cmp_universos(df, com_materia=True):
    s = st.session_state
    base = cmp_base(df)
    univs = []
    for i in range(s["cmp_n"]):
        modo = cmp_modo(i)
        if modo == CMP_MODOS[1]:
            u = base
        elif modo == CMP_MODOS[2]:
            u = base.drop(univs[0].index, errors="ignore")
        else:
            u = cmp_filtrar(base, i)
        univs.append(u)
    if com_materia and s["cmp_materia"] != "Todas":
        univs = [u[u["materia"] == s["cmp_materia"]] for u in univs]
    return univs


def anos_txt(anos):
    anos = sorted({int(a) for a in anos})
    faixas, ini, prev = [], None, None
    for a in anos:
        if ini is None:
            ini = prev = a
        elif a == prev + 1:
            prev = a
        else:
            faixas.append((ini, prev))
            ini = prev = a
    if ini is not None:
        faixas.append((ini, prev))
    return ", ".join(str(a) if a == b else (f"{a}, {b}" if b == a + 1 else f"{a}–{b}") for a, b in faixas)


def cmp_rotulo_auto(i):
    s = st.session_state
    modo = cmp_modo(i)
    if modo == CMP_MODOS[1]:
        return "Total da base"
    if modo == CMP_MODOS[2]:
        return "Resto (fora do cenário 1)"
    partes = []
    if s[ck(i, "escol")]:
        partes.append("/".join(s[ck(i, "escol")]))
    if s[ck(i, "cargo")]:
        partes.append("/".join(c.split()[0] for c in s[ck(i, "cargo")]))
    org = s[ck(i, "orgao")]
    if org:
        partes.append(org[0] if len(org) == 1 else f"{len(org)} órgãos")
    if s[ck(i, "ano")]:
        partes.append(anos_txt(s[ck(i, "ano")]))
    return " · ".join(partes) or "Todas as questões"


def cmp_rotulos(n):
    s = st.session_state
    rot = []
    for i in range(n):
        base = s[f"cmp_{i}_nome"].strip() or cmp_rotulo_auto(i)
        nome, k = base, 2
        while nome in rot:
            nome, k = f"{base} ({k})", k + 1
        rot.append(nome)
    return rot


# ----- estatísticas: contagens por nó da árvore de assuntos, por cenário e por prova -----
def construir_stats(univs):
    """cnt[prefixo][cenário][prova] = nº de questões; provas[cenário][prova] = tamanho da prova no universo.
    "Prova" = (escolaridade, órgão, ano, cargo)."""
    cnt, provas = {}, []
    for si, u in enumerate(univs):
        chaves = list(zip(u["escolaridade"], u["orgao"], u["ano"], u["Cargo"]))
        tam = Counter(chaves)
        provas.append(dict(tam))
        for linha, k in zip(u[NIVEIS].itertuples(index=False), chaves):
            c = caminho(linha)
            for j in range(1, len(c) + 1):
                lst = cnt.setdefault(c[:j], [defaultdict(int) for _ in univs])
                lst[si][k] += 1
    return cnt, provas


def n_no(cnt, pref, si):
    return sum(cnt[pref][si].values()) if pref in cnt else 0


def partilhas(cnt, provas, prefs, si):
    """Fração de cada prova do cenário si que pertence aos nós `prefs`."""
    return [sum(cnt[p][si].get(k, 0) for p in prefs if p in cnt) / tam for k, tam in provas[si].items()]


def diferenca(cnt, provas, tot, prefs, ref, k):
    """(diferença em p.p. do cenário k contra o ref, é maior que ≈2 erros-padrão entre provas?)
    O 2º valor é None quando não dá para estimar (menos de 2 provas em algum cenário)."""
    if not tot[ref] or not tot[k]:
        return None, None
    d = (sum(n_no(cnt, p, k) for p in prefs) / tot[k] - sum(n_no(cnt, p, ref) for p in prefs) / tot[ref]) * 100
    a, b = partilhas(cnt, provas, prefs, ref), partilhas(cnt, provas, prefs, k)
    if len(a) < 2 or len(b) < 2:
        return d, None
    if abs(d) < 1e-9:
        return d, False
    se = math.sqrt(statistics.variance(a) / len(a) + statistics.variance(b) / len(b)) * 100
    return d, abs(d) >= 2 * se


def grupos_cmp(cnt, tot, nivel):
    """Grupos do nível (1 = matéria, 2 = nível 2): fortes (>= limiar em algum cenário) e pequenos (vão para 'Outras')."""
    todos = [p for p in cnt if len(p) == nivel]
    mx = {p: max(n_no(cnt, p, si) / tot[si] * 100 for si in range(len(tot)) if tot[si]) for p in todos}
    fortes = sorted([p for p in todos if mx[p] >= LIMIAR_OUTRAS], key=lambda p: (-mx[p], p))
    pequenos = sorted([p for p in todos if mx[p] < LIMIAR_OUTRAS], key=lambda p: (-mx[p], p))
    return fortes, pequenos


# ----- figuras -----
def figura_comparacao(rot, cats, pct_cat, qtd_cat):
    N = len(rot)
    maxp = max(max(v) for v in pct_cat.values())
    passo = passo_eixo(maxp)
    marcas = [round(i * passo, 4) for i in range(int(maxp // passo) + 1)]
    fig = go.Figure()
    for si in range(N):
        fig.add_bar(
            y=cats, x=[pct_cat[c][si] for c in cats], name=rot[si], orientation="h",
            marker=dict(color=CMP_CORES[si]), customdata=[qtd_cat[c][si] for c in cats],
            text=[fmt_pct(pct_cat[c][si]) for c in cats], textposition="outside", cliponaxis=False,
            hovertemplate="<b>%{y}</b><br>" + html.escape(rot[si]) + ": %{x:.1f}% · %{customdata} questões<extra></extra>",
        )
    fig.update_layout(
        barmode="group", bargap=0.22, bargroupgap=0.05, showlegend=False, separators=",.",
        height=max(360, len(cats) * (N * 26 + 18) + 80), margin=dict(l=10, r=80, t=10, b=40),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(range=[0, maxp], tickvals=marcas, ticksuffix="%", tickformat=".1f" if passo < 1 else ".0f",
                   gridcolor="rgba(128,128,128,.25)", zeroline=False),
        yaxis=dict(autorange="reversed", automargin=True), hoverlabel=HOVER,
    )
    fig.update_traces(hoverlabel=HOVER)
    return fig


def figura_diferenca(k, cats, dif):
    ordem = sorted([c for c in cats if c != "Outras"], key=lambda c: -dif[c][0])
    if "Outras" in cats:
        ordem.append("Outras")
    x = [dif[c][0] for c in ordem]
    cor = [CMP_CORES[k] if v >= 0 else CMP_CORES[0] for v in x]
    opac = [1.0 if dif[c][1] else 0.35 for c in ordem]
    txt = [("+" if v > 0.05 else "−" if v < -0.05 else "") + f"{abs(v):.1f}".replace(".", ",") + " p.p." for v in x]
    m = max([abs(v) for v in x] + [1.0]) * 1.3
    fig = go.Figure(go.Bar(
        y=ordem, x=x, orientation="h", marker=dict(color=cor, opacity=opac), text=txt,
        textposition="outside", cliponaxis=False, hovertemplate="<b>%{y}</b><br>%{text}<extra></extra>",
    ))
    fig.update_layout(
        showlegend=False, bargap=0.25, separators=",.", height=max(320, 46 * len(ordem) + 90),
        margin=dict(l=10, r=30, t=10, b=40), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(range=[-m, m], ticksuffix=" p.p.", tickformat=".1f" if m < 2 else ".0f", gridcolor="rgba(128,128,128,.25)",
                   zeroline=True, zerolinecolor="rgba(128,128,128,.8)", zerolinewidth=2),
        yaxis=dict(autorange="reversed", automargin=True), hoverlabel=HOVER,
    )
    fig.update_traces(hoverlabel=HOVER)
    return fig


def figura_pontos(rot, cats, prefs_cat, cnt, provas, tot):
    N = len(rot)
    fig = go.Figure()
    xmax = 1.0
    for si in range(N):
        off = (si - (N - 1) / 2) * 0.26
        xs, ys, txt = [], [], []
        for ci, cat in enumerate(cats):
            for k, tam in provas[si].items():
                c = sum(cnt[p][si].get(k, 0) for p in prefs_cat[cat] if p in cnt)
                v = c / tam * 100
                jit = ((zlib.crc32(str(k).encode()) % 9) - 4) * 0.02
                xs.append(v)
                ys.append(ci + off + jit)
                txt.append(f"{k[1]} {k[2]} · {k[3].split()[0]} · {k[0]}<br>{fmt_pct(v)} ({c} de {tam})")
        xmax = max([xmax] + xs)
        fig.add_scatter(x=xs, y=ys, mode="markers", name=rot[si], text=txt, hovertemplate="%{text}<extra></extra>",
                        marker=dict(color=CMP_CORES[si], size=7, opacity=0.55))
        pooled = [sum(n_no(cnt, p, si) for p in prefs_cat[cat]) / tot[si] * 100 for cat in cats]
        fig.add_scatter(x=pooled, y=[ci + off for ci in range(len(cats))], mode="markers", name=f"{rot[si]} (total)",
                        text=[f"{rot[si]}<br>cenário inteiro: {fmt_pct(v)}" for v in pooled],
                        hovertemplate="%{text}<extra></extra>",
                        marker=dict(color=CMP_CORES[si], size=13, symbol="diamond", line=dict(color="white", width=1.5)))
    passo = passo_eixo(xmax)
    fig.update_layout(
        showlegend=False, separators=",.", height=max(360, 58 * len(cats) + 90), margin=dict(l=10, r=30, t=10, b=40),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(range=[0, xmax * 1.05], tickvals=[round(i * passo, 4) for i in range(int(xmax // passo) + 1)],
                   ticksuffix="%", tickformat=".1f" if passo < 1 else ".0f", gridcolor="rgba(128,128,128,.25)", zeroline=False),
        yaxis=dict(tickmode="array", tickvals=list(range(len(cats))), ticktext=cats, range=[len(cats) - 0.5, -0.5],
                   gridcolor="rgba(128,128,128,.15)", automargin=True),
        hoverlabel=HOVER,
    )
    fig.update_traces(hoverlabel=HOVER)
    return fig


def arvore_cmp_html(rot, cnt, provas, tot, por_materia):
    """Menu expansível com os cenários lado a lado: % (n) de cada um e diferença contra o cenário 1."""
    N = len(rot)
    filhos = {}
    for p in cnt:
        filhos.setdefault(p[:-1], []).append(p)
    corte = 1 if por_materia else 2

    def pct_de(prefs, si):
        return sum(n_no(cnt, p, si) for p in prefs) / tot[si] * 100 if tot[si] else 0.0

    def maxpct(prefs):
        return max(pct_de(prefs, si) for si in range(N))

    def celulas(prefs):
        out = ""
        for si in range(N):
            n = sum(n_no(cnt, p, si) for p in prefs)
            out += f'<span class="c">{fmt_pct(pct_de(prefs, si))}<small> ({fmt_int(n)})</small></span>'
        for si in range(1, N):
            d, nota = diferenca(cnt, provas, tot, prefs, 0, si)
            if d is None:
                txt, cl = "–", "nd"
            else:
                txt = ("+" if d > 0.05 else "−" if d < -0.05 else "") + f"{abs(d):.1f}".replace(".", ",")
                cl = "pos" if d > 0.05 else "neg" if d < -0.05 else "z"
                if not nota:                      # dentro do ruído ou sem provas suficientes para estimar
                    cl += " ruido"
            out += f'<span class="d {cl}">{txt}</span>'
        return out

    def render(prefixos, agrupar, topo=False):
        itens = [{"nome": p[-1], "prefs": [p], "kids": filhos.get(p, []), "outras": False} for p in prefixos]
        itens.sort(key=lambda it: (-maxpct(it["prefs"]), it["nome"]))
        if agrupar:
            fortes = [it for it in itens if maxpct(it["prefs"]) >= LIMIAR_OUTRAS]
            peq = [it for it in itens if maxpct(it["prefs"]) < LIMIAR_OUTRAS]
            itens = fortes
            if peq:
                membros = [p for it in peq for p in it["prefs"]]
                itens.append({"nome": "Outras", "prefs": membros, "kids": membros, "outras": True})
        partes = []
        for it in itens:
            cl = ("raiz " if topo else "no ") + ("outras" if it["outras"] else "")
            nome, cel = html.escape(it["nome"]), celulas(it["prefs"])
            if it["kids"]:
                ag = (not it["outras"]) and len(it["kids"][0]) == corte
                abre = " open" if (topo and len(itens) == 1) else ""
                partes.append(f'<details class="{cl}"{abre}><summary><span class="nm">{nome}</span>{cel}</summary>'
                              f'<div class="ch">{render(it["kids"], ag)}</div></details>')
            else:
                partes.append(f'<div class="fol {cl}"><span class="nm">{nome}</span>{cel}</div>')
        return "".join(partes)

    cab = '<div class="cab"><span class="nm">Matéria / assunto</span>'
    for i, r in enumerate(rot):
        cab += (f'<span class="c" title="{html.escape(r)}"><span class="lg-c" style="background:{CMP_CORES[i]};'
                f'display:inline-block;margin-right:5px;vertical-align:middle"></span>{html.escape(r)}</span>')
    for i in range(1, N):
        cab += f'<span class="d" title="Cenário {i + 1} menos cenário 1, em pontos percentuais">Δ {i + 1}−1</span>'
    cab += "</div>"
    return f'<div class="tt tc">{cab}{render(filhos.get((), []), por_materia, topo=True)}</div>'


# ----- página -----
def pagina_comparar(df):
    s = st.session_state
    cmp_inicializar()
    cmp_limpar_selecoes(df)
    base = cmp_base(df)

    univs0 = cmp_universos(df, com_materia=False)
    idx = set().union(*[set(u.index) for u in univs0])
    cont_mat = df.loc[list(idx), "materia"].value_counts() if idx else pd.Series(dtype=int)
    if s["cmp_materia"] != "Todas" and s["cmp_materia"] not in cont_mat.index:
        s["cmp_materia"] = "Todas"
    n = s["cmp_n"]
    if not 1 <= s["cmp_k"] < n:
        s["cmp_k"] = 1

    univs = cmp_universos(df)
    tot = [len(u) for u in univs]
    rot = cmp_rotulos(n)
    cnt, provas = construir_stats(univs)

    st.subheader("⚖️ Comparar cenários")
    st.caption("Cada cenário tem os seus próprios filtros. Os percentuais são sempre relativos ao universo do próprio cenário.")
    c1, c2, c3 = st.columns([1.3, 2.6, 1.4])
    c1.radio("Número de cenários", [2, 3, 4], key="cmp_n", horizontal=True)
    c2.selectbox("Matéria (opcional)", ["Todas"] + cont_mat.index.tolist(), key="cmp_materia",
                 format_func=lambda v: "Todas" if v == "Todas" else f"{v} ({cont_mat[v]})")
    c2.caption("⚠️ Apenas uma matéria por vez. Com uma matéria, a comparação passa a ser pelos assuntos de nível 2 dela.")
    c3.checkbox("Incluir questões anuladas", key="cmp_anuladas")

    cols = st.columns(n)
    for i, c in enumerate(cols):
        with c:
            st.markdown(f"<span style='color:{CMP_CORES[i]}'>■</span> **Cenário {i + 1}**", unsafe_allow_html=True)
            if i > 0:
                st.radio(f"Fonte do cenário {i + 1}", CMP_MODOS, key=f"cmp_{i}_modo", label_visibility="collapsed")
            if cmp_modo(i) == CMP_MODOS[0]:
                for campo, col, rotulo in CMP_CAMPOS:
                    cont = cmp_filtrar(base, i, ignorar=campo)[col].value_counts()
                    ordem = chave_orgao if col == "orgao" else ((lambda a: -a) if col == "ano" else str)
                    st.multiselect(rotulo, sorted(cont.index.tolist(), key=ordem), key=ck(i, campo),
                                   format_func=lambda v, cont=cont: f"{v} ({cont[v]})", placeholder="Todos")
                st.button("↩ usar os filtros das outras páginas", key=f"cmp_{i}_copiar", on_click=cmp_copiar, args=(i,))
            st.text_input("Nome (opcional)", key=f"cmp_{i}_nome", placeholder=cmp_rotulo_auto(i))
            u = univs[i]
            st.caption(f"{fmt_int(len(u))} questões · {len(provas[i])} provas · {u['orgao'].nunique()} órgãos · "
                       f"{anos_txt(u['ano'].unique()) if len(u) else '–'}")

    vazios = [rot[i] for i in range(n) if tot[i] == 0]
    if vazios:
        st.warning("Cenário sem questões: " + "; ".join(vazios))
        return

    por_materia = s["cmp_materia"] == "Todas"
    fortes, pequenos = grupos_cmp(cnt, tot, 1 if por_materia else 2)
    prefs_cat = {p[-1]: [p] for p in fortes}
    if pequenos:
        prefs_cat["Outras"] = pequenos
    cats = list(prefs_cat)
    qtd_cat = {c: [sum(n_no(cnt, p, si) for p in prefs_cat[c]) for si in range(n)] for c in cats}
    pct_cat = {c: [qtd_cat[c][si] / tot[si] * 100 for si in range(n)] for c in cats}

    st.markdown("---")
    st.subheader("📊 Distribuição por matéria (% de cada cenário)" if por_materia
                 else f"📊 Distribuição por assunto (nível 2) — {s['cmp_materia']}")
    html_md(legenda_html([(CMP_CORES[i], rot[i], f"{fmt_int(tot[i])} questões") for i in range(n)]))
    mostrar_plotly(figura_comparacao(rot, cats, pct_cat, qtd_cat))
    st.caption(f"“Outras” reúne o que tem menos de {fmt_pct(LIMIAR_OUTRAS, 0)} em todos os cenários.")

    st.markdown("---")
    st.subheader("↔️ Diferença entre cenários (pontos percentuais)")
    k = 1
    if n > 2:
        k = st.selectbox("Mostrar a diferença do cenário", list(range(1, n)), key="cmp_k",
                         format_func=lambda x: f"Cenário {x + 1} − Cenário 1 ({rot[x]})")
    dif = {c: diferenca(cnt, provas, tot, prefs_cat[c], 0, k) for c in cats}
    html_md(legenda_html([(CMP_CORES[k], f"Mais frequente em: {rot[k]}", ""), (CMP_CORES[0], f"Mais frequente em: {rot[0]}", "")]))
    mostrar_plotly(figura_diferenca(k, cats, dif))
    st.caption("Barras claras: diferença dentro da variação entre provas (menos de ≈2 erros-padrão, tomando cada prova como "
               "uma observação, ou sem provas suficientes para estimar). Barras sólidas: diferença maior que essa variação. "
               "Com poucas provas por cenário, trate como indício.")

    st.markdown("---")
    st.subheader("🔵 Variação entre provas")
    st.caption("Cada ponto é uma prova (órgão · ano · cargo · escolaridade) e mostra o % daquele grupo dentro da prova; "
               "o losango é o % do cenário inteiro. Pontos espalhados indicam que a diferença depende da prova.")
    mostrar_plotly(figura_pontos(rot, cats, prefs_cat, cnt, provas, tot))

    st.markdown("---")
    st.subheader("📚 Matérias e assuntos — comparação")
    st.caption("Cada célula mostra % (nº de questões). Δ = cenário − cenário 1, em pontos percentuais "
               "(verde: mais frequente no cenário; vermelho: menos; esmaecido: dentro da variação entre provas ou não estimável).")
    html_md(arvore_cmp_html(rot, cnt, provas, tot, por_materia))


# ───────────────────────────── página ─────────────────────────────
def main():
    html_md(CSS)
    st.title("📊 Painel Interativo de Questões")
    st.markdown(
        "**Banco de dados de questões previamente tratado com os seguintes filtros:**  \n"
        "Banca: FCC  \nÓrgãos: TRT (todos)  \nCargos: Técnico e Analista"
    )
    try:
        df = carregar()
    except Exception as e:
        st.error(f"Não foi possível carregar {ARQUIVO.name}: {e}")
        st.stop()

    for k in chaves_persistentes():          # mantém o estado dos filtros de páginas que não estão na tela
        if k in st.session_state:
            st.session_state[k] = st.session_state[k]

    barra_lateral(df)
    if st.session_state["pagina"] == PAG_COMPARAR:
        pagina_comparar(df)
        return
    df_u = universo(df)

    st.markdown("---")
    metricas(df_u, df)
    if df_u.empty:
        st.warning("Nenhuma questão com esses filtros.")
        return

    if st.session_state["pagina"] == "🥧 Pizza":
        pagina_pizza(df_u)
    else:
        pagina_barras(df_u)
    bloco_tabelas(df_u)


main()
