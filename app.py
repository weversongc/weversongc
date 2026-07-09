import streamlit as st
import pandas as pd
import io
from supabase import create_client, Client
from conversor import ConversorPDF

# =========================================================================
# 1. CONFIGURAÇÃO DE PÁGINA
# =========================================================================
st.set_page_config(page_title="Weversongc - Painel de Controle", layout="wide", page_icon="🤖")

# =========================================================================
# 2. CREDENCIAIS (st.secrets)
# =========================================================================
SUPABASE_URL = st.secrets.get("SUPABASE_URL", "")
SUPABASE_KEY = st.secrets.get("SUPABASE_KEY", "")
NOME_BUCKET = st.secrets.get("NOME_BUCKET", "arquivo-usuários")


@st.cache_resource
def iniciar_supabase() -> Client | None:
    if not SUPABASE_URL or not SUPABASE_KEY:
        return None
    try:
        return create_client(SUPABASE_URL, SUPABASE_KEY)
    except Exception:
        return None


supabase = iniciar_supabase()


# =========================================================================
# 3. CSS CUSTOMIZADO (Glassmorphism + Dark Theme)
# =========================================================================
CSS = """
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">

<style>
/* ----------------------- BASE / RESET ----------------------- */
html, body, [class*="css"] {
    font-family: 'Inter', 'Segoe UI', sans-serif;
}
.stApp {
    background: radial-gradient(1200px 600px at 80% -10%, rgba(56,189,248,0.08), transparent 60%),
                radial-gradient(900px 500px at -10% 110%, rgba(168,85,247,0.08), transparent 60%),
                #121316;
    color: #E5E7EB;
}

/* Esconde o menu / rodapé padrão */
#MainMenu, footer { visibility: hidden; }
.stDeployButton { display: none; }

/* ----------------------- SCROLLBAR ----------------------- */
::-webkit-scrollbar { width: 8px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb { background: rgba(255,255,255,0.12); border-radius: 8px; }

/* ----------------------- SIDEBAR ----------------------- */
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, rgba(28,29,34,0.85), rgba(18,19,22,0.92));
    backdrop-filter: blur(18px);
    -webkit-backdrop-filter: blur(18px);
    border-right: 1px solid rgba(255,255,255,0.07) !important;
}
section[data-testid="stSidebar"] .stMarkdown { color: #D1D5DB; }

/* Cabeçalho da sidebar */
.sb-title {
    display: flex; align-items: center; gap: 10px;
    margin-bottom: 2px;
}
.sb-title .ic { font-size: 22px; }
.sb-title h1 {
    font-size: 19px; font-weight: 800; color: #FFFFFF; margin: 0; letter-spacing: -0.3px;
}
.sb-sub {
    font-size: 12.5px; color: #9CA3AF; font-weight: 500; margin: 0 0 18px 0;
}

/* ----------------------- NAV CARDS (radio) ----------------------- */
section[data-testid="stSidebar"] [role="radiogroup"] { gap: 10px; }
section[data-testid="stSidebar"] [role="radiogroup"] label,
section[data-testid="stSidebar"] div[role="radio"] {
    background: rgba(255,255,255,0.03);
    border: 1px solid rgba(255,255,255,0.07);
    border-radius: 12px;
    padding: 13px 14px;
    margin-bottom: 10px;
    transition: all .2s ease;
}
section[data-testid="stSidebar"] [role="radiogroup"] label:hover,
section[data-testid="stSidebar"] div[role="radio"]:hover {
    background: rgba(255,255,255,0.06);
    border-color: rgba(255,255,255,0.14);
    transform: translateY(-1px);
}
/* Item ATIVO - brilho/gradiente */
section[data-testid="stSidebar"] [role="radiogroup"] [aria-checked="true"],
section[data-testid="stSidebar"] [role="radio"][aria-checked="true"] {
    background: linear-gradient(135deg, rgba(34,197,94,0.18), rgba(56,189,248,0.10));
    border: 1px solid rgba(74,222,128,0.55);
    box-shadow: 0 0 0 1px rgba(74,222,128,0.15), 0 8px 24px -8px rgba(34,197,94,0.35);
}
/* Esconde o círculo do radio padrão */
section[data-testid="stSidebar"] [role="radiogroup"] [aria-checked="true"] > div:first-child,
section[data-testid="stSidebar"] [role="radio"] > div:first-child { visibility: hidden; }
section[data-testid="stSidebar"] [role="radiogroup"] label p,
section[data-testid="stSidebar"] div[role="radio"] p {
    font-size: 14px; font-weight: 600; color: #E5E7EB;
}

/* ----------------------- STATUS FOOTER SIDEBAR ----------------------- */
.sb-status {
    display: flex; align-items: center; gap: 9px;
    margin-top: 18px; padding: 11px 13px;
    background: rgba(255,255,255,0.03);
    border: 1px solid rgba(255,255,255,0.07);
    border-radius: 12px;
}
.sb-status .dot {
    width: 9px; height: 9px; border-radius: 50%;
    background: #22c55e; box-shadow: 0 0 10px #22c55e;
    animation: pulse 2s infinite;
}
@keyframes pulse { 0%{opacity:1} 50%{opacity:.45} 100%{opacity:1} }
.sb-status.off .dot { background:#ef4444; box-shadow:0 0 10px #ef4444; }
.sb-status span { font-size: 12.5px; font-weight: 600; color:#D1D5DB; }
.sb-status small { font-size: 11px; color:#6B7280; display:block; }

/* ----------------------- HEADER SUPERIOR ----------------------- */
.app-header {
    display: flex; align-items: center; justify-content: space-between;
    margin-bottom: 22px;
}
.app-header .left { display:flex; align-items:center; gap:16px; }
.app-header .big-ic {
    font-size: 42px;
    background: linear-gradient(135deg, rgba(239,68,68,0.22), rgba(34,197,94,0.22));
    border: 1px solid rgba(255,255,255,0.10);
    border-radius: 16px; padding: 12px 14px;
    box-shadow: 0 10px 30px -12px rgba(0,0,0,0.6);
}
.app-header h1 { font-size: 26px; font-weight: 800; color:#FFFFFF; margin:0; letter-spacing:-0.5px; }
.app-header .desc { font-size: 13.5px; color:#9CA3AF; margin-top:2px; font-weight:500; }
.app-header .right { display:flex; align-items:center; gap:10px; }

/* Botões do header */
.btn-share {
    background: linear-gradient(135deg, #2563EB, #3B82F6);
    border: 1px solid rgba(96,165,250,0.6);
    color:#fff; border-radius:10px; padding:9px 18px;
    font-size:13.5px; font-weight:700; cursor:pointer;
    box-shadow: 0 8px 22px -8px rgba(37,99,235,0.7);
    transition: transform .15s ease;
}
.btn-share:hover { transform: translateY(-1px); }
.icon-btn {
    width:38px; height:38px; border-radius:10px;
    background: rgba(255,255,255,0.05);
    border: 1px solid rgba(255,255,255,0.09);
    display:flex; align-items:center; justify-content:center;
    font-size:16px; color:#D1D5DB; cursor:pointer; transition: all .15s ease;
}
.icon-btn:hover { background: rgba(255,255,255,0.10); }
.avatar {
    width:38px; height:38px; border-radius:50%;
    background: linear-gradient(135deg,#a855f7,#22d3ee);
    display:flex; align-items:center; justify-content:center;
    font-weight:800; color:#fff; font-size:14px;
    border: 2px solid rgba(255,255,255,0.15);
}

/* ----------------------- GLASS CONTAINERS ----------------------- */
.glass {
    background: rgba(255,255,255,0.035);
    backdrop-filter: blur(14px); -webkit-backdrop-filter: blur(14px);
    border: 1px solid rgba(255,255,255,0.08);
    border-radius: 18px;
    box-shadow: 0 20px 50px -25px rgba(0,0,0,0.7);
}

/* ----------------------- DROP ZONE ----------------------- */
.dropzone { padding: 30px; position: relative; }
.tag-rule {
    position:absolute; top:16px; right:18px;
    background: rgba(255,255,255,0.05);
    border: 1px solid rgba(255,255,255,0.10);
    color:#9CA3AF; font-size:11px; font-weight:600;
    padding:5px 11px; border-radius:20px;
}
.flow {
    display:flex; align-items:center; justify-content:center; gap:18px;
    padding: 10px 0 22px;
}
.flow .node {
    width:78px; height:78px; border-radius:18px;
    display:flex; flex-direction:column; align-items:center; justify-content:center;
    font-size:30px; gap:2px;
}
.flow .node small { font-size:10px; font-weight:700; letter-spacing:.5px; color:#D1D5DB; }
.flow .pdf { background: rgba(239,68,68,0.12); border:1px solid rgba(239,68,68,0.35); }
.flow .xls { background: rgba(34,197,94,0.12); border:1px solid rgba(34,197,94,0.35); }
.flow .arrow {
    color:#6B7280; font-size:20px; letter-spacing:-3px;
    animation: dash 1.6s linear infinite;
}
@keyframes dash { 0%{opacity:.3} 50%{opacity:1} 100%{opacity:.3} }

/* Estiliza o file_uploader dentro do glass */
.dropzone .stFileUploader { background: transparent; }
.dropzone [data-testid="stFileUploaderDropzone"] {
    background: rgba(56,189,248,0.05);
    border: 1.5px dashed rgba(96,165,250,0.45);
    border-radius: 14px;
    transition: all .2s ease;
}
.dropzone [data-testid="stFileUploaderDropzone"]:hover {
    border-color: rgba(96,165,250,0.8);
    background: rgba(56,189,248,0.09);
}

/* ----------------------- PREVIEW CARDS ----------------------- */
.io-grid { display:flex; gap:18px; margin-top:22px; }
.io-card {
    flex:1; padding:18px; display:flex; gap:14px; align-items:center;
}
.io-card .ic-box {
    width:52px; height:52px; border-radius:13px; flex-shrink:0;
    display:flex; align-items:center; justify-content:center; font-size:24px;
}
.io-card.in .ic-box { background: rgba(239,68,68,0.13); border:1px solid rgba(239,68,68,0.30); }
.io-card.out .ic-box { background: rgba(34,197,94,0.13); border:1px solid rgba(34,197,94,0.30); }
.io-card h4 { margin:0; font-size:14px; color:#F3F4F6; font-weight:700; }
.io-card p { margin:3px 0 0; font-size:11.5px; color:#9CA3AF; }
.io-thumb {
    margin-top:10px; height:8px; border-radius:4px;
    background: repeating-linear-gradient(90deg, rgba(255,255,255,0.10) 0 40%, transparent 40% 52%);
}

/* ----------------------- COMPONENTES STREAMLIT ----------------------- */
.stButton > button {
    background: linear-gradient(135deg, rgba(56,189,248,0.18), rgba(168,85,247,0.18));
    border: 1px solid rgba(96,165,250,0.45);
    color:#fff; border-radius:12px; font-weight:700;
    transition: all .18s ease;
}
.stButton > button:hover {
    transform: translateY(-1px);
    box-shadow: 0 10px 26px -10px rgba(56,189,248,0.6);
    border-color: rgba(96,165,250,0.8);
}
/* Download button */
.stDownloadButton > button {
    background: linear-gradient(135deg, #16a34a, #22c55e) !important;
    border: 1px solid rgba(74,222,128,0.6) !important;
    color:#fff !important; border-radius:12px; font-weight:700;
}
.stDownloadButton > button:hover { transform: translateY(-1px); }

/* Dataframes */
.stDataFrame { border-radius: 14px; overflow:hidden; }
[data-testid="stDataFrame"] {
    border: 1px solid rgba(255,255,255,0.08);
    border-radius: 14px;
}
/* Métricas */
[data-testid="stMetric"] {
    background: rgba(255,255,255,0.035);
    border: 1px solid rgba(255,255,255,0.08);
    border-radius: 14px; padding: 14px 16px;
}
[data-testid="stMetricLabel"] { color:#9CA3AF !important; font-size:12px; }
[data-testid="stMetricValue"] { color:#FFFFFF !important; }

/* Headings color */
h1, h2, h3 { color:#FFFFFF !important; }
hr { border-color: rgba(255,255,255,0.08); }
</style>
"""


def nav_card_html():
    """HTML do cabeçalho e status da sidebar."""
    conectado = supabase is not None
    status_cls = "" if conectado else " off"
    txt = "Conectado ao Supabase" if conectado else "Supabase offline (local)"
    sub = NOME_BUCKET if conectado else "configure suas secrets"
    return f"""
    <div class="sb-title">
        <span class="ic">⚙️</span>
        <h1>Painel de Controle</h1>
    </div>
    <p class="sb-sub">Weversongc</p>
    <div class="sb-status{status_cls}">
        <span class="dot"></span>
        <div>
            <span>{txt}</span>
            <small>{sub}</small>
        </div>
    </div>
    """


def header_html(icone, titulo, desc, icone_cor=None):
    """Cabeçalho principal com botões de ação à direita."""
    return f"""
    <div class="app-header">
        <div class="left">
            <div class="big-ic">{icone}</div>
            <div>
                <h1>{titulo}</h1>
                <div class="desc">{desc}</div>
            </div>
        </div>
        <div class="right">
            <button class="btn-share">🔗 Compartilhar</button>
            <div class="icon-btn" title="Editar">✏️</div>
            <a href="https://github.com/weversongc/weversongc" target="_blank" class="icon-btn" title="GitHub">🐙</a>
            <div class="avatar" title="Weversongc">W</div>
        </div>
    </div>
    """


# =========================================================================
# 4. FUNÇÃO AUXILIAR SUPABASE
# =========================================================================
def enviar_para_supabase(nome_arquivo: str, dados_bytes: bytes) -> None:
    if supabase is None:
        st.info("ℹ️ Supabase não configurado. Arquivo disponível apenas para download.")
        return
    try:
        supabase.storage.from_(NOME_BUCKET).upload(
            path=nome_arquivo,
            file=dados_bytes,
            file_options={
                "content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            },
        )
        st.success(f"☁️ Arquivo '{nome_arquivo}' salvo no Supabase Storage!")
    except Exception as e:
        st.warning(f"Aviso ao salvar na nuvem: {e}")


# =========================================================================
# 5. RENDER
# =========================================================================
st.markdown(CSS, unsafe_allow_html=True)

# ---- SIDEBAR ----
with st.sidebar:
    st.markdown(nav_card_html(), unsafe_allow_html=True)
    st.markdown("###")
    opcao = st.radio(
        "menu",
        [
            "📄  Conversor PDF → Excel",
            "📊  Conciliação Bancária",
            "💸  Diferenças Braslog × MH",
        ],
        label_visibility="collapsed",
    )

# ---- HEADER + CONTEÚDO ----
if opcao.startswith("📄"):
    st.markdown(header_html("📄", "Conversor Inteligente de PDF → XLSX",
                            "Transforme relatórios em PDF em abas estruturadas no Excel."), unsafe_allow_html=True)

    st.markdown('<div class="glass dropzone">', unsafe_allow_html=True)
    st.markdown('<div class="tag-rule">200MB per file • PDF</div>', unsafe_allow_html=True)
    st.markdown("""
        <div class="flow">
            <div class="node pdf">📕<small>PDF</small></div>
            <div class="arrow">• • ▸</div>
            <div class="node xls">📗<small>XLSX</small></div>
        </div>
    """, unsafe_allow_html=True)
    arquivo_pdf = st.file_uploader("Arraste seu arquivo PDF aqui", type=["pdf"], label_visibility="collapsed")
    st.markdown('</div>', unsafe_allow_html=True)

    # Cards Input/Output
    st.markdown("""
        <div class="io-grid">
            <div class="glass io-card in">
                <div class="ic-box">📕</div>
                <div><h4>Input PDF</h4><p>Documento de origem</p><div class="io-thumb"></div></div>
            </div>
            <div class="glass io-card out">
                <div class="ic-box">📗</div>
                <div><h4>Output Excel</h4><p>Planilha gerada</p><div class="io-thumb"></div></div>
            </div>
        </div>
    """, unsafe_allow_html=True)

    if arquivo_pdf is not None:
        st.markdown("###")
        senha = st.text_input("Senha do PDF (se protegido)", type="password", placeholder="Deixe em branco se não tiver senha")
        if st.button("🚀 Iniciar Conversão", use_container_width=True):
            with st.spinner("Processando páginas do PDF e montando uma única aba..."):
                conv = ConversorPDF()
                res = conv.converter_para_excel(arquivo_pdf, senha=senha or "")
                if not res["ok"]:
                    st.error(f"Erro na conversão: {res['erro']}")
                else:
                    dados_excel = res["dados"]
                    st.success("🎉 PDF convertido com sucesso!")
                    m1, m2, m3 = st.columns(3)
                    m1.metric("Páginas", res["paginas"])
                    m2.metric("Como Tabela", res["tabelas"])
                    m3.metric("Como Texto", res["paginas_texto"])
                    st.caption("📄 Tudo consolidado em **uma única aba** preservando o layout do PDF.")
                    nome_final = arquivo_pdf.name.replace(".pdf", ".xlsx")
                    st.download_button("📥 Baixar Arquivo Excel", dados_excel, file_name=nome_final,
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                       use_container_width=True)
                    enviar_para_supabase(f"conversoes/{nome_final}", dados_excel)

# ---- CONCILIAÇÃO ----
elif opcao.startswith("📊"):
    st.markdown(header_html("📊", "Conciliação de Lançamentos Financeiros",
                            "Cruza os dados da Braslog com o extrato bancário do Sicoob."), unsafe_allow_html=True)

    st.markdown('<div class="glass dropzone">', unsafe_allow_html=True)
    st.markdown('<div class="tag-rule">XLSX + CSV</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        arq_braslog = st.file_uploader("📕 Balanço Braslog (XLSX)", type=["xlsx"])
    with c2:
        arq_sicoob = st.file_uploader("🧾 Extrato Sicoob (CSV)", type=["csv"])
    st.markdown('</div>', unsafe_allow_html=True)

    if arq_braslog and arq_sicoob:
        st.markdown("###")
        if st.button("🔄 Executar Conciliação", use_container_width=True):
            with st.spinner("Cruzando dados..."):
                try:
                    df_plan = pd.read_excel(arq_braslog, sheet_name="3", header=5)
                    df_plan.columns = [str(c).strip() for c in df_plan.columns]
                    df_plan = df_plan[df_plan["CONTA"] == "SIC"].copy()
                    df_plan["DATA"] = pd.to_datetime(df_plan["DATA"], errors="coerce").dt.normalize()
                    df_plan["Recebido"] = pd.to_numeric(df_plan["Recebido"], errors="coerce").fillna(0)
                    df_plan["Pago"] = pd.to_numeric(df_plan["Pago"], errors="coerce").fillna(0)

                    def compoe(row):
                        if row["Recebido"] > 0:
                            return pd.Series({"Valor": row["Recebido"], "Direcao": "C"})
                        if row["Pago"] > 0:
                            return pd.Series({"Valor": row["Pago"], "Direcao": "D"})
                        return pd.Series({"Valor": 0.0, "Direcao": None})

                    df_plan[["Valor", "Direcao"]] = df_plan.apply(compoe, axis=1)
                    df_plan = df_plan[df_plan["Valor"] > 0].copy()
                    df_plan["Valor"] = df_plan["Valor"].round(2)

                    df_ext = pd.read_csv(arq_sicoob, sep=",", decimal=",", thousands=".", encoding="utf-8")
                    df_ext.columns = [str(c).strip() for c in df_ext.columns]
                    df_ext["data"] = pd.to_datetime(df_ext["data"], format="%d/%m/%Y", errors="coerce").dt.normalize()
                    df_ext["valor"] = pd.to_numeric(df_ext["valor"], errors="coerce")
                    df_ext = df_ext.rename(columns={"data": "DATA", "valor": "Valor"})
                    df_ext["Direcao"] = df_ext["tipo"]
                    df_ext["Valor"] = df_ext["Valor"].round(2)

                    dmin, dmax = df_ext["DATA"].min(), df_ext["DATA"].max()
                    df_plan = df_plan[(df_plan["DATA"] >= dmin) & (df_plan["DATA"] <= dmax)].copy()
                    df_plan["Seq"] = df_plan.groupby(["DATA", "Valor", "Direcao"]).cumcount()
                    df_ext["Seq"] = df_ext.groupby(["DATA", "Valor", "Direcao"]).cumcount()

                    cols_plan = ["FORNECEDOR", "DESTINO", "Categoria", "SUB Categoria", "OBSERVAÇÃO"]
                    cols_ext = ["historico", "detalhes"]
                    dk_plan = df_plan[["DATA", "Valor", "Direcao", "Seq"] + [c for c in cols_plan if c in df_plan.columns]].copy()
                    dk_ext = df_ext[["DATA", "Valor", "Direcao", "Seq"] + [c for c in cols_ext if c in df_ext.columns]].copy()
                    df = pd.merge(dk_plan, dk_ext, on=["DATA", "Valor", "Direcao", "Seq"], how="outer", suffixes=("_p", "_e"), indicator=True)
                    df["Status"] = df["_merge"].map({"both": "Ok", "left_only": "Falta no Extrato", "right_only": "Falta na Planilha"})
                    df = df.drop(columns=["_merge"])
                    df["Direcao"] = df["Direcao"].map({"C": "Entrada", "D": "Saída"})
                    meses = {1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 6: "Junho",
                             7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"}
                    df["Mes"] = df["DATA"].dt.month.map(meses)
                    ordem = ["Status", "Mes", "DATA", "Valor", "Direcao", "FORNECEDOR", "DESTINO",
                             "Categoria", "SUB Categoria", "OBSERVAÇÃO", "historico", "detalhes"]
                    df = df[[c for c in ordem if c in df.columns]]
                    df = df.sort_values(["DATA", "Status", "Valor"]).reset_index(drop=True)

                    st.success("Conciliação efetuada com sucesso!")
                    st.markdown("#### Resumo")
                    st.dataframe(df["Status"].value_counts().reset_index().rename(columns={"index": "Status", "Status": "Quantidade"}),
                                 use_container_width=True)
                    st.markdown("#### Prévia dos Resultados")
                    st.dataframe(df, use_container_width=True)
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine="openpyxl") as w:
                        df.to_excel(w, index=False, sheet_name="Resultado")
                    dados = output.getvalue()
                    st.download_button("📥 Baixar Resultado da Conciliação", dados, file_name="Resultado_Conciliacao.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                    enviar_para_supabase("conciliacoes/Resultado_Conciliacao.xlsx", dados)
                except Exception as e:
                    st.error(f"Erro ao processar conciliação: {e}")

# ---- DIFERENÇAS ----
elif opcao.startswith("💸"):
    st.markdown(header_html("💸", "Análise de Transferências",
                            "Filtre e calcule o saldo das transferências entre Braslog e MH Gestão."), unsafe_allow_html=True)

    st.markdown('<div class="glass dropzone">', unsafe_allow_html=True)
    st.markdown('<div class="tag-rule">XLSX • Aba 3</div>', unsafe_allow_html=True)
    arq_diferenca = st.file_uploader("Upload: Planilha de Lançamentos (XLSX)", type=["xlsx"], label_visibility="collapsed")
    st.markdown('</div>', unsafe_allow_html=True)

    if arq_diferenca:
        st.markdown("###")
        if st.button("🔍 Analisar Diferenças", use_container_width=True):
            with st.spinner("Filtrando e calculando saldos..."):
                try:
                    df = pd.read_excel(arq_diferenca, sheet_name="3", header=5)
                    df.columns = [str(c).strip() for c in df.columns]
                    col_obs = [c for c in df.columns if "OBS" in c.upper()][0]
                    df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce").dt.normalize()
                    df["Recebido"] = pd.to_numeric(df["Recebido"], errors="coerce").fillna(0)
                    df["Pago"] = pd.to_numeric(df["Pago"], errors="coerce").fillna(0)
                    cols_busca = ["DESTINO", "FORNECEDOR", "Categoria", "SUB Categoria", "Centro de custo", col_obs]
                    texto = df[[c for c in cols_busca if c in df.columns]].fillna("").astype(str).agg(" ".join, axis=1).str.upper()
                    mh = texto.str.contains("MH GEST", na=False) | texto.str.contains("MHGEST", na=False)
                    aport = (df["Categoria"].fillna("").str.upper() == "APORTES") | texto.str.contains("EMPREST", na=False)
                    adiant = texto.str.contains("ADIANT|ANTECIP|SALARI|SALÁR", na=True, regex=True)
                    df_f = df[mh & aport & ~adiant].copy()

                    def dirx(r):
                        if r["Pago"] > 0:
                            return "Braslog → MH (MH deve a Braslog)"
                        if r["Recebido"] > 0:
                            return "MH → Braslog (Braslog deve a MH)"
                        return ""
                    df_f["Direção"] = df_f.apply(dirx, axis=1)
                    tp, tr, saldo = df_f["Pago"].sum(), df_f["Recebido"].sum(), df_f["Pago"].sum() - df_f["Recebido"].sum()
                    if saldo > 0.005:
                        dev, cred, val = "MH", "Braslog", saldo
                    elif saldo < -0.005:
                        dev, cred, val = "Braslog", "MH", abs(saldo)
                    else:
                        dev, cred, val = None, None, 0.0
                    ordem = ["Direção", "DATA", "CONTA", "BANCO", "DESTINO", "FORNECEDOR", "Centro de custo",
                             "Categoria", "SUB Categoria", "Recebido", "Pago", col_obs]
                    df_f = df_f[[c for c in ordem if c in df_f.columns]].sort_values("DATA").reset_index(drop=True)
                    resumo = [
                        {c: "" for c in df_f.columns},
                        {**{c: "" for c in df_f.columns}, "Direção": "RESUMO DAS TRANSFERÊNCIAS BRASLOG ↔ MH"},
                        {**{c: "" for c in df_f.columns}, "Direção": "Total pago por Braslog a MH:", "Pago": tp},
                        {**{c: "" for c in df_f.columns}, "Direção": "Total recebido de MH:", "Recebido": tr},
                        {**{c: "" for c in df_f.columns}, "Direção": "Saldo líquido:", "Pago": saldo},
                    ]
                    if dev:
                        resumo.append({**{c: "" for c in df_f.columns}, "Direção": f"=> {dev} DEVE a {cred}:", "Pago": val})
                    df_saida = pd.concat([df_f, pd.DataFrame(resumo)], ignore_index=True)

                    st.success("Análise concluída!")
                    st.markdown("#### Resumo Financeiro")
                    m1, m2, m3 = st.columns(3)
                    m1.metric("Total Pago (→ MH)", f"R$ {tp:,.2f}")
                    m2.metric("Total Recebido (← MH)", f"R$ {tr:,.2f}")
                    m3.metric("Saldo Líquido", f"R$ {saldo:,.2f}")
                    if dev:
                        st.info(f"➡️ **{dev} DEVE a {cred}: R$ {val:,.2f}**")
                    else:
                        st.info("➡️ Contas zeradas.")
                    st.markdown("#### Detalhamento")
                    st.dataframe(df_saida, use_container_width=True)
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine="openpyxl") as w:
                        df_saida.to_excel(w, index=False, sheet_name="Diferenças")
                    dados = output.getvalue()
                    st.download_button("📥 Baixar Relatório de Diferenças", dados, file_name="Diferencas_Abas.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                    enviar_para_supabase("diferencas/Diferencas_Abas.xlsx", dados)
                except Exception as e:
                    st.error(f"Erro ao processar análise: {e}")
