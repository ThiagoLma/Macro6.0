import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
import io

# ==============================
# CONFIGURAÇÃO DA PÁGINA
# ==============================
st.set_page_config(
    page_title="Auditoria FATURA x APEX",
    page_icon="📊",
    layout="wide"
)

st.markdown("""
<style>
.stApp { background-color: #f4f6fb; }
h1, h2, h3 { color: #ff8c00; }
div[data-testid="stMetric"] {
    background-color: white;
    padding: 20px;
    border-radius: 12px;
    border-left: 6px solid #ff8c00;
    box-shadow: 0px 4px 12px rgba(0,0,0,0.08);
}
.stButton>button {
    background-color: #ff8c00;
    color: white;
    border-radius: 8px;
    border: none;
}
.upload-box {
    background: white;
    border-radius: 12px;
    padding: 16px;
    box-shadow: 0px 3px 10px rgba(0,0,0,0.07);
    margin-bottom: 10px;
}
</style>
""", unsafe_allow_html=True)

# ==============================
# TÍTULO
# ==============================
st.title("📊 Auditoria FATURA x APEX")
st.caption("Carregue os arquivos abaixo para iniciar a comparação.")
st.divider()

# ==============================
# FUNÇÃO: Comparar dados (substitui a macro VBA)
# ==============================
def comparar_fatura_apex(df_fatura, df_apex):
    """
    Replica a lógica da macro VBA CompararFaturaApex:
    - Cruza FATURA e APEX pelo código/tombo
    - Compara valores e classifica o status
    """
    df_fatura = df_fatura.copy()
    df_apex = df_apex.copy()

    # Padroniza nomes das colunas geradas pelo ler_fatura / ler_apex
    df_fatura.columns = ["PAT", "Valor FATURA", "Num Série"]
    df_apex.columns = ["Tombo", "Valor APEX"]

    # Garante tipos string para o merge
    df_fatura["PAT"] = df_fatura["PAT"].astype(str).str.strip()
    df_apex["Tombo"] = df_apex["Tombo"].astype(str).str.strip()

    # Merge pelo código
    df_merged = pd.merge(
        df_fatura,
        df_apex,
        left_on="PAT",
        right_on="Tombo",
        how="outer",
        indicator=True
    )

    # Classifica o status
    def classificar(row):
        if row["_merge"] == "left_only":
            return "Falta no APEX"
        if row["_merge"] == "right_only":
            return "Falta na FATURA"
        try:
            fat = float(row["Valor FATURA"])
            apex = float(row["Valor APEX"])
            if abs(fat - apex) < 0.01:
                return "Valores iguais"
            else:
                return "Divergente"
        except Exception:
            return "Falta informação"

    df_merged["Status"] = df_merged.apply(classificar, axis=1)
    df_merged["Código"] = df_merged["PAT"].combine_first(df_merged["Tombo"])

    # Reorganiza colunas
    colunas = ["Código", "Num Série", "Valor FATURA", "Valor APEX", "Status"]
    df_resultado = df_merged[[c for c in colunas if c in df_merged.columns]].copy()
    df_resultado["Valor FATURA"] = pd.to_numeric(df_resultado["Valor FATURA"], errors="coerce")
    df_resultado["Valor APEX"]   = pd.to_numeric(df_resultado["Valor APEX"],   errors="coerce")
    df_resultado["Diferença"]    = df_resultado["Valor FATURA"] - df_resultado["Valor APEX"]

    return df_resultado


# ==============================
# FUNÇÃO: Gerar Excel de saída
# ==============================
def gerar_excel(df_resultado, df_fatura_raw, df_apex_raw):
    wb = load_workbook(filename=io.BytesIO(b""))  # wb vazio
    # Cria do zero com openpyxl puro
    from openpyxl import Workbook
    wb = Workbook()

    # ---- Aba RESULTADO ----
    ws_res = wb.active
    ws_res.title = "RESULTADO"

    header_fill = PatternFill("solid", fgColor="FF8C00")
    header_font = Font(bold=True, color="FFFFFF")

    for col_idx, col_name in enumerate(df_resultado.columns, start=1):
        cell = ws_res.cell(row=1, column=col_idx, value=col_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    fill_igual      = PatternFill("solid", fgColor="EAFAF1")
    fill_divergente = PatternFill("solid", fgColor="FDECEA")
    fill_falta      = PatternFill("solid", fgColor="FEF9E7")

    for r_idx, row in enumerate(df_resultado.itertuples(index=False), start=2):
        status = getattr(row, "Status", "")
        for c_idx, value in enumerate(row, start=1):
            cell = ws_res.cell(row=r_idx, column=c_idx, value=value)
            if status == "Valores iguais":
                cell.fill = fill_igual
            elif status == "Divergente":
                cell.fill = fill_divergente
            else:
                cell.fill = fill_falta

    for col in ws_res.columns:
        max_len = max((len(str(c.value)) if c.value else 0) for c in col)
        ws_res.column_dimensions[col[0].column_letter].width = min(max_len + 4, 40)

    # ---- Aba FATURA ----
    ws_fat = wb.create_sheet("FATURA")
    for col_idx, col_name in enumerate(["PAT", "Tot Geral", "Num Série"], start=1):
        ws_fat.cell(row=1, column=col_idx, value=col_name).font = Font(bold=True)
    for r_idx, row in enumerate(df_fatura_raw.itertuples(index=False), start=2):
        for c_idx, value in enumerate(row, start=1):
            ws_fat.cell(row=r_idx, column=c_idx, value=value)

    # ---- Aba APEX ----
    ws_apex = wb.create_sheet("APEX")
    for col_idx, col_name in enumerate(["Tombo", "Vr Loc"], start=1):
        ws_apex.cell(row=1, column=col_idx, value=col_name).font = Font(bold=True)
    for r_idx, row in enumerate(df_apex_raw.itertuples(index=False), start=2):
        for c_idx, value in enumerate(row, start=1):
            ws_apex.cell(row=r_idx, column=c_idx, value=value)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


# ==============================
# UPLOADS
# ==============================
col1, col2 = st.columns(2)

with col1:
    st.markdown("#### 📄 Arquivo FATURA")
    file_fatura = st.file_uploader(
        "Selecione o arquivo da FATURA",
        type=["xlsx", "xlsm", "xls"],
        key="fatura"
    )

with col2:
    st.markdown("#### 📄 Arquivo APEX")
    file_apex = st.file_uploader(
        "Selecione o arquivo do APEX",
        type=["xlsx", "xlsm"],
        key="apex"
    )

st.divider()

# ==============================
# PROCESSAMENTO
# ==============================
if file_fatura and file_apex:

    with st.spinner("Lendo arquivos e comparando dados..."):

        # Lê FATURA — colunas: B(idx1)=PAT, E(idx4)=Tot Geral, A(idx0)=Num Série
        df_fat_raw_full = pd.read_excel(file_fatura, header=0)
        df_fatura_dest = pd.DataFrame({
            "PAT":       df_fat_raw_full.iloc[:, 1],
            "Tot Geral": df_fat_raw_full.iloc[:, 4],
            "Num Série": df_fat_raw_full.iloc[:, 0]
        })

        # Lê APEX — colunas: I(idx8)=Tombo, K(idx10)=Vr Loc
        df_apex_raw_full = pd.read_excel(file_apex, header=0)
        df_apex_dest = pd.DataFrame({
            "Tombo":  df_apex_raw_full.iloc[:, 8],
            "Vr Loc": df_apex_raw_full.iloc[:, 10]
        })

        # Executa comparação (substitui a macro VBA)
        df_resultado = comparar_fatura_apex(df_fatura_dest, df_apex_dest)

    # ==============================
    # MÉTRICAS
    # ==============================
    total      = len(df_resultado)
    iguais     = (df_resultado["Status"] == "Valores iguais").sum()
    divergente = (df_resultado["Status"] == "Divergente").sum()
    faltando   = total - iguais - divergente

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total de Registros", total)
    m2.metric("✅ Valores Iguais",   iguais)
    m3.metric("❌ Divergentes",      divergente)
    m4.metric("⚠️ Com Pendência",    faltando)

    st.divider()

    # ==============================
    # FILTROS
    # ==============================
    colA, colB = st.columns([2, 1])
    busca  = colA.text_input("🔎 Buscar por Código")
    status = colB.radio(
        "Filtrar Status",
        ["Todos"] + list(df_resultado["Status"].unique()),
        horizontal=True
    )

    df_view = df_resultado.copy()
    if busca:
        df_view = df_view[df_view["Código"].astype(str).str.contains(busca, case=False)]
    if status != "Todos":
        df_view = df_view[df_view["Status"] == status]

    # ==============================
    # TABELA
    # ==============================
    def color_status(row):
        if row["Status"] == "Valores iguais":
            return ["background-color:#eafaf1"] * len(row)
        if row["Status"] == "Divergente":
            return ["background-color:#fdecea"] * len(row)
        return ["background-color:#fef9e7"] * len(row)

    st.subheader("📋 Resultado da Auditoria")
    st.dataframe(
        df_view.style.apply(color_status, axis=1),
        use_container_width=True,
        height=500
    )

    st.divider()

    # ==============================
    # DOWNLOADS
    # ==============================
    col_dl1, col_dl2 = st.columns(2)

    with col_dl1:
        st.download_button(
            "📥 Baixar resultado (CSV)",
            df_view.to_csv(index=False).encode("utf-8"),
            file_name="auditoria_resultado.csv",
            mime="text/csv"
        )

    with col_dl2:
        excel_buffer = gerar_excel(df_resultado, df_fatura_dest, df_apex_dest)
        st.download_button(
            "📥 Baixar Excel completo (FATURA + APEX + RESULTADO)",
            excel_buffer,
            file_name="auditoria_completa.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("⬆️ Carregue os dois arquivos acima para iniciar a auditoria.")
