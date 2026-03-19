import streamlit as st
import pandas as pd
import plotly.express as px

# CONFIGURAÇÃO DA PÁGINA
st.set_page_config(
    page_title="Auditoria FATURA x APEX",
    page_icon="📊",
    layout="wide"
)

# TEMA
st.markdown("""
<style>

.stApp {
background-color: #f4f6fb;
}

h1, h2, h3 {
color: #ff8c00;
}

div[data-testid="stMetric"] {
background-color: white;
padding: 20px;
border-radius: 12px;
border-left: 6px solid #ff8c00;
box-shadow: 0px 4px 12px rgba(0,0,0,0.08);
}

[data-testid="stDataFrame"] {
background-color: white;
border-radius: 10px;
box-shadow: 0px 3px 10px rgba(0,0,0,0.08);
}

.stButton>button {
background-color: #ff8c00;
color: white;
border-radius: 8px;
border: none;
}

</style>
""", unsafe_allow_html=True)

# TÍTULO
st.title("📊 Auditoria FATURA x APEX")

st.divider()

# CONTROLES DE VISUALIZAÇÃO
colA, colB = st.columns(2)

mostrar_dash = colA.toggle("Mostrar Dashboards", True)
mostrar_graficos = colB.toggle("Mostrar Gráficos", True)

st.divider()

# UPLOAD
file = st.file_uploader("Carregar planilha Excel", type=["xlsx","xlsm","xls"])

if file:

    df = pd.read_excel(file)

    # BUSCA
    busca = st.text_input("🔎 Buscar Código")

    if busca:
        df = df[df["Código"].astype(str).str.contains(busca)]

    # FILTRO STATUS
    status = st.radio(
        "Filtrar Status",
        ["Todos"] + list(df["Status"].unique()),
        horizontal=True
    )

    if status != "Todos":
        df = df[df["Status"] == status]

    # MÉTRICAS
    if mostrar_dash:

        total = len(df)
        iguais = (df["Status"] == "Valores iguais").sum()
        divergente = (df["Status"] != "Valores iguais").sum()

        col1, col2, col3 = st.columns(3)

        col1.metric("Total", total)
        col2.metric("Valores Iguais", iguais)
        col3.metric("Divergentes", divergente)

        st.divider()

    # GRÁFICOS
    if mostrar_graficos:

        col1, col2 = st.columns(2)

        cores = {
            "Valores iguais": "#2ecc71",
            "Divergente": "#e74c3c",
            "Falta informação": "#f39c12"
        }

        with col1:

            grafico = df["Status"].value_counts().reset_index()
            grafico.columns = ["Status","Quantidade"]

            fig = px.bar(
                grafico,
                x="Status",
                y="Quantidade",
                color="Status",
                text="Quantidade",
                color_discrete_map=cores,
                title="Distribuição de Status"
            )

            st.plotly_chart(fig, use_container_width=True)

        with col2:

            fig2 = px.scatter(
                df,
                x="Valor FATURA",
                y="Valor APEX",
                color="Status",
                color_discrete_map=cores,
                hover_data=["Código"],
                title="Comparação FATURA x APEX"
            )

            st.plotly_chart(fig2, use_container_width=True)

        st.divider()

    # DIFERENÇA
    df["Diferença"] = df["Valor FATURA"] - df["Valor APEX"]

    # TABELA
    def color_status(row):

        if row["Status"] == "Valores iguais":
            return ['background-color:#eafaf1'] * len(row)

        if row["Status"] == "Divergente":
            return ['background-color:#fdecea'] * len(row)

        return [''] * len(row)

    st.subheader("📋 Dados")

    st.dataframe(
        df.style.apply(color_status, axis=1),
        use_container_width=True,
        height=800
    )

    # DOWNLOAD
    st.download_button(
        "📥 Baixar dados filtrados",
        df.to_csv(index=False),
        file_name="dados_filtrados.csv"
    )