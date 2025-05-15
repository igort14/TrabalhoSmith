import streamlit as st
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
from sklearn.linear_model import LinearRegression
import numpy as np
import io

st.set_page_config(page_title="Analisador de Investimentos - Secretaria de Esporte", layout="wide")
st.title("🏢 Análise de Investimentos e ROI - Secretaria de Esporte e Lazer do DF")

# Upload do arquivo Excel
uploaded_file = st.file_uploader("📂 Envie a planilha de investimentos (2014-2024)", type=["xlsx"])

# Leitura do arquivo local padrão, se nenhum for enviado
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.success("✅ Arquivo enviado carregado com sucesso!")
else:
    try:
        df = pd.read_excel("Investimentos_SEL_DF_2014_2024_com_ROI.xlsx")
        st.info("ℹ️ Nenhum arquivo enviado. Carregado arquivo local padrão.")
    except FileNotFoundError:
        st.error("❌ Nenhum arquivo enviado e o arquivo local não foi encontrado.")
        st.stop()

# Mostrar colunas para depuração
st.write("🔍 Colunas detectadas na planilha:", df.columns.tolist())

# Função de conversão BRL para float
def brl_to_float(value):
    if isinstance(value, str):
        return float(value.replace('.', '').replace(',', '.'))
    return value

# Padroniza nomes das colunas (tira espaços e capitaliza)
df.columns = df.columns.str.strip()

# Lista de colunas esperadas
colunas_esperadas = ['Liquidado (R$)', 'ROI 1,5x (R$)', 'ROI 2,0x (R$)', 'ROI 2,5x (R$)']

# Confirma se todas as colunas estão presentes
for col in colunas_esperadas:
    if col in df.columns:
        df[col] = df[col].apply(brl_to_float)
    else:
        st.warning(f"⚠️ A coluna '{col}' não foi encontrada na planilha.")

st.dataframe(df.style.format("{:.0f}"))

# Gráfico 1 - Evolução dos investimentos
if 'Ano' in df.columns and 'Liquidado (R$)' in df.columns:
    st.subheader("📈 Evolução dos Investimentos")
    fig1, ax1 = plt.subplots()
    sns.lineplot(data=df, x='Ano', y='Liquidado (R$)', marker='o', ax=ax1)
    st.pyplot(fig1)

# Gráfico 2 - Comparativo dos ROIs
if all(col in df.columns for col in ['Ano'] + colunas_esperadas[1:]):
    st.subheader("🔄 Comparativo de ROI por Cenário")
    fig2, ax2 = plt.subplots()
    sns.lineplot(data=df, x='Ano', y='ROI 1,5x (R$)', label='Conservador', marker='o', ax=ax2)
    sns.lineplot(data=df, x='Ano', y='ROI 2,0x (R$)', label='Moderado', marker='o', ax=ax2)
    sns.lineplot(data=df, x='Ano', y='ROI 2,5x (R$)', label='Otimista', marker='o', ax=ax2)
    st.pyplot(fig2)

# Gráfico 3 - Crescimento percentual
if 'Ano' in df.columns and 'Liquidado (R$)' in df.columns:
    st.subheader("📊 Crescimento Percentual Anual")
    df['Crescimento (%)'] = df['Liquidado (R$)'].pct_change() * 100
    fig3, ax3 = plt.subplots()
    sns.barplot(data=df, x='Ano', y='Crescimento (%)', palette='coolwarm', ax=ax3)
    st.pyplot(fig3)

# Previsão de Investimentos futuros
if 'Ano' in df.columns and 'Liquidado (R$)' in df.columns:
    st.subheader("🤖 Previsão de Investimentos e ROI (10 anos futuros)")

    modelo_inv = LinearRegression()
    modelo_inv.fit(df[['Ano']], df['Liquidado (R$)'])

    ultimo_ano = int(df['Ano'].max())
    anos_futuros = pd.DataFrame({'Ano': np.arange(ultimo_ano + 1, ultimo_ano + 11)})
    anos_futuros['Investimento Previsto (R$)'] = modelo_inv.predict(anos_futuros).round(2)
    anos_futuros['Ano'] = anos_futuros['Ano'].astype(int)

    # Cálculo de ROIs
    anos_futuros['ROI 1,5x (R$)'] = (anos_futuros['Investimento Previsto (R$)'] * 1.5).round(2)
    anos_futuros['ROI 2,0x (R$)'] = (anos_futuros['Investimento Previsto (R$)'] * 2.0).round(2)
    anos_futuros['ROI 2,5x (R$)'] = (anos_futuros['Investimento Previsto (R$)'] * 2.5).round(2)

    st.dataframe(anos_futuros.style.format({
        'Ano': "{:d}",
        'Investimento Previsto (R$)': "{:.2f}",
        'ROI 1,5x (R$)': "{:.2f}",
        'ROI 2,0x (R$)': "{:.2f}",
        'ROI 2,5x (R$)': "{:.2f}"
    }))

    # Gráfico da previsão
    fig4, ax4 = plt.subplots()
    sns.lineplot(data=anos_futuros, x='Ano', y='Investimento Previsto (R$)', label='Investimento Previsto', marker='o', ax=ax4)
    sns.lineplot(data=anos_futuros, x='Ano', y='ROI 1,5x (R$)', label='ROI 1.5x', linestyle='--', ax=ax4)
    sns.lineplot(data=anos_futuros, x='Ano', y='ROI 2,0x (R$)', label='ROI 2.0x', linestyle='--', ax=ax4)
    sns.lineplot(data=anos_futuros, x='Ano', y='ROI 2,5x (R$)', label='ROI 2.5x', linestyle='--', ax=ax4)
    st.pyplot(fig4)

    # Exportação para Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        anos_futuros.to_excel(writer, index=False, sheet_name='Previsoes_Investimentos', float_format="%.2f")
        workbook  = writer.book
        worksheet = writer.sheets['Previsoes_Investimentos']
        for i, col in enumerate(anos_futuros.columns):
            column_len = max(anos_futuros[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, column_len)
    st.download_button(label="📄 Baixar planilha com previsões",
                    data=output.getvalue(),
                    file_name="Previsoes_Investimentos_ROI_futuro.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")