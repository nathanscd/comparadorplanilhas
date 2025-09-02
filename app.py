import streamlit as st
import pandas as pd
import difflib
from io import BytesIO
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook

def comparar_resumido_explicito(texto1, texto2):
    if pd.isna(texto1) or pd.isna(texto2):
        return "Um dos textos está vazio"

    d = difflib.Differ()
    diff = list(d.compare(str(texto1).split(), str(texto2).split()))

    removidos = sorted(set(x[2:] for x in diff if x.startswith("- ")))
    adicionados = sorted(set(x[2:] for x in diff if x.startswith("+ ")))

    if not removidos and not adicionados:
        return "Nenhuma diferença encontrada"

    resumo = ""
    if removidos:
        resumo += "Removido: " + ", ".join(removidos[:20])
        if len(removidos) > 20:
            resumo += ", ..."
        resumo += " | "
    if adicionados:
        resumo += "Adicionado: " + ", ".join(adicionados[:20])
        if len(adicionados) > 20:
            resumo += ", ..."

    return resumo

def processar_excel(df, col1, col2):
    df["Diferenças"] = df.apply(lambda row: comparar_resumido_explicito(row[col1], row[col2]), axis=1)

    buffer = BytesIO()
    df.to_excel(buffer, index=False)
    buffer.seek(0)

    wb = load_workbook(buffer)
    ws = wb.active

    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = 50 
        
        for i, cell in enumerate(col):
            if i == 0:
                cell.fill = PatternFill(start_color="ededed", end_color="ededed", fill_type="solid")
                cell.font = Font(bold=True)
            else:
                cell.alignment = Alignment(wrap_text=True, vertical='top')

    buffer_final = BytesIO()
    wb.save(buffer_final)
    buffer_final.seek(0)
    return buffer_final

st.set_page_config(
    page_title="Comparador de Planilhas",  # título da aba
    page_icon="📊",  # ícone da aba (pode ser emoji ou caminho para imagem .ico/.png)
    layout="wide"    # opcional: pode ser "centered" ou "wide"
)

st.title("Comparador de Planilhas")
st.write("Faça upload de um arquivo Excel e selecione as colunas que deseja comparar.")

uploaded_file = st.file_uploader(
    label="Envie aqui sua planilha para comparar",
    type=["xlsx", "xls"]
)

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.write("Pré-visualização dos dados:")
    st.dataframe(df.head())

    colunas = df.columns.tolist()
    st.write("Selecione as colunas que deseja comparar:")

    col1 = st.selectbox("Coluna 1", colunas, key="col1")
    col2 = st.selectbox("Coluna 2", colunas, key="col2")

    if col1 and col2 and col1 != col2:
        arquivo_processado = processar_excel(df, col1, col2)

        st.success(f"Comparação concluída entre '{col1}' e '{col2}'.")
        st.download_button(
            label="📥 Baixar Diferenças.xlsx",
            data=arquivo_processado,
            file_name="Diferenças.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("Selecione duas colunas diferentes para comparar.")
