import streamlit as st
import pandas as pd
import difflib
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Comparador de Planilhas", page_icon="📊")
st.title("Comparador de Planilhas")
st.markdown("Faça upload de duas planilhas para comparar os itens e gerar as diferenças.")

uploaded_file1 = st.file_uploader("Selecione a primeira planilha (base de referência)", type=["xlsx"])
uploaded_file2 = st.file_uploader("Selecione a segunda planilha (para buscar similaridades)", type=["xlsx"])

if uploaded_file1 and uploaded_file2:
    planilha1 = pd.ExcelFile(uploaded_file1)
    planilha2 = pd.ExcelFile(uploaded_file2)

    aba1 = st.selectbox("Escolha a aba da primeira planilha", planilha1.sheet_names)
    aba2 = st.selectbox("Escolha a aba da segunda planilha", planilha2.sheet_names)

    df1 = planilha1.parse(aba1)
    df2 = planilha2.parse(aba2)

    st.subheader("Visualização da primeira planilha")
    st.dataframe(df1)
    st.subheader("Visualização da segunda planilha")
    st.dataframe(df2)

    st.subheader("Configuração das colunas")
    col1 = st.selectbox("Coluna da primeira planilha", df1.columns)
    col2 = st.selectbox("Coluna da segunda planilha", df2.columns)

    def encontrar_mais_similar(texto, lista_textos):
        texto = str(texto)
        lista_textos = [str(t) for t in lista_textos]

        similar = difflib.get_close_matches(texto, lista_textos, n=1, cutoff=0.0)
        if not similar:
            return "Nenhum encontrado", "N/A"
        mais_similar = similar[0]

        d = difflib.Differ()
        diff = list(d.compare(texto.split(), mais_similar.split()))
        diffs = [x for x in diff if x.startswith("+ ") or x.startswith("- ")]
        diff_text = "Nenhuma diferença" if not diffs else " | ".join(diffs)
        return mais_similar, diff_text

    if st.button("Iniciar Conversão"):
        n = len(df1)
        progresso_bar = st.progress(0)
        progresso_text = st.empty()

        resultados = []
        for i, valor in enumerate(df1[col1]):
            mais_similar, diff_text = encontrar_mais_similar(valor, df2[col2].tolist())
            resultados.append((mais_similar, diff_text))

            progresso_bar.progress((i + 1) / n)
            progresso_text.text(f"Processando {i + 1} de {n} itens...")

        df1["Mais Similar"], df1["Diferenças"] = zip(*resultados)

        output = BytesIO()
        df1.to_excel(output, index=False)
        output.seek(0)

        wb = load_workbook(output)
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

        output_final = BytesIO()
        wb.save(output_final)
        output_final.seek(0)

        st.success("Conversão concluída!")
        st.download_button(
            label="Baixar planilha com diferenças",
            data=output_final,
            file_name="Diferenças.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
