import streamlit as st
import pandas as pd

# Importa as funções dinâmicas do excel_cleaner.py
from testecalibracaoschulz import process_and_validate_sheet, gerar_pdf_bytes

st.set_page_config(page_title="Validador de Certificados", layout="wide")

st.title("📐 Gerador e Validador de Certificados de Calibração")
st.write("Insira qualquer planilha de calibração para identificar automaticamente os metadados, calcular erros e limites máximo/mínimo, e gerar o PDF oficial.")

uploaded_file = st.file_uploader("Selecione a planilha (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Processa a planilha limpando rodapés espúrios
        info_dict, df_medicoes = process_and_validate_sheet(uploaded_file, target_sheet_name="Registro")

        status_geral = info_dict['status_geral']

        # Destaque do Parecer do Equipamento
        if status_geral == "APROVADO":
            st.success(f"### STATUS GERAL: INSTRUMENTO APROVADO ✅")
        else:
            st.error(f"### STATUS GERAL: INSTRUMENTO REPROVADO ❌")

        # Exibição dos Metadados Extraídos
        st.markdown("### 📋 Metadados Identificados")
        col1, col2 = st.columns(2)
        with col1:
            st.write(f"**Cliente:** {info_dict['cliente']}")
            st.write(f"**Técnico Executante:** {info_dict['tecnico']}")
            st.write(f"**Data da Calibração:** {info_dict['data']}")
        with col2:
            st.write(f"**Código do Instrumento:** {info_dict['codigo']}")
            st.write(f"**Denominação:** {info_dict['denominacao']}")

        # Tabela Detalhada no Streamlit
        st.markdown("### 📊 Validação Detalhada (Limites Mín/Máx, Erros e Status)")

        df_display = df_medicoes[['Parâmetro', 'Nominal', 'Limite Mín', 'Limite Máx', 'Valor Obtido', 'Erro', 'Status_Txt']].copy()
        df_display.columns = ['Parâmetro', 'Nominal (mm)', 'Limite Mínimo', 'Limite Máximo', 'Valor Obtido (mm)', 'Erro Encontrado', 'Parecer']

        st.dataframe(df_display, use_container_width=True)

        # GERAÇÃO E BOTÃO DE DOWNLOAD DO PDF
        pdf_bytes = gerar_pdf_bytes(info_dict, df_medicoes)

        st.markdown("---")
        st.download_button(
            label="📥 Baixar Certificado Oficial em PDF",
            data=pdf_bytes,
            file_name=f"Certificado_{info_dict['codigo']}_{status_geral}.pdf",
            mime="application/pdf"
        )

    except Exception as e:
        st.error(f"Erro ao processar a planilha: {e}")