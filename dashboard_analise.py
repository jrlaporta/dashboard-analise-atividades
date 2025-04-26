import pandas as pd
import plotly.express as px
import streamlit as st
import os
from pathlib import Path
import numpy as np
from datetime import datetime
import unicodedata
import openpyxl

# Upload do arquivo Excel
uploaded_file = st.sidebar.file_uploader("Escolha um arquivo Excel", type="xlsx")
if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        st.write("Colunas do arquivo:", df.columns.tolist())
        st.dataframe(df.head())
        df_final_geral = df.copy()  # ou seu processamento customizado

        # Análise de Equipamentos e Causas
        if 'Equipamento' in df_final_geral.columns and 'Categorização de Resolução 1' in df_final_geral.columns:
            st.subheader("Equipamentos Mais Frequentes")
            contagem_equipamento = df_final_geral['Equipamento'].value_counts().reset_index()
            contagem_equipamento.columns = ['Equipamento', 'Total Atividades']
            top_n = 10
            contagem_equipamento_top = contagem_equipamento.head(top_n)

            fig5 = px.bar(contagem_equipamento_top, x='Equipamento', y='Total Atividades',
                          text='Total Atividades',
                          title=f'Top {top_n} Equipamentos Mais Frequentes')
            fig5.update_layout(xaxis_title="Equipamento", yaxis_title="Número de Atividades")
            fig5.update_traces(textposition='outside')
            st.plotly_chart(fig5, use_container_width=True)

            st.subheader(f"Resolução Causa 1 para os Top {top_n} Equipamentos")
            for equipamento in contagem_equipamento_top['Equipamento']:
                df_equipamento = df_final_geral[df_final_geral['Equipamento'] == equipamento].copy()
                if not df_equipamento.empty:
                    contagem_causas = df_equipamento['Categorização de Resolução 1'].value_counts().reset_index()
                    contagem_causas.columns = ['Resolução Causa 1', 'Total']
                    top_causas_n = 5
                    contagem_causas_top = contagem_causas.head(top_causas_n)
                    if not contagem_causas_top.empty:
                        st.write(f"**Causas mais comuns para '{equipamento}':**")
                        fig6 = px.bar(contagem_causas_top, x='Resolução Causa 1', y='Total',
                                      text='Total',
                                      title=f'Top {top_causas_n} Resoluções para {equipamento}')
                        fig6.update_layout(xaxis_title="Resolução Causa 1", yaxis_title="Número de Ocorrências")
                        fig6.update_traces(textposition='outside')
                        st.plotly_chart(fig6, use_container_width=True)
                    else:
                        st.write(f"Não há informações de 'Resolução Causa 1' para '{equipamento}' com os filtros atuais.")
        else:
            st.warning("O arquivo não possui as colunas necessárias: 'Equipamento' e 'Categorização de Resolução 1'.")
    except Exception as e:
        st.error(f"Erro ao ler o arquivo Excel: {e}")
        st.stop()
else:
    st.warning("Por favor, faça upload de um arquivo Excel para visualizar o dashboard.")