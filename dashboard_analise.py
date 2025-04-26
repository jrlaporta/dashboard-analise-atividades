import pandas as pd
import plotly.express as px
import streamlit as st
import os
from pathlib import Path
import numpy as np
from datetime import datetime
import unicodedata
import openpyxl

if 'EQUIPAMENTO' in df_final_geral.columns and 'Categorização de Resolução 1' in df_final_geral.columns:
    st.subheader("Equipamentos Mais Frequentes")
    contagem_equipamento = df_final_geral['EQUIPAMENTO'].value_counts().reset_index()
    contagem_equipamento.columns = ['Equipamento', 'Total Atividades']
    # Mostrar apenas os top N equipamentos, por exemplo, top 10
    top_n = 10
    contagem_equipamento_top = contagem_equipamento.head(top_n)

    fig5 = px.bar(contagem_equipamento_top, x='Equipamento', y='Total Atividades',
                  text='Total Atividades',
                  title=f'Top {top_n} Equipamentos Mais Frequentes')
    fig5.update_layout(xaxis_title="Equipamento", yaxis_title="Número de Atividades")
    fig5.update_traces(textposition='outside')
    st.plotly_chart(fig5, use_container_width=True)

    st.subheader(f"Resolução Causa 1 para os Top {top_n} Equipamentos")
    # Para cada top equipamento, mostrar as causas mais comuns
    for equipamento in contagem_equipamento_top['Equipamento']:
        df_equipamento = df_final_geral[df_final_geral['EQUIPAMENTO'] == equipamento].copy()
        if not df_equipamento.empty:
            contagem_causas = df_equipamento['Categorização de Resolução 1'].value_counts().reset_index()
            contagem_causas.columns = ['Resolução Causa 1', 'Total']
            # Mostrar top N causas para este equipamento
            top_causas_n = 5 # Pode ajustar
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