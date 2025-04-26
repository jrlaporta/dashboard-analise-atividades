import pandas as pd
import plotly.express as px
import streamlit as st
import os
from pathlib import Path
import numpy as np
from datetime import datetime
import unicodedata
import openpyxl

# Configuração da página
st.set_page_config(page_title="Análise de atividades", layout="wide")

# Leitura direta do arquivo Excel
try:
    df = pd.read_excel("Pordutividade_Geral.xlsx")  # Nome do seu arquivo Excel
    df_final_geral = df.copy()

    # Filtros de segmentação
    # Filtro de Técnico
    tecnicos = sorted(df_final_geral['TÉCNICO'].dropna().unique())
    tecnico_selecionado = st.sidebar.multiselect('Filtrar por Técnico', tecnicos, default=tecnicos)

    # Filtro de Mês/Ano
    if pd.api.types.is_datetime64_any_dtype(df_final_geral['Data']):
        df_final_geral['Mes_Ano'] = df_final_geral['Data'].dt.strftime('%m/%Y')
    else:
        df_final_geral['Mes_Ano'] = pd.to_datetime(df_final_geral['Data'], errors='coerce').dt.strftime('%m/%Y')
    meses_ano = sorted(df_final_geral['Mes_Ano'].dropna().unique())
    mes_ano_selecionado = st.sidebar.multiselect('Filtrar por Mês/Ano', meses_ano, default=meses_ano)

    # Aplicar filtros
    df_final_geral = df_final_geral[df_final_geral['TÉCNICO'].isin(tecnico_selecionado)]
    df_final_geral = df_final_geral[df_final_geral['Mes_Ano'].isin(mes_ano_selecionado)]

    # Título principal
    st.title('Análise de atividades')

    # Configurações padrão para todos os gráficos
    ROTULO_TECNICO = 24    # Tamanho dos rótulos para gráficos de técnico
    ROTULO_EQUIP = 16      # Tamanho dos rótulos para gráficos de equipamento
    ESCALA_TAMANHO = 8     # Tamanho das escalas
    TITULO_TAMANHO = 24    # Tamanho dos títulos
    LEGENDA_TAMANHO = 16   # Tamanho das legendas
    ALTURA_GRAFICO = 600   # Altura dos gráficos

    # --- INÍCIO: Análises relacionadas a TÉCNICO ---
    # 1. Produtividade total de cada técnico
    if 'TÉCNICO' in df_final_geral.columns:
        st.subheader("Produtividade Total por Técnico")
        produtividade_tecnico = df_final_geral['TÉCNICO'].value_counts().reset_index()
        produtividade_tecnico.columns = ['Técnico', 'Total Atividades']
        produtividade_tecnico['%'] = 100 * produtividade_tecnico['Total Atividades'] / produtividade_tecnico['Total Atividades'].sum()
        fig_prod = px.bar(produtividade_tecnico, x='Técnico', y='Total Atividades', 
                         text=produtividade_tecnico.apply(lambda row: f"{row['Total Atividades']} ({row['%']:.1f}%)", axis=1), 
                         title='Produtividade Total por Técnico')
        fig_prod.update_traces(textposition='outside', textfont=dict(size=ROTULO_TECNICO, family="Arial Black"))  # Negrito
        fig_prod.update_layout(
            xaxis_title="Técnico", 
            yaxis_title="Total de Atividades",
            xaxis=dict(tickfont=dict(size=ESCALA_TAMANHO)),
            yaxis=dict(tickfont=dict(size=ESCALA_TAMANHO)),
            title_font=dict(size=TITULO_TAMANHO),
            height=ALTURA_GRAFICO
        )
        st.plotly_chart(fig_prod, use_container_width=True)

    # 2. SEM SINAL NO PRAZO por técnico
    if 'TÉCNICO' in df_final_geral.columns and 'SEM SINAL NO PRAZO ' in df_final_geral.columns:
        st.subheader('SEM SINAL NO PRAZO por Técnico')
        sem_sinal = df_final_geral[['TÉCNICO', 'SEM SINAL NO PRAZO ']].dropna()
        sem_sinal = sem_sinal[sem_sinal['SEM SINAL NO PRAZO '].str.strip() != '']
        sem_sinal_count = sem_sinal.groupby(['TÉCNICO', 'SEM SINAL NO PRAZO ']).size().reset_index(name='Total')
        total_por_tecnico = sem_sinal.groupby('TÉCNICO').size().reset_index(name='Total Técnico')
        sem_sinal_count = sem_sinal_count.merge(total_por_tecnico, on='TÉCNICO')
        sem_sinal_count['%'] = 100 * sem_sinal_count['Total'] / sem_sinal_count['Total Técnico']
        fig_sem_sinal = px.bar(sem_sinal_count, x='TÉCNICO', y='Total', color='SEM SINAL NO PRAZO ',
                               text=sem_sinal_count.apply(lambda row: f"{row['Total']} ({row['%']:.1f}%)", axis=1),
                               barmode='group',
                               title='SEM SINAL NO PRAZO por Técnico')
        fig_sem_sinal.update_traces(textposition='outside', textfont=dict(size=ROTULO_TECNICO, family="Arial Black"))  # Negrito
        fig_sem_sinal.update_layout(
            xaxis_title="Técnico", 
            yaxis_title="Total",
            legend_title_text='SEM SINAL NO PRAZO',
            xaxis=dict(tickfont=dict(size=ESCALA_TAMANHO)),
            yaxis=dict(tickfont=dict(size=ESCALA_TAMANHO)),
            legend=dict(font=dict(size=LEGENDA_TAMANHO)),
            title_font=dict(size=TITULO_TAMANHO),
            height=ALTURA_GRAFICO
        )
        st.plotly_chart(fig_sem_sinal, use_container_width=True)

    # 3. DEGRADAÇÃO NO PRAZO por técnico
    if 'TÉCNICO' in df_final_geral.columns and 'DEGRADAÇÃO NO PRAZO ' in df_final_geral.columns:
        st.subheader('DEGRADAÇÃO NO PRAZO por Técnico')
        degradacao = df_final_geral[['TÉCNICO', 'DEGRADAÇÃO NO PRAZO ']].dropna()
        degradacao = degradacao[degradacao['DEGRADAÇÃO NO PRAZO '].str.strip() != '']
        degradacao_count = degradacao.groupby(['TÉCNICO', 'DEGRADAÇÃO NO PRAZO ']).size().reset_index(name='Total')
        total_por_tecnico_deg = degradacao.groupby('TÉCNICO').size().reset_index(name='Total Técnico')
        degradacao_count = degradacao_count.merge(total_por_tecnico_deg, on='TÉCNICO')
        degradacao_count['%'] = 100 * degradacao_count['Total'] / degradacao_count['Total Técnico']
        fig_degradacao = px.bar(degradacao_count, x='TÉCNICO', y='Total', color='DEGRADAÇÃO NO PRAZO ',
                                text=degradacao_count.apply(lambda row: f"{row['Total']} ({row['%']:.1f}%)", axis=1),
                                barmode='group',
                                title='DEGRADAÇÃO NO PRAZO por Técnico')
        fig_degradacao.update_traces(textposition='outside', textfont=dict(size=ROTULO_TECNICO, family="Arial Black"))  # Negrito
        fig_degradacao.update_layout(
            xaxis_title="Técnico", 
            yaxis_title="Total",
            legend_title_text='DEGRADAÇÃO NO PRAZO',
            xaxis=dict(tickfont=dict(size=ESCALA_TAMANHO)),
            yaxis=dict(tickfont=dict(size=ESCALA_TAMANHO)),
            legend=dict(font=dict(size=LEGENDA_TAMANHO)),
            title_font=dict(size=TITULO_TAMANHO),
            height=ALTURA_GRAFICO
        )
        st.plotly_chart(fig_degradacao, use_container_width=True)
    # --- FIM: Análises relacionadas a TÉCNICO ---

    # --- INÍCIO: Análises relacionadas a EQUIPAMENTO ---
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
        fig5.update_traces(textposition='outside', textfont=dict(size=ROTULO_EQUIP))  # Tamanho 16 normal
        fig5.update_layout(
            xaxis_title="Equipamento", 
            yaxis_title="Número de Atividades",
            xaxis=dict(tickfont=dict(size=ESCALA_TAMANHO)),
            yaxis=dict(tickfont=dict(size=ESCALA_TAMANHO)),
            title_font=dict(size=TITULO_TAMANHO),
            height=ALTURA_GRAFICO
        )
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
                    fig6.update_traces(textposition='outside', textfont=dict(size=ROTULO_EQUIP))  # Tamanho 16 normal
                    fig6.update_layout(
                        xaxis_title="Resolução Causa 1", 
                        yaxis_title="Número de Ocorrências",
                        xaxis=dict(tickfont=dict(size=ESCALA_TAMANHO)),
                        yaxis=dict(tickfont=dict(size=ESCALA_TAMANHO)),
                        title_font=dict(size=TITULO_TAMANHO),
                        height=ALTURA_GRAFICO
                    )
                    st.plotly_chart(fig6, use_container_width=True)
                else:
                    st.write(f"Não há informações de 'Resolução Causa 1' para '{equipamento}' com os filtros atuais.")
    else:
        st.warning("O arquivo não possui as colunas necessárias: 'Equipamento' e 'Categorização de Resolução 1'.")
    # --- FIM: Análises relacionadas a EQUIPAMENTO ---
except Exception as e:
    st.error(f"Erro ao carregar os dados: {e}")
    st.stop()