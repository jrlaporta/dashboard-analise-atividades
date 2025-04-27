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
    df = pd.read_excel("Pordutividade_Geral.xlsx")  #  arquivo Excel
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
        
        # Calcular totais por técnico
        total_por_tecnico = df_final_geral['TÉCNICO'].value_counts().reset_index()
        total_por_tecnico.columns = ['TÉCNICO', 'Total Atividades']
        
        # Gráfico para respostas "Sim"
        sem_sinal_sim = sem_sinal[sem_sinal['SEM SINAL NO PRAZO '] == 'Sim']
        sem_sinal_sim_count = sem_sinal_sim.groupby('TÉCNICO').size().reset_index(name='Total Sim')
        sem_sinal_sim_count = sem_sinal_sim_count.merge(total_por_tecnico, on='TÉCNICO')
        sem_sinal_sim_count['% Sim'] = 100 * sem_sinal_sim_count['Total Sim'] / sem_sinal_sim_count['Total Atividades']
        
        fig_sem_sinal_sim = px.bar(sem_sinal_sim_count, x='TÉCNICO', y='Total Atividades',
                                  text=sem_sinal_sim_count.apply(lambda row: f"{row['Total Atividades']} ({row['% Sim']:.1f}% Sim)", axis=1),
                                  title='SEM SINAL NO PRAZO - Total de Atividades com % de "Sim"')
        fig_sem_sinal_sim.update_traces(textposition='outside', textfont=dict(size=ROTULO_TECNICO, family="Arial Black"))
        fig_sem_sinal_sim.update_layout(
            xaxis_title="Técnico", 
            yaxis_title="Total de Atividades",
            xaxis=dict(tickfont=dict(size=ESCALA_TAMANHO)),
            yaxis=dict(tickfont=dict(size=ESCALA_TAMANHO)),
            title_font=dict(size=TITULO_TAMANHO),
            height=ALTURA_GRAFICO
        )
        st.plotly_chart(fig_sem_sinal_sim, use_container_width=True)
        
        # Gráfico para respostas "Não"
        sem_sinal_nao = sem_sinal[sem_sinal['SEM SINAL NO PRAZO '] == 'Não']
        sem_sinal_nao_count = sem_sinal_nao.groupby('TÉCNICO').size().reset_index(name='Total Não')
        sem_sinal_nao_count = sem_sinal_nao_count.merge(total_por_tecnico, on='TÉCNICO')
        sem_sinal_nao_count['% Não'] = 100 * sem_sinal_nao_count['Total Não'] / sem_sinal_nao_count['Total Atividades']
        
        fig_sem_sinal_nao = px.bar(sem_sinal_nao_count, x='TÉCNICO', y='Total Atividades',
                                  text=sem_sinal_nao_count.apply(lambda row: f"{row['Total Atividades']} ({row['% Não']:.1f}% Não)", axis=1),
                                  title='SEM SINAL NO PRAZO - Total de Atividades com % de "Não"')
        fig_sem_sinal_nao.update_traces(textposition='outside', textfont=dict(size=ROTULO_TECNICO, family="Arial Black"))
        fig_sem_sinal_nao.update_layout(
            xaxis_title="Técnico", 
            yaxis_title="Total de Atividades",
            xaxis=dict(tickfont=dict(size=ESCALA_TAMANHO)),
            yaxis=dict(tickfont=dict(size=ESCALA_TAMANHO)),
            title_font=dict(size=TITULO_TAMANHO),
            height=ALTURA_GRAFICO
        )
        st.plotly_chart(fig_sem_sinal_nao, use_container_width=True)

    # 3. DEGRADAÇÃO NO PRAZO por técnico
    if 'TÉCNICO' in df_final_geral.columns and 'DEGRADAÇÃO NO PRAZO ' in df_final_geral.columns:
        st.subheader('DEGRADAÇÃO NO PRAZO por Técnico')
        degradacao = df_final_geral[['TÉCNICO', 'DEGRADAÇÃO NO PRAZO ']].dropna()
        degradacao = degradacao[degradacao['DEGRADAÇÃO NO PRAZO '].str.strip() != '']
        
        # Calcular totais por técnico
        total_por_tecnico = df_final_geral['TÉCNICO'].value_counts().reset_index()
        total_por_tecnico.columns = ['TÉCNICO', 'Total Atividades']
        
        # Gráfico para respostas "Sim"
        degradacao_sim = degradacao[degradacao['DEGRADAÇÃO NO PRAZO '] == 'Sim']
        degradacao_sim_count = degradacao_sim.groupby('TÉCNICO').size().reset_index(name='Total Sim')
        degradacao_sim_count = degradacao_sim_count.merge(total_por_tecnico, on='TÉCNICO')
        degradacao_sim_count['% Sim'] = 100 * degradacao_sim_count['Total Sim'] / degradacao_sim_count['Total Atividades']
        
        fig_degradacao_sim = px.bar(degradacao_sim_count, x='TÉCNICO', y='Total Atividades',
                                   text=degradacao_sim_count.apply(lambda row: f"{row['Total Atividades']} ({row['% Sim']:.1f}% Sim)", axis=1),
                                   title='DEGRADAÇÃO NO PRAZO - Total de Atividades com % de "Sim"')
        fig_degradacao_sim.update_traces(textposition='outside', textfont=dict(size=ROTULO_TECNICO, family="Arial Black"))
        fig_degradacao_sim.update_layout(
            xaxis_title="Técnico", 
            yaxis_title="Total de Atividades",
            xaxis=dict(tickfont=dict(size=ESCALA_TAMANHO)),
            yaxis=dict(tickfont=dict(size=ESCALA_TAMANHO)),
            title_font=dict(size=TITULO_TAMANHO),
            height=ALTURA_GRAFICO
        )
        st.plotly_chart(fig_degradacao_sim, use_container_width=True)
        
        # Gráfico para respostas "Não"
        degradacao_nao = degradacao[degradacao['DEGRADAÇÃO NO PRAZO '] == 'Não']
        degradacao_nao_count = degradacao_nao.groupby('TÉCNICO').size().reset_index(name='Total Não')
        degradacao_nao_count = degradacao_nao_count.merge(total_por_tecnico, on='TÉCNICO')
        degradacao_nao_count['% Não'] = 100 * degradacao_nao_count['Total Não'] / degradacao_nao_count['Total Atividades']
        
        fig_degradacao_nao = px.bar(degradacao_nao_count, x='TÉCNICO', y='Total Atividades',
                                   text=degradacao_nao_count.apply(lambda row: f"{row['Total Atividades']} ({row['% Não']:.1f}% Não)", axis=1),
                                   title='DEGRADAÇÃO NO PRAZO - Total de Atividades com % de "Não"')
        fig_degradacao_nao.update_traces(textposition='outside', textfont=dict(size=ROTULO_TECNICO, family="Arial Black"))
        fig_degradacao_nao.update_layout(
            xaxis_title="Técnico", 
            yaxis_title="Total de Atividades",
            xaxis=dict(tickfont=dict(size=ESCALA_TAMANHO)),
            yaxis=dict(tickfont=dict(size=ESCALA_TAMANHO)),
            title_font=dict(size=TITULO_TAMANHO),
            height=ALTURA_GRAFICO
        )
        st.plotly_chart(fig_degradacao_nao, use_container_width=True)
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