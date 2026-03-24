import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO

# Configuração da página
st.set_page_config(
    page_title="Sistema de Gestão de Estoque e Movimentações",
    layout="wide"
)

st.title("📊 Sistema de Gestão de Estoque e Movimentações")
st.markdown("---")

# ==================== CARREGAMENTO DOS ARQUIVOS ====================
st.sidebar.header("📂 Upload dos Arquivos")

uploaded_estoque = st.sidebar.file_uploader(
    "Estoque com o Técnico (estoque_com_o_tecnico.xlsx)",
    type=['xlsx', 'xls']
)

uploaded_movimentacao = st.sidebar.file_uploader(
    "Movimentação por Ordem de Serviço (movimentacao_por_ordem_servico.xlsx)",
    type=['xlsx', 'xls']
)

if uploaded_estoque is not None and uploaded_movimentacao is not None:
    try:
        # Carregar planilhas COMPLETAS
        df_estoque = pd.read_excel(uploaded_estoque, sheet_name='Resultado da consulta')
        df_movimentacao = pd.read_excel(uploaded_movimentacao, sheet_name='Resultado da consulta')
        
        # Mostrar informações sobre os dados carregados
        st.sidebar.info(f"📊 Estoque: {len(df_estoque)} registros carregados")
        st.sidebar.info(f"📊 Movimentações: {len(df_movimentacao)} registros carregados")
        
        # Verificar se as colunas existem e renomear adequadamente
        # Para o arquivo de estoque
        if len(df_estoque.columns) >= 7:
            df_estoque.columns = ['Tipo_Produto', 'Funcao', 'Tecnico', 'Produto', 'Quantidade_Total', 'Unidade_Medida', 'Tipo_Utilizacao']
        else:
            st.error("O arquivo de estoque não possui o formato esperado. Verifique as colunas.")
            st.stop()
        
        # Para o arquivo de movimentação
        if len(df_movimentacao.columns) >= 13:
            df_movimentacao.columns = ['Numero_OS', 'POP', 'Tipo_OS', 'Tecnico_Fechamento', 'Data_OS', 'Data_Movimento', 
                                       'ID_Movimento', 'Natureza_Movimento', 'Destino_Movimento', 'ID_Cliente_Servico', 
                                       'Produto', 'Cidade', 'Quantidade']
        else:
            st.error("O arquivo de movimentação não possui o formato esperado. Verifique as colunas.")
            st.stop()
        
        # Converter datas
        df_movimentacao['Data_OS'] = pd.to_datetime(df_movimentacao['Data_OS'], errors='coerce')
        df_movimentacao['Data_Movimento'] = pd.to_datetime(df_movimentacao['Data_Movimento'], errors='coerce')
        
        # Remover linhas com datas inválidas
        df_movimentacao = df_movimentacao.dropna(subset=['Data_OS'])
        
        st.sidebar.success("✅ Arquivos carregados com sucesso!")
        
        # ==================== PAINEL 1: MOVIMENTAÇÕES ====================
        st.header("📦 Relatório de Movimentações por Ordem de Serviço")
        
        # Filtros interativos
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            pop_options = sorted(df_movimentacao['POP'].dropna().unique())
            pop_filter = st.multiselect(
                "POP",
                options=pop_options,
                default=pop_options if len(pop_options) <= 10 else pop_options[:10]
            )
        
        with col2:
            tipo_os_options = sorted(df_movimentacao['Tipo_OS'].dropna().unique())
            tipo_os_filter = st.multiselect(
                "Tipo de OS",
                options=tipo_os_options,
                default=tipo_os_options
            )
        
        with col3:
            tecnico_options = sorted(df_movimentacao['Tecnico_Fechamento'].dropna().unique())
            tecnico_filter = st.multiselect(
                "Técnico",
                options=tecnico_options,
                default=tecnico_options if len(tecnico_options) <= 10 else tecnico_options[:10]
            )
        
        with col4:
            if len(df_movimentacao) > 0:
                min_date = df_movimentacao['Data_OS'].min().date()
                max_date = df_movimentacao['Data_OS'].max().date()
                data_range = st.date_input(
                    "Período",
                    value=(min_date, max_date)
                )
        
        # Aplicar filtros
        df_mov_filtered = df_movimentacao.copy()
        
        if pop_filter:
            df_mov_filtered = df_mov_filtered[df_mov_filtered['POP'].isin(pop_filter)]
        
        if tipo_os_filter:
            df_mov_filtered = df_mov_filtered[df_mov_filtered['Tipo_OS'].isin(tipo_os_filter)]
        
        if tecnico_filter:
            df_mov_filtered = df_mov_filtered[df_mov_filtered['Tecnico_Fechamento'].isin(tecnico_filter)]
        
        if 'data_range' in locals() and len(data_range) == 2:
            df_mov_filtered = df_mov_filtered[
                (df_mov_filtered['Data_OS'].dt.date >= data_range[0]) &
                (df_mov_filtered['Data_OS'].dt.date <= data_range[1])
            ]
        
        # ==================== TABELA DE SOMA POR PRODUTO ====================
        st.subheader("📊 Soma de Quantidade por Produto")
        
        # Agrupar dados por produto para visualização da soma
        df_soma_produto = df_mov_filtered.groupby('Produto', as_index=False)['Quantidade'].sum()
        df_soma_produto = df_soma_produto.sort_values('Quantidade', ascending=False)
        
        # Exibir tabela de soma por produto
        col_table1, col_table2 = st.columns([2, 1])
        
        with col_table1:
            st.dataframe(
                df_soma_produto,
                use_container_width=True,
                hide_index=True
            )
        
        with col_table2:
            # Resumo rápido
            total_produtos = len(df_soma_produto)
            total_quantidade = df_soma_produto['Quantidade'].sum()
            media_por_produto = total_quantidade / total_produtos if total_produtos > 0 else 0
            
            st.metric("Total de Produtos", total_produtos)
            st.metric("Quantidade Total Movimentada", f"{total_quantidade:,.0f}")
            st.metric("Média por Produto", f"{media_por_produto:,.2f}")
        
        # Gráfico de barras da soma por produto (Top 15)
        if len(df_soma_produto) > 0:
            top_produtos = df_soma_produto.head(15)
            fig_soma = px.bar(
                top_produtos,
                x='Produto',
                y='Quantidade',
                title='Top 15 Produtos por Quantidade Movimentada',
                labels={'Produto': 'Produto', 'Quantidade': 'Quantidade'}
            )
            fig_soma.update_layout(xaxis_tickangle=-45, height=500)
            st.plotly_chart(fig_soma, use_container_width=True)
        
        st.markdown("---")
        
        # ==================== TABELA DE MOVIMENTAÇÕES AGRUPADAS ====================
        st.subheader("📋 Movimentações Agrupadas por POP, Tipo OS e Técnico")
        
        # Agrupar dados por POP, Tipo_OS, Tecnico e Produto
        if len(df_mov_filtered) > 0:
            df_grouped = df_mov_filtered.groupby(
                ['POP', 'Tipo_OS', 'Tecnico_Fechamento', 'Produto'], 
                as_index=False
            )['Quantidade'].sum()
            
            # Exibir tabela
            st.dataframe(
                df_grouped.sort_values(['POP', 'Tipo_OS', 'Tecnico_Fechamento']),
                use_container_width=True,
                hide_index=True
            )
            
            # Gráficos
            st.subheader("📊 Análises Gráficas")
            
            col1_graph, col2_graph = st.columns(2)
            
            with col1_graph:
                # Gráfico de barras - Top produtos agrupados
                top_produtos_agrupados = df_grouped.groupby('Produto')['Quantidade'].sum().nlargest(10).reset_index()
                if len(top_produtos_agrupados) > 0:
                    fig1 = px.bar(
                        top_produtos_agrupados,
                        x='Produto',
                        y='Quantidade',
                        title='Top 10 Produtos Mais Movimentados',
                        labels={'Produto': 'Produto', 'Quantidade': 'Quantidade'}
                    )
                    fig1.update_layout(xaxis_tickangle=-45)
                    st.plotly_chart(fig1, use_container_width=True)
            
            with col2_graph:
                # Gráfico de pizza - Distribuição por Tipo de OS
                os_dist = df_grouped.groupby('Tipo_OS')['Quantidade'].sum().reset_index()
                if len(os_dist) > 0:
                    fig2 = px.pie(
                        os_dist,
                        values='Quantidade',
                        names='Tipo_OS',
                        title='Distribuição por Tipo de OS'
                    )
                    st.plotly_chart(fig2, use_container_width=True)
            
            # Gráfico de linha - Evolução temporal
            df_temporal = df_mov_filtered.groupby(df_mov_filtered['Data_OS'].dt.date)['Quantidade'].sum().reset_index()
            df_temporal.columns = ['Data', 'Quantidade']
            if len(df_temporal) > 0:
                fig3 = px.line(
                    df_temporal,
                    x='Data',
                    y='Quantidade',
                    title='Evolução Temporal das Movimentações',
                    labels={'Data': 'Data', 'Quantidade': 'Quantidade Movimentada'}
                )
                st.plotly_chart(fig3, use_container_width=True)
        
        else:
            st.warning("Nenhum dado encontrado com os filtros selecionados.")
        
        st.markdown("---")
        
        # ==================== PAINEL 2: PREVISÃO DE RESSUPRIMENTO ====================
        st.header("🔄 Previsão de Ressuprimento Técnico")
        
        # Usar os técnicos filtrados nas movimentações
        tecnicos_filtrados = tecnico_filter if tecnico_filter else df_movimentacao['Tecnico_Fechamento'].unique()
        
        st.info(f"📌 **Analisando técnicos:** {', '.join(tecnicos_filtrados[:5])}{'...' if len(tecnicos_filtrados) > 5 else ''}")
        
        # Calcular último dia de movimentação
        ultimo_dia = df_movimentacao['Data_OS'].max()
        data_90_dias = ultimo_dia - timedelta(days=90)
        
        st.info(f"📅 **Período analisado:** {data_90_dias.strftime('%d/%m/%Y')} a {ultimo_dia.strftime('%d/%m/%Y')} (últimos 90 dias)")
        
        # Filtrar movimentações dos últimos 90 dias APENAS para os técnicos selecionados
        df_ultimos_90 = df_movimentacao[
            (df_movimentacao['Data_OS'] >= data_90_dias) &
            (df_movimentacao['Tecnico_Fechamento'].isin(tecnicos_filtrados))
        ]
        
        if len(df_ultimos_90) > 0:
            # Calcular consumo por técnico e produto nos últimos 90 dias
            consumo_tecnico = df_ultimos_90.groupby(['Tecnico_Fechamento', 'Produto'])['Quantidade'].sum().reset_index()
            consumo_tecnico.columns = ['Tecnico', 'Produto', 'Consumo_90_dias']
            
            # Calcular média de 45 dias (dividir por 2)
            consumo_tecnico['Media_45_dias'] = consumo_tecnico['Consumo_90_dias'] / 2
            
            # Obter estoque atual por técnico - APENAS para os técnicos filtrados
            estoque_tecnico = df_estoque[df_estoque['Tecnico'].isin(tecnicos_filtrados)]
            estoque_tecnico = estoque_tecnico[['Tecnico', 'Produto', 'Quantidade_Total']].copy()
            estoque_tecnico.columns = ['Tecnico', 'Produto', 'Estoque_Atual']
            
            # Mesclar dados de consumo e estoque
            df_previsao = pd.merge(
                consumo_tecnico,
                estoque_tecnico,
                on=['Tecnico', 'Produto'],
                how='outer'
            ).fillna(0)
            
            # Calcular necessidade de reposição
            df_previsao['Diferenca'] = df_previsao['Estoque_Atual'] - df_previsao['Media_45_dias']
            df_previsao['Necessita_Reposicao'] = df_previsao['Diferenca'].apply(lambda x: '❌ Sim (Reposição Necessária)' if x < 0 else '✅ Não (Estoque Suficiente)')
            df_previsao['Quantidade_Necessaria'] = df_previsao['Media_45_dias'] - df_previsao['Estoque_Atual']
            df_previsao['Quantidade_Necessaria'] = df_previsao['Quantidade_Necessaria'].apply(lambda x: max(0, x))
            
            # Arredondar valores
            df_previsao['Media_45_dias'] = df_previsao['Media_45_dias'].round(2)
            df_previsao['Quantidade_Necessaria'] = df_previsao['Quantidade_Necessaria'].round(0)
            
            # Filtros adicionais para previsão
            st.subheader("🔍 Filtros para Análise de Ressuprimento")
            
            col_filter1, col_filter2 = st.columns(2)
            
            with col_filter1:
                tecnicos_previsao = sorted(df_previsao['Tecnico'].unique())
                tecnico_previsao_filter = st.multiselect(
                    "Técnico (Previsão)",
                    options=tecnicos_previsao,
                    default=tecnicos_previsao
                )
            
            with col_filter2:
                situacao_options = ['Todos', '❌ Sim (Reposição Necessária)', '✅ Não (Estoque Suficiente)']
                situacao_filter = st.multiselect(
                    "Situação",
                    options=situacao_options,
                    default=['Todos']
                )
            
            # Aplicar filtros
            df_previsao_filtered = df_previsao[df_previsao['Tecnico'].isin(tecnico_previsao_filter)]
            
            if 'Todos' not in situacao_filter:
                df_previsao_filtered = df_previsao_filtered[df_previsao_filtered['Necessita_Reposicao'].isin(situacao_filter)]
            
            # Exibir tabela de previsão
            st.subheader(f"📋 Análise de Estoque por Técnico ({len(df_previsao_filtered)} registros)")
            
            # Formatar para exibição
            df_display = df_previsao_filtered.copy()
            df_display['Media_45_dias'] = df_display['Media_45_dias'].apply(lambda x: f"{x:.2f}")
            df_display['Quantidade_Necessaria'] = df_display['Quantidade_Necessaria'].apply(lambda x: f"{x:.0f}")
            df_display['Diferenca'] = df_display['Diferenca'].apply(lambda x: f"{x:.2f}")
            
            st.dataframe(
                df_display[['Tecnico', 'Produto', 'Estoque_Atual', 'Consumo_90_dias', 'Media_45_dias', 'Diferenca', 'Quantidade_Necessaria', 'Necessita_Reposicao']],
                use_container_width=True,
                hide_index=True
            )
            
            # Gráficos da previsão
            st.subheader("📊 Análise de Ressuprimento")
            
            col1_prev, col2_prev = st.columns(2)
            
            with col1_prev:
                # Gráfico de barras - Produtos que precisam de reposição
                df_reposicao = df_previsao_filtered[df_previsao_filtered['Diferenca'] < 0].copy()
                if len(df_reposicao) > 0:
                    df_reposicao = df_reposicao.nlargest(10, 'Quantidade_Necessaria')
                    fig4 = px.bar(
                        df_reposicao,
                        x='Produto',
                        y='Quantidade_Necessaria',
                        color='Tecnico',
                        title='Top Produtos que Necessitam Reposição',
                        labels={'Produto': 'Produto', 'Quantidade_Necessaria': 'Quantidade Necessária'}
                    )
                    fig4.update_layout(xaxis_tickangle=-45)
                    st.plotly_chart(fig4, use_container_width=True)
                else:
                    st.info("✅ Todos os produtos estão com estoque suficiente!")
            
            with col2_prev:
                # Gráfico de pizza - Distribuição por situação
                situacao_dist = df_previsao_filtered['Necessita_Reposicao'].value_counts().reset_index()
                situacao_dist.columns = ['Situação', 'Quantidade']
                if len(situacao_dist) > 0:
                    fig5 = px.pie(
                        situacao_dist,
                        values='Quantidade',
                        names='Situação',
                        title='Distribuição por Situação de Estoque'
                    )
                    st.plotly_chart(fig5, use_container_width=True)
            
            # Resumo estatístico
            st.subheader("📈 Resumo da Análise")
            
            col_sum1, col_sum2, col_sum3, col_sum4 = st.columns(4)
            
            with col_sum1:
                total_produtos = len(df_previsao_filtered)
                st.metric("Total de Produtos Analisados", total_produtos)
            
            with col_sum2:
                produtos_criticos = len(df_previsao_filtered[df_previsao_filtered['Diferenca'] < 0])
                st.metric("Produtos com Estoque Crítico", produtos_criticos)
            
            with col_sum3:
                valor_total_reposicao = df_previsao_filtered[df_previsao_filtered['Quantidade_Necessaria'] > 0]['Quantidade_Necessaria'].sum()
                st.metric("Total de Unidades para Reposição", f"{valor_total_reposicao:.0f}")
            
            with col_sum4:
                tecnicos_com_critico = df_previsao_filtered[df_previsao_filtered['Diferenca'] < 0]['Tecnico'].nunique()
                st.metric("Técnicos com Estoque Crítico", tecnicos_com_critico)
            
        else:
            st.warning(f"Não há movimentações nos últimos 90 dias para os técnicos selecionados.")
        
        # Botão para exportar dados
        st.markdown("---")
        st.subheader("📥 Exportar Dados")
        
        col_export1, col_export2, col_export3 = st.columns(3)
        
        with col_export1:
            if 'df_soma_produto' in locals() and len(df_soma_produto) > 0:
                output1 = BytesIO()
                with pd.ExcelWriter(output1, engine='openpyxl') as writer:
                    df_soma_produto.to_excel(writer, sheet_name='Soma_por_Produto', index=False)
                st.download_button(
                    label="📊 Download Soma por Produto",
                    data=output1.getvalue(),
                    file_name="soma_por_produto.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        
        with col_export2:
            if 'df_grouped' in locals() and len(df_grouped) > 0:
                output2 = BytesIO()
                with pd.ExcelWriter(output2, engine='openpyxl') as writer:
                    df_grouped.to_excel(writer, sheet_name='Movimentacoes_Agrupadas', index=False)
                st.download_button(
                    label="📊 Download Movimentações Agrupadas",
                    data=output2.getvalue(),
                    file_name="movimentacoes_agrupadas.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        
        with col_export3:
            if 'df_previsao_filtered' in locals() and len(df_previsao_filtered) > 0:
                output3 = BytesIO()
                with pd.ExcelWriter(output3, engine='openpyxl') as writer:
                    df_previsao_filtered.to_excel(writer, sheet_name='Previsao_Ressuprimento', index=False)
                st.download_button(
                    label="🔄 Download Previsão de Ressuprimento",
                    data=output3.getvalue(),
                    file_name="previsao_ressuprimento.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    
    except Exception as e:
        st.error(f"Erro ao processar os arquivos: {str(e)}")
        st.error("Verifique se os arquivos estão no formato correto e contêm as abas esperadas.")

else:
    st.info("📂 **Por favor, faça upload dos arquivos Excel para iniciar a análise.**")
    
    st.markdown("""
    ### Instruções:
    
    1. **Estoque com o Técnico** - Arquivo contendo:
       - Coluna A: Tipo de Produto
       - Coluna B: Função
       - Coluna C: Técnico
       - Coluna D: Produto  
       - Coluna E: Quantidade Total
       - Coluna F: Unidade de Medida
       - Coluna G: Tipo Utilização
    
    2. **Movimentação por Ordem de Serviço** - Arquivo contendo:
       - Coluna A: Número OS
       - Coluna B: POP
       - Coluna C: Tipo de OS
       - Coluna D: Técnico de Fechamento
       - Coluna E: Data OS
       - Coluna F: Data Movimento
       - Coluna G: ID Movimento
       - Coluna H: Natureza Movimento
       - Coluna I: Destino Movimento
       - Coluna J: ID Cliente Servico
       - Coluna K: Produto
       - Coluna L: Cidade
       - Coluna M: Quantidade
    
    ### Funcionalidades:
    
    - **Soma por Produto:** Visualize a quantidade total movimentada por produto com filtros interativos
    - **Filtros Dinâmicos:** Filtre por POP, Tipo de OS, Técnico e período
    - **Movimentações Agrupadas:** Visualize as movimentações agrupadas por POP, Tipo OS e Técnico
    - **Previsão de Ressuprimento:** Análise automática baseada nos técnicos filtrados, considerando os últimos 90 dias
    - **Exportação de Dados:** Exporte todas as análises para Excel
    """)
