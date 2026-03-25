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
        if len(df_estoque.columns) >= 7:
            df_estoque.columns = ['Tipo_Produto', 'Funcao', 'Tecnico', 'Produto', 'Quantidade_Total', 'Unidade_Medida', 'Tipo_Utilizacao']
        else:
            st.error("O arquivo de estoque não possui o formato esperado. Verifique as colunas.")
            st.stop()
        
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
        df_movimentacao = df_movimentacao.dropna(subset=['Data_OS'])
        
        st.sidebar.success("✅ Arquivos carregados com sucesso!")
        
        # ==================== PAINEL 1: MOVIMENTAÇÕES ====================
        st.header("📦 Relatório de Movimentações por Ordem de Serviço")
        
        # Filtros interativos
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            pop_options = sorted(df_movimentacao['POP'].dropna().unique())
            pop_filter = st.multiselect("POP", options=pop_options, default=pop_options, key="pop_filter_main")
        
        with col2:
            tipo_os_options = sorted(df_movimentacao['Tipo_OS'].dropna().unique())
            tipo_os_filter = st.multiselect("Tipo de OS", options=tipo_os_options, default=tipo_os_options, key="tipo_os_filter_main")
        
        with col3:
            tecnico_options = sorted(df_movimentacao['Tecnico_Fechamento'].dropna().unique())
            tecnico_filter = st.multiselect("Técnico", options=tecnico_options, default=tecnico_options, key="tecnico_filter_main")
        
        with col4:
            if len(df_movimentacao) > 0:
                min_date = df_movimentacao['Data_OS'].min().date()
                max_date = df_movimentacao['Data_OS'].max().date()
                data_range = st.date_input("Período", value=(min_date, max_date), key="data_range_main")
        
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
        
        # ==================== RELATÓRIO COMPLETO ====================
        st.subheader("📋 Relatório Completo de Movimentações")
        
        tab1, tab2, tab3, tab4 = st.tabs(["📊 Visão Geral", "📈 Por Produto", "👤 Por Técnico", "🏢 Por POP"])
        
        with tab1:
            st.write(f"**Total de registros:** {len(df_mov_filtered)}")
            if len(df_mov_filtered) > 0:
                st.write(f"**Período analisado:** {df_mov_filtered['Data_OS'].min().strftime('%d/%m/%Y')} a {df_mov_filtered['Data_OS'].max().strftime('%d/%m/%Y')}")
            
            col_met1, col_met2, col_met3, col_met4 = st.columns(4)
            with col_met1:
                st.metric("Quantidade Total", f"{df_mov_filtered['Quantidade'].sum():,.0f}")
            with col_met2:
                st.metric("Produtos Distintos", df_mov_filtered['Produto'].nunique())
            with col_met3:
                st.metric("POPs Distintos", df_mov_filtered['POP'].nunique())
            with col_met4:
                st.metric("Técnicos", df_mov_filtered['Tecnico_Fechamento'].nunique())
            
            st.dataframe(df_mov_filtered.sort_values('Data_OS', ascending=False), use_container_width=True, hide_index=True)
        
        with tab2:
            produto_options = sorted(df_mov_filtered['Produto'].dropna().unique())
            produto_filter_tab = st.multiselect("Selecione os produtos", options=produto_options, default=produto_options, key="produto_filter_tab")
            
            df_produto_filtrado = df_mov_filtered if not produto_filter_tab else df_mov_filtered[df_mov_filtered['Produto'].isin(produto_filter_tab)]
            
            df_produto_grouped = df_produto_filtrado.groupby('Produto', as_index=False).agg({
                'Quantidade': 'sum',
                'Numero_OS': 'count',
                'Tecnico_Fechamento': lambda x: x.nunique(),
                'POP': lambda x: x.nunique()
            }).rename(columns={'Numero_OS': 'Total_Movimentacoes', 'Tecnico_Fechamento': 'Qtd_Tecnicos', 'POP': 'Qtd_POPs'})
            df_produto_grouped = df_produto_grouped.sort_values('Quantidade', ascending=False)
            
            st.dataframe(df_produto_grouped, use_container_width=True, hide_index=True)
            
            fig_produto = px.bar(df_produto_grouped.head(15), x='Produto', y='Quantidade', title='Top 15 Produtos')
            fig_produto.update_layout(xaxis_tickangle=-45)
            st.plotly_chart(fig_produto, use_container_width=True)
        
        with tab3:
            tecnico_options_tab = sorted(df_mov_filtered['Tecnico_Fechamento'].dropna().unique())
            tecnico_filter_tab = st.multiselect("Selecione os técnicos", options=tecnico_options_tab, default=tecnico_options_tab, key="tecnico_filter_tab")
            
            df_tecnico_filtrado = df_mov_filtered if not tecnico_filter_tab else df_mov_filtered[df_mov_filtered['Tecnico_Fechamento'].isin(tecnico_filter_tab)]
            
            df_tecnico_grouped = df_tecnico_filtrado.groupby('Tecnico_Fechamento', as_index=False).agg({
                'Quantidade': 'sum',
                'Numero_OS': 'count',
                'Produto': 'nunique',
                'POP': 'nunique'
            }).rename(columns={'Numero_OS': 'Total_Movimentacoes', 'Produto': 'Qtd_Produtos', 'POP': 'Qtd_POPs'})
            df_tecnico_grouped = df_tecnico_grouped.sort_values('Quantidade', ascending=False)
            
            st.dataframe(df_tecnico_grouped, use_container_width=True, hide_index=True)
            
            fig_tecnico = px.bar(df_tecnico_grouped.head(15), x='Tecnico_Fechamento', y='Quantidade', title='Top 15 Técnicos')
            fig_tecnico.update_layout(xaxis_tickangle=-45)
            st.plotly_chart(fig_tecnico, use_container_width=True)
        
        with tab4:
            pop_options_tab = sorted(df_mov_filtered['POP'].dropna().unique())
            pop_filter_tab = st.multiselect("Selecione os POPs", options=pop_options_tab, default=pop_options_tab, key="pop_filter_tab")
            
            df_pop_filtrado = df_mov_filtered if not pop_filter_tab else df_mov_filtered[df_mov_filtered['POP'].isin(pop_filter_tab)]
            
            df_pop_grouped = df_pop_filtrado.groupby('POP', as_index=False).agg({
                'Quantidade': 'sum',
                'Numero_OS': 'count',
                'Produto': 'nunique',
                'Tecnico_Fechamento': 'nunique'
            }).rename(columns={'Numero_OS': 'Total_Movimentacoes', 'Produto': 'Qtd_Produtos', 'Tecnico_Fechamento': 'Qtd_Tecnicos'})
            df_pop_grouped = df_pop_grouped.sort_values('Quantidade', ascending=False)
            
            st.dataframe(df_pop_grouped, use_container_width=True, hide_index=True)
            
            fig_pop = px.bar(df_pop_grouped.head(15), x='POP', y='Quantidade', title='Top 15 POPs')
            fig_pop.update_layout(xaxis_tickangle=-45)
            st.plotly_chart(fig_pop, use_container_width=True)
        
        st.markdown("---")
        
        # ==================== TABELA DE SOMA POR PRODUTO ====================
        st.subheader("📊 Soma de Quantidade por Produto")
        
        df_soma_produto = df_mov_filtered.groupby('Produto', as_index=False)['Quantidade'].sum().sort_values('Quantidade', ascending=False)
        
        col_table1, col_table2 = st.columns([2, 1])
        with col_table1:
            st.dataframe(df_soma_produto, use_container_width=True, hide_index=True)
        with col_table2:
            st.metric("Total de Produtos", len(df_soma_produto))
            st.metric("Quantidade Total", f"{df_soma_produto['Quantidade'].sum():,.0f}")
        
        if len(df_soma_produto) > 0:
            fig_soma = px.bar(df_soma_produto.head(15), x='Produto', y='Quantidade', title='Top 15 Produtos')
            fig_soma.update_layout(xaxis_tickangle=-45, height=500)
            st.plotly_chart(fig_soma, use_container_width=True)
        
        st.markdown("---")
        
        # ==================== MOVIMENTAÇÕES AGRUPADAS ====================
        st.subheader("📋 Movimentações Agrupadas por POP, Tipo OS e Técnico")
        
        col_f1, col_f2, col_f3, col_f4 = st.columns(4)
        
        with col_f1:
            produto_agrupado_options = sorted(df_mov_filtered['Produto'].dropna().unique())
            produto_agrupado_filter = st.multiselect("🔍 **Foco: Produto**", options=produto_agrupado_options, default=produto_agrupado_options, key="produto_agrupado_filter")
        
        with col_f2:
            pop_agrupado_options = sorted(df_mov_filtered['POP'].dropna().unique())
            pop_agrupado_filter = st.multiselect("POP", options=pop_agrupado_options, default=pop_agrupado_options, key="pop_agrupado_filter")
        
        with col_f3:
            tipo_os_agrupado_options = sorted(df_mov_filtered['Tipo_OS'].dropna().unique())
            tipo_os_agrupado_filter = st.multiselect("Tipo de OS", options=tipo_os_agrupado_options, default=tipo_os_agrupado_options, key="tipo_os_agrupado_filter")
        
        with col_f4:
            tecnico_agrupado_options = sorted(df_mov_filtered['Tecnico_Fechamento'].dropna().unique())
            tecnico_agrupado_filter = st.multiselect("Técnico", options=tecnico_agrupado_options, default=tecnico_agrupado_options, key="tecnico_agrupado_filter")
        
        df_mov_agrupado = df_mov_filtered.copy()
        if produto_agrupado_filter:
            df_mov_agrupado = df_mov_agrupado[df_mov_agrupado['Produto'].isin(produto_agrupado_filter)]
        if pop_agrupado_filter:
            df_mov_agrupado = df_mov_agrupado[df_mov_agrupado['POP'].isin(pop_agrupado_filter)]
        if tipo_os_agrupado_filter:
            df_mov_agrupado = df_mov_agrupado[df_mov_agrupado['Tipo_OS'].isin(tipo_os_agrupado_filter)]
        if tecnico_agrupado_filter:
            df_mov_agrupado = df_mov_agrupado[df_mov_agrupado['Tecnico_Fechamento'].isin(tecnico_agrupado_filter)]
        
        if len(df_mov_agrupado) > 0:
            df_grouped = df_mov_agrupado.groupby(['POP', 'Tipo_OS', 'Tecnico_Fechamento', 'Produto'], as_index=False)['Quantidade'].sum()
            st.dataframe(df_grouped.sort_values(['POP', 'Tipo_OS', 'Tecnico_Fechamento']), use_container_width=True, hide_index=True)
            st.info(f"**Resumo:** {len(df_grouped)} registros | Produtos: {df_grouped['Produto'].nunique()} | Técnicos: {df_grouped['Tecnico_Fechamento'].nunique()} | POPs: {df_grouped['POP'].nunique()}")
            
            col_g1, col_g2 = st.columns(2)
            with col_g1:
                if len(produto_agrupado_filter) <= 1:
                    df_prod_agrupado = df_grouped.groupby('Produto')['Quantidade'].sum().nlargest(10).reset_index()
                else:
                    df_prod_agrupado = df_grouped.groupby('Produto')['Quantidade'].sum().reset_index()
                fig_g1 = px.bar(df_prod_agrupado, x='Produto', y='Quantidade', title='Produtos por Quantidade')
                fig_g1.update_layout(xaxis_tickangle=-45)
                st.plotly_chart(fig_g1, use_container_width=True)
            
            with col_g2:
                os_dist_agrupado = df_grouped.groupby('Tipo_OS')['Quantidade'].sum().reset_index()
                if len(os_dist_agrupado) > 0:
                    fig_g2 = px.pie(os_dist_agrupado, values='Quantidade', names='Tipo_OS', title='Distribuição por Tipo de OS')
                    st.plotly_chart(fig_g2, use_container_width=True)
            
            # ==================== TABELAS DINÂMICAS ====================
            st.subheader("📊 Tabelas Dinâmicas")
            st.info("💡 **Dica:** Utilize os filtros acima para reduzir a quantidade de dados e melhorar a visualização")
            
            tab_pivot1, tab_pivot2 = st.tabs(["📊 Produtos vs Técnicos", "📊 Produtos vs POP"])
            
            with tab_pivot1:
                st.write("**Matriz: Produtos vs Técnicos**")
                
                # Filtros para a tabela dinâmica
                col_pivot1, col_pivot2 = st.columns(2)
                with col_pivot1:
                    tecnicos_para_pivot = st.multiselect(
                        "Selecione os técnicos para visualizar na matriz",
                        options=sorted(df_grouped['Tecnico_Fechamento'].unique()),
                        default=sorted(df_grouped['Tecnico_Fechamento'].unique())[:15] if len(df_grouped['Tecnico_Fechamento'].unique()) > 15 else sorted(df_grouped['Tecnico_Fechamento'].unique()),
                        key="tecnicos_pivot"
                    )
                
                with col_pivot2:
                    top_produtos_pivot = st.number_input(
                        "Número de produtos a exibir (ordenados por quantidade)",
                        min_value=5,
                        max_value=len(df_grouped['Produto'].unique()),
                        value=min(20, len(df_grouped['Produto'].unique())),
                        key="top_produtos_pivot"
                    )
                
                if tecnicos_para_pivot:
                    df_pivot_filtered = df_grouped[df_grouped['Tecnico_Fechamento'].isin(tecnicos_para_pivot)]
                    
                    # Criar pivot
                    pivot_tecnicos = pd.pivot_table(
                        df_pivot_filtered,
                        values='Quantidade',
                        index='Produto',
                        columns='Tecnico_Fechamento',
                        aggfunc='sum',
                        fill_value=0
                    )
                    
                    # Selecionar top produtos
                    produto_total = df_pivot_filtered.groupby('Produto')['Quantidade'].sum()
                    top_produtos = produto_total.nlargest(top_produtos_pivot).index
                    pivot_tecnicos = pivot_tecnicos.loc[top_produtos]
                    
                    st.dataframe(pivot_tecnicos, use_container_width=True, height=500)
                    
                    col_r1, col_r2, col_r3 = st.columns(3)
                    with col_r1:
                        st.metric("Produtos Exibidos", len(pivot_tecnicos))
                    with col_r2:
                        st.metric("Técnicos Exibidos", len(pivot_tecnicos.columns))
                    with col_r3:
                        st.metric("Quantidade Total", f"{pivot_tecnicos.sum().sum():,.0f}")
                else:
                    st.warning("Selecione pelo menos um técnico para visualizar a matriz")
            
            with tab_pivot2:
                st.write("**Matriz: Produtos vs POP**")
                
                # Filtros para a tabela dinâmica
                col_pivot3, col_pivot4 = st.columns(2)
                with col_pivot3:
                    pops_para_pivot = st.multiselect(
                        "Selecione os POPs para visualizar na matriz",
                        options=sorted(df_grouped['POP'].unique()),
                        default=sorted(df_grouped['POP'].unique())[:15] if len(df_grouped['POP'].unique()) > 15 else sorted(df_grouped['POP'].unique()),
                        key="pops_pivot"
                    )
                
                with col_pivot4:
                    top_produtos_pivot_pop = st.number_input(
                        "Número de produtos a exibir (ordenados por quantidade)",
                        min_value=5,
                        max_value=len(df_grouped['Produto'].unique()),
                        value=min(20, len(df_grouped['Produto'].unique())),
                        key="top_produtos_pivot_pop"
                    )
                
                if pops_para_pivot:
                    df_pivot_filtered_pop = df_grouped[df_grouped['POP'].isin(pops_para_pivot)]
                    
                    # Criar pivot
                    pivot_pop = pd.pivot_table(
                        df_pivot_filtered_pop,
                        values='Quantidade',
                        index='Produto',
                        columns='POP',
                        aggfunc='sum',
                        fill_value=0
                    )
                    
                    # Selecionar top produtos
                    produto_total_pop = df_pivot_filtered_pop.groupby('Produto')['Quantidade'].sum()
                    top_produtos_pop = produto_total_pop.nlargest(top_produtos_pivot_pop).index
                    pivot_pop = pivot_pop.loc[top_produtos_pop]
                    
                    st.dataframe(pivot_pop, use_container_width=True, height=500)
                    
                    col_r1, col_r2, col_r3 = st.columns(3)
                    with col_r1:
                        st.metric("Produtos Exibidos", len(pivot_pop))
                    with col_r2:
                        st.metric("POPs Exibidos", len(pivot_pop.columns))
                    with col_r3:
                        st.metric("Quantidade Total", f"{pivot_pop.sum().sum():,.0f}")
                else:
                    st.warning("Selecione pelo menos um POP para visualizar a matriz")
        
        else:
            st.warning("Nenhum dado encontrado com os filtros selecionados.")
        
        st.markdown("---")
        
        # ==================== PREVISÃO DE RESSUPRIMENTO ====================
        st.header("🔄 Previsão de Ressuprimento Técnico")
        
        tecnicos_filtrados = tecnico_filter if tecnico_filter else df_movimentacao['Tecnico_Fechamento'].unique()
        st.info(f"📌 **Analisando técnicos:** {', '.join(tecnicos_filtrados[:5])}{'...' if len(tecnicos_filtrados) > 5 else ''}")
        
        ultimo_dia = df_movimentacao['Data_OS'].max()
        data_90_dias = ultimo_dia - timedelta(days=90)
        st.info(f"📅 **Período analisado:** {data_90_dias.strftime('%d/%m/%Y')} a {ultimo_dia.strftime('%d/%m/%Y')} (últimos 90 dias)")
        
        df_ultimos_90 = df_movimentacao[(df_movimentacao['Data_OS'] >= data_90_dias) & (df_movimentacao['Tecnico_Fechamento'].isin(tecnicos_filtrados))]
        
        if len(df_ultimos_90) > 0:
            consumo_tecnico = df_ultimos_90.groupby(['Tecnico_Fechamento', 'Produto'])['Quantidade'].sum().reset_index()
            consumo_tecnico.columns = ['Tecnico', 'Produto', 'Consumo_90_dias']
            consumo_tecnico['Media_45_dias'] = consumo_tecnico['Consumo_90_dias'] / 2
            
            estoque_tecnico = df_estoque[df_estoque['Tecnico'].isin(tecnicos_filtrados)][['Tecnico', 'Produto', 'Quantidade_Total']].copy()
            estoque_tecnico.columns = ['Tecnico', 'Produto', 'Estoque_Atual']
            
            df_previsao = pd.merge(consumo_tecnico, estoque_tecnico, on=['Tecnico', 'Produto'], how='outer').fillna(0)
            df_previsao['Diferenca'] = df_previsao['Estoque_Atual'] - df_previsao['Media_45_dias']
            df_previsao['Necessita_Reposicao'] = df_previsao['Diferenca'].apply(lambda x: '❌ Sim' if x < 0 else '✅ Não')
            df_previsao['Quantidade_Necessaria'] = (df_previsao['Media_45_dias'] - df_previsao['Estoque_Atual']).apply(lambda x: max(0, x)).round(0)
            df_previsao['Media_45_dias'] = df_previsao['Media_45_dias'].round(2)
            
            st.subheader("🔍 Filtros para Análise")
            col_filter1, col_filter2 = st.columns(2)
            
            with col_filter1:
                tecnicos_previsao = sorted(df_previsao['Tecnico'].unique())
                tecnico_previsao_filter = st.multiselect("Técnico", options=tecnicos_previsao, default=tecnicos_previsao, key="tecnico_previsao_filter")
            
            with col_filter2:
                situacao_options = ['Todos', '❌ Sim', '✅ Não']
                situacao_filter = st.multiselect("Situação", options=situacao_options, default=['Todos'], key="situacao_filter")
            
            df_previsao_filtered = df_previsao[df_previsao['Tecnico'].isin(tecnico_previsao_filter)]
            if 'Todos' not in situacao_filter:
                df_previsao_filtered = df_previsao_filtered[df_previsao_filtered['Necessita_Reposicao'].isin(situacao_filter)]
            
            st.dataframe(df_previsao_filtered[['Tecnico', 'Produto', 'Estoque_Atual', 'Consumo_90_dias', 'Media_45_dias', 'Diferenca', 'Quantidade_Necessaria', 'Necessita_Reposicao']], use_container_width=True, hide_index=True)
            
            col1_prev, col2_prev = st.columns(2)
            with col1_prev:
                df_reposicao = df_previsao_filtered[df_previsao_filtered['Diferenca'] < 0].nlargest(10, 'Quantidade_Necessaria')
                if len(df_reposicao) > 0:
                    fig4 = px.bar(df_reposicao, x='Produto', y='Quantidade_Necessaria', color='Tecnico', title='Top Produtos que Necessitam Reposição')
                    fig4.update_layout(xaxis_tickangle=-45)
                    st.plotly_chart(fig4, use_container_width=True)
            
            with col2_prev:
                situacao_dist = df_previsao_filtered['Necessita_Reposicao'].value_counts().reset_index()
                situacao_dist.columns = ['Situação', 'Quantidade']
                fig5 = px.pie(situacao_dist, values='Quantidade', names='Situação', title='Distribuição por Situação')
                st.plotly_chart(fig5, use_container_width=True)
        
        # ==================== EXPORTAR DADOS ====================
        st.markdown("---")
        st.subheader("📥 Exportar Dados")
        
        col_export1, col_export2, col_export3, col_export4 = st.columns(4)
        
        if 'df_soma_produto' in locals() and len(df_soma_produto) > 0:
            output1 = BytesIO()
            with pd.ExcelWriter(output1, engine='openpyxl') as writer:
                df_soma_produto.to_excel(writer, sheet_name='Soma_por_Produto', index=False)
            st.download_button("📊 Soma por Produto", data=output1.getvalue(), file_name="soma_por_produto.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="export1")
        
        if 'df_grouped' in locals() and len(df_grouped) > 0:
            output2 = BytesIO()
            with pd.ExcelWriter(output2, engine='openpyxl') as writer:
                df_grouped.to_excel(writer, sheet_name='Movimentacoes_Agrupadas', index=False)
            st.download_button("📊 Movimentações Agrupadas", data=output2.getvalue(), file_name="movimentacoes_agrupadas.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="export2")
        
        if 'df_previsao_filtered' in locals() and len(df_previsao_filtered) > 0:
            output3 = BytesIO()
            with pd.ExcelWriter(output3, engine='openpyxl') as writer:
                df_previsao_filtered.to_excel(writer, sheet_name='Previsao_Ressuprimento', index=False)
            st.download_button("🔄 Previsão de Ressuprimento", data=output3.getvalue(), file_name="previsao_ressuprimento.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="export3")
        
        if 'df_mov_filtered' in locals() and len(df_mov_filtered) > 0:
            output4 = BytesIO()
            with pd.ExcelWriter(output4, engine='openpyxl') as writer:
                df_mov_filtered.to_excel(writer, sheet_name='Relatorio_Completo', index=False)
            st.download_button("📋 Relatório Completo", data=output4.getvalue(), file_name="relatorio_completo.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="export4")
    
    except Exception as e:
        st.error(f"Erro ao processar os arquivos: {str(e)}")
        st.error("Verifique se os arquivos estão no formato correto e contêm as abas esperadas.")

else:
    st.info("📂 **Por favor, faça upload dos arquivos Excel para iniciar a análise.**")
    
    st.markdown("""
    ### Instruções:
    
    1. **Estoque com o Técnico** - Arquivo com colunas: Tipo de Produto, Função, Técnico, Produto, Quantidade Total, Unidade de Medida, Tipo Utilização
    
    2. **Movimentação por Ordem de Serviço** - Arquivo com colunas: Número OS, POP, Tipo de OS, Técnico de Fechamento, Data OS, Produto, Quantidade, entre outras
    
    ### Funcionalidades:
    
    - **Relatório Completo:** Visualize todas as movimentações com abas por visão geral, produto, técnico e POP
    - **Soma por Produto:** Quantidade total movimentada por produto
    - **Filtros Dinâmicos:** Filtre por POP, Tipo de OS, Técnico e período (todos preselecionados)
    - **Tabelas Dinâmicas:** Selecione quais técnicos/POPs visualizar para evitar limitação de colunas
    - **Previsão de Ressuprimento:** Análise baseada nos últimos 90 dias
    - **Exportação de Dados:** Exporte todas as análises para Excel
    """)
