import streamlit as st
import pandas as pd
import plotly.express as px
from utils import formatar_cpf_cnpj_brasileiro
from vencimentos_utils import (
    carregar_dados_vencimentos,
    extrair_ano_vencimento,
    gerar_excel_vencimentos_formatado,
    remover_duplicados_manter_mais_antiga,
)

# Configuração básica da página.
st.set_page_config(
    page_title="Sistema de Análise por Ano de Vencimento",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilo visual do app.
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f4e79;
        text-align: center;
        padding: 1rem 0;
        margin-bottom: 2rem;
    }
    .metric-card {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #1f4e79;
    }
    .stButton>button {
        width: 100%;
        background-color: #1f4e79;
        color: white;
        font-weight: bold;
    }
    .success-box {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 0.5rem;
        padding: 1rem;
        margin: 1rem 0;
    }
    .warning-box {
        background-color: #fff3cd;
        border: 1px solid #ffeaa7;
        border-radius: 0.5rem;
        padding: 1rem;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

# Cabeçalho principal.
st.markdown('<div class="main-header">📅 Sistema de Análise por Ano de Vencimento</div>', unsafe_allow_html=True)
st.markdown('<p style="text-align: center; color: #666; font-size: 1.1rem;">Filtre e visualize autos de infração por ano de vencimento</p>', unsafe_allow_html=True)

# Barra lateral com upload e colunas esperadas.
with st.sidebar:
    st.image("https://via.placeholder.com/200x80/1f4e79/ffffff?text=ANTT", use_container_width=True)
    st.markdown("### 📁 Upload de Arquivo")
    
    arquivo = st.file_uploader(
        "Selecione a planilha",
        type=['xlsx', 'xls', 'csv'],
        key='arquivo_vencimentos'
    )
    
    st.markdown("---")
    st.markdown("### ⚙️ Configurações")
    
    st.markdown("#### 🔑 Colunas Obrigatórias")
    # Esse campo precisa bater exatamente com o nome da coluna no arquivo.
    coluna_auto = st.text_input(
        "Nome da coluna Auto de Infração",
        value="Identificador do Débito",
        help="⚠️ OBRIGATÓRIO: Digite o nome exato da coluna que contém os Autos de Infração"
    )
    
    # A coluna de vencimento é obrigatória para montar a análise por ano.
    coluna_vencimento = st.text_input(
        "Nome da coluna Vencimento",
        value="Data do Vencimento",
        help="⚠️ OBRIGATÓRIO: Digite o nome exato da coluna que contém as datas de vencimento"
    )
    
    st.markdown("---")
    st.markdown("#### 📋 Colunas Opcionais")
    # O protocolo ajuda na exportação, mas não trava o uso do app.
    coluna_protocolo = st.text_input(
        "Nome da coluna Nº do Processo (Opcional)",
        value="Nº do Processo",
        help="Digite o nome exato da coluna que contém os números de protocolos (opcional)"
    )
    
    # Modal é opcional e entra mais na leitura do relatório/exportação.
    coluna_modal = st.text_input(
        "Nome da coluna Subtipo de Débito (Opcional)",
        value="Subtipo de Débito",
        help="Digite o nome exato da coluna que contém os modais (opcional)"
    )
    
    # CPF/CNPJ aparece na exportação quando a base trouxer esse dado.
    coluna_cpf_cnpj = st.text_input(
        "Nome da coluna CPF/CNPJ (Opcional)",
        value="CPF/CNPJ",
        help="Digite o nome exato da coluna que contém CPF/CNPJ (opcional)"
    )

# As rotinas de leitura, ano e exportação ficaram em `vencimentos_utils.py`.

# Fluxo principal do app.
if arquivo:
    st.markdown("---")
    
    # Primeiro carrega a base para validar colunas antes de montar o resto da tela.
    with st.spinner("Carregando planilha..."):
        df = carregar_dados_vencimentos(arquivo)
    
    if df is not None:
        # Mostra um preview simples para o usuário conferir se o arquivo certo foi enviado.
        st.markdown("### 📋 Preview da Planilha")
        st.dataframe(df.head(), use_container_width=True)
        st.caption(f"Total de registros: {len(df):,}")
        st.caption(f"Colunas: {', '.join(df.columns.tolist()[:10])}...")
        
        st.markdown("---")
        
        # Não segue para a análise se as colunas obrigatórias não estiverem na base.
        if coluna_auto not in df.columns:
            st.error(f"⚠️ Coluna '{coluna_auto}' não encontrada na planilha. Esta coluna é OBRIGATÓRIA!")
        elif coluna_vencimento not in df.columns:
            st.error(f"⚠️ Coluna '{coluna_vencimento}' não encontrada na planilha. Esta coluna é OBRIGATÓRIA!")
        else:
            # A análise só roda quando o usuário confirma.
            if st.button("🚀 Analisar por Ano de Vencimento", type="primary", use_container_width=True):
                with st.spinner("Analisando dados por ano de vencimento..."):
                    # Se o auto repetir, fica só a linha com vencimento mais antigo.
                    df_sem_duplicados = remover_duplicados_manter_mais_antiga(df, coluna_auto, coluna_vencimento)
                    
                    # Depois disso o app quebra a base por ano.
                    df_com_ano = extrair_ano_vencimento(df_sem_duplicados, coluna_vencimento)
                    
                    # Linha sem data válida não entra na contagem por ano.
                    df_com_ano = df_com_ano[df_com_ano['ANO_VENCIMENTO'].notna()].copy()
                    
                    # Monta a lista de anos encontrados na base.
                    anos_encontrados = sorted(df_com_ano['ANO_VENCIMENTO'].unique())
                    anos_encontrados = [int(ano) for ano in anos_encontrados if not pd.isna(ano)]
                    
                    # Guarda os DataFrames por ano para usar no dashboard e na exportação.
                    stats_por_ano = {}
                    for ano in anos_encontrados:
                        df_ano = df_com_ano[df_com_ano['ANO_VENCIMENTO'] == ano].copy()
                        stats_por_ano[ano] = {
                            'quantidade': len(df_ano),
                            'dataframe': df_ano
                        }
                    
                    # Esse número vai para a interface para o usuário entender o que foi limpo.
                    total_original = len(df)
                    total_sem_duplicados = len(df_sem_duplicados)
                    duplicados_removidos = total_original - total_sem_duplicados
                    
                    # Salva o resultado para a tela não recalcular tudo a cada interação.
                    st.session_state['df_com_ano'] = df_com_ano
                    st.session_state['anos_encontrados'] = anos_encontrados
                    st.session_state['stats_por_ano'] = stats_por_ano
                    st.session_state['coluna_auto'] = coluna_auto
                    st.session_state['coluna_vencimento'] = coluna_vencimento
                    st.session_state['coluna_cpf_cnpj'] = coluna_cpf_cnpj
                    st.session_state['coluna_modal'] = coluna_modal
                    st.session_state['coluna_protocolo'] = coluna_protocolo
                    st.session_state['duplicados_removidos'] = duplicados_removidos
                    st.session_state['total_original'] = total_original
                    
                    st.success("✅ Análise concluída com sucesso!")
                    if duplicados_removidos > 0:
                        st.info(f"ℹ️ {duplicados_removidos:,} autos duplicados foram removidos (mantida a data de vencimento mais antiga).")
                    st.rerun()

# Exibição dos resultados já calculados.
if 'df_com_ano' in st.session_state:
    df_com_ano = st.session_state['df_com_ano']
    anos_encontrados = st.session_state['anos_encontrados']
    stats_por_ano = st.session_state['stats_por_ano']
    coluna_auto = st.session_state['coluna_auto']
    coluna_vencimento = st.session_state['coluna_vencimento']
    coluna_cpf_cnpj = st.session_state.get('coluna_cpf_cnpj', '')
    coluna_modal = st.session_state.get('coluna_modal', '')
    coluna_protocolo = st.session_state.get('coluna_protocolo', '')
    
    st.markdown("---")
    st.markdown("## 📊 Dashboard - Análise por Ano de Vencimento")
    
    # Se teve duplicado, o app avisa logo no topo.
    duplicados_removidos = st.session_state.get('duplicados_removidos', 0)
    total_original = st.session_state.get('total_original', 0)
    if duplicados_removidos > 0:
        st.info(f"ℹ️ **{duplicados_removidos:,} autos duplicados foram removidos** (de {total_original:,} registros originais). Mantida a data de vencimento mais antiga para cada auto.")
    
    # Resumo rápido antes dos gráficos.
    total_autos = len(df_com_ano)
    total_anos = len(anos_encontrados)
    ano_mais_autos = max(anos_encontrados, key=lambda x: stats_por_ano[x]['quantidade']) if anos_encontrados else None
    qtd_mais_autos = stats_por_ano[ano_mais_autos]['quantidade'] if ano_mais_autos else 0
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric(
            "Total de Autos",
            f"{total_autos:,}",
            delta="Autos com vencimento válido"
        )
    
    with col2:
        st.metric(
            "Anos Encontrados",
            f"{total_anos}",
            delta="Anos diferentes"
        )
    
    with col3:
        st.metric(
            "Ano com Mais Autos",
            f"{ano_mais_autos}" if ano_mais_autos else "N/A",
            delta=f"{qtd_mais_autos:,} autos" if ano_mais_autos else ""
        )
    
    with col4:
        ano_menos_autos = min(anos_encontrados, key=lambda x: stats_por_ano[x]['quantidade']) if anos_encontrados else None
        qtd_menos_autos = stats_por_ano[ano_menos_autos]['quantidade'] if ano_menos_autos else 0
        st.metric(
            "Ano com Menos Autos",
            f"{ano_menos_autos}" if ano_menos_autos else "N/A",
            delta=f"{qtd_menos_autos:,} autos" if ano_menos_autos else ""
        )
    
    # Gráfico principal de distribuição por ano.
    st.markdown("---")
    st.markdown("### 📈 Distribuição de Autos por Ano")
    
    anos_ordenados = sorted(anos_encontrados)
    quantidades = [stats_por_ano[ano]['quantidade'] for ano in anos_ordenados]
    
    fig_barras = go.Figure(data=[
        go.Bar(
            x=anos_ordenados,
            y=quantidades,
            marker_color='#1f4e79',
            text=quantidades,
            textposition='auto'
        )
    ])
    fig_barras.update_layout(
        title="Quantidade de Autos por Ano de Vencimento",
        xaxis_title="Ano",
        yaxis_title="Quantidade de Autos",
        height=400,
        showlegend=False
    )
    st.plotly_chart(fig_barras, use_container_width=True)
    
    # Gráfico de pizza
    col1, col2 = st.columns(2)
    
    with col1:
        fig_pizza = go.Figure(data=[go.Pie(
            labels=[str(ano) for ano in anos_ordenados],
            values=quantidades,
            hole=0.4,
            marker_colors=px.colors.qualitative.Set3
        )])
        fig_pizza.update_layout(
            title="Distribuição Percentual por Ano",
            height=400
        )
        st.plotly_chart(fig_pizza, use_container_width=True)
    
    with col2:
        # Resumo em tabela para leitura mais direta.
        st.markdown("### 📋 Resumo por Ano")
        df_resumo = pd.DataFrame({
            'Ano': anos_ordenados,
            'Quantidade de Autos': quantidades,
            'Percentual': [f"{(qtd/total_autos)*100:.2f}%" for qtd in quantidades]
        })
        st.dataframe(df_resumo, use_container_width=True, hide_index=True)
    
    # Cada ano vira um arquivo separado para facilitar o uso depois.
    st.markdown("---")
    st.markdown("### 📥 Exportar por Ano de Vencimento")
    st.info("💡 Baixe planilhas separadas por ano de vencimento. Cada arquivo contém apenas os autos daquele ano específico.")
    
    # O nome do arquivo leva a data para não sobrescrever exportações anteriores.
    data_arquivo = datetime.now().strftime('%d %m %Y %H:%M')
    
    # Organiza os botões em até 3 colunas para a tela não ficar comprida demais.
    num_colunas_export = min(3, len(anos_encontrados))  # Máximo 3 colunas
    cols_export = st.columns(num_colunas_export)
    
    for idx, ano in enumerate(sorted(anos_encontrados, reverse=False)):  # Do mais antigo para o mais recente
        col_idx = idx % num_colunas_export
        with cols_export[col_idx]:
            st.markdown(f"##### 📅 Ano {ano}")
            
            df_ano = stats_por_ano[ano]['dataframe'].copy()
            qtd_autos = stats_por_ano[ano]['quantidade']
            
            # A exportação segue a ordem que o pessoal já usa no dia a dia.
            colunas_export = {}
            
            # Essa coluna sempre entra.
            colunas_export['IDENTIFICADOR DE DÉBITO'] = df_ano[coluna_auto].fillna('').astype(str).str.strip()
            
            # As demais entram só se estiverem presentes na base.
            if coluna_protocolo and coluna_protocolo in df_ano.columns:
                colunas_export['Nº DE PROCESSO'] = df_ano[coluna_protocolo].fillna('').astype(str).str.strip()
            
            if coluna_modal and coluna_modal in df_ano.columns:
                colunas_export['MODAL'] = df_ano[coluna_modal].fillna('').astype(str).str.strip()
            
            if coluna_cpf_cnpj and coluna_cpf_cnpj in df_ano.columns:
                colunas_export['CNPJ'] = df_ano[coluna_cpf_cnpj].fillna('').astype(str).str.strip().apply(formatar_cpf_cnpj_brasileiro)
            
            # Vencimento entra sempre, mesmo quando precisa cair para texto.
            try:
                vencimento_dt = pd.to_datetime(
                    df_ano[coluna_vencimento],
                    errors='coerce',
                    dayfirst=True
                )
                colunas_export['DATA DE VENCIMENTO'] = vencimento_dt.dt.strftime('%d/%m/%Y').fillna('')
            except:
                colunas_export['DATA DE VENCIMENTO'] = df_ano[coluna_vencimento].fillna('').astype(str).str.strip()
            
            # Monta o DataFrame final na mesma ordem da planilha exportada.
            ordem_colunas = ['IDENTIFICADOR DE DÉBITO']
            if 'Nº DE PROCESSO' in colunas_export:
                ordem_colunas.append('Nº DE PROCESSO')
            if 'MODAL' in colunas_export:
                ordem_colunas.append('MODAL')
            if 'CNPJ' in colunas_export:
                ordem_colunas.append('CNPJ')
            ordem_colunas.append('DATA DE VENCIMENTO')
            
            dados_exportacao = pd.DataFrame({col: colunas_export[col] for col in ordem_colunas}, index=df_ano.index)
            
            # Evita carregar linha vazia para o Excel final.
            dados_exportacao = dados_exportacao.dropna(how='all')
            
            if not dados_exportacao.empty:
                try:
                    excel_ano = gerar_excel_vencimentos_formatado(
                        dados_exportacao,
                        f'Autos_{ano}',
                        f'Autos Vencimento {ano} {data_arquivo}.xlsx'
                    )
                    
                    st.download_button(
                        label=f"📥 Download Autos {ano}",
                        data=excel_ano,
                        file_name=f"Autos Vencimento {ano} {data_arquivo}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key=f"download_ano_{ano}",
                        help=f"Arquivo Excel com todos os autos de vencimento {ano}"
                    )
                    
                    st.success(f"✅ {qtd_autos:,} autos")
                    st.caption(f"📊 Ano {ano}")
                except Exception as e:
                    st.error(f"⚠️ Erro ao gerar arquivo: {str(e)}")
            else:
                st.warning(f"⚠️ Nenhum auto encontrado para {ano}")
    
    # Área para inspecionar um ano específico na própria tela.
    st.markdown("---")
    st.markdown("### 📋 Visualização Detalhada por Ano")
    
    # O mais recente aparece primeiro porque costuma ser o mais consultado.
    ano_selecionado = st.selectbox(
        "Selecione um ano para visualizar os autos:",
        options=sorted(anos_encontrados, reverse=True),
        key="ano_visualizacao"
    )
    
    if ano_selecionado:
        df_ano_selecionado = stats_por_ano[ano_selecionado]['dataframe'].copy()
        
        st.markdown(f"#### 📅 Autos de Vencimento {ano_selecionado}")
        st.info(f"💡 Mostrando {len(df_ano_selecionado):,} autos com vencimento em {ano_selecionado}")
        
        # Mostra só o que fizer sentido para leitura rápida.
        colunas_exibicao = [coluna_auto]
        if coluna_protocolo and coluna_protocolo in df_ano_selecionado.columns:
            colunas_exibicao.append(coluna_protocolo)
        if coluna_modal and coluna_modal in df_ano_selecionado.columns:
            colunas_exibicao.append(coluna_modal)
        if coluna_cpf_cnpj and coluna_cpf_cnpj in df_ano_selecionado.columns:
            colunas_exibicao.append(coluna_cpf_cnpj)
        if coluna_vencimento in df_ano_selecionado.columns:
            colunas_exibicao.append(coluna_vencimento)
        
        # Garante que a tela não quebre se alguma coluna opcional não existir.
        colunas_exibicao = [col for col in colunas_exibicao if col in df_ano_selecionado.columns]
        
        if colunas_exibicao:
            st.dataframe(
                df_ano_selecionado[colunas_exibicao],
                use_container_width=True,
                height=400
            )
        else:
            st.warning("⚠️ Nenhuma coluna disponível para exibição")
    
    # Rodapé simples com o fechamento da análise.
    st.markdown("---")
    st.caption(f"📅 Data de extração: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    st.caption(f"✅ **Total de autos analisados:** {total_autos:,}")
    st.caption(f"📊 **Anos encontrados:** {', '.join([str(ano) for ano in sorted(anos_encontrados)])}")

else:
    st.info("👆 Por favor, faça o upload da planilha na barra lateral para iniciar a análise.")
    
    # Instruções rápidas para primeiro uso.
    with st.expander("ℹ️ Como usar o sistema"):
        st.markdown("""
        ### Instruções de Uso:
        
        1. **Upload de Arquivo**: Faça o upload da planilha na barra lateral
        2. **Configuração de Colunas**: Informe os nomes exatos das colunas:
           - ⚠️ **Coluna Auto de Infração** (OBRIGATÓRIA)
           - ⚠️ **Coluna Vencimento** (OBRIGATÓRIA)
           - Coluna CPF/CNPJ (Opcional)
           - Coluna Valor (Opcional)
           - Coluna Nº do Processo (Opcional)
        3. **Executar Análise**: Clique no botão "Analisar por Ano de Vencimento"
        4. **Visualizar Resultados**: Explore o dashboard com métricas e gráficos
        5. **Exportar**: Baixe planilhas separadas por ano de vencimento
        
        ### Funcionalidades:
        - ✅ Análise automática por ano de vencimento
        - ✅ Dashboard com métricas e gráficos
        - ✅ Exportação separada por ano
        - ✅ Visualização detalhada por ano
        - ✅ Formatação profissional das planilhas Excel
        """)

