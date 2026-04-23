import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import json
import hashlib
import sqlite3

from app_helpers import (
    carregar_dados,
    deduplicar_por_protocolo_ou_auto,
    descrever_periodo_vencimento,
    filtrar_por_periodo_vencimento,
    formatar_periodo_analise,
    normalizar_intervalo_anos,
)
from comparison_analysis import analisar_bases
from utils import (
    normalizar_cpf_cnpj, normalizar_auto, converter_valor_sql,
    formatar_cpf_cnpj_brasileiro, formatar_valor_br,
    normalizar_e_mesclar_modais, resolver_coluna_vencimento,
)
from classificacao import (
    DEFAULT_CLASSIFICACAO_CONFIG,
    carregar_config_classificacao, salvar_config_classificacao,
    parse_lista_multilinha, obter_config_classificacao_ativa,
    classificar_autuado_detalhado,
    filtrar_autuados_cobraveis, set_session_config_getter,
)
from decadencia import (
    calcular_situacao_decadente, _resolver_coluna_data,
)
from exportacao import gerar_excel_formatado
from historico_db import save_run, list_runs, get_run, excluir_run

# Configuração básica da página.
st.set_page_config(
    page_title="Sistema de Análise SERASA x Dívida Ativa",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilo visual principal.
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

# Cabeçalho da aplicação.
st.markdown('<div class="main-header">📊 Sistema de Análise de Autos de Infração ANTT</div>', unsafe_allow_html=True)
st.markdown('<p style="text-align: center; color: #666; font-size: 1.1rem;">SERASA × Dívida Ativa - Análise Inteligente de Dados</p>', unsafe_allow_html=True)

# Barra lateral com upload, colunas e regras de apoio.
with st.sidebar:
    st.image("https://via.placeholder.com/200x80/1f4e79/ffffff?text=ANTT", use_container_width=True)
    st.markdown("### 📁 Upload de Arquivos")
    
    st.markdown("#### Base SERASA")
    arquivo_serasa = st.file_uploader(
        "Selecione a planilha SERASA",
        type=['xlsx', 'xls', 'csv'],
        key='serasa'
    )
    
    st.markdown("#### Base Dívida Ativa")
    arquivo_divida = st.file_uploader(
        "Selecione a planilha Dívida Ativa",
        type=['xlsx', 'xls', 'csv'],
        key='divida'
    )
    
    st.markdown("---")
    st.markdown("### ⚙️ Configurações")
    
    st.markdown("#### 🔑 Coluna Principal (Obrigatória)")
    # Essa é a coluna mais importante da comparação.
    coluna_auto = st.text_input(
        "Nome da coluna Auto de Infração",
        value="Identificador do Débito",
        help="⚠️ OBRIGATÓRIO: Digite o nome exato da coluna que contém os Autos de Infração (ex: CRGPF00074552019)"
    )
    
    st.markdown("#### 📋 Colunas Adicionais")
    # CPF/CNPJ ajuda nas análises adicionais e nas exportações.
    coluna_cpf_cnpj = st.text_input(
        "Nome da coluna CPF/CNPJ",
        value="CPF/CNPJ",
        help="Digite o nome exato da coluna que contém CPF/CNPJ"
    )
    
    # Valor da base SERASA.
    coluna_valor = st.text_input(
        "Nome da coluna Valor (SERASA)",
        value="Valor Multa Atualizado",
        help="Digite o nome exato da coluna de valor na base SERASA"
    )
    
    # Valor equivalente na base de Dívida Ativa.
    coluna_valor_divida = st.text_input(
        "Nome da coluna Valor (Dívida Ativa)",
        value="Valor Atualizado do Débito",
        help="Digite o nome exato da coluna de valor na base Dívida Ativa"
    )
    
    # O filtro por período de vencimento parte dessa coluna.
    coluna_vencimento = st.text_input(
        "Nome da coluna Vencimento",
        value="Data do Vencimento",
        help="Digite o nome exato da coluna que contém as datas de vencimento"
    )
    
    # O protocolo ajuda na deduplicação e na exportação.
    coluna_protocolo = st.text_input(
        "Nome da coluna Nº do Processo",
        value="Nº do Processo",
        help="Digite o nome exato da coluna que contém os números de protocolos"
    )

    st.markdown("#### 🚛 Colunas de Modais")
    # Modal da base SERASA.
    coluna_modal_serasa = st.text_input(
        "Nome da coluna Modal (SERASA)",
        value="Tipo Modal",
        help="Digite o nome exato da coluna que contém os modais na base SERASA"
    )
    
    # Modal/Subtipo vindo da Dívida Ativa.
    coluna_modal_divida = st.text_input(
        "Nome da coluna Modal (Dívida Ativa)",
        value="Subtipo de Débito",
        help="Digite o nome exato da coluna que contém os modais na base Dívida Ativa"
    )

    st.markdown("#### 📅 Período de Análise")
    ano_padrao = datetime.now().year
    col_ano_inicio, col_ano_fim = st.columns(2)
    with col_ano_inicio:
        ano_analise_inicial = st.number_input(
            "Ano inicial",
            min_value=1900,
            max_value=2035,
            value=ano_padrao,
            step=1,
            help="Ano inicial usado para filtrar autos por data de vencimento"
        )
    with col_ano_fim:
        ano_analise_final = st.number_input(
            "Ano final",
            min_value=1900,
            max_value=2035,
            value=ano_padrao,
            step=1,
            help="Ano final usado para filtrar autos por data de vencimento"
        )

    periodo_analise_valido = int(ano_analise_final) >= int(ano_analise_inicial)
    if periodo_analise_valido:
        st.caption(f"Período selecionado: {formatar_periodo_analise(ano_analise_inicial, ano_analise_final)}")
    else:
        st.warning("⚠️ O ano final deve ser maior ou igual ao ano inicial.")

    st.markdown("#### 👤 Classificação de Autuados (exportação base SERASA)")
    coluna_nome_autuado = st.text_input(
        "Nome da coluna do autuado (SERASA)",
        value="Nome Autuado",
        help="Coluna usada para classificar se o autuado pode ou não ser cobrado (órgãos, bancos, leasing)"
    )

if "classificacao_config" not in st.session_state:
    st.session_state["classificacao_config"] = carregar_config_classificacao()

with st.sidebar.expander("🛠️ Regras da classificação de autuados", expanded=False):
    st.caption("Adicione exceções e termos extras sem precisar alterar o código.")
    config_atual = st.session_state.get("classificacao_config", DEFAULT_CLASSIFICACAO_CONFIG.copy())
    extras_orgao_text = st.text_area(
        "Termos extras para Órgão",
        value="\n".join(config_atual.get("extras_orgao", [])),
        help="Um termo por linha. Ex.: POLICIA CIENTIFICA"
    )
    extras_banco_text = st.text_area(
        "Termos extras para Banco",
        value="\n".join(config_atual.get("extras_banco", [])),
        help="Um termo por linha. Ex.: BANCO MERCEDES"
    )
    extras_leasing_text = st.text_area(
        "Termos extras para Leasing",
        value="\n".join(config_atual.get("extras_leasing", [])),
        help="Um termo por linha. Ex.: LEASING OPERACIONAL"
    )
    excecoes_text = st.text_area(
        "Exceções - sempre Pode cobrar",
        value="\n".join(config_atual.get("excecoes_pode_cobrar", [])),
        help="Use aqui nomes/termos que estavam gerando falso positivo. Um por linha."
    )
    if st.button("💾 Salvar regras da classificação", use_container_width=True):
        novo_config = {
            "extras_orgao": parse_lista_multilinha(extras_orgao_text),
            "extras_banco": parse_lista_multilinha(extras_banco_text),
            "extras_leasing": parse_lista_multilinha(extras_leasing_text),
            "excecoes_pode_cobrar": parse_lista_multilinha(excecoes_text),
        }
        salvar_config_classificacao(novo_config)
        st.session_state["classificacao_config"] = novo_config
        st.success("✅ Regras salvas com sucesso.")

# A classificação do autuado usa a configuração que estiver ativa na sessão.
set_session_config_getter(lambda: st.session_state.get("classificacao_config"))

# A lógica mais pesada ficou separada em módulos para o app não concentrar tudo.

def obter_dataframe_cacheado_sessao(chave_cache, builder):
    """Guarda DataFrames intermediários na sessão para evitar reconstruções repetidas."""
    df_cache = st.session_state.get(chave_cache)
    if isinstance(df_cache, pd.DataFrame):
        return df_cache.copy()

    df_novo = builder()
    if isinstance(df_novo, pd.DataFrame):
        st.session_state[chave_cache] = df_novo.copy()
    return df_novo

# Fluxo principal.
if arquivo_serasa and arquivo_divida:
    st.markdown("---")
    
    # Primeiro carrega as bases para validar se tudo faz sentido antes da análise.
    with st.spinner("Carregando bases de dados..."):
        df_serasa = carregar_dados(arquivo_serasa, "SERASA")
        df_divida = carregar_dados(arquivo_divida, "Dívida Ativa")
    
    if df_serasa is not None and df_divida is not None:
        # O preview ajuda a conferir rapidamente se o arquivo certo foi enviado.
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("### 📋 Preview - Base SERASA")
            st.dataframe(df_serasa.head(), use_container_width=True)
            st.caption(f"Total de registros: {len(df_serasa)}")
            st.caption(f"Colunas: {', '.join(df_serasa.columns.tolist()[:5])}...")
        
        with col2:
            st.markdown("### 📋 Preview - Base Dívida Ativa")
            st.dataframe(df_divida.head(), use_container_width=True)
            st.caption(f"Total de registros: {len(df_divida)}")
            st.caption(f"Colunas: {', '.join(df_divida.columns.tolist()[:5])}...")
        
        st.markdown("---")
        
        def _hash_analise(df_s, df_d, params):
            """Monta uma assinatura simples para saber quando precisa recalcular."""
            h = hashlib.md5()
            h.update(str(len(df_s)).encode())
            h.update(str(len(df_d)).encode())
            h.update(str(list(df_s.columns)).encode())
            h.update(str(list(df_d.columns)).encode())
            for p in params:
                h.update(str(p).encode())
            return h.hexdigest()
        
        params_analise = [
            coluna_auto,
            coluna_cpf_cnpj,
            coluna_valor,
            coluna_vencimento,
            coluna_protocolo,
            coluna_modal_serasa,
            coluna_modal_divida,
            int(ano_analise_inicial),
            int(ano_analise_final),
        ]
        hash_atual = _hash_analise(df_serasa, df_divida, params_analise)
        
        cache_valido = (
            'resultados' in st.session_state and
            st.session_state.get('_analise_hash') == hash_atual
        )
        if cache_valido:
            st.info("💡 Resultados anteriores carregados do cache. Clique em \"Executar Análise Completa\" para re-analisar.")

        resultados = None
        if st.button("🚀 Executar Análise Completa", type="primary", use_container_width=True):
            if not coluna_auto or not coluna_auto.strip():
                st.error("⚠️ Por favor, informe o nome da coluna de Auto de Infração!")
            elif not periodo_analise_valido:
                st.error("⚠️ O ano final deve ser maior ou igual ao ano inicial para executar a análise.")
            else:
                with st.spinner("Analisando bases de dados por Auto de Infração... Isso pode levar alguns instantes."):
                    resultados = analisar_bases(
                        df_serasa, 
                        df_divida, 
                        coluna_auto,
                        coluna_cpf_cnpj, 
                        coluna_valor, 
                        coluna_vencimento,
                        coluna_protocolo,
                        coluna_modal_serasa,
                        coluna_modal_divida,
                        ano_analise_inicial=ano_analise_inicial,
                        ano_analise_final=ano_analise_final
                    )
            
            if resultados:
                st.session_state['resultados'] = resultados
                st.session_state['_analise_hash'] = hash_atual
                st.session_state['coluna_auto'] = coluna_auto
                st.session_state['coluna_cpf_cnpj'] = coluna_cpf_cnpj
                st.session_state['coluna_valor'] = coluna_valor
                st.session_state['coluna_valor_divida'] = coluna_valor_divida
                st.session_state['coluna_vencimento'] = coluna_vencimento
                st.session_state['coluna_protocolo'] = coluna_protocolo
                st.session_state['coluna_modal_serasa'] = coluna_modal_serasa
                st.session_state['coluna_modal_divida'] = coluna_modal_divida
                st.session_state['coluna_nome_autuado'] = coluna_nome_autuado
                st.session_state['ano_analise_inicial'] = int(ano_analise_inicial)
                st.session_state['ano_analise_final'] = int(ano_analise_final)
                st.session_state['ano_analise'] = formatar_periodo_analise(ano_analise_inicial, ano_analise_final)
                st.session_state['export_run_id'] = datetime.now().strftime('%Y%m%d%H%M%S%f')
                st.session_state['export_run_label'] = datetime.now().strftime('%d %m %Y %H:%M')
                st.session_state['nome_arquivo_serasa'] = arquivo_serasa.name
                st.session_state['nome_arquivo_divida'] = arquivo_divida.name
                st.session_state['_historico_salvo'] = False
                st.success("✅ Análise concluída com sucesso!")
                st.rerun()

# Exibição dos resultados que já ficaram em sessão.
if 'resultados' in st.session_state:
    resultados = st.session_state['resultados']
    # Recupera as colunas usadas na última análise.
    coluna_auto = st.session_state.get('coluna_auto', 'Identificador do Débito')
    coluna_cpf_cnpj = st.session_state.get('coluna_cpf_cnpj', 'CPF/CNPJ')
    # Valor principal da análise.
    coluna_valor = st.session_state.get('coluna_valor', 'Valor Multa Atualizado')
    # Valor da base espelho.
    coluna_valor_divida = st.session_state.get('coluna_valor_divida', 'Valor Atualizado do Débito')
    coluna_vencimento = st.session_state.get('coluna_vencimento', 'Data do Vencimento')
    coluna_protocolo = st.session_state.get('coluna_protocolo', 'Nº do Processo')
    coluna_modal_serasa = st.session_state.get('coluna_modal_serasa', 'Tipo Modal')
    coluna_modal_divida = st.session_state.get('coluna_modal_divida', 'Subtipo de Débito')
    coluna_nome_autuado = st.session_state.get('coluna_nome_autuado', 'Nome Autuado')
    ano_analise_inicial = st.session_state.get('ano_analise_inicial', st.session_state.get('ano_analise', 2025))
    ano_analise_final = st.session_state.get('ano_analise_final', ano_analise_inicial)
    ano_analise_inicial, ano_analise_final = normalizar_intervalo_anos(ano_analise_inicial, ano_analise_final)
    ano_analise = formatar_periodo_analise(ano_analise_inicial, ano_analise_final)
    descricao_periodo_vencimento = descrever_periodo_vencimento(ano_analise_inicial, ano_analise_final)
    
    st.markdown("---")
    st.markdown("## 📊 Dashboard de Resultados - Análise por Auto de Infração")
    
    st.markdown("### 🔑 Análise Principal: Autos de Infração")
    # Resumo principal olhando os autos.
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric(
            "Total Registros SERASA",
            f"{resultados['total_registros_serasa']:,}",
            delta=f"{resultados['total_autos_serasa']:,} autos únicos válidos"
        )
        st.caption(f"📋 {len(resultados['df_serasa_filtrado']):,} registros após filtros ({descricao_periodo_vencimento})")
    
    with col2:
        st.metric(
            "Total Registros Dívida Ativa",
            f"{resultados['total_registros_divida']:,}",
            delta=f"{resultados['total_autos_divida']:,} autos únicos válidos"
        )
        st.caption(f"📋 {len(resultados['df_divida_filtrado']):,} registros após filtros ({descricao_periodo_vencimento})")
    
    with col3:
        # Aqui entram as linhas que sustentam a exportação comparada.
        st.metric(
            "Autos em Ambas as Bases",
            f"{resultados['autos_em_ambas']:,}",
            delta=f"{resultados.get('autos_em_ambas_unicos', 0):,} autos únicos"
        )
    
    with col4:
        taxa_match_autos = (resultados['autos_em_ambas'] / max(resultados['total_autos_serasa'], 1)) * 100
        st.metric(
            "Taxa de Correspondência",
            f"{taxa_match_autos:.1f}%",
            delta="Autos SERASA vs Dívida Ativa"
        )
    
    # Faixas de valor montadas a partir da lógica acumulativa por CPF/CNPJ.
    st.markdown(f"#### 💰 Autos por Faixa de Valor (SERASA - {descricao_periodo_vencimento})")
    col5, col6, col7 = st.columns(3)
    
    with col5:
        st.metric(
            "Autos até R$ 999,99",
            f"{resultados.get('qtd_autos_ate_999', 0):,}",
            delta="Soma acumulativa por CNPJ"
        )
    
    with col6:
        st.metric(
            "Autos R$ 500,00 a R$ 999,99",
            f"{resultados.get('qtd_autos_500_999', 0):,}",
            delta="Soma acumulativa por CNPJ sem decadentes"
        )
    
    with col7:
        st.metric(
            "Autos acima de R$ 1.000,00",
            f"{resultados.get('qtd_autos_acima_1000', 0):,}",
            delta="Soma acumulativa por CNPJ"
        )
    
    # Visão geral sem recorte de vencimento, só para comparar com o filtrado.
    st.markdown("---")
    st.markdown("#### 📊 Comparação Geral - Todos os Autos (Sem Filtro de Vencimento)")
    st.info("💡 Esta seção mostra a comparação de TODOS os autos em ambas as bases, **independente da data de vencimento**. Mantém todos os outros filtros (valores > 0, remoção de duplicados, etc.).")
    
    # Totais gerais da comparação.
    autos_geral = resultados.get('autos_em_ambas_geral', 0)
    autos_geral_unicos = resultados.get('autos_em_ambas_geral_unicos', 0)
    
    # Soma total da visão sem filtro de data.
    df_final_geral = resultados.get('df_final_geral', pd.DataFrame())
    valor_total_geral = 0.0
    qtd_autos_geral_valor = 0
    if not df_final_geral.empty and 'Valor' in df_final_geral.columns:
        valores_validos_geral = df_final_geral['Valor'][df_final_geral['Valor'].notna() & (df_final_geral['Valor'] > 0)]
        valor_total_geral = float(valores_validos_geral.sum()) if len(valores_validos_geral) > 0 else 0.0
        qtd_autos_geral_valor = len(valores_validos_geral)
    
    col_geral1, col_geral2, col_geral3, col_geral4 = st.columns(4)
    
    with col_geral1:
        st.metric(
            "Autos em Ambas (Geral)",
            f"{autos_geral:,}",
            delta=f"{autos_geral_unicos:,} autos únicos"
        )
        st.caption("📋 Todos os autos (sem filtro de data)")
    
    with col_geral2:
        st.metric(
            "Valor Total (Geral)",
            f"R$ {valor_total_geral:,.2f}",
            delta=f"{qtd_autos_geral_valor:,} autos"
        )
        st.caption("💰 Soma de todos os valores")
    
    with col_geral3:
        # Diferença entre o geral e o recorte atual.
        autos_filtrado = resultados.get('autos_em_ambas', 0)
        diferenca_autos = autos_geral - autos_filtrado
        st.metric(
            "Diferença (Geral vs Filtrado)",
            f"{diferenca_autos:,} autos",
            delta=f"Geral - Filtrado ({descricao_periodo_vencimento})"
        )
        st.caption("📊 Diferença entre geral e filtrado")
    
    with col_geral4:
        # Quanto do geral ficou coberto pelo período selecionado.
        if autos_geral > 0:
            taxa_cobertura = (autos_filtrado / autos_geral) * 100
            st.metric(
                f"Cobertura {ano_analise}",
                f"{taxa_cobertura:.1f}%",
                delta=f"{autos_filtrado:,} de {autos_geral:,}"
            )
            st.caption(f"📈 % dos autos com {descricao_periodo_vencimento}")
        else:
            st.metric(f"Cobertura {ano_analise}", "N/A")
    
    # Gráfico simples para bater o olho na diferença entre geral e filtrado.
    st.markdown("---")
    st.markdown(f"##### 📊 Comparação Visual: Geral vs Filtrado ({descricao_periodo_vencimento})")
    fig_comparacao = go.Figure(data=[
        go.Bar(name='Todos os Autos (Geral)', x=['Comparação'], y=[autos_geral], marker_color='#3498db', text=f"{autos_geral:,}", textposition='auto'),
        go.Bar(name=f'Autos no período {ano_analise} (Filtrado)', x=['Comparação'], y=[autos_filtrado], marker_color='#2ecc71', text=f"{autos_filtrado:,}", textposition='auto')
    ])
    fig_comparacao.update_layout(
        title=f"Comparação: Todos os Autos vs Autos com Vencimento no Período {ano_analise}",
        xaxis_title="",
        yaxis_title="Quantidade de Autos",
        barmode='group',
        height=400,
        showlegend=True
    )
    st.plotly_chart(fig_comparacao, use_container_width=True)
    st.caption(f"💡 Comparação entre todos os autos ({autos_geral:,}) e apenas os com vencimento no período {ano_analise} ({autos_filtrado:,})")
    
    # Bloco de conferência para validar se os agrupamentos fecharam.
    st.markdown("---")
    st.markdown("#### ✅ Validação de Exatidão dos Dados")
    st.info("💡 Esta seção valida se os cálculos estão 100% corretos conforme os dados da planilha.")
    
    # Aqui a ideia é checar se a soma das faixas bate com o total esperado.
    total_autos_em_ambas = resultados.get('autos_em_ambas', 0)
    total_ate_999 = resultados.get('qtd_autos_ate_999', 0)
    total_acima_1000 = resultados.get('qtd_autos_acima_1000', 0)
    soma_grupos = total_ate_999 + total_acima_1000
    diferenca = total_autos_em_ambas - soma_grupos
    
    col_val1, col_val2, col_val3 = st.columns(3)
    
    with col_val1:
        st.metric(
            "Total Autos em Ambas",
            f"{total_autos_em_ambas:,}",
            delta=f"Linhas válidas ({descricao_periodo_vencimento} e valor > 0)"
        )
    
    with col_val2:
        st.metric(
            "Soma dos Grupos",
            f"{soma_grupos:,}",
            delta=f"Até 999,99: {total_ate_999:,} + Acima 1000: {total_acima_1000:,}"
        )
    
    with col_val3:
        if diferenca == 0:
            st.metric(
                "✅ Validação",
                "CORRETO",
                delta="Soma dos grupos = Total"
            )
        else:
            st.metric(
                "⚠️ Validação",
                f"Diferença: {diferenca:,}",
                delta="Valores entre 999 e 1000 (não contados nos grupos)"
            )
            st.caption(f"💡 {diferenca:,} autos têm valores entre R$ 999,99 e R$ 1.000,00 (não incluídos nos grupos filtrados)")
    
    # Gráficos principais dos autos.
    col1, col2 = st.columns(2)
    
    with col1:
        # Comparação direta entre o que bateu e o que sobrou em cada base.
        fig_corresp = go.Figure(data=[
            go.Bar(name='Em Ambas', x=['Autos de Infração'], y=[resultados['autos_em_ambas']], marker_color='#2ecc71'),
            go.Bar(name='Apenas SERASA', x=['Autos de Infração'], y=[resultados['autos_apenas_serasa']], marker_color='#3498db'),
            go.Bar(name='Apenas Dívida Ativa', x=['Autos de Infração'], y=[resultados['autos_apenas_divida']], marker_color='#e74c3c')
        ])
        fig_corresp.update_layout(
            title='Análise de Correspondência de Autos de Infração',
            barmode='group',
            height=400,
            showlegend=True
        )
        st.plotly_chart(fig_corresp, use_container_width=True)
    
    with col2:
        # A mesma informação em formato proporcional.
        labels = ['Em Ambas', 'Apenas SERASA', 'Apenas Dívida Ativa']
        values = [resultados['autos_em_ambas'], resultados['autos_apenas_serasa'], resultados['autos_apenas_divida']]
        colors = ['#2ecc71', '#3498db', '#e74c3c']
        
        fig_pizza = go.Figure(data=[go.Pie(
            labels=labels,
            values=values,
            hole=0.4,
            marker_colors=colors
        )])
        fig_pizza.update_layout(
            title='Distribuição de Autos de Infração',
            height=400
        )
        st.plotly_chart(fig_pizza, use_container_width=True)
    
    # Visão complementar por CPF/CNPJ.
    st.markdown("### 📋 Análise Adicional: CPF/CNPJ")
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric(
            "CPF/CNPJ SERASA",
            f"{resultados['total_cpf_serasa']:,}",
            delta="CPF/CNPJ únicos"
        )
    
    with col2:
        st.metric(
            "CPF/CNPJ Dívida Ativa",
            f"{resultados['total_cpf_divida']:,}",
            delta="CPF/CNPJ únicos"
        )
    
    with col3:
        st.metric(
            "CPF/CNPJ em Ambas",
            f"{resultados['cpf_em_ambas']:,}",
            delta="CPF/CNPJ correspondentes"
        )
    
    with col4:
        if resultados['total_cpf_serasa'] > 0:
            taxa_match_cpf = (resultados['cpf_em_ambas'] / resultados['total_cpf_serasa']) * 100
            st.metric(
                "Taxa CPF/CNPJ",
                f"{taxa_match_cpf:.1f}%",
                delta="Correspondência"
            )
        else:
            st.metric("Taxa CPF/CNPJ", "N/A", delta="Sem dados")
    
    # A decadência fica separada porque depende de datas e modais específicos.
    df_final_sql = resultados.get('df_final_sql', pd.DataFrame())
    if not df_final_sql.empty and 'Situação decadente' in df_final_sql.columns:
        st.markdown("---")
        st.markdown("### ⏱️ Análise de Decadência")
        st.caption("Regra: o prazo começa no **primeiro dia útil** a partir da Data Infração (a própria data se for dia útil). A partir desse dia 1, a contagem é em **dias corridos**: a notificação de autuação (ajustada em +4 dias) não pode ultrapassar **31 dias corridos**, e a notificação de multa (ajustada em +4 dias) não pode ultrapassar **181 dias corridos**. Autos que ultrapassam são decadentes.")
        sit = df_final_sql['Situação decadente'].fillna('').astype(str).str.strip()
        qtd_decad_autuacao = int(((sit == 'Decadente autuação') | (sit == 'Decadente autuação e multa')).sum())
        qtd_decad_multa = int(((sit == 'Decadente multa') | (sit == 'Decadente autuação e multa')).sum())
        qtd_decad_ambos = int((sit == 'Decadente autuação e multa').sum())
        col_d1, col_d2, col_d3 = st.columns(3)
        with col_d1:
            st.metric("Decadentes (autuação)", f"{qtd_decad_autuacao:,}", delta="> 31 dias corridos")
        with col_d2:
            st.metric("Decadentes (multa)", f"{qtd_decad_multa:,}", delta="> 181 dias corridos")
        with col_d3:
            st.metric("Decadentes (autuação e multa)", f"{qtd_decad_ambos:,}", delta="ambos prazos ultrapassados")
        # Quebra por ano da infração para ajudar na leitura do volume.
        col_infracao = _resolver_coluna_data(df_final_sql, "Data Infração")
        if col_infracao:
            data_infracao_dt = pd.to_datetime(df_final_sql[col_infracao], errors='coerce', dayfirst=True)
            df_final_sql = df_final_sql.copy()
            df_final_sql['_ano_infracao'] = data_infracao_dt.dt.year
            df_decad = df_final_sql[sit != ''].copy()
            if not df_decad.empty and df_decad['_ano_infracao'].notna().any():
                df_decad['Decadente autuação'] = (df_decad['Situação decadente'].fillna('').str.contains('autuação', na=False)).astype(int)
                df_decad['Decadente multa'] = (df_decad['Situação decadente'].fillna('').str.contains('multa', na=False)).astype(int)
                por_ano = df_decad.groupby('_ano_infracao').agg({'Decadente autuação': 'sum', 'Decadente multa': 'sum'}).reset_index()
                por_ano = por_ano.rename(columns={'_ano_infracao': 'Ano'})
                fig_decad = go.Figure()
                fig_decad.add_trace(go.Bar(name='Decadente autuação', x=por_ano['Ano'].astype(int), y=por_ano['Decadente autuação'], marker_color='#e74c3c'))
                fig_decad.add_trace(go.Bar(name='Decadente multa', x=por_ano['Ano'].astype(int), y=por_ano['Decadente multa'], marker_color='#f39c12'))
                fig_decad.update_layout(
                    title='Decadentes por ano (Data Infração)',
                    barmode='group',
                    xaxis_title='Ano',
                    yaxis_title='Quantidade',
                    height=400,
                    showlegend=True
                )
                st.plotly_chart(fig_decad, use_container_width=True)
    
    # A partir daqui entra o detalhamento por tema.
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "🔑 Autos de Infração",
        "📈 Agrupamento por CPF/CNPJ",
        "💰 Separação por Valores",
        "⚠️ Divergências",
        "📜 Histórico"
    ])
    
    with tab1:
        st.markdown("### 🔑 Análise de Autos de Infração")
        st.info("💡 Esta é a análise PRINCIPAL. Os autos de infração são a chave de comparação entre as bases.")
        
        st.markdown(f"#### ✅ Autos Presentes em Ambas as Bases ({descricao_periodo_vencimento})")
        st.success(f"Total de {resultados['autos_em_ambas']:,} autos de infração encontrados em ambas as bases com vencimento no período {ano_analise}")
        
        # Monta uma visão mais útil para conferência na tela.
        if not resultados['df_serasa_filtrado'].empty and not resultados['df_divida_filtrado'].empty:
            # Trabalha em cópias para não mexer no que veio da análise.
            df_serasa_comp = resultados['df_serasa_filtrado'].copy()
            df_divida_comp = resultados['df_divida_filtrado'].copy()
            
            # Auto, valor e vencimento vêm primeiro porque são o que mais importa aqui.
            colunas_importantes = []
            if coluna_auto in df_serasa_comp.columns:
                colunas_importantes.append(coluna_auto)
            if coluna_valor in df_serasa_comp.columns:
                colunas_importantes.append(coluna_valor)
            if coluna_vencimento in df_serasa_comp.columns:
                colunas_importantes.append(coluna_vencimento)
            
            # O restante entra depois, só para manter a leitura mais organizada.
            outras_colunas = [c for c in df_serasa_comp.columns if c not in colunas_importantes and c not in ['AUTO_NORM', 'CPF_CNPJ_NORM']]
            colunas_ordenadas = colunas_importantes + outras_colunas
            
            # A ideia é deixar a tabela pronta para bater o olho sem precisar reordenar manualmente.
            df_serasa_display = df_serasa_comp[[c for c in colunas_ordenadas if c in df_serasa_comp.columns]]
            df_divida_display = df_divida_comp[[c for c in colunas_ordenadas if c in df_divida_comp.columns]]
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("##### 📊 Base SERASA - Autos Correspondentes")
                st.markdown(f"**Valores e vencimentos no período {ano_analise}**")
                st.info("💡 Mostrando apenas autos que estão em **AMBAS as bases** (SERASA e Dívida Ativa)")
                st.dataframe(df_serasa_display, use_container_width=True, height=400)
                
                if coluna_valor in df_serasa_comp.columns:
                    try:
                        # Se o valor veio como texto, converte antes de somar.
                        if df_serasa_comp[coluna_valor].dtype not in ['int64', 'float64']:
                            valores_serasa = pd.to_numeric(df_serasa_comp[coluna_valor], errors='coerce')
                        else:
                            valores_serasa = df_serasa_comp[coluna_valor]
                        
                        # Faz a conta só com o que realmente virou número.
                        valores_serasa_validos = valores_serasa[valores_serasa.notna()]
                        if len(valores_serasa_validos) > 0:
                            # A soma daqui precisa bater com a leitura feita na planilha.
                            soma_serasa = float(valores_serasa_validos.sum())
                            media_serasa = float(valores_serasa_validos.mean())
                            st.metric("💰 Valor Total SERASA (Em Ambas)", f"R$ {soma_serasa:,.2f}", 
                                     delta=f"Média: R$ {media_serasa:,.2f}")
                    except Exception as e:
                        st.error(f"Erro ao calcular valores SERASA: {str(e)}")
                        pass
                
                st.caption(f"📋 Total: {len(resultados['df_serasa_filtrado'])} autos de infração em ambas as bases (período {ano_analise})")
            
            with col2:
                st.markdown("##### 📊 Base Dívida Ativa - Autos Correspondentes")
                st.markdown(f"**Valores e vencimentos no período {ano_analise}**")
                st.info("💡 Mostrando apenas autos que estão em **AMBAS as bases** (SERASA e Dívida Ativa)")
                st.dataframe(df_divida_display, use_container_width=True, height=400)
                
                # Repete a mesma leitura para a base de Dívida Ativa.
                if coluna_valor in df_divida_comp.columns:
                    try:
                        # Se o valor veio como texto, converte antes de somar.
                        if df_divida_comp[coluna_valor].dtype not in ['int64', 'float64']:
                            valores_divida = pd.to_numeric(df_divida_comp[coluna_valor], errors='coerce')
                        else:
                            valores_divida = df_divida_comp[coluna_valor]
                        
                        # Faz a conta só com o que realmente virou número.
                        valores_divida_validos = valores_divida[valores_divida.notna()]
                        if len(valores_divida_validos) > 0:
                            # A soma daqui precisa bater com a leitura feita na planilha.
                            soma_divida = float(valores_divida_validos.sum())
                            media_divida = float(valores_divida_validos.mean())
                            st.metric("💰 Valor Total Dívida Ativa (Em Ambas)", f"R$ {soma_divida:,.2f}", 
                                     delta=f"Média: R$ {media_divida:,.2f}")
                    except Exception as e:
                        st.error(f"Erro ao calcular valores Dívida Ativa: {str(e)}")
                        pass
                
                st.caption(f"📋 Total: {len(resultados['df_divida_filtrado'])} autos de infração em ambas as bases (período {ano_analise})")
            
            st.markdown("---")
            st.markdown(f"#### 💰 Totais Gerais (TODOS os autos com {descricao_periodo_vencimento})")
            st.warning(f"⚠️ **IMPORTANTE:** Estes são os totais de TODOS os autos de cada base com vencimento no período {ano_analise}, não apenas os que estão em ambas as bases.")
            
            if coluna_valor in resultados['df_serasa_total_ano'].columns and coluna_valor in resultados['df_divida_total_ano'].columns:
                col1, col2, col3 = st.columns(3)
                
                try:
                    # Aqui entra a base inteira do período, não só a interseção.
                    df_serasa_total = resultados['df_serasa_total_ano'].copy()
                    if df_serasa_total[coluna_valor].dtype not in ['int64', 'float64']:
                        valores_serasa_total = pd.to_numeric(df_serasa_total[coluna_valor], errors='coerce')
                    else:
                        valores_serasa_total = df_serasa_total[coluna_valor]
                    
                    valores_serasa_total_validos = valores_serasa_total[valores_serasa_total.notna()]
                    soma_serasa_total = float(valores_serasa_total_validos.sum()) if len(valores_serasa_total_validos) > 0 else 0.0
                    
                    # Aqui vale a mesma lógica para a Dívida Ativa.
                    df_divida_total = resultados['df_divida_total_ano'].copy()
                    if df_divida_total[coluna_valor].dtype not in ['int64', 'float64']:
                        valores_divida_total = pd.to_numeric(df_divida_total[coluna_valor], errors='coerce')
                    else:
                        valores_divida_total = df_divida_total[coluna_valor]
                    
                    valores_divida_total_validos = valores_divida_total[valores_divida_total.notna()]
                    soma_divida_total = float(valores_divida_total_validos.sum()) if len(valores_divida_total_validos) > 0 else 0.0
                    
                    with col1:
                        st.metric("💰 Total SERASA (Todos)", f"R$ {soma_serasa_total:,.2f}", 
                                 delta=f"{len(valores_serasa_total_validos):,} autos")
                        st.caption(f"Soma de TODOS os valores da coluna de valor da SERASA com vencimento no período {ano_analise}")
                    
                    with col2:
                        st.metric("💰 Total Dívida Ativa (Todos)", f"R$ {soma_divida_total:,.2f}", 
                                 delta=f"{len(valores_divida_total_validos):,} autos")
                        st.caption(f"Soma de TODOS os valores da coluna de valor da Dívida Ativa com vencimento no período {ano_analise}")
                    
                    with col3:
                        diferenca_total = soma_serasa_total - soma_divida_total
                        st.metric("Diferença", f"R$ {diferenca_total:,.2f}", 
                                 delta="SERASA - Dívida Ativa")
                except Exception as e:
                    st.error(f"Erro ao calcular totais: {str(e)}")
            
            # Nesta parte a conta olha só o recorte que está nas duas bases.
            st.markdown("---")
            st.markdown("#### 💰 Comparação de Valores (Apenas autos em AMBAS as bases)")
            st.info(f"💡 Estes são os totais apenas dos autos que estão presentes em AMBAS as bases (SERASA e Dívida Ativa) com vencimento no período {ano_analise}.")
            
            if coluna_valor in df_serasa_comp.columns and coluna_valor in df_divida_comp.columns:
                col1, col2, col3 = st.columns(3)
                
                try:
                    # Converte os dois lados antes de comparar.
                    if df_serasa_comp[coluna_valor].dtype not in ['int64', 'float64']:
                        valores_serasa = pd.to_numeric(df_serasa_comp[coluna_valor], errors='coerce')
                    else:
                        valores_serasa = df_serasa_comp[coluna_valor]
                    
                    if df_divida_comp[coluna_valor].dtype not in ['int64', 'float64']:
                        valores_divida = pd.to_numeric(df_divida_comp[coluna_valor], errors='coerce')
                    else:
                        valores_divida = df_divida_comp[coluna_valor]
                    
                    # Ignora o que não virou número.
                    valores_serasa_validos = valores_serasa[valores_serasa.notna()]
                    valores_divida_validos = valores_divida[valores_divida.notna()]
                    
                    # A soma precisa refletir só os valores válidos.
                    soma_serasa = float(valores_serasa_validos.sum()) if len(valores_serasa_validos) > 0 else 0.0
                    soma_divida = float(valores_divida_validos.sum()) if len(valores_divida_validos) > 0 else 0.0
                    
                    with col1:
                        st.metric("Total SERASA (Em Ambas)", f"R$ {soma_serasa:,.2f}", 
                                 delta=f"{len(valores_serasa_validos):,} autos")
                        st.caption("Soma dos valores dos autos que estão em ambas as bases")
                    
                    with col2:
                        st.metric("Total Dívida Ativa (Em Ambas)", f"R$ {soma_divida:,.2f}", 
                                 delta=f"{len(valores_divida_validos):,} autos")
                        st.caption("Soma dos valores dos autos que estão em ambas as bases")
                    
                    with col3:
                        diferenca = soma_serasa - soma_divida
                        st.metric("Diferença", f"R$ {diferenca:,.2f}", 
                                 delta="SERASA - Dívida Ativa")
                except Exception as e:
                    st.error(f"Erro ao calcular comparação: {str(e)}")
            
            # Faixa de vencimentos do período inteiro em cada base.
            st.markdown("---")
            st.markdown(f"#### 📅 Resumo de Vencimentos - Totais Gerais ({descricao_periodo_vencimento})")
            st.warning(f"⚠️ **IMPORTANTE:** Estes são os totais de TODOS os autos de cada base com {descricao_periodo_vencimento}.")
            col1, col2 = st.columns(2)
            
            with col1:
                try:
                    df_serasa_total = resultados['df_serasa_total_ano'].copy()
                    if coluna_vencimento in df_serasa_total.columns:
                        venc_serasa = pd.to_datetime(df_serasa_total[coluna_vencimento], errors='coerce')
                        data_limite = pd.Timestamp(f'{ano_analise_inicial}-01-01')
                        data_limite_fim = pd.Timestamp(f'{ano_analise_final}-12-31')
                        venc_serasa_filtrado = venc_serasa[(venc_serasa >= data_limite) & (venc_serasa <= data_limite_fim)]
                        if len(venc_serasa_filtrado) > 0:
                            st.markdown("**SERASA (Todos os autos):**")
                            st.write(f"- Primeiro vencimento: {venc_serasa_filtrado.min().strftime('%d/%m/%Y')}")
                            st.write(f"- Último vencimento: {venc_serasa_filtrado.max().strftime('%d/%m/%Y')}")
                            st.write(f"- **Total com vencimento no período {ano_analise}: {len(df_serasa_total):,} autos**")
                except Exception as e:
                    st.error(f"Erro ao processar vencimentos SERASA: {str(e)}")
            
            with col2:
                try:
                    df_divida_total = resultados['df_divida_total_ano'].copy()
                    if coluna_vencimento in df_divida_total.columns:
                        venc_divida = pd.to_datetime(df_divida_total[coluna_vencimento], errors='coerce')
                        data_limite = pd.Timestamp(f'{ano_analise_inicial}-01-01')
                        data_limite_fim = pd.Timestamp(f'{ano_analise_final}-12-31')
                        venc_divida_filtrado = venc_divida[(venc_divida >= data_limite) & (venc_divida <= data_limite_fim)]
                        if len(venc_divida_filtrado) > 0:
                            st.markdown("**Dívida Ativa (Todos os autos):**")
                            st.write(f"- Primeiro vencimento: {venc_divida_filtrado.min().strftime('%d/%m/%Y')}")
                            st.write(f"- Último vencimento: {venc_divida_filtrado.max().strftime('%d/%m/%Y')}")
                            st.write(f"- **Total com vencimento no período {ano_analise}: {len(df_divida_total):,} autos**")
                except Exception as e:
                    st.error(f"Erro ao processar vencimentos Dívida Ativa: {str(e)}")
            
            st.markdown("---")
            
            # A exportação principal sai daqui.
            st.markdown("#### 📥 Exportar Base Comparada")
            # O texto muda um pouco quando a exportação vai levar protocolo.
            tem_protocolo = coluna_protocolo in resultados['df_serasa_filtrado'].columns if not resultados['df_serasa_filtrado'].empty else False
            if tem_protocolo:
                st.info("💡 Exporte os autos que estão em **AMBAS as bases** (SERASA e Dívida Ativa). **IMPORTANTE:** Antes da exportação, o sistema **exclui autuados classificados como Órgão, Banco ou Leasing** pela coluna de nome do autuado da SERASA, preservando exceções permitidas como **SAFRA**. A classificação por faixa de valor (até R$ 999,99 ou acima de R$ 1.000,00) é feita pela **SOMA dos valores por CPF/CNPJ**. Se um CPF/CNPJ tiver soma <= R$ 999,99, TODOS os seus autos vão para 'até R$ 999,99'. Se tiver soma >= R$ 1.000,00, TODOS os seus autos vão para 'acima de R$ 1.000,00'. Cada arquivo contém: Auto de Infração, **Número de Protocolo**, CPF/CNPJ e **Valor Individual de cada auto**, ordenados por CPF/CNPJ (do maior para menor número de autos).")
            else:
                st.info("💡 Exporte os autos que estão em **AMBAS as bases** (SERASA e Dívida Ativa). **IMPORTANTE:** Antes da exportação, o sistema **exclui autuados classificados como Órgão, Banco ou Leasing** pela coluna de nome do autuado da SERASA, preservando exceções permitidas como **SAFRA**. A classificação por faixa de valor (até R$ 999,99 ou acima de R$ 1.000,00) é feita pela **SOMA dos valores por CPF/CNPJ**. Se um CPF/CNPJ tiver soma <= R$ 999,99, TODOS os seus autos vão para 'até R$ 999,99'. Se tiver soma >= R$ 1.000,00, TODOS os seus autos vão para 'acima de R$ 1.000,00'. Cada arquivo contém: Auto de Infração, CPF/CNPJ e **Valor Individual de cada auto**, ordenados por CPF/CNPJ (do maior para menor número de autos).")
            
            # A base de exportação parte do resultado já tratado pela análise principal.
            if 'df_final_sql' in resultados and not resultados['df_final_sql'].empty:
                    cache_df_export_key = f"cache_df_export::{st.session_state.get('_analise_hash', 'default')}"

                    def montar_df_export():
                        df_final_work = resultados['df_final_sql'].copy()

                        # Tenta reconstruir a estrutura final usando primeiro o lado da SERASA.
                        df_export_local = pd.DataFrame()

                        # O auto é a coluna-base da exportação.
                        if f"{coluna_auto}_serasa" in df_final_work.columns:
                            df_export_local[coluna_auto] = df_final_work[f"{coluna_auto}_serasa"]
                        elif f"{coluna_auto}_divida" in df_final_work.columns:
                            df_export_local[coluna_auto] = df_final_work[f"{coluna_auto}_divida"]
                        elif 'AUTO_NORM' in df_final_work.columns:
                            autos_finais = set(df_final_work['AUTO_NORM'].unique())
                            df_export_local = resultados['df_serasa_filtrado'][resultados['df_serasa_filtrado']['AUTO_NORM'].isin(autos_finais)].copy()
                        else:
                            df_export_local = resultados['df_serasa_filtrado'].copy()

                        if coluna_auto in df_export_local.columns and len(df_export_local) == len(df_final_work):
                            if 'Valor' in df_final_work.columns:
                                df_export_local[coluna_valor] = df_final_work['Valor']

                            if coluna_cpf_cnpj in resultados['df_serasa_filtrado'].columns:
                                if f"{coluna_cpf_cnpj}_serasa" in df_final_work.columns:
                                    df_export_local[coluna_cpf_cnpj] = df_final_work[f"{coluna_cpf_cnpj}_serasa"]
                                elif f"{coluna_cpf_cnpj}_divida" in df_final_work.columns:
                                    df_export_local[coluna_cpf_cnpj] = df_final_work[f"{coluna_cpf_cnpj}_divida"]

                            vencimento_adicionado = False
                            col_venc_df_final = resolver_coluna_vencimento(df_final_work, coluna_vencimento)
                            if col_venc_df_final:
                                df_export_local[coluna_vencimento] = df_final_work[col_venc_df_final]
                                vencimento_adicionado = True

                            if not vencimento_adicionado:
                                col_venc_serasa = resolver_coluna_vencimento(resultados['df_serasa_filtrado'], coluna_vencimento)
                                if col_venc_serasa and 'AUTO_NORM' in df_final_work.columns and 'AUTO_NORM' in resultados['df_serasa_filtrado'].columns:
                                    vencimento_map = resultados['df_serasa_filtrado'].set_index('AUTO_NORM')[col_venc_serasa].to_dict()
                                    df_export_local[coluna_vencimento] = df_final_work['AUTO_NORM'].map(vencimento_map)
                                    vencimento_adicionado = True

                            if not vencimento_adicionado:
                                if coluna_vencimento in resultados['df_serasa_filtrado'].columns and 'AUTO_NORM' in df_export_local.columns:
                                    vencimento_map = resultados['df_serasa_filtrado'].set_index('AUTO_NORM')[coluna_vencimento].to_dict()
                                    df_export_local[coluna_vencimento] = df_export_local['AUTO_NORM'].map(vencimento_map).fillna('')
                                else:
                                    df_export_local[coluna_vencimento] = pd.Series([''] * len(df_export_local), index=df_export_local.index)

                            if coluna_protocolo in resultados['df_serasa_filtrado'].columns:
                                if f"{coluna_protocolo}_serasa" in df_final_work.columns:
                                    df_export_local[coluna_protocolo] = df_final_work[f"{coluna_protocolo}_serasa"]
                                elif f"{coluna_protocolo}_divida" in df_final_work.columns:
                                    df_export_local[coluna_protocolo] = df_final_work[f"{coluna_protocolo}_divida"]

                            if 'Modais' in df_final_work.columns:
                                df_export_local['Modais'] = df_final_work['Modais']
                            elif 'Modais' in resultados['df_serasa_filtrado'].columns and 'AUTO_NORM' in df_export_local.columns:
                                modais_map = resultados['df_serasa_filtrado'].set_index('AUTO_NORM')['Modais'].to_dict()
                                df_export_local['Modais'] = df_export_local['AUTO_NORM'].map(modais_map).fillna('')

                            if coluna_nome_autuado in resultados['df_serasa_filtrado'].columns:
                                if f"{coluna_nome_autuado}_serasa" in df_final_work.columns:
                                    df_export_local[coluna_nome_autuado] = df_final_work[f"{coluna_nome_autuado}_serasa"]
                                elif 'AUTO_NORM' in df_export_local.columns and 'AUTO_NORM' in resultados['df_serasa_filtrado'].columns:
                                    nome_autuado_map = resultados['df_serasa_filtrado'].set_index('AUTO_NORM')[coluna_nome_autuado].to_dict()
                                    df_export_local[coluna_nome_autuado] = df_export_local['AUTO_NORM'].map(nome_autuado_map).fillna('')

                            if 'AUTO_NORM' in df_final_work.columns:
                                df_export_local['AUTO_NORM'] = df_final_work['AUTO_NORM']

                            if coluna_cpf_cnpj in df_export_local.columns:
                                if 'CPF_CNPJ_NORM' not in df_export_local.columns:
                                    df_export_local['CPF_CNPJ_NORM'] = df_export_local[coluna_cpf_cnpj].apply(normalizar_cpf_cnpj)
                            elif 'CPF_CNPJ_NORM' in resultados['df_serasa_filtrado'].columns and 'AUTO_NORM' in df_export_local.columns:
                                cpf_norm_map = resultados['df_serasa_filtrado'].set_index('AUTO_NORM')['CPF_CNPJ_NORM'].to_dict()
                                df_export_local['CPF_CNPJ_NORM'] = df_export_local['AUTO_NORM'].map(cpf_norm_map)

                            if 'Situação decadente' in df_final_work.columns:
                                df_export_local['Situação decadente'] = df_final_work['Situação decadente']
                        else:
                            df_export_local = resultados['df_serasa_filtrado'].copy()
                            if 'CPF_CNPJ_NORM' not in df_export_local.columns and coluna_cpf_cnpj in df_export_local.columns:
                                df_export_local['CPF_CNPJ_NORM'] = df_export_local[coluna_cpf_cnpj].apply(normalizar_cpf_cnpj)

                        return df_export_local

                    df_export = obter_dataframe_cacheado_sessao(cache_df_export_key, montar_df_export)
            
            if df_export.empty:
                st.warning("⚠️ Nenhum dado encontrado para exportar!")
            else:
                # Esse preparo concentra toda a regra antes de gerar os arquivos.
                def preparar_dados_exportacao(df_base, filtro_valor=None):
                    """Prepara a base final da exportação, com ou sem filtro de valor."""
                    # Primeiro tira quem não deve seguir para cobrança.
                    df_base = filtrar_autuados_cobraveis(df_base, coluna_nome_autuado)
                    if df_base is None or df_base.empty:
                        return None

                    # Nessa faixa, os decadentes saem antes da soma para continuar batendo com o uso do time.
                    if filtro_valor == 'entre_500_999' and 'Situação decadente' in df_base.columns:
                        situacao_decadente = df_base['Situação decadente'].fillna('').astype(str).str.strip()
                        df_base = df_base[situacao_decadente == ''].copy()
                        if df_base.empty:
                            return None

                    # Garante que o valor esteja numérico antes de filtrar.
                    if coluna_valor in df_base.columns:
                        if df_base[coluna_valor].dtype not in ['int64', 'float64']:
                            df_base[coluna_valor] = pd.to_numeric(df_base[coluna_valor], errors='coerce')
                    
                    # Valor vazio ou zero não entra na exportação.
                    if coluna_valor in df_base.columns:
                        df_base_sem_zero = df_base[
                            (df_base[coluna_valor].notna()) & 
                            (df_base[coluna_valor] > 0)
                        ].copy()
                    else:
                        df_base_sem_zero = df_base.copy()
                    
                    # Deduplica antes da soma para não inflar CPF/CNPJ com repetição.
                    df_base_preparada = deduplicar_por_protocolo_ou_auto(
                        df_base_sem_zero,
                        coluna_auto,
                        coluna_protocolo=coluna_protocolo,
                        coluna_vencimento=coluna_vencimento
                    )

                    # O filtro por faixa é feito em cima da soma por CPF/CNPJ.
                    if filtro_valor is not None and coluna_valor in df_base_preparada.columns:
                        # Se ainda não existir o documento normalizado, monta aqui.
                        if 'CPF_CNPJ_NORM' not in df_base_preparada.columns and coluna_cpf_cnpj in df_base_preparada.columns:
                            df_base_preparada['CPF_CNPJ_NORM'] = df_base_preparada[coluna_cpf_cnpj].apply(normalizar_cpf_cnpj)

                        if 'CPF_CNPJ_NORM' in df_base_preparada.columns:
                            # Documento vazio não ajuda no agrupamento.
                            df_base_para_agrupar = df_base_preparada[df_base_preparada['CPF_CNPJ_NORM'].notna()].copy()

                            if not df_base_para_agrupar.empty:
                                # Aqui sai a soma total de cada documento.
                                agrupado_export = df_base_para_agrupar.groupby('CPF_CNPJ_NORM').agg({
                                    coluna_valor: 'sum'
                                }).reset_index()
                                agrupado_export.columns = ['CPF_CNPJ_NORM', 'VALOR_TOTAL']

                                if filtro_valor == 'ate_999':
                                    # Até 999,99 fica neste grupo; 1000 já sobe para a faixa seguinte.
                                    cpf_ate_999 = set(
                                        agrupado_export[agrupado_export['VALOR_TOTAL'] < 1000]['CPF_CNPJ_NORM'].unique()
                                    )
                                    # Depois volta para os autos ligados a esses documentos.
                                    df_filtrado = df_base_preparada[df_base_preparada['CPF_CNPJ_NORM'].isin(cpf_ate_999)].copy()
                                elif filtro_valor == 'acima_1000':
                                    # A faixa de cima começa em 1000.
                                    cpf_acima_1000 = set(
                                        agrupado_export[agrupado_export['VALOR_TOTAL'] >= 1000]['CPF_CNPJ_NORM'].unique()
                                    )
                                    # Depois volta para os autos ligados a esses documentos.
                                    df_filtrado = df_base_preparada[df_base_preparada['CPF_CNPJ_NORM'].isin(cpf_acima_1000)].copy()
                                elif filtro_valor == 'entre_500_999':
                                    cpf_500_999 = set(
                                        agrupado_export[
                                            (agrupado_export['VALOR_TOTAL'] >= 500)
                                            & (agrupado_export['VALOR_TOTAL'] <= 999.99)
                                        ]['CPF_CNPJ_NORM'].unique()
                                    )
                                    df_filtrado = df_base_preparada[df_base_preparada['CPF_CNPJ_NORM'].isin(cpf_500_999)].copy()
                                else:
                                    df_filtrado = df_base_preparada.copy()
                            else:
                                # Se não houver documento aproveitável, cai para a leitura por linha.
                                if filtro_valor == 'ate_999':
                                    df_filtrado = df_base_preparada[df_base_preparada[coluna_valor] < 1000].copy()
                                elif filtro_valor == 'acima_1000':
                                    df_filtrado = df_base_preparada[df_base_preparada[coluna_valor] >= 1000].copy()
                                elif filtro_valor == 'entre_500_999':
                                    df_filtrado = df_base_preparada[
                                        (df_base_preparada[coluna_valor] >= 500)
                                        & (df_base_preparada[coluna_valor] <= 999.99)
                                    ].copy()
                                else:
                                    df_filtrado = df_base_preparada.copy()
                        else:
                            # Sem documento normalizado, sobra a leitura individual.
                            if filtro_valor == 'ate_999':
                                df_filtrado = df_base_preparada[df_base_preparada[coluna_valor] < 1000].copy()
                            elif filtro_valor == 'acima_1000':
                                df_filtrado = df_base_preparada[df_base_preparada[coluna_valor] >= 1000].copy()
                            elif filtro_valor == 'entre_500_999':
                                df_filtrado = df_base_preparada[
                                    (df_base_preparada[coluna_valor] >= 500)
                                    & (df_base_preparada[coluna_valor] <= 999.99)
                                ].copy()
                            else:
                                df_filtrado = df_base_preparada.copy()
                    else:
                        # Sem filtro de faixa, leva tudo que já passou pela limpeza.
                        df_filtrado = df_base_preparada.copy()
                    
                    if df_filtrado.empty:
                        return None
                    
                    # A partir daqui é só montar o layout final sem perder o índice original.
                    colunas_export = {
                        'Auto de Infração': df_filtrado[coluna_auto].fillna('').astype(str).str.strip() if coluna_auto in df_filtrado.columns else ''
                    }
                    
                    # O protocolo pode vir com nome puro ou com sufixo da base.
                    col_protocolo_encontrada = None
                    if coluna_protocolo in df_filtrado.columns:
                        col_protocolo_encontrada = coluna_protocolo
                    elif f"{coluna_protocolo}_serasa" in df_filtrado.columns:
                        col_protocolo_encontrada = f"{coluna_protocolo}_serasa"
                    elif f"{coluna_protocolo}_divida" in df_filtrado.columns:
                        col_protocolo_encontrada = f"{coluna_protocolo}_divida"
                    
                    if col_protocolo_encontrada:
                        colunas_export['Número de Protocolo'] = df_filtrado[col_protocolo_encontrada].fillna('').astype(str).str.strip()
                    
                    # O vencimento sempre entra, mesmo quando precisar ficar vazio.
                    col_vencimento_encontrada = resolver_coluna_vencimento(df_filtrado, coluna_vencimento)
                    
                    if col_vencimento_encontrada:
                        # Se der para converter, já sai no formato que o pessoal usa.
                        try:
                            vencimento_dt = pd.to_datetime(
                                df_filtrado[col_vencimento_encontrada],
                                errors='coerce',
                                dayfirst=True
                            )
                            colunas_export['Data de Vencimento'] = vencimento_dt.dt.strftime('%d/%m/%Y').fillna('')
                        except (ValueError, TypeError):
                            colunas_export['Data de Vencimento'] = df_filtrado[col_vencimento_encontrada].fillna('').astype(str).str.strip()
                    else:
                        # Se não aparecer, mantém a coluna vazia para preservar o layout.
                        colunas_export['Data de Vencimento'] = pd.Series([''] * len(df_filtrado), index=df_filtrado.index)
                    
                    # Modal entra só quando existir.
                    if 'Modais' in df_filtrado.columns:
                        colunas_export['Modais'] = df_filtrado['Modais'].fillna('').astype(str).str.strip()
                    
                    # CPF/CNPJ e valor são a base do arquivo.
                    colunas_export['CPF_CNPJ'] = df_filtrado[coluna_cpf_cnpj].fillna('').astype(str).str.strip() if coluna_cpf_cnpj in df_filtrado.columns else ''
                    colunas_export['Valor'] = df_filtrado[coluna_valor] if coluna_valor in df_filtrado.columns else None
                    # Se a análise já calculou decadência, ela segue junto.
                    if 'Situação decadente' in df_filtrado.columns:
                        colunas_export['Situação decadente'] = df_filtrado['Situação decadente'].fillna('').astype(str).str.strip()
                    
                    dados_exportacao = pd.DataFrame(colunas_export, index=df_filtrado.index)  # PRESERVAR ÍNDICE ORIGINAL
                    
                    # Reforça a coluna numérica para evitar surpresa no Excel.
                    if 'Valor' in dados_exportacao.columns:
                        dados_exportacao['Valor'] = pd.to_numeric(dados_exportacao['Valor'], errors='coerce')
                    
                    # Ordena agrupando quem tem mais autos, sem mexer nos valores individuais.
                    if coluna_cpf_cnpj in df_filtrado.columns and 'CPF_CNPJ_NORM' in df_filtrado.columns:
                        contagem_autos = df_filtrado.groupby('CPF_CNPJ_NORM').size().to_dict()
                        # O índice preservado ajuda a manter o vínculo certo na ordenação.
                        dados_exportacao['_QTD_AUTOS'] = df_filtrado['CPF_CNPJ_NORM'].map(contagem_autos).fillna(0)
                        # Primeiro vem quem tem mais autos; depois, agrupa pelo documento.
                        dados_exportacao['_CPF_CNPJ_NORM'] = df_filtrado['CPF_CNPJ_NORM']
                        dados_exportacao = dados_exportacao.sort_values(['_QTD_AUTOS', '_CPF_CNPJ_NORM'], ascending=[False, True])
                        dados_exportacao = dados_exportacao.drop(columns=['_QTD_AUTOS', '_CPF_CNPJ_NORM'])
                    
                    # No fim, formata o documento para leitura humana.
                    dados_exportacao['CPF_CNPJ'] = dados_exportacao['CPF_CNPJ'].apply(formatar_cpf_cnpj_brasileiro)
                    
                    # Limpa linhas vazias antes de gravar.
                    dados_exportacao = dados_exportacao.dropna(how='all')
                    
                    # Linha sem auto não faz sentido no arquivo final.
                    if 'Auto de Infração' in dados_exportacao.columns:
                        dados_exportacao = dados_exportacao[dados_exportacao['Auto de Infração'].str.strip() != '']
                    
                    # Faz uma última checagem para não deixar zero ou vazio escapar.
                    if 'Valor' in dados_exportacao.columns:
                        # Aqui fica só o que realmente vale para cobrança.
                        dados_exportacao = dados_exportacao[
                            (dados_exportacao['Valor'].notna()) & 
                            (dados_exportacao['Valor'] > 0)
                        ]
                    
                    return dados_exportacao
                
                # Identificadores da geração atual.
                data_extracao = datetime.now().strftime('%d/%m/%Y')
                data_arquivo = st.session_state.get('export_run_label', datetime.now().strftime('%d %m %Y %H:%M'))
                export_run_id = st.session_state.get('export_run_id', 'default')
                
                # Resume as colunas entregues em cada arquivo.
                def get_colunas_msg(dados_df):
                    """Monta a legenda curta com as colunas do arquivo."""
                    colunas = ["📊 Auto de Infração"]
                    if 'Número de Protocolo' in dados_df.columns:
                        colunas.append("Número de Protocolo")
                    if 'Situação Dívida' in dados_df.columns:
                        colunas.append("Situação Dívida")
                    if 'Situação Congelamento' in dados_df.columns:
                        colunas.append("Situação Congelamento")
                    if 'Nome Autuado' in dados_df.columns:
                        colunas.append("Nome Autuado")
                    if 'Classificação Autuado' in dados_df.columns:
                        colunas.append("Classificação Autuado")
                    if 'Motivo Classificação' in dados_df.columns:
                        colunas.append("Motivo Classificação")
                    if 'Termo Identificado' in dados_df.columns:
                        colunas.append("Termo Identificado")
                    if 'Data de Vencimento' in dados_df.columns:
                        colunas.append("Data de Vencimento")
                    if 'Data Infração' in dados_df.columns:
                        colunas.append("Data Infração")
                    if 'Data Pagamento' in dados_df.columns:
                        colunas.append("Data Pagamento")
                    if 'Modais' in dados_df.columns:
                        colunas.append("Modais")
                    colunas.append("CPF/CNPJ")
                    if 'Valor' in dados_df.columns:
                        colunas.append("Valor Individual")
                    if 'Valor (R$)' in dados_df.columns:
                        colunas.append("Valor (R$)")
                    if 'Situação decadente' in dados_df.columns:
                        colunas.append("Situação decadente")
                    return " | ".join(colunas)

                def get_resumo_classificacao_autuado(dados_df):
                    """Monta um resumo curto da classificação dos autuados."""
                    if dados_df is None or dados_df.empty or 'Classificação Autuado' not in dados_df.columns:
                        return None
                    partes = []
                    contagem = dados_df['Classificação Autuado'].value_counts()
                    if not contagem.empty:
                        partes.append(" | ".join([f"{k}: {v:,}" for k, v in contagem.items()]))
                    if 'Termo Identificado' in dados_df.columns:
                        termos = dados_df['Termo Identificado'].fillna('').astype(str).str.strip()
                        termos = termos[termos != '']
                        if not termos.empty:
                            top_termos = termos.value_counts().head(5)
                            partes.append("Top termos: " + " | ".join([f"{k}: {v:,}" for k, v in top_termos.items()]))
                    return " || ".join(partes) if partes else None

                def render_exportacao_excel(chave, titulo_gerar, label_download, nome_aba, nome_arquivo, help_download, producer, empty_warning, success_template, extra_caption_fn=None):
                    """Gera o arquivo só quando o usuário pedir para evitar trabalho desnecessário."""
                    state_key = f"export_payload::{export_run_id}::{chave}"
                    state_df_key = f"export_df::{export_run_id}::{chave}"
                    status_placeholder = st.empty()
                    if st.button(f"⚡ {titulo_gerar}", key=f"btn_generate::{export_run_id}::{chave}", use_container_width=True):
                        status_placeholder.info("Preparando arquivo Excel...")
                        try:
                            dados_df = st.session_state.get(state_df_key)
                            if dados_df is None:
                                dados_df = producer()
                                if dados_df is not None and not dados_df.empty:
                                    st.session_state[state_df_key] = dados_df
                            if dados_df is None or dados_df.empty:
                                st.session_state.pop(state_key, None)
                                st.session_state.pop(state_df_key, None)
                                status_placeholder.warning(empty_warning)
                            else:
                                excel_bytes = gerar_excel_formatado(dados_df, nome_aba, nome_arquivo)
                                if excel_bytes is None:
                                    st.session_state.pop(state_key, None)
                                    status_placeholder.warning(empty_warning)
                                else:
                                    payload = {
                                        'excel': excel_bytes,
                                        'qtd': len(dados_df),
                                        'colunas_msg': get_colunas_msg(dados_df),
                                    }
                                    if extra_caption_fn:
                                        payload['extra_caption'] = extra_caption_fn(dados_df)
                                    st.session_state[state_key] = payload
                                    status_placeholder.success("Arquivo pronto para download.")
                        except Exception as e:
                            st.session_state.pop(state_key, None)
                            status_placeholder.error(f"⚠️ Erro ao gerar arquivo: {str(e)}")

                    payload = st.session_state.get(state_key)
                    if payload:
                        st.download_button(
                            label=label_download,
                            data=payload['excel'],
                            file_name=nome_arquivo,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                            key=f"btn_download::{export_run_id}::{chave}",
                            help=help_download
                        )
                        st.success(success_template.format(qtd=payload['qtd']))
                        st.caption(payload['colunas_msg'])
                        if payload.get('extra_caption'):
                            st.caption(payload['extra_caption'])
                
                def preparar_dados_apenas_serasa(df_apenas_serasa):
                    """Prepara a exportação dos autos que ficaram só na SERASA."""
                    if df_apenas_serasa is None or df_apenas_serasa.empty:
                        return None
                    df = df_apenas_serasa.copy()
                    col_venc_ord = resolver_coluna_vencimento(df, coluna_vencimento)
                    if col_venc_ord and coluna_auto in df.columns:
                        df['_VENC_ORD'] = pd.to_datetime(df[col_venc_ord], errors='coerce', dayfirst=True)
                        df = df.sort_values(by='_VENC_ORD', ascending=False, na_position='last')
                        df = df.drop_duplicates(subset=[coluna_auto], keep='first').copy()
                        df = df.drop(columns=['_VENC_ORD'], errors='ignore')
                    elif coluna_auto in df.columns:
                        df = df.drop_duplicates(subset=[coluna_auto], keep='first').copy()
                    if 'CPF_CNPJ_NORM' not in df.columns and coluna_cpf_cnpj in df.columns:
                        df['CPF_CNPJ_NORM'] = df[coluna_cpf_cnpj].apply(normalizar_cpf_cnpj)

                    # Aqui saem autos zerados ou sem valor útil.
                    if coluna_valor in df.columns:
                        df['_VALOR_NUM'] = df[coluna_valor].apply(converter_valor_sql)
                    else:
                        df['_VALOR_NUM'] = None
                    df = df[(df['_VALOR_NUM'].notna()) & (df['_VALOR_NUM'] > 0)].copy()
                    if df.empty:
                        return None

                    # A decadência segue junto quando fizer sentido para o modal.
                    df['Situação decadente'] = calcular_situacao_decadente(df, coluna_modal=coluna_modal_serasa)

                    colunas_export = {}
                    colunas_export['Auto de Infração'] = df[coluna_auto].fillna('').astype(str).str.strip() if coluna_auto in df.columns else pd.Series([''] * len(df), index=df.index)
                    if coluna_protocolo in df.columns:
                        colunas_export['Número de Protocolo'] = df[coluna_protocolo].fillna('').astype(str).str.strip()
                    if 'Situação Dívida' in df.columns:
                        colunas_export['Situação Dívida'] = df['Situação Dívida'].fillna('').astype(str).str.strip()
                    if 'Situação Congelamento' in df.columns:
                        colunas_export['Situação Congelamento'] = df['Situação Congelamento'].fillna('').astype(str).str.strip()
                    col_venc = resolver_coluna_vencimento(df, coluna_vencimento)
                    if col_venc:
                        try:
                            venc_dt = pd.to_datetime(df[col_venc], errors='coerce', dayfirst=True)
                            colunas_export['Data de Vencimento'] = venc_dt.dt.strftime('%d/%m/%Y').fillna('')
                        except Exception:
                            colunas_export['Data de Vencimento'] = df[col_venc].fillna('').astype(str).str.strip()
                    else:
                        colunas_export['Data de Vencimento'] = pd.Series([''] * len(df), index=df.index)
                    col_infracao = _resolver_coluna_data(df, "Data Infração")
                    if col_infracao:
                        try:
                            infracao_dt = pd.to_datetime(df[col_infracao], errors='coerce', dayfirst=True)
                            colunas_export['Data Infração'] = infracao_dt.dt.strftime('%d/%m/%Y').fillna('')
                        except Exception:
                            colunas_export['Data Infração'] = df[col_infracao].fillna('').astype(str).str.strip()
                    else:
                        colunas_export['Data Infração'] = pd.Series([''] * len(df), index=df.index)
                    if 'Data Pagamento' in df.columns:
                        try:
                            pagamento_dt = pd.to_datetime(df['Data Pagamento'], errors='coerce', dayfirst=True)
                            colunas_export['Data Pagamento'] = pagamento_dt.dt.strftime('%d/%m/%Y').fillna('')
                        except Exception:
                            colunas_export['Data Pagamento'] = df['Data Pagamento'].fillna('').astype(str).str.strip()
                    if coluna_modal_serasa and coluna_modal_serasa in df.columns:
                        colunas_export['Modais'] = df[coluna_modal_serasa].fillna('').astype(str).str.strip()
                    colunas_export['CPF_CNPJ'] = df[coluna_cpf_cnpj].fillna('').astype(str).str.strip() if coluna_cpf_cnpj in df.columns else pd.Series([''] * len(df), index=df.index)
                    colunas_export['Valor (R$)'] = df['_VALOR_NUM'].apply(formatar_valor_br)
                    colunas_export['Situação decadente'] = df['Situação decadente'].fillna('').astype(str).str.strip()
                    ordem = ['Auto de Infração']
                    if 'Número de Protocolo' in colunas_export:
                        ordem.append('Número de Protocolo')
                    if 'Situação Dívida' in colunas_export:
                        ordem.append('Situação Dívida')
                    if 'Situação Congelamento' in colunas_export:
                        ordem.append('Situação Congelamento')
                    ordem.append('Data de Vencimento')
                    ordem.append('Data Infração')
                    if 'Data Pagamento' in colunas_export:
                        ordem.append('Data Pagamento')
                    if 'Modais' in colunas_export:
                        ordem.append('Modais')
                    ordem.append('CPF_CNPJ')
                    ordem.append('Valor (R$)')
                    ordem.append('Situação decadente')
                    dados = pd.DataFrame({k: colunas_export[k] for k in ordem if k in colunas_export}, index=df.index)
                    dados['CPF_CNPJ'] = dados['CPF_CNPJ'].apply(formatar_cpf_cnpj_brasileiro)
                    if 'CPF_CNPJ_NORM' in df.columns and not dados.empty:
                        dados['_CPF'] = dados.index.map(lambda i: df.loc[i, 'CPF_CNPJ_NORM'] if i in df.index else '')
                        contagem = dados.groupby('_CPF').size().to_dict()
                        dados['_QTD'] = dados['_CPF'].map(contagem).fillna(0)
                        dados = dados.sort_values(['_QTD', '_CPF'], ascending=[False, True]).drop(columns=['_QTD', '_CPF'], errors='ignore')
                    # Reindexa no fim para o Excel não ganhar linhas em branco.
                    dados = dados.replace('', np.nan).dropna(how='all').reset_index(drop=True)
                    return dados

                def preparar_dados_serasa_classificacao(df_serasa):
                    """Prepara a base SERASA inteira já com a classificação do autuado."""
                    if df_serasa is None or df_serasa.empty:
                        return None
                    df = df_serasa.copy()
                    config_classificacao = obter_config_classificacao_ativa()
                    if 'CPF_CNPJ_NORM' not in df.columns and coluna_cpf_cnpj in df.columns:
                        df['CPF_CNPJ_NORM'] = df[coluna_cpf_cnpj].apply(normalizar_cpf_cnpj)
                    coluna_valor_serasa = coluna_valor
                    if coluna_valor_serasa in df.columns:
                        df['_VALOR_NUM'] = df[coluna_valor_serasa].apply(converter_valor_sql)
                    else:
                        df['_VALOR_NUM'] = 0.0
                    df['_VALOR_NUM'] = df['_VALOR_NUM'].fillna(0)
                    colunas_export = {}
                    colunas_export['Auto de Infração'] = df[coluna_auto].fillna('').astype(str).str.strip() if coluna_auto in df.columns else pd.Series([''] * len(df), index=df.index)
                    if coluna_protocolo in df.columns:
                        colunas_export['Número de Protocolo'] = df[coluna_protocolo].fillna('').astype(str).str.strip()
                    if coluna_nome_autuado in df.columns:
                        colunas_export['Nome Autuado'] = df[coluna_nome_autuado].fillna('').astype(str).str.strip()
                        nomes_autuado = df[coluna_nome_autuado].fillna('').astype(str)
                        mapa_classificacao = {
                            nome: classificar_autuado_detalhado(nome, config=config_classificacao)
                            for nome in nomes_autuado.unique()
                        }
                        classificacoes = nomes_autuado.map(mapa_classificacao)
                        detalhes_df = pd.DataFrame(classificacoes.tolist(), index=df.index)
                        detalhes_df.columns = ['Classificação Autuado', 'Motivo Classificação', 'Termo Identificado']
                        colunas_export['Classificação Autuado'] = detalhes_df['Classificação Autuado']
                        colunas_export['Motivo Classificação'] = detalhes_df['Motivo Classificação']
                        colunas_export['Termo Identificado'] = detalhes_df['Termo Identificado']
                    else:
                        colunas_export['Nome Autuado'] = pd.Series([''] * len(df), index=df.index)
                        colunas_export['Classificação Autuado'] = pd.Series(['Pode cobrar'] * len(df), index=df.index)
                        colunas_export['Motivo Classificação'] = pd.Series(['Coluna Nome Autuado não encontrada'] * len(df), index=df.index)
                        colunas_export['Termo Identificado'] = pd.Series([''] * len(df), index=df.index)
                    if 'Situação Dívida' in df.columns:
                        colunas_export['Situação Dívida'] = df['Situação Dívida'].fillna('').astype(str).str.strip()
                    if 'Situação Congelamento' in df.columns:
                        colunas_export['Situação Congelamento'] = df['Situação Congelamento'].fillna('').astype(str).str.strip()
                    col_venc = resolver_coluna_vencimento(df, coluna_vencimento)
                    if col_venc:
                        try:
                            venc_dt = pd.to_datetime(df[col_venc], errors='coerce', dayfirst=True)
                            colunas_export['Data de Vencimento'] = venc_dt.dt.strftime('%d/%m/%Y').fillna('')
                        except Exception:
                            colunas_export['Data de Vencimento'] = df[col_venc].fillna('').astype(str).str.strip()
                    else:
                        colunas_export['Data de Vencimento'] = pd.Series([''] * len(df), index=df.index)
                    col_infracao = _resolver_coluna_data(df, "Data Infração")
                    if col_infracao:
                        try:
                            infracao_dt = pd.to_datetime(df[col_infracao], errors='coerce', dayfirst=True)
                            colunas_export['Data Infração'] = infracao_dt.dt.strftime('%d/%m/%Y').fillna('')
                        except Exception:
                            colunas_export['Data Infração'] = df[col_infracao].fillna('').astype(str).str.strip()
                    else:
                        colunas_export['Data Infração'] = pd.Series([''] * len(df), index=df.index)
                    if 'Data Pagamento' in df.columns:
                        try:
                            pagamento_dt = pd.to_datetime(df['Data Pagamento'], errors='coerce', dayfirst=True)
                            colunas_export['Data Pagamento'] = pagamento_dt.dt.strftime('%d/%m/%Y').fillna('')
                        except Exception:
                            colunas_export['Data Pagamento'] = df['Data Pagamento'].fillna('').astype(str).str.strip()
                    if coluna_modal_serasa and coluna_modal_serasa in df.columns:
                        colunas_export['Modais'] = df[coluna_modal_serasa].fillna('').astype(str).str.strip()
                    colunas_export['CPF_CNPJ'] = df[coluna_cpf_cnpj].fillna('').astype(str).str.strip() if coluna_cpf_cnpj in df.columns else pd.Series([''] * len(df), index=df.index)
                    colunas_export['Valor (R$)'] = df['_VALOR_NUM'].apply(formatar_valor_br)
                    ordem = ['Auto de Infração']
                    if 'Número de Protocolo' in colunas_export:
                        ordem.append('Número de Protocolo')
                    ordem.append('Nome Autuado')
                    ordem.append('Classificação Autuado')
                    ordem.append('Motivo Classificação')
                    ordem.append('Termo Identificado')
                    if 'Situação Dívida' in colunas_export:
                        ordem.append('Situação Dívida')
                    if 'Situação Congelamento' in colunas_export:
                        ordem.append('Situação Congelamento')
                    ordem.append('Data de Vencimento')
                    ordem.append('Data Infração')
                    if 'Data Pagamento' in colunas_export:
                        ordem.append('Data Pagamento')
                    if 'Modais' in colunas_export:
                        ordem.append('Modais')
                    ordem.append('CPF_CNPJ')
                    ordem.append('Valor (R$)')
                    dados = pd.DataFrame({k: colunas_export[k] for k in ordem if k in colunas_export}, index=df.index)
                    dados['CPF_CNPJ'] = dados['CPF_CNPJ'].apply(formatar_cpf_cnpj_brasileiro)
                    dados = dados.replace('', np.nan).dropna(how='all').reset_index(drop=True)
                    return dados

                # Mantém o botão principal em destaque e as faixas logo abaixo.
                st.markdown("##### 📥 Todos os Autos")
                st.markdown("**Sem filtro de valor**")
                render_exportacao_excel(
                    chave="todos_autos",
                    titulo_gerar="Gerar arquivo Todos os Autos",
                    label_download="📥 Download Todos os Autos",
                    nome_aba="Todos_Autos",
                    nome_arquivo=f"Base Comparada Todos {data_arquivo}.xlsx",
                    help_download="Arquivo Excel com TODOS os autos que estão em ambas as bases (sem filtro de valor)",
                    producer=lambda: preparar_dados_exportacao(df_export, filtro_valor=None),
                    empty_warning="⚠️ Nenhum auto encontrado",
                    success_template="✅ {qtd:,} autos (todos)"
                )

                col_exp1, col_exp2, col_exp3 = st.columns(3)

                # Faixa até 999,99.
                with col_exp1:
                    st.markdown("##### 📥 Autos até R$ 999,99")
                    render_exportacao_excel(
                        chave="autos_ate_999",
                        titulo_gerar="Gerar arquivo Autos ≤ R$ 999,99",
                        label_download="📥 Download Autos ≤ R$ 999,99",
                        nome_aba="Autos_Ate_999",
                        nome_arquivo=f"Base Comparada Ate 999 {data_arquivo}.xlsx",
                        help_download="Arquivo Excel com autos de valor até R$ 999,99",
                        producer=lambda: preparar_dados_exportacao(df_export, filtro_valor='ate_999'),
                        empty_warning="⚠️ Nenhum auto encontrado até R$ 999,99",
                        success_template="✅ {qtd:,} autos até R$ 999,99"
                    )

                # Faixa de 500 a 999,99 usando a soma por CPF/CNPJ.
                with col_exp2:
                    st.markdown("##### 📥 Autos R$ 500 a R$ 999,99")
                    render_exportacao_excel(
                        chave="autos_500_999",
                        titulo_gerar="Gerar arquivo Autos R$ 500–R$ 999,99",
                        label_download="📥 Download Autos R$ 500–R$ 999,99",
                        nome_aba="Autos_500_999",
                        nome_arquivo=f"Base Comparada 500 a 999 {data_arquivo}.xlsx",
                        help_download="Arquivo Excel com autos não decadentes cuja soma por CPF/CNPJ está entre R$ 500,00 e R$ 999,99",
                        producer=lambda: preparar_dados_exportacao(df_export, filtro_valor='entre_500_999'),
                        empty_warning="⚠️ Nenhum auto encontrado na faixa R$ 500–R$ 999,99",
                        success_template="✅ {qtd:,} autos na faixa R$ 500–R$ 999,99"
                    )

                # Faixa acima de 1.000.
                with col_exp3:
                    st.markdown("##### 📥 Autos acima de R$ 1.000,00")
                    render_exportacao_excel(
                        chave="autos_acima_1000",
                        titulo_gerar="Gerar arquivo Autos > R$ 1.000,00",
                        label_download="📥 Download Autos > R$ 1.000,00",
                        nome_aba="Autos_Acima_1000",
                        nome_arquivo=f"Base Comparada Acima 1000 {data_arquivo}.xlsx",
                        help_download="Arquivo Excel com autos de valor acima de R$ 1.000,00",
                        producer=lambda: preparar_dados_exportacao(df_export, filtro_valor='acima_1000'),
                        empty_warning="⚠️ Nenhum auto encontrado acima de R$ 1.000,00",
                        success_template="✅ {qtd:,} autos acima de R$ 1.000,00"
                    )
                
                # A mesma lógica, mas sem cortar por vencimento.
                st.markdown("---")
                st.markdown("#### 📥 Exportar Base Sem Filtro de Data")
                st.info("💡 Exporte TODOS os autos que estão em ambas as bases, **independente da data de vencimento**. Mantém todos os outros filtros (valores > 0, remoção de duplicados, etc.) e também **exclui autuados classificados como Órgão, Banco ou Leasing**, preservando exceções permitidas como **SAFRA**.")
                
                if 'df_final_geral' in resultados and not resultados['df_final_geral'].empty:
                    cache_df_export_geral_key = f"cache_df_export_geral::{st.session_state.get('_analise_hash', 'default')}"

                    def montar_df_export_geral():
                        df_final_geral_work = resultados['df_final_geral'].copy()
                        df_export_geral_local = pd.DataFrame()

                        if f"{coluna_auto}_serasa" in df_final_geral_work.columns:
                            df_export_geral_local[coluna_auto] = df_final_geral_work[f"{coluna_auto}_serasa"]
                        elif f"{coluna_auto}_divida" in df_final_geral_work.columns:
                            df_export_geral_local[coluna_auto] = df_final_geral_work[f"{coluna_auto}_divida"]
                        elif 'AUTO_NORM' in df_final_geral_work.columns:
                            autos_gerais = set(df_final_geral_work['AUTO_NORM'].unique())
                            df_export_geral_local = resultados['df_serasa_original'][resultados['df_serasa_original']['AUTO_NORM'].isin(autos_gerais)].copy()
                        else:
                            autos_gerais = set(df_final_geral_work['AUTO_NORM'].unique())
                            df_export_geral_local = resultados['df_serasa_original'][resultados['df_serasa_original']['AUTO_NORM'].isin(autos_gerais)].copy()

                        if coluna_auto in df_export_geral_local.columns and len(df_export_geral_local) == len(df_final_geral_work):
                            if 'Valor' in df_final_geral_work.columns:
                                df_export_geral_local[coluna_valor] = df_final_geral_work['Valor']

                            if coluna_cpf_cnpj in resultados['df_serasa_original'].columns:
                                if f"{coluna_cpf_cnpj}_serasa" in df_final_geral_work.columns:
                                    df_export_geral_local[coluna_cpf_cnpj] = df_final_geral_work[f"{coluna_cpf_cnpj}_serasa"]
                                elif f"{coluna_cpf_cnpj}_divida" in df_final_geral_work.columns:
                                    df_export_geral_local[coluna_cpf_cnpj] = df_final_geral_work[f"{coluna_cpf_cnpj}_divida"]

                                if coluna_cpf_cnpj in df_export_geral_local.columns:
                                    df_export_geral_local['CPF_CNPJ_NORM'] = df_export_geral_local[coluna_cpf_cnpj].apply(normalizar_cpf_cnpj)

                            vencimento_adicionado_geral = False
                            col_venc_geral = resolver_coluna_vencimento(df_final_geral_work, coluna_vencimento)
                            if col_venc_geral:
                                df_export_geral_local[coluna_vencimento] = df_final_geral_work[col_venc_geral]
                                vencimento_adicionado_geral = True

                            if not vencimento_adicionado_geral:
                                if coluna_vencimento in resultados['df_serasa_original'].columns and 'AUTO_NORM' in df_export_geral_local.columns:
                                    vencimento_map_geral = resultados['df_serasa_original'].set_index('AUTO_NORM')[coluna_vencimento].to_dict()
                                    df_export_geral_local[coluna_vencimento] = df_export_geral_local['AUTO_NORM'].map(vencimento_map_geral).fillna('')
                                else:
                                    df_export_geral_local[coluna_vencimento] = pd.Series([''] * len(df_export_geral_local), index=df_export_geral_local.index)

                            if coluna_protocolo in resultados['df_serasa_original'].columns:
                                if f"{coluna_protocolo}_serasa" in df_final_geral_work.columns:
                                    df_export_geral_local[coluna_protocolo] = df_final_geral_work[f"{coluna_protocolo}_serasa"]
                                elif f"{coluna_protocolo}_divida" in df_final_geral_work.columns:
                                    df_export_geral_local[coluna_protocolo] = df_final_geral_work[f"{coluna_protocolo}_divida"]

                            if 'Modais' in df_final_geral_work.columns:
                                df_export_geral_local['Modais'] = df_final_geral_work['Modais']
                            elif 'Modais' in resultados['df_serasa_original'].columns and 'AUTO_NORM' in df_export_geral_local.columns:
                                modais_map_geral = resultados['df_serasa_original'].set_index('AUTO_NORM')['Modais'].to_dict()
                                df_export_geral_local['Modais'] = df_export_geral_local['AUTO_NORM'].map(modais_map_geral).fillna('')

                            if coluna_nome_autuado in resultados['df_serasa_original'].columns:
                                if f"{coluna_nome_autuado}_serasa" in df_final_geral_work.columns:
                                    df_export_geral_local[coluna_nome_autuado] = df_final_geral_work[f"{coluna_nome_autuado}_serasa"]
                                elif 'AUTO_NORM' in df_export_geral_local.columns and 'AUTO_NORM' in resultados['df_serasa_original'].columns:
                                    nome_autuado_map_geral = resultados['df_serasa_original'].set_index('AUTO_NORM')[coluna_nome_autuado].to_dict()
                                    df_export_geral_local[coluna_nome_autuado] = df_export_geral_local['AUTO_NORM'].map(nome_autuado_map_geral).fillna('')

                            if 'AUTO_NORM' in df_final_geral_work.columns:
                                df_export_geral_local['AUTO_NORM'] = df_final_geral_work['AUTO_NORM']
                            if 'Situação decadente' in df_final_geral_work.columns:
                                df_export_geral_local['Situação decadente'] = df_final_geral_work['Situação decadente']
                        else:
                            df_export_geral_local = df_final_geral_work.copy()

                        return df_export_geral_local

                    df_export_geral = obter_dataframe_cacheado_sessao(cache_df_export_geral_key, montar_df_export_geral)
                    
                    render_exportacao_excel(
                        chave="todos_sem_filtro_data",
                        titulo_gerar="Gerar arquivo Todos os Autos (Sem Filtro de Data)",
                        label_download="📥 Download Todos os Autos (Sem Filtro de Data)",
                        nome_aba="Todos_Sem_Filtro_Data",
                        nome_arquivo=f"Base Comparada Sem Filtro Data {data_arquivo}.xlsx",
                        help_download="Arquivo Excel com TODOS os autos que estão em ambas as bases (sem filtro de data de vencimento)",
                        producer=lambda: preparar_dados_exportacao(df_export_geral, filtro_valor=None),
                        empty_warning="⚠️ Nenhum auto encontrado sem filtro de data",
                        success_template="✅ {qtd:,} autos únicos (sem filtro de data)",
                        extra_caption_fn=lambda dados_df: "📋 Todos os autos (sem filtro de data)"
                    )
                    
                    col_geral_exp1, col_geral_exp2, col_geral_exp3 = st.columns(3)
                    
                    # Faixa até 999,99 sem filtro de data.
                    with col_geral_exp1:
                        st.markdown("##### 📥 Autos até R$ 999,99 (Sem Filtro de Data)")
                        render_exportacao_excel(
                            chave="ate_999_sem_filtro_data",
                            titulo_gerar="Gerar arquivo Autos ≤ R$ 999,99 (Sem Filtro de Data)",
                            label_download="📥 Download Autos ≤ R$ 999,99 (Sem Filtro de Data)",
                            nome_aba="Autos_Ate_999_Sem_Data",
                            nome_arquivo=f"Base Comparada Ate 999 Sem Filtro Data {data_arquivo}.xlsx",
                            help_download="Arquivo Excel com autos de valor até R$ 999,99 (sem filtro de data de vencimento)",
                            producer=lambda: preparar_dados_exportacao(df_export_geral, filtro_valor='ate_999'),
                            empty_warning="⚠️ Nenhum auto encontrado até R$ 999,99 (sem filtro de data)",
                            success_template="✅ {qtd:,} autos até R$ 999,99 (sem filtro de data)"
                        )
                    
                    with col_geral_exp2:
                        st.markdown("##### 📥 Autos R$ 500–R$ 999,99 (Sem Filtro de Data)")
                        render_exportacao_excel(
                            chave="500_999_sem_filtro_data",
                            titulo_gerar="Gerar arquivo Autos R$ 500–R$ 999,99 (Sem Filtro de Data)",
                            label_download="📥 Download Autos R$ 500–R$ 999,99 (Sem Filtro de Data)",
                            nome_aba="Autos_500_999_Sem_Data",
                            nome_arquivo=f"Base Comparada 500 a 999 Sem Filtro Data {data_arquivo}.xlsx",
                            help_download="Arquivo Excel com autos não decadentes na faixa R$ 500–R$ 999,99 por soma de CPF/CNPJ (sem filtro de data)",
                            producer=lambda: preparar_dados_exportacao(df_export_geral, filtro_valor='entre_500_999'),
                            empty_warning="⚠️ Nenhum auto encontrado na faixa (sem filtro de data)",
                            success_template="✅ {qtd:,} autos na faixa R$ 500–R$ 999,99 (sem filtro de data)"
                        )
                    
                    # Faixa acima de 1.000 sem filtro de data.
                    with col_geral_exp3:
                        st.markdown("##### 📥 Autos acima de R$ 1.000,00 (Sem Filtro de Data)")
                        render_exportacao_excel(
                            chave="acima_1000_sem_filtro_data",
                            titulo_gerar="Gerar arquivo Autos > R$ 1.000,00 (Sem Filtro de Data)",
                            label_download="📥 Download Autos > R$ 1.000,00 (Sem Filtro de Data)",
                            nome_aba="Autos_Acima_1000_Sem_Data",
                            nome_arquivo=f"Base Comparada Acima 1000 Sem Filtro Data {data_arquivo}.xlsx",
                            help_download="Arquivo Excel com autos de valor acima de R$ 1.000,00 (sem filtro de data de vencimento)",
                            producer=lambda: preparar_dados_exportacao(df_export_geral, filtro_valor='acima_1000'),
                            empty_warning="⚠️ Nenhum auto encontrado acima de R$ 1.000,00 (sem filtro de data)",
                            success_template="✅ {qtd:,} autos acima de R$ 1.000,00 (sem filtro de data)"
                        )
                
                # Exportação do que sobrou só na SERASA.
                st.markdown("---")
                st.markdown("#### 📥 Exportar Apenas na SERASA")
                st.info("💡 Esta exportação contém apenas os autos que estão na base SERASA e **não** constam na Dívida Ativa. O sistema compara as duas bases e exclui da relação todos os autos que têm em comum; o arquivo traz somente o que restou da SERASA após essa exclusão. **Sem filtro de vencimento nem de valor:** todos os anos e todos os valores. **Uma linha por auto:** duplicados por Auto de Infração são removidos, mantendo a linha com data de vencimento mais recente.")
                st.caption("📐 **Como é feita a conta:** A comparação usa **autos únicos** (cada Auto de Infração conta uma vez). O \"Total Registros SERASA\" é o total de **linhas** da planilha; o número de **autos únicos** pode ser um pouco menor quando o mesmo auto aparece em mais de uma linha. Apenas na SERASA = autos únicos SERASA − autos únicos em ambas.")
                if 'df_autos_apenas_serasa' in resultados and not resultados['df_autos_apenas_serasa'].empty:
                    render_exportacao_excel(
                        chave="apenas_serasa",
                        titulo_gerar="Gerar arquivo Apenas na SERASA",
                        label_download="📥 Download Apenas na SERASA (sem autos da Dívida)",
                        nome_aba="Apenas_SERASA",
                        nome_arquivo=f"Base Apenas SERASA {data_arquivo}.xlsx",
                        help_download="Arquivo Excel com autos que estão somente na SERASA (excluídos os que constam na Dívida Ativa), com coluna Data Infração.",
                        producer=lambda: preparar_dados_apenas_serasa(resultados['df_autos_apenas_serasa']),
                        empty_warning="⚠️ Nenhum dado disponível para exportar (apenas na SERASA).",
                        success_template="✅ {qtd:,} autos apenas na SERASA"
                    )
                else:
                    st.warning("⚠️ Nenhum auto encontrado apenas na SERASA (todos os autos da SERASA constam na Dívida Ativa).")

                # Exportação da base inteira já classificada.
                st.markdown("---")
                st.markdown("#### 📥 Exportar base SERASA com classificação de autuado")
                st.info("💡 Exporta **toda a base SERASA** sem filtro (sem comparação com Dívida). Adiciona as colunas **Classificação Autuado**, **Motivo Classificação** e **Termo Identificado** para explicar por que o nome foi marcado como **Não pode cobrar - Órgão**, **Não pode cobrar - Banco**, **Não pode cobrar - Leasing** ou **Pode cobrar**.")
                if 'df_serasa_original' in resultados and resultados['df_serasa_original'] is not None and not resultados['df_serasa_original'].empty:
                    render_exportacao_excel(
                        chave="serasa_classificacao_autuado",
                        titulo_gerar="Gerar arquivo Base SERASA com classificação",
                        label_download="📥 Download Base SERASA com classificação de autuado",
                        nome_aba="SERASA_Classificacao",
                        nome_arquivo=f"Base SERASA Classificacao Autuados {data_arquivo}.xlsx",
                        help_download="Planilha com toda a base SERASA e coluna Classificação Autuado (Órgão, Banco, Leasing, Pode cobrar)",
                        producer=lambda: preparar_dados_serasa_classificacao(resultados['df_serasa_original']),
                        empty_warning="⚠️ Nenhum dado disponível para exportar.",
                        success_template="✅ {qtd:,} registros com classificação",
                        extra_caption_fn=get_resumo_classificacao_autuado
                    )
                else:
                    st.warning("⚠️ Base SERASA não disponível para exportação.")
                
                # Fechamento do bloco de exportação.
                st.markdown("---")
                st.caption(f"📅 Data de extração: {data_extracao}")
                st.caption(f"✅ **Dados comparados:** Estes são os autos que estão presentes em AMBAS as bases (SERASA e Dívida Ativa)")
                st.caption(f"💡 **Classificação por faixa de valor:** Baseada na SOMA dos valores por CPF/CNPJ. Cada arquivo contém valores individuais de cada auto, mas a classificação (até R$ 999,99 ou acima de R$ 1.000,00) é feita pela soma total do CPF/CNPJ.")
        
        # Lista curta só para conferência rápida na tela.
        if resultados['autos_em_ambas'] > 0:
            st.markdown("---")
            st.markdown("#### 📝 Lista de Autos de Infração em Ambas as Bases")
            autos_lista = sorted(list(resultados['autos_em_ambas_lista']))
            # Divide em colunas para a leitura ficar menos cansativa.
            num_cols = 4
            cols = st.columns(num_cols)
            for idx, auto in enumerate(autos_lista[:100]):  # Mantém um limite para a tela continuar leve.
                with cols[idx % num_cols]:
                    st.text(auto)
            if len(autos_lista) > 100:
                st.caption(f"Mostrando 100 de {len(autos_lista)} autos. Use a exportação para ver a lista completa.")
    
    with tab2:
        st.markdown("### 📈 Agrupamento por CPF/CNPJ")
        st.info("💡 CPF/CNPJ ordenados do **MAIOR** para o **MENOR** número de autos de infração")
        
        if not resultados['agrupado_serasa'].empty:
            # Resumo rápido do agrupamento.
            total_cpf = len(resultados['agrupado_serasa'])
            total_autos = resultados['agrupado_serasa']['QTD_AUTOS'].sum()
            maior_qtd = resultados['agrupado_serasa']['QTD_AUTOS'].max()
            menor_qtd = resultados['agrupado_serasa']['QTD_AUTOS'].min()
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Total CPF/CNPJ", f"{total_cpf:,}")
            with col2:
                st.metric("Total Autos", f"{total_autos:,}")
            with col3:
                st.metric("Maior Quantidade", f"{int(maior_qtd)} autos")
            with col4:
                st.metric("Menor Quantidade", f"{int(menor_qtd)} autos")
            
            st.markdown("---")
            
            # A tabela já sai no mesmo sentido usado pelo time no Excel.
            st.markdown("#### 📊 Lista Completa (Ordenada: Maior → Menor)")
            st.success(f"✅ Ordenação: Do CPF/CNPJ com **{int(maior_qtd)} autos** até o com **{int(menor_qtd)} autos**")
            
            # Mostra a posição no ranking.
            agrupado_display = resultados['agrupado_serasa'].copy()
            agrupado_display.insert(0, 'Posição', range(1, len(agrupado_display) + 1))
            
            # Valor total fica mais legível assim.
            if 'VALOR_TOTAL' in agrupado_display.columns:
                agrupado_display['VALOR_TOTAL_FORMAT'] = agrupado_display['VALOR_TOTAL'].apply(
                    lambda x: f"R$ {x:,.2f}" if pd.notna(x) else "N/A"
                )
            
            # Deixa a leitura mais próxima do jeito que o pessoal costuma conferir.
            st.markdown("**💡 Dica:** Use a barra de rolagem para navegar por todos os registros. Os dados estão ordenados do maior para o menor número de autos, como no Excel.")
            
            # Ajusta a altura sem deixar a página pesada demais.
            altura_tabela = min(800, max(400, total_cpf * 30))
            
            st.dataframe(
                agrupado_display, 
                use_container_width=True, 
                height=altura_tabela,
                hide_index=True
            )
            
            st.caption(f"📊 Exibindo {total_cpf:,} CPF/CNPJ ordenados do maior ({int(maior_qtd)} autos) para o menor ({int(menor_qtd)} autos)")
            st.info("💡 Para exportar os dados comparados, use a aba **Autos de Infração** onde está o download principal.")
        else:
            st.warning("⚠️ Não foi possível realizar o agrupamento. Verifique se a coluna de valor e CPF/CNPJ estão corretas.")
    
    with tab3:
        st.markdown("### Separação por Valores - Autos de Infração ANTT")
        
        # As exportações da aba ficam no topo para não obrigar a rolar tudo antes.
        st.markdown("#### 📥 Exportar Planilhas por Valor")
        st.info("💡 Baixe as planilhas separadas por valor em formato Excel (.xlsx) com formatação completa")
        
        col_exp1, col_exp2, col_exp3 = st.columns(3)
        
        def _preparar_df_tab3(dados_base):
            """Prepara DataFrame para exportação simplificada na Tab3."""
            if dados_base is None or dados_base.empty:
                return None
            df = dados_base.copy()
            if 'CPF_CNPJ_NORM' in df.columns:
                contagem = df.groupby('CPF_CNPJ_NORM').size()
                df['_QTD'] = df['CPF_CNPJ_NORM'].map(contagem).fillna(0)
                df = df.sort_values(['_QTD', 'CPF_CNPJ_NORM'], ascending=[False, True]).drop(columns=['_QTD'])
            if not (coluna_auto in df.columns and coluna_cpf_cnpj in df.columns and coluna_valor in df.columns):
                return None
            export = pd.DataFrame({'Auto de Infração': df[coluna_auto].fillna('').astype(str).str.strip()})
            if coluna_protocolo in df.columns:
                export['Número de Protocolo'] = df[coluna_protocolo].fillna('').astype(str).str.strip()
            if coluna_vencimento in df.columns:
                try:
                    vdt = pd.to_datetime(df[coluna_vencimento], errors='coerce', dayfirst=True)
                    export['Data de Vencimento'] = vdt.dt.strftime('%d/%m/%Y').fillna('')
                except (ValueError, TypeError):
                    export['Data de Vencimento'] = df[coluna_vencimento].fillna('').astype(str).str.strip()
            export['CPF_CNPJ'] = df[coluna_cpf_cnpj].fillna('').astype(str).str.strip().apply(formatar_cpf_cnpj_brasileiro)
            export['Valor'] = pd.to_numeric(df[coluna_valor], errors='coerce')
            return export

        with col_exp1:
            st.markdown("##### 📥 Autos ≤ R$ 999,99")
            if not resultados['serasa_abaixo_1000_ind'].empty:
                render_exportacao_excel(
                    chave="tab3_abaixo_999",
                    titulo_gerar="Gerar planilha Autos ≤ R$ 999,99",
                    label_download="📥 Download Autos ≤ R$ 999,99 (Excel)",
                    nome_aba="Autos_Ate_999",
                    nome_arquivo=f"Autos Ate 999 {datetime.now().strftime('%d %m %Y %H:%M')}.xlsx",
                    help_download="Planilha simplificada dos autos até R$ 999,99.",
                    producer=lambda: _preparar_df_tab3(resultados['serasa_abaixo_1000_ind']),
                    empty_warning="⚠️ Colunas necessárias não encontradas ou sem dados para exportar",
                    success_template="✅ {qtd:,} autos de infração"
                )
            else:
                st.warning("Nenhum auto encontrado nesta faixa de valor")
        
        with col_exp2:
            st.markdown("##### 📥 Autos R$ 500–R$ 999,99")
            st.caption("Autos correspondentes aos CPF/CNPJ cuja soma sem decadentes ficou na faixa.")
            if not resultados['serasa_500_999_acum_autos'].empty:
                render_exportacao_excel(
                    chave="tab3_500_999",
                    titulo_gerar="Gerar planilha Autos R$ 500–R$ 999,99",
                    label_download="📥 Download Autos R$ 500–R$ 999,99 (Excel)",
                    nome_aba="Autos_500_999",
                    nome_arquivo=f"Autos 500 a 999 {datetime.now().strftime('%d %m %Y %H:%M')}.xlsx",
                    help_download="Planilha simplificada dos autos correspondentes aos CPF/CNPJ da faixa R$ 500–R$ 999,99.",
                    producer=lambda: _preparar_df_tab3(resultados['serasa_500_999_acum_autos']),
                    empty_warning="⚠️ Colunas necessárias não encontradas ou sem dados para exportar",
                    success_template="✅ {qtd:,} autos de infração"
                )
            else:
                st.warning("Nenhum auto encontrado nesta faixa de valor")
        
        with col_exp3:
            st.markdown("##### 📥 Autos > R$ 1.000,00")
            if not resultados['serasa_acima_1000_ind'].empty:
                render_exportacao_excel(
                    chave="tab3_acima_1000",
                    titulo_gerar="Gerar planilha Autos > R$ 1.000,00",
                    label_download="📥 Download Autos > R$ 1.000,00 (Excel)",
                    nome_aba="Autos_Acima_1000",
                    nome_arquivo=f"Autos Acima 1000 {datetime.now().strftime('%d %m %Y %H:%M')}.xlsx",
                    help_download="Planilha simplificada dos autos acima de R$ 1.000,00.",
                    producer=lambda: _preparar_df_tab3(resultados['serasa_acima_1000_ind']),
                    empty_warning="⚠️ Colunas necessárias não encontradas ou sem dados para exportar",
                    success_template="✅ {qtd:,} autos de infração"
                )
            else:
                st.warning("Nenhum auto encontrado nesta faixa de valor")
        
        st.markdown("---")
        st.markdown("#### 📊 Visualização dos Dados")
        st.info("💡 Visualize os dados abaixo antes de exportar")
        col1, col2, col_mid = st.columns(3)
        
        with col1:
            st.markdown("**≤ R$ 999,99 (Individual)**")
            if not resultados['serasa_abaixo_1000_ind'].empty:
                st.dataframe(resultados['serasa_abaixo_1000_ind'], use_container_width=True, height=350)
                st.caption(f"📋 Total: {len(resultados['serasa_abaixo_1000_ind'])} autos de infração")
            else:
                st.warning("Nenhum auto encontrado nesta faixa de valor")
        
        with col_mid:
            st.markdown("**R$ 500–R$ 999,99 (Autos Correspondentes)**")
            st.caption("Autos nao decadentes dos CPF/CNPJ cuja soma ficou na faixa.")
            if not resultados['serasa_500_999_acum_autos'].empty:
                st.dataframe(resultados['serasa_500_999_acum_autos'], use_container_width=True, height=350)
                st.caption(f"📋 Total: {len(resultados['serasa_500_999_acum_autos'])} autos de infração")
            else:
                st.warning("Nenhum auto encontrado nesta faixa de valor")
        
        with col2:
            st.markdown("**> R$ 1.000,00 (Individual)**")
            if not resultados['serasa_acima_1000_ind'].empty:
                st.dataframe(resultados['serasa_acima_1000_ind'], use_container_width=True, height=350)
                st.caption(f"📋 Total: {len(resultados['serasa_acima_1000_ind'])} autos de infração")
            else:
                st.warning("Nenhum auto encontrado nesta faixa de valor")
        
        st.markdown("---")
        st.markdown("#### 📊 SERASA - Valores Acumulativos (por CPF/CNPJ)")
        st.info("💡 A soma de todos os autos do mesmo CPF/CNPJ é considerada")
        
        col3, col4, col5 = st.columns(3)
        
        with col3:
            st.markdown("**≤ R$ 999,99 (Acumulado)**")
            if not resultados['serasa_abaixo_1000_acum'].empty:
                st.markdown("**Resumo por CPF/CNPJ:**")
                st.dataframe(resultados['serasa_abaixo_1000_acum'], use_container_width=True, height=200)
                st.caption(f"📊 Total: {len(resultados['serasa_abaixo_1000_acum'])} CPF/CNPJ")
                
                st.markdown("**Autos de Infração Correspondentes:**")
                if not resultados['serasa_abaixo_1000_acum_autos'].empty:
                    st.dataframe(resultados['serasa_abaixo_1000_acum_autos'], use_container_width=True, height=300)
                    st.caption(f"📋 Total: {len(resultados['serasa_abaixo_1000_acum_autos'])} autos")
                else:
                    st.info("Nenhum auto encontrado")
            else:
                st.warning("Nenhum CPF/CNPJ encontrado nesta faixa de valor")
        
        with col4:
            st.markdown("**R$ 500–R$ 999,99 (Acumulado)**")
            st.caption("Soma por CPF/CNPJ considerando apenas autos nao decadentes.")
            if not resultados['serasa_500_999_acum'].empty:
                st.markdown("**Resumo por CPF/CNPJ:**")
                st.dataframe(resultados['serasa_500_999_acum'], use_container_width=True, height=200)
                st.caption(f"📊 Total: {len(resultados['serasa_500_999_acum'])} CPF/CNPJ")
                
                st.markdown("**Autos de Infração Correspondentes:**")
                if not resultados['serasa_500_999_acum_autos'].empty:
                    st.dataframe(resultados['serasa_500_999_acum_autos'], use_container_width=True, height=300)
                    st.caption(f"📋 Total: {len(resultados['serasa_500_999_acum_autos'])} autos")
                else:
                    st.info("Nenhum auto encontrado")
            else:
                st.warning("Nenhum CPF/CNPJ encontrado nesta faixa de valor")
        
        with col5:
            st.markdown("**> R$ 1.000,00 (Acumulado)**")
            if not resultados['serasa_acima_1000_acum'].empty:
                st.markdown("**Resumo por CPF/CNPJ:**")
                st.dataframe(resultados['serasa_acima_1000_acum'], use_container_width=True, height=200)
                st.caption(f"📊 Total: {len(resultados['serasa_acima_1000_acum'])} CPF/CNPJ")
                
                st.markdown("**Autos de Infração Correspondentes:**")
                if not resultados['serasa_acima_1000_acum_autos'].empty:
                    st.dataframe(resultados['serasa_acima_1000_acum_autos'], use_container_width=True, height=300)
                    st.caption(f"📋 Total: {len(resultados['serasa_acima_1000_acum_autos'])} autos")
                else:
                    st.info("Nenhum auto encontrado")
            else:
                st.warning("Nenhum CPF/CNPJ encontrado nesta faixa de valor")
    
    with tab4:
        st.markdown("### ⚠️ Registros com Divergências")
        st.info("💡 Análise baseada em **Autos de Infração** (principal) e CPF/CNPJ (adicional)")
        
        st.markdown("#### 🔑 Divergências por Auto de Infração (PRINCIPAL)")
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("##### 🔴 Autos Apenas na Base SERASA")
            st.warning(f"Total de {resultados['autos_apenas_serasa']:,} autos de infração que não constam na Dívida Ativa")
            if resultados['autos_apenas_serasa'] > 0 and not resultados['df_autos_apenas_serasa'].empty:
                st.dataframe(resultados['df_autos_apenas_serasa'], use_container_width=True, height=400)
                st.caption(f"📋 {len(resultados['df_autos_apenas_serasa'])} registros")
            elif resultados['autos_apenas_serasa'] == 0:
                st.success("✅ Todos os autos da SERASA possuem correspondência na Dívida Ativa!")
        
        with col2:
            st.markdown("##### 🔴 Autos Apenas na Base Dívida Ativa")
            st.warning(f"Total de {resultados['autos_apenas_divida']:,} autos de infração que não constam na SERASA")
            if resultados['autos_apenas_divida'] > 0 and not resultados['df_autos_apenas_divida'].empty:
                st.dataframe(resultados['df_autos_apenas_divida'], use_container_width=True, height=400)
                st.caption(f"📋 {len(resultados['df_autos_apenas_divida'])} registros")
            elif resultados['autos_apenas_divida'] == 0:
                st.success("✅ Todos os autos da Dívida Ativa possuem correspondência na SERASA!")
        
        st.markdown("---")
        st.markdown("#### 📋 Divergências por CPF/CNPJ (ADICIONAL)")
        col3, col4 = st.columns(2)
        
        with col3:
            st.markdown("##### 🔴 CPF/CNPJ Apenas na Base SERASA")
            st.info(f"Total de {resultados['cpf_apenas_serasa']:,} CPF/CNPJ que não constam na Dívida Ativa")
            if resultados['cpf_apenas_serasa'] > 0 and not resultados['df_cpf_apenas_serasa'].empty:
                st.dataframe(resultados['df_cpf_apenas_serasa'], use_container_width=True, height=300)
            elif resultados['cpf_apenas_serasa'] == 0:
                st.success("✅ Todos os CPF/CNPJ da SERASA possuem correspondência na Dívida Ativa!")
        
        with col4:
            st.markdown("##### 🔴 CPF/CNPJ Apenas na Base Dívida Ativa")
            st.info(f"Total de {resultados['cpf_apenas_divida']:,} CPF/CNPJ que não constam na SERASA")
            if resultados['cpf_apenas_divida'] > 0 and not resultados['df_cpf_apenas_divida'].empty:
                st.dataframe(resultados['df_cpf_apenas_divida'], use_container_width=True, height=300)
            elif resultados['cpf_apenas_divida'] == 0:
                st.success("✅ Todos os CPF/CNPJ da Dívida Ativa possuem correspondência na SERASA!")

    with tab5:
        st.markdown("### 📜 Histórico de Comparações")
        st.info("Registros das análises anteriores e download dos arquivos exportados.")

        if not st.session_state.get('_historico_salvo', True) and resultados:
            try:
                config_hist = {
                    "coluna_auto": st.session_state.get('coluna_auto', ''),
                    "coluna_cpf_cnpj": st.session_state.get('coluna_cpf_cnpj', ''),
                    "coluna_valor": st.session_state.get('coluna_valor', ''),
                    "coluna_vencimento": st.session_state.get('coluna_vencimento', ''),
                    "coluna_protocolo": st.session_state.get('coluna_protocolo', ''),
                    "coluna_modal_serasa": st.session_state.get('coluna_modal_serasa', ''),
                    "coluna_modal_divida": st.session_state.get('coluna_modal_divida', ''),
                    "ano_analise_inicial": st.session_state.get('ano_analise_inicial'),
                    "ano_analise_final": st.session_state.get('ano_analise_final'),
                }
                excel_hist = {}
                export_run_id = st.session_state.get('export_run_id', '')
                for key, val in st.session_state.items():
                    if key.startswith(f"export_payload::{export_run_id}::") and isinstance(val, dict) and val.get('excel'):
                        nome = key.split("::")[-1]
                        excel_hist[f"{nome}.xlsx"] = val['excel']
                save_run(
                    resultados,
                    st.session_state.get('nome_arquivo_serasa', 'SERASA'),
                    st.session_state.get('nome_arquivo_divida', 'Dívida Ativa'),
                    st.session_state.get('ano_analise', str(datetime.now().year)),
                    config_hist,
                    excel_dict=excel_hist if excel_hist else None,
                )
                st.session_state['_historico_salvo'] = True
            except (OSError, sqlite3.Error, TypeError, ValueError) as e:
                st.warning(f"Não foi possível salvar no histórico: {e}")

        try:
            runs = list_runs()
        except (sqlite3.Error, OSError):
            runs = []

        if not runs:
            st.caption("Nenhuma comparação registrada ainda.")
        else:
            for run in runs:
                titulo = f"{run['data_hora']}  |  {run['nome_base_serasa']} × {run['nome_base_divida']}  |  Período {run.get('ano_analise', '?')}"
                with st.expander(titulo, expanded=False):
                    c1, c2, c3 = st.columns(3)
                    c1.metric("Autos SERASA", f"{run['total_serasa']:,}")
                    c2.metric("Autos Dívida", f"{run['total_divida']:,}")
                    c3.metric("Em ambas", f"{run['autos_em_ambas']:,}")

                    detalhes = get_run(run['id'])
                    if detalhes and detalhes.get('arquivos'):
                        st.markdown("**Arquivos exportados:**")
                        for arq in detalhes['arquivos']:
                            if arq.suffix in ('.xlsx', '.xls'):
                                try:
                                    arq_bytes = arq.read_bytes()
                                    st.download_button(
                                        label=f"📥 {arq.name}",
                                        data=arq_bytes,
                                        file_name=arq.name,
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                        key=f"hist_{run['id']}_{arq.name}",
                                    )
                                except Exception:
                                    st.caption(f"Arquivo indisponível: {arq.name}")

                    if st.button("🗑️ Excluir", key=f"del_{run['id']}"):
                        if excluir_run(run['id']):
                            st.success("Registro excluído.")
                            st.rerun()
                        else:
                            st.error("Erro ao excluir.")

else:
    st.info("👆 Por favor, faça o upload das duas planilhas na barra lateral para iniciar a análise.")
    
    # Mostrar instruções
    with st.expander("ℹ️ Como usar o sistema"):
        st.markdown("""
        ### Instruções de Uso:
        
        1. **Upload de Arquivos**: Faça o upload das planilhas SERASA e Dívida Ativa na barra lateral
        2. **Configuração de Colunas**: Informe os nomes exatos das colunas:
           - ⚠️ **Coluna Auto de Infração** (OBRIGATÓRIA - ex: "Auto de Infração")
           - Coluna CPF/CNPJ (adicional)
           - Coluna Valor (adicional)
           - Coluna Vencimento (adicional)
        3. **Executar Análise**: Clique no botão "Executar Análise Completa"
        4. **Visualizar Resultados**: Explore as abas com os resultados detalhados
        5. **Exportar**: Baixe a planilha formatada na aba "Agrupamento por CPF/CNPJ"
        
        ### Funcionalidades:
        - ✅ **Análise principal por Auto de Infração** (chave de comparação)
        - ✅ Cruzamento automático entre bases baseado em autos
        - ✅ Filtro por ano de vencimento (configurável na sidebar)
        - ✅ Agrupamento por CPF/CNPJ (análise adicional)
        - ✅ Separação por valores SERASA (≤ R$ 1.000 e > R$ 1.000)
        - ✅ Visualização de autos de infração correspondentes
        - ✅ Análise individual e acumulativa
        - ✅ Identificação de divergências (autos e CPF/CNPJ)
        - ✅ Dashboard visual com gráficos
        - ✅ Exportação de resultados
        """)

