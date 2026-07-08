"""Funções de apoio do app principal.

Aqui ficam as partes mais operacionais de carregamento, filtro e deduplicação
que o `app.py` usa o tempo todo.
"""

import pandas as pd
import streamlit as st

from utils import resolver_coluna_vencimento


def normalizar_intervalo_anos(ano_inicial, ano_final=None):
    """Ajusta o intervalo de anos para trabalhar sempre com valores válidos."""
    ano_inicial = int(ano_inicial)
    ano_final = ano_inicial if ano_final is None else int(ano_final)
    if ano_final < ano_inicial:
        raise ValueError("O ano final deve ser maior ou igual ao ano inicial.")
    return ano_inicial, ano_final


def formatar_periodo_analise(ano_inicial, ano_final=None):
    """Monta o texto que aparece na interface para o período escolhido."""
    ano_inicial, ano_final = normalizar_intervalo_anos(ano_inicial, ano_final)
    return str(ano_inicial) if ano_inicial == ano_final else f"{ano_inicial} a {ano_final}"


def descrever_periodo_vencimento(ano_inicial, ano_final=None):
    """Descreve o recorte de vencimento de um jeito mais legível para a tela."""
    ano_inicial, ano_final = normalizar_intervalo_anos(ano_inicial, ano_final)
    if ano_inicial == ano_final:
        return f"vencimento em {ano_inicial}: 01/01 a 31/12"
    return f"vencimento de {ano_inicial} a {ano_final}: 01/01/{ano_inicial} a 31/12/{ano_final}"


def filtrar_por_periodo_vencimento(df_base, col_vencimento, ano_inicial, ano_final=None):
    """Filtra pelo intervalo de vencimento.

    Primeiro tenta usar a coluna como data de verdade. Se isso não funcionar,
    cai para uma leitura mais simples do ano dentro do texto.
    """
    ano_inicial, ano_final = normalizar_intervalo_anos(ano_inicial, ano_final)
    if col_vencimento not in df_base.columns:
        return df_base.copy()

    data_limite = pd.Timestamp(f"{ano_inicial}-01-01")
    data_limite_fim = pd.Timestamp(f"{ano_final}-12-31")
    df_tmp = df_base.copy()

    try:
        if df_tmp[col_vencimento].dtype == "datetime64[ns]":
            vencimentos = df_tmp[col_vencimento]
        else:
            vencimentos = pd.to_datetime(
                df_tmp[col_vencimento],
                errors="coerce",
                dayfirst=True,
            )

        resultado = df_tmp[
            vencimentos.notna() &
            (vencimentos >= data_limite) &
            (vencimentos <= data_limite_fim)
        ].copy()

        if vencimentos.notna().any():
            return resultado
    except Exception:
        pass

    serie_texto = df_tmp[col_vencimento].astype(str)
    anos_extraidos = pd.to_numeric(
        serie_texto.str.extract(r"(\d{4})", expand=False),
        errors="coerce",
    )
    resultado = df_tmp[anos_extraidos.between(ano_inicial, ano_final, inclusive="both")].copy()
    return resultado[~serie_texto.isin(["NaT", "nan", "None", ""])].copy()


def deduplicar_por_protocolo_ou_auto(df_base, coluna_auto, coluna_protocolo=None, coluna_vencimento=None):
    """Remove duplicados usando protocolo quando existir.

    Quando o protocolo não vem na base, o fallback continua sendo o auto.
    """
    if df_base is None or df_base.empty:
        return df_base.copy()

    df = df_base.copy()
    col_duplicidade = None

    if coluna_protocolo:
        for candidato in (coluna_protocolo, f"{coluna_protocolo}_serasa", f"{coluna_protocolo}_divida"):
            if candidato in df.columns:
                col_duplicidade = candidato
                break

    if col_duplicidade is None:
        for candidato in (coluna_auto, f"{coluna_auto}_serasa", f"{coluna_auto}_divida", "AUTO_NORM"):
            if candidato in df.columns:
                col_duplicidade = candidato
                break

    if col_duplicidade is None:
        return df

    col_vencimento_ordenacao = resolver_coluna_vencimento(df, coluna_vencimento) if coluna_vencimento else None
    if col_vencimento_ordenacao and col_vencimento_ordenacao in df.columns:
        df_temp = df.copy()
        df_temp["_VENCIMENTO_ORD"] = pd.to_datetime(
            df_temp[col_vencimento_ordenacao],
            errors="coerce",
            dayfirst=True,
        )
        df_temp = df_temp.sort_values(by="_VENCIMENTO_ORD", ascending=False, na_position="last")
        return df_temp.drop_duplicates(subset=[col_duplicidade], keep="first").drop(columns=["_VENCIMENTO_ORD"])

    return df.drop_duplicates(subset=[col_duplicidade], keep="first").copy()


@st.cache_data
def carregar_dados(arquivo, nome_base):
    """Carrega a planilha e guarda o total original para a interface."""
    try:
        if arquivo.name.endswith(".csv"):
            df = pd.read_csv(arquivo, encoding="utf-8", sep=";", decimal=",", header=0)
        else:
            df = pd.read_excel(arquivo, header=0)

        total_original = len(df)
        df = df.dropna(how="all")
        df.attrs["total_original"] = total_original
        return df
    except Exception as exc:
        st.error(f"Erro ao carregar {nome_base}: {exc}")
        return None
