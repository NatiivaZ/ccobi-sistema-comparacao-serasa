"""Núcleo da análise do comparador SERASA x Dívida Ativa."""

import pandas as pd
import streamlit as st
from streamlit import session_state as _ss

from app_helpers import (
    deduplicar_por_protocolo_ou_auto,
    filtrar_por_periodo_vencimento,
    normalizar_intervalo_anos,
)
from decadencia import calcular_situacao_decadente
from utils import (
    converter_valor_sql,
    normalizar_auto,
    normalizar_cpf_cnpj,
    normalizar_e_mesclar_modais,
)


def _construir_mapa_modal(df_base, col_auto_norm, col_modal):
    """Monta um mapa simples de auto para modal."""
    if df_base is None or df_base.empty or col_modal not in df_base.columns or col_auto_norm not in df_base.columns:
        return {}
    df_valido = df_base[[col_auto_norm, col_modal]].dropna(subset=[col_auto_norm])
    df_com_modal = df_valido[df_valido[col_modal].notna()]
    if not df_com_modal.empty:
        return df_com_modal.drop_duplicates(subset=[col_auto_norm], keep="first").set_index(col_auto_norm)[col_modal].to_dict()
    return df_valido.drop_duplicates(subset=[col_auto_norm], keep="first").set_index(col_auto_norm)[col_modal].to_dict()


def analisar_bases(
    df_serasa,
    df_divida,
    col_auto,
    col_cpf,
    col_valor,
    col_vencimento,
    coluna_protocolo,
    coluna_modal_serasa=None,
    coluna_modal_divida=None,
    ano_analise_inicial=2025,
    ano_analise_final=None,
):
    """Executa a análise principal e devolve o resultado no formato esperado pelo app."""
    resultados = {}
    ano_analise_inicial, ano_analise_final = normalizar_intervalo_anos(
        ano_analise_inicial,
        ano_analise_final,
    )

    if col_auto not in df_serasa.columns:
        st.error(f"⚠️ Coluna '{col_auto}' não encontrada na base SERASA. Esta coluna é OBRIGATÓRIA!")
        return None

    if col_auto not in df_divida.columns:
        st.error(f"⚠️ Coluna '{col_auto}' não encontrada na base Dívida Ativa. Esta coluna é OBRIGATÓRIA!")
        return None

    total_registros_serasa_original = df_serasa.attrs.get("total_original", len(df_serasa))
    total_registros_divida_original = df_divida.attrs.get("total_original", len(df_divida))

    # O auto normalizado é a chave de tudo daqui para frente.
    df_serasa["AUTO_NORM"] = df_serasa[col_auto].apply(normalizar_auto)
    df_divida["AUTO_NORM"] = df_divida[col_auto].apply(normalizar_auto)

    df_serasa_clean = df_serasa[df_serasa["AUTO_NORM"].notna()].copy()
    df_divida_clean = df_divida[df_divida["AUTO_NORM"].notna()].copy()

    # Neste ponto ficam só os autos que batem entre as duas bases.
    df_joined = pd.merge(
        df_serasa_clean,
        df_divida_clean,
        on="AUTO_NORM",
        how="inner",
        suffixes=("_serasa", "_divida"),
    )

    if df_joined.empty:
        st.warning("⚠️ Nenhum auto encontrado em ambas as bases!")
        return None

    col_valor_divida_cfg = _ss.get("coluna_valor_divida", col_valor)
    col_valor_serasa = f"{col_valor}_serasa" if f"{col_valor}_serasa" in df_joined.columns else col_valor
    col_valor_divida = (
        f"{col_valor_divida_cfg}_divida"
        if f"{col_valor_divida_cfg}_divida" in df_joined.columns
        else col_valor_divida_cfg
    )
    col_valor_usar = col_valor_serasa if col_valor_serasa in df_joined.columns else col_valor_divida

    # Dependendo da carga, o valor mais confiável pode vir de um lado ou do outro.
    if col_valor_usar in df_joined.columns:
        df_joined["Valor"] = df_joined[col_valor_usar].apply(converter_valor_sql)
    elif col_valor in df_joined.columns:
        df_joined["Valor"] = df_joined[col_valor].apply(converter_valor_sql)
    else:
        df_joined["Valor"] = None

    df_sem_zero = df_joined[
        (df_joined["Valor"].notna())
        & (df_joined["Valor"] != 0.00)
        & (df_joined["Valor"] > 0)
    ].copy()

    col_vencimento_para_ordenacao = None
    if f"{col_vencimento}_serasa" in df_sem_zero.columns:
        col_vencimento_para_ordenacao = f"{col_vencimento}_serasa"
    elif f"{col_vencimento}_divida" in df_sem_zero.columns:
        col_vencimento_para_ordenacao = f"{col_vencimento}_divida"
    elif col_vencimento in df_sem_zero.columns:
        col_vencimento_para_ordenacao = col_vencimento

    # Se houver protocolo, ele vira a referência da deduplicação; sem isso, fica o auto.
    if coluna_protocolo in df_sem_zero.columns:
        if col_vencimento_para_ordenacao and col_vencimento_para_ordenacao in df_sem_zero.columns:
            df_sem_zero_temp = df_sem_zero.copy()
            df_sem_zero_temp["_VENCIMENTO_ORD"] = pd.to_datetime(
                df_sem_zero_temp[col_vencimento_para_ordenacao],
                errors="coerce",
            )
            df_sem_zero_temp = df_sem_zero_temp.sort_values(
                by="_VENCIMENTO_ORD",
                ascending=False,
                na_position="last",
            )
            df_sem_duplicados = df_sem_zero_temp.drop_duplicates(
                subset=[coluna_protocolo],
                keep="first",
            ).copy()
            df_sem_duplicados = df_sem_duplicados.drop(columns=["_VENCIMENTO_ORD"])
        else:
            df_sem_duplicados = df_sem_zero.drop_duplicates(subset=[coluna_protocolo], keep="first").copy()
    else:
        if col_vencimento_para_ordenacao and col_vencimento_para_ordenacao in df_sem_zero.columns:
            df_sem_zero_temp = df_sem_zero.copy()
            df_sem_zero_temp["_VENCIMENTO_ORD"] = pd.to_datetime(
                df_sem_zero_temp[col_vencimento_para_ordenacao],
                errors="coerce",
            )
            df_sem_zero_temp = df_sem_zero_temp.sort_values(
                by="_VENCIMENTO_ORD",
                ascending=False,
                na_position="last",
            )
            df_sem_duplicados = df_sem_zero_temp.drop_duplicates(subset=["AUTO_NORM"], keep="first").copy()
            df_sem_duplicados = df_sem_duplicados.drop(columns=["_VENCIMENTO_ORD"])
        else:
            df_sem_duplicados = df_sem_zero.drop_duplicates(subset=["AUTO_NORM"], keep="first").copy()

    col_vencimento_usar = f"{col_vencimento}_serasa" if f"{col_vencimento}_serasa" in df_sem_duplicados.columns else col_vencimento
    if col_vencimento_usar not in df_sem_duplicados.columns:
        col_vencimento_usar = f"{col_vencimento}_divida" if f"{col_vencimento}_divida" in df_sem_duplicados.columns else col_vencimento

    if col_vencimento_usar in df_sem_duplicados.columns:
        df_final = filtrar_por_periodo_vencimento(
            df_sem_duplicados,
            col_vencimento_usar,
            ano_analise_inicial,
            ano_analise_final,
        )
    else:
        df_final = df_sem_duplicados.copy()

    autos_em_ambas = set(df_final["AUTO_NORM"].unique())
    df_serasa_filtrado = df_serasa_clean[df_serasa_clean["AUTO_NORM"].isin(autos_em_ambas)].copy()
    df_divida_filtrado = df_divida_clean[df_divida_clean["AUTO_NORM"].isin(autos_em_ambas)].copy()

    autos_serasa = set(df_serasa_clean["AUTO_NORM"].unique())
    autos_divida = set(df_divida_clean["AUTO_NORM"].unique())
    autos_apenas_serasa = autos_serasa - autos_divida
    autos_apenas_divida = autos_divida - autos_serasa

    autos_serasa_filtrado = set(df_serasa_filtrado["AUTO_NORM"].unique())
    autos_divida_filtrado = set(df_divida_filtrado["AUTO_NORM"].unique())
    if autos_serasa_filtrado != autos_divida_filtrado:
        autos_em_ambas_validados = autos_serasa_filtrado.intersection(autos_divida_filtrado)
        df_serasa_filtrado = df_serasa_filtrado[df_serasa_filtrado["AUTO_NORM"].isin(autos_em_ambas_validados)].copy()
        df_divida_filtrado = df_divida_filtrado[df_divida_filtrado["AUTO_NORM"].isin(autos_em_ambas_validados)].copy()
        autos_em_ambas = autos_em_ambas_validados

    if col_cpf in df_serasa_filtrado.columns:
        df_serasa_filtrado["CPF_CNPJ_NORM"] = df_serasa_filtrado[col_cpf].apply(normalizar_cpf_cnpj)
    else:
        df_serasa_filtrado["CPF_CNPJ_NORM"] = None

    if col_cpf in df_divida_filtrado.columns:
        df_divida_filtrado["CPF_CNPJ_NORM"] = df_divida_filtrado[col_cpf].apply(normalizar_cpf_cnpj)
    else:
        df_divida_filtrado["CPF_CNPJ_NORM"] = None

    modal_serasa_map = {}
    modal_divida_map = {}
    # Já deixa os modais combinados prontos para a interface e para os exports.
    if coluna_modal_serasa and coluna_modal_divida:
        modal_serasa_map = _construir_mapa_modal(
            df_serasa_clean if coluna_modal_serasa in df_serasa_clean.columns else df_serasa_filtrado,
            "AUTO_NORM",
            coluna_modal_serasa,
        )
        modal_divida_map = _construir_mapa_modal(
            df_divida_clean if coluna_modal_divida in df_divida_clean.columns else df_divida_filtrado,
            "AUTO_NORM",
            coluna_modal_divida,
        )

        if "AUTO_NORM" in df_serasa_filtrado.columns:
            df_serasa_filtrado["Modais"] = df_serasa_filtrado["AUTO_NORM"].apply(
                lambda auto: normalizar_e_mesclar_modais(
                    modal_serasa_map.get(auto),
                    modal_divida_map.get(auto),
                )
            )

        if "AUTO_NORM" in df_divida_filtrado.columns:
            df_divida_filtrado["Modais"] = df_divida_filtrado["AUTO_NORM"].apply(
                lambda auto: normalizar_e_mesclar_modais(
                    modal_serasa_map.get(auto),
                    modal_divida_map.get(auto),
                )
            )

        if "AUTO_NORM" in df_final.columns:
            df_final["Modais"] = df_final["AUTO_NORM"].apply(
                lambda auto: normalizar_e_mesclar_modais(
                    modal_serasa_map.get(auto),
                    modal_divida_map.get(auto),
                )
            )

    cpf_serasa = set(df_serasa_filtrado[df_serasa_filtrado["CPF_CNPJ_NORM"].notna()]["CPF_CNPJ_NORM"].unique())
    cpf_divida = set(df_divida_filtrado[df_divida_filtrado["CPF_CNPJ_NORM"].notna()]["CPF_CNPJ_NORM"].unique())
    cpf_em_ambas = cpf_serasa.intersection(cpf_divida)
    cpf_apenas_serasa = cpf_serasa - cpf_divida
    cpf_apenas_divida = cpf_divida - cpf_serasa

    df_serasa_total_ano = filtrar_por_periodo_vencimento(
        df_serasa_clean,
        col_vencimento,
        ano_analise_inicial,
        ano_analise_final,
    )
    df_divida_total_ano = filtrar_por_periodo_vencimento(
        df_divida_clean,
        col_vencimento,
        ano_analise_inicial,
        ano_analise_final,
    )

    df_serasa_2025 = df_serasa_filtrado.copy()
    df_divida_2025 = df_divida_filtrado.copy()

    if col_vencimento in df_serasa_2025.columns and df_serasa_2025[col_vencimento].dtype != "datetime64[ns]":
        df_serasa_2025[col_vencimento] = pd.to_datetime(df_serasa_2025[col_vencimento], errors="coerce")
    if col_vencimento in df_divida_2025.columns and df_divida_2025[col_vencimento].dtype != "datetime64[ns]":
        df_divida_2025[col_vencimento] = pd.to_datetime(df_divida_2025[col_vencimento], errors="coerce")

    if col_valor in df_serasa_2025.columns:
        try:
            df_serasa_2025[col_valor] = pd.to_numeric(df_serasa_2025[col_valor], errors="coerce")
            df_serasa_2025_valido = df_serasa_2025[df_serasa_2025[col_valor].notna()].copy()

            if not df_serasa_2025_valido.empty:
                agrupado_serasa = df_serasa_2025_valido.groupby("CPF_CNPJ_NORM").agg({
                    col_valor: ["sum", "count"],
                }).reset_index()
                agrupado_serasa.columns = ["CPF_CNPJ_NORM", "VALOR_TOTAL", "QTD_AUTOS"]
                agrupado_serasa["VALOR_TOTAL"] = pd.to_numeric(agrupado_serasa["VALOR_TOTAL"], errors="coerce").fillna(0)
                agrupado_serasa["CPF_CNPJ"] = agrupado_serasa["CPF_CNPJ_NORM"]
                agrupado_serasa = agrupado_serasa.sort_values("QTD_AUTOS", ascending=False)
            else:
                agrupado_serasa = pd.DataFrame()
        except Exception:
            agrupado_serasa = pd.DataFrame()
    else:
        agrupado_serasa = pd.DataFrame()

    if col_valor in df_serasa_2025.columns:
        try:
            if df_serasa_2025[col_valor].dtype not in ["int64", "float64"]:
                df_serasa_2025[col_valor] = pd.to_numeric(df_serasa_2025[col_valor], errors="coerce")

            df_serasa_2025_sem_zero = df_serasa_2025[
                (df_serasa_2025[col_valor].notna()) &
                (df_serasa_2025[col_valor] > 0)
            ].copy()

            col_venc_dedup = col_vencimento if col_vencimento in df_serasa_2025_sem_zero.columns else None
            df_serasa_2025_valido = deduplicar_por_protocolo_ou_auto(
                df_serasa_2025_sem_zero,
                col_auto,
                coluna_protocolo=coluna_protocolo,
                coluna_vencimento=col_venc_dedup,
            )

            serasa_abaixo_1000_ind = df_serasa_2025_valido[df_serasa_2025_valido[col_valor] <= 999.99].copy()
            serasa_500_999_ind = df_serasa_2025_valido[
                (df_serasa_2025_valido[col_valor] >= 500) & (df_serasa_2025_valido[col_valor] <= 999.99)
            ].copy()
            serasa_acima_1000_ind = df_serasa_2025_valido[df_serasa_2025_valido[col_valor] >= 1000].copy()

            if not df_serasa_2025_valido.empty and col_valor in df_serasa_2025_valido.columns and "CPF_CNPJ_NORM" in df_serasa_2025_valido.columns:
                agrupado_serasa = df_serasa_2025_valido.groupby("CPF_CNPJ_NORM").agg({
                    col_valor: ["sum", "count"],
                }).reset_index()
                agrupado_serasa.columns = ["CPF_CNPJ_NORM", "VALOR_TOTAL", "QTD_AUTOS"]
                agrupado_serasa["VALOR_TOTAL"] = pd.to_numeric(agrupado_serasa["VALOR_TOTAL"], errors="coerce").fillna(0)
                agrupado_serasa["CPF_CNPJ"] = agrupado_serasa["CPF_CNPJ_NORM"]
                agrupado_serasa = agrupado_serasa.sort_values("QTD_AUTOS", ascending=False)

            if not df_serasa_2025_valido.empty:
                df_serasa_500_999_base = df_serasa_2025_valido.copy()
                df_serasa_500_999_base["Situação decadente"] = calcular_situacao_decadente(df_serasa_500_999_base)
                situacao_decadente_500_999 = df_serasa_500_999_base["Situação decadente"].fillna("").astype(str).str.strip()
                df_serasa_500_999_sem_decad = df_serasa_500_999_base[situacao_decadente_500_999 == ""].copy()
            else:
                df_serasa_500_999_sem_decad = pd.DataFrame()

            serasa_500_999_ind = pd.DataFrame()
            serasa_500_999_acum = pd.DataFrame()
            serasa_500_999_acum_autos = pd.DataFrame()
            if not df_serasa_500_999_sem_decad.empty:
                serasa_500_999_ind = df_serasa_500_999_sem_decad[
                    (df_serasa_500_999_sem_decad[col_valor] >= 500) & (df_serasa_500_999_sem_decad[col_valor] <= 999.99)
                ].copy()

                if "CPF_CNPJ_NORM" in df_serasa_500_999_sem_decad.columns:
                    agrupado_serasa_500_999 = df_serasa_500_999_sem_decad.groupby("CPF_CNPJ_NORM").agg({
                        col_valor: "sum",
                    }).reset_index()
                    agrupado_serasa_500_999.columns = ["CPF_CNPJ_NORM", "VALOR_TOTAL"]
                    agrupado_serasa_500_999["VALOR_TOTAL"] = pd.to_numeric(
                        agrupado_serasa_500_999["VALOR_TOTAL"],
                        errors="coerce",
                    ).fillna(0)
                    agrupado_serasa_500_999["CPF_CNPJ"] = agrupado_serasa_500_999["CPF_CNPJ_NORM"]

                    serasa_500_999_acum = agrupado_serasa_500_999[
                        (agrupado_serasa_500_999["VALOR_TOTAL"] >= 500)
                        & (agrupado_serasa_500_999["VALOR_TOTAL"] <= 999.99)
                    ].copy()
                    cpf_500_999_acum = set(serasa_500_999_acum["CPF_CNPJ_NORM"].unique())
                    serasa_500_999_acum_autos = df_serasa_500_999_sem_decad[
                        df_serasa_500_999_sem_decad["CPF_CNPJ_NORM"].isin(cpf_500_999_acum)
                    ].copy()

            if not agrupado_serasa.empty:
                serasa_abaixo_1000_acum = agrupado_serasa[agrupado_serasa["VALOR_TOTAL"] < 1000].copy()
                cpf_abaixo_1000_acum = set(serasa_abaixo_1000_acum["CPF_CNPJ_NORM"].unique())
                serasa_abaixo_1000_acum_autos = df_serasa_2025_valido[
                    df_serasa_2025_valido["CPF_CNPJ_NORM"].isin(cpf_abaixo_1000_acum)
                ].copy()
            else:
                serasa_abaixo_1000_acum = pd.DataFrame()
                serasa_abaixo_1000_acum_autos = pd.DataFrame()

            if not agrupado_serasa.empty:
                serasa_acima_1000_acum = agrupado_serasa[agrupado_serasa["VALOR_TOTAL"] >= 1000].copy()
                cpf_acima_1000_acum = set(serasa_acima_1000_acum["CPF_CNPJ_NORM"].unique())
                serasa_acima_1000_acum_autos = df_serasa_2025_valido[
                    df_serasa_2025_valido["CPF_CNPJ_NORM"].isin(cpf_acima_1000_acum)
                ].copy()
            else:
                serasa_acima_1000_acum = pd.DataFrame()
                serasa_acima_1000_acum_autos = pd.DataFrame()
        except Exception:
            serasa_abaixo_1000_ind = pd.DataFrame()
            serasa_500_999_ind = pd.DataFrame()
            serasa_acima_1000_ind = pd.DataFrame()
            serasa_abaixo_1000_acum = pd.DataFrame()
            serasa_acima_1000_acum = pd.DataFrame()
            serasa_500_999_acum = pd.DataFrame()
            serasa_abaixo_1000_acum_autos = pd.DataFrame()
            serasa_acima_1000_acum_autos = pd.DataFrame()
            serasa_500_999_acum_autos = pd.DataFrame()
    else:
        serasa_abaixo_1000_ind = pd.DataFrame()
        serasa_500_999_ind = pd.DataFrame()
        serasa_acima_1000_ind = pd.DataFrame()
        serasa_abaixo_1000_acum = pd.DataFrame()
        serasa_acima_1000_acum = pd.DataFrame()
        serasa_500_999_acum = pd.DataFrame()
        serasa_abaixo_1000_acum_autos = pd.DataFrame()
        serasa_acima_1000_acum_autos = pd.DataFrame()
        serasa_500_999_acum_autos = pd.DataFrame()

    df_autos_apenas_serasa = df_serasa_clean[df_serasa_clean["AUTO_NORM"].isin(autos_apenas_serasa)].copy()
    df_autos_apenas_divida = df_divida_clean[df_divida_clean["AUTO_NORM"].isin(autos_apenas_divida)].copy()
    df_cpf_apenas_serasa = (
        df_serasa_filtrado[
            df_serasa_filtrado["CPF_CNPJ_NORM"].notna()
            & df_serasa_filtrado["CPF_CNPJ_NORM"].isin(cpf_apenas_serasa)
        ].copy()
        if col_cpf in df_serasa_filtrado.columns
        else pd.DataFrame()
    )
    df_cpf_apenas_divida = (
        df_divida_filtrado[
            df_divida_filtrado["CPF_CNPJ_NORM"].notna()
            & df_divida_filtrado["CPF_CNPJ_NORM"].isin(cpf_apenas_divida)
        ].copy()
        if col_cpf in df_divida_filtrado.columns
        else pd.DataFrame()
    )

    try:
        qtd_autos_ate_999 = len(serasa_abaixo_1000_acum_autos) if not serasa_abaixo_1000_acum_autos.empty else 0
    except Exception:
        qtd_autos_ate_999 = 0

    try:
        qtd_autos_acima_1000 = len(serasa_acima_1000_acum_autos) if not serasa_acima_1000_acum_autos.empty else 0
    except Exception:
        qtd_autos_acima_1000 = 0

    try:
        qtd_autos_500_999 = len(serasa_500_999_acum_autos) if not serasa_500_999_acum_autos.empty else 0
    except Exception:
        qtd_autos_500_999 = 0

    autos_em_ambas_linhas = len(df_final)
    df_final_geral = df_sem_duplicados.copy()
    autos_em_ambas_geral = set(df_final_geral["AUTO_NORM"].unique())
    autos_em_ambas_geral_linhas = len(df_final_geral)

    if coluna_modal_serasa and coluna_modal_divida and "AUTO_NORM" in df_final_geral.columns:
        df_final_geral["Modais"] = df_final_geral["AUTO_NORM"].apply(
            lambda auto: normalizar_e_mesclar_modais(
                modal_serasa_map.get(auto),
                modal_divida_map.get(auto),
            )
        )

    col_cpf_serasa_geral = f"{col_cpf}_serasa" if f"{col_cpf}_serasa" in df_final_geral.columns else col_cpf
    if col_cpf_serasa_geral in df_final_geral.columns:
        df_final_geral["CPF_CNPJ_NORM"] = df_final_geral[col_cpf_serasa_geral].apply(normalizar_cpf_cnpj)
    elif col_cpf in df_final_geral.columns:
        df_final_geral["CPF_CNPJ_NORM"] = df_final_geral[col_cpf].apply(normalizar_cpf_cnpj)
    else:
        df_final_geral["CPF_CNPJ_NORM"] = None

    df_final["Situação decadente"] = calcular_situacao_decadente(df_final)
    df_final_geral["Situação decadente"] = calcular_situacao_decadente(df_final_geral)

    resultados = {
        "df_serasa_original": df_serasa,
        "df_divida_original": df_divida,
        "df_serasa_filtrado": df_serasa_2025,
        "df_divida_filtrado": df_divida_2025,
        "df_final_sql": df_final,
        "df_final_geral": df_final_geral,
        "df_serasa_total_ano": df_serasa_total_ano,
        "df_divida_total_ano": df_divida_total_ano,
        "autos_em_ambas": autos_em_ambas_linhas,
        "autos_em_ambas_unicos": len(autos_em_ambas),
        "autos_em_ambas_geral": autos_em_ambas_geral_linhas,
        "autos_em_ambas_geral_unicos": len(autos_em_ambas_geral),
        "autos_apenas_serasa": len(autos_apenas_serasa),
        "autos_apenas_divida": len(autos_apenas_divida),
        "total_registros_serasa": total_registros_serasa_original,
        "total_registros_divida": total_registros_divida_original,
        "total_autos_serasa": len(autos_serasa),
        "total_autos_divida": len(autos_divida),
        "df_autos_apenas_serasa": df_autos_apenas_serasa,
        "df_autos_apenas_divida": df_autos_apenas_divida,
        "cpf_em_ambas": len(cpf_em_ambas),
        "cpf_apenas_serasa": len(cpf_apenas_serasa),
        "cpf_apenas_divida": len(cpf_apenas_divida),
        "total_cpf_serasa": len(cpf_serasa),
        "total_cpf_divida": len(cpf_divida),
        "df_cpf_apenas_serasa": df_cpf_apenas_serasa,
        "df_cpf_apenas_divida": df_cpf_apenas_divida,
        "agrupado_serasa": agrupado_serasa,
        "serasa_abaixo_1000_ind": serasa_abaixo_1000_ind,
        "serasa_500_999_ind": serasa_500_999_ind,
        "serasa_acima_1000_ind": serasa_acima_1000_ind,
        "serasa_abaixo_1000_acum": serasa_abaixo_1000_acum,
        "serasa_acima_1000_acum": serasa_acima_1000_acum,
        "serasa_500_999_acum": serasa_500_999_acum,
        "serasa_abaixo_1000_acum_autos": serasa_abaixo_1000_acum_autos,
        "serasa_acima_1000_acum_autos": serasa_acima_1000_acum_autos,
        "serasa_500_999_acum_autos": serasa_500_999_acum_autos,
        "qtd_autos_ate_999": qtd_autos_ate_999,
        "qtd_autos_500_999": qtd_autos_500_999,
        "qtd_autos_acima_1000": qtd_autos_acima_1000,
        "autos_em_ambas_lista": autos_em_ambas,
        "autos_apenas_serasa_lista": autos_apenas_serasa,
        "autos_apenas_divida_lista": autos_apenas_divida,
    }

    return resultados
