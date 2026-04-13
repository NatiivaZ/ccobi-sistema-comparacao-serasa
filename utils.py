"""Funções utilitárias de normalização e formatação para o Sistema de Comparação SERASA."""

import pandas as pd
import numpy as np
import re
import unicodedata


def normalizar_cpf_cnpj(valor):
    """Remove formatação de CPF/CNPJ, mantendo apenas dígitos."""
    if pd.isna(valor):
        return None
    valor_str = str(valor).replace('.', '').replace('-', '').replace('/', '').strip()
    return valor_str if valor_str else None


def normalizar_auto(valor):
    """Normaliza Auto de Infração para comparação (strip + upper + espaços)."""
    if pd.isna(valor):
        return None
    valor_str = str(valor).strip().upper()
    valor_str = ' '.join(valor_str.split())
    return valor_str if valor_str else None


def converter_valor_sql(valor):
    """
    Converte valor de texto para decimal seguindo a lógica do SQL:
    TRY_CONVERT(decimal(18,2), REPLACE(REPLACE(REPLACE(REPLACE([Valor], 'R$', ''), ' ', ''), '.', ''), ',', '.'))

    Se o valor já for numérico (int/float), retorna diretamente SEM
    manipulação de string, evitando o bug onde "241.11" -> remove '.' -> "24111".
    """
    if pd.isna(valor):
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    try:
        valor_str = str(valor).strip()
        valor_str = valor_str.replace('R$', '').replace('r$', '').replace('R$', '')
        valor_str = valor_str.replace(' ', '')
        if ',' in valor_str:
            valor_str = valor_str.replace('.', '')
            valor_str = valor_str.replace(',', '.')
        return float(valor_str)
    except (ValueError, TypeError):
        return None


def formatar_cpf_cnpj_brasileiro(valor):
    """Formata CPF (XXX.XXX.XXX-XX) ou CNPJ (XX.XXX.XXX/XXXX-XX)."""
    if pd.isna(valor) or valor == '' or valor is None:
        return ''
    valor_str = str(valor).replace('.', '').replace('-', '').replace('/', '').strip()
    if not valor_str or not valor_str.isdigit():
        return str(valor)
    if len(valor_str) == 11:
        return f"{valor_str[0:3]}.{valor_str[3:6]}.{valor_str[6:9]}-{valor_str[9:11]}"
    elif len(valor_str) == 14:
        return f"{valor_str[0:2]}.{valor_str[2:5]}.{valor_str[5:8]}/{valor_str[8:12]}-{valor_str[12:14]}"
    return str(valor)


def formatar_valor_br(valor):
    """Formata valor numérico no padrão brasileiro: R$ 1.234,56 ou R$ 0,00 para zero/nulo."""
    if valor is None or pd.isna(valor):
        return "R$ 0,00"
    try:
        v = float(valor)
        if v == 0:
            return "R$ 0,00"
        s = f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        return "R$ " + s
    except (TypeError, ValueError):
        return "R$ 0,00"


def normalizar_e_mesclar_modais(modal_serasa, modal_divida):
    """Normaliza e mescla modais de ambas as bases, padronizando para maiúsculas."""
    modal_serasa_str = str(modal_serasa).strip().upper() if pd.notna(modal_serasa) and str(modal_serasa).strip() != '' else None
    modal_divida_str = str(modal_divida).strip().upper() if pd.notna(modal_divida) and str(modal_divida).strip() != '' else None

    if modal_serasa_str and modal_divida_str:
        if modal_serasa_str == modal_divida_str:
            return modal_serasa_str
        else:
            return f"{modal_serasa_str} / {modal_divida_str}"
    elif modal_serasa_str:
        return modal_serasa_str
    elif modal_divida_str:
        return modal_divida_str
    else:
        return ''


def resolver_coluna_vencimento(df, coluna_vencimento):
    """Resolve a coluna de vencimento com variações de nome e sufixo (_serasa/_divida)."""
    if df is None or df.empty:
        return None
    colunas = list(df.columns)
    colunas_map = {c.lower(): c for c in colunas}

    candidatos = []
    if coluna_vencimento:
        candidatos.append(coluna_vencimento)
        if " do " in coluna_vencimento:
            candidatos.append(coluna_vencimento.replace(" do ", " de "))
        if " de " in coluna_vencimento:
            candidatos.append(coluna_vencimento.replace(" de ", " do "))
    candidatos += ["Data de Vencimento", "Data do Vencimento"]

    for cand in candidatos:
        for sufixo in ["_serasa", "_divida", ""]:
            chave = f"{cand}{sufixo}".lower()
            if chave in colunas_map:
                return colunas_map[chave]

    cols_venc = [c for c in colunas if "vencimento" in c.lower()]
    if cols_venc:
        for pref in ["serasa", "divida"]:
            for c in cols_venc:
                if pref in c.lower():
                    return c
        return cols_venc[0]

    return None
