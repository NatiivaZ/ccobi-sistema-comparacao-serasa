"""Utilitários mais usados no comparador SERASA."""

import pandas as pd


def _limpar_documento(valor):
    """Tira a pontuação mais comum de CPF/CNPJ sem inventar regra nova."""
    return str(valor).replace(".", "").replace("-", "").replace("/", "").strip()


def _normalizar_modal(valor):
    """Padroniza o modal em caixa alta quando vier algo aproveitável."""
    if pd.notna(valor) and str(valor).strip() != "":
        return str(valor).strip().upper()
    return None


def normalizar_cpf_cnpj(valor):
    """Deixa o documento só com números para facilitar comparação."""
    if pd.isna(valor):
        return None
    valor_str = _limpar_documento(valor)
    return valor_str if valor_str else None


def normalizar_auto(valor):
    """Padroniza o auto para evitar diferença só por espaço ou caixa."""
    if pd.isna(valor):
        return None
    valor_str = str(valor).strip().upper()
    valor_str = ' '.join(valor_str.split())
    return valor_str if valor_str else None


def converter_valor_sql(valor):
    """Converte o valor seguindo a mesma ideia da consulta usada antes.

    O cuidado aqui é não tratar número que já veio como `int` ou `float`
    como se fosse texto, porque isso distorce valores com ponto decimal.
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
    """Formata CPF ou CNPJ para exibição no padrão brasileiro."""
    if pd.isna(valor) or valor == '' or valor is None:
        return ''
    valor_str = _limpar_documento(valor)
    if not valor_str or not valor_str.isdigit():
        return str(valor)
    if len(valor_str) == 11:
        return f"{valor_str[0:3]}.{valor_str[3:6]}.{valor_str[6:9]}-{valor_str[9:11]}"
    elif len(valor_str) == 14:
        return f"{valor_str[0:2]}.{valor_str[2:5]}.{valor_str[5:8]}/{valor_str[8:12]}-{valor_str[12:14]}"
    return str(valor)


def formatar_valor_br(valor):
    """Formata número no padrão brasileiro sem complicar os casos nulos."""
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
    """Junta os modais das duas bases em um texto só, já padronizado."""
    modal_serasa_str = _normalizar_modal(modal_serasa)
    modal_divida_str = _normalizar_modal(modal_divida)

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
    """Tenta achar a coluna de vencimento mesmo quando o nome varia um pouco."""
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
