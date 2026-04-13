"""Cálculo de decadência de autos de infração para o Sistema de Comparação SERASA."""

import pandas as pd
import unicodedata
from datetime import date, timedelta

PRAZO_DIAS_AUTUACAO = 31
PRAZO_DIAS_MULTA = 181
AJUSTE_DIAS_AUTUACAO = 4
AJUSTE_DIAS_MULTA = 4

_MODAIS_COM_DECADENCIA = ['EXCESSO DE PESO', 'EVASAO DE PEDAGIO']


def _easter_year(ano):
    """Retorna a data do domingo de Páscoa no ano dado (algoritmo de Butcher/Meeus)."""
    a = ano % 19
    b = ano // 100
    c = ano % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    mes = (h + l - 7 * m + 114) // 31
    dia = ((h + l - 7 * m + 114) % 31) + 1
    return date(ano, mes, dia)


def _feriados_nacionais_brasil(ano):
    """Retorna set de datas dos feriados nacionais do Brasil no ano."""
    pascoa = _easter_year(ano)
    feriados = {
        date(ano, 1, 1), date(ano, 4, 21), date(ano, 5, 1),
        date(ano, 9, 7), date(ano, 10, 12), date(ano, 11, 2),
        date(ano, 11, 15), date(ano, 12, 25),
    }
    feriados.add(pascoa - timedelta(days=2))
    feriados.add(pascoa - timedelta(days=48))
    feriados.add(pascoa - timedelta(days=47))
    feriados.add(pascoa + timedelta(days=60))
    return feriados


def _resolver_coluna_data(df, nome_base):
    """Retorna o nome da coluna de data no DataFrame (com ou sem sufixo _serasa/_divida)."""
    if nome_base in df.columns:
        return nome_base
    if f"{nome_base}_serasa" in df.columns:
        return f"{nome_base}_serasa"
    if f"{nome_base}_divida" in df.columns:
        return f"{nome_base}_divida"
    return None


def _modal_tem_decadencia(modal_str):
    """Retorna True se o modal se enquadra nos tipos com cálculo de decadência."""
    if not modal_str or (isinstance(modal_str, float)):
        return False
    m = unicodedata.normalize('NFD', str(modal_str).upper().strip())
    m_sem_acento = ''.join(c for c in m if unicodedata.category(c) != 'Mn')
    return any(kw in m_sem_acento for kw in _MODAIS_COM_DECADENCIA)


def calcular_situacao_decadente(df, coluna_modal=None):
    """
    Calcula a coluna [Situação decadente] com base nas datas da SERASA.
    Retorna Series: '' | 'Decadente autuação' | 'Decadente multa' | 'Decadente autuação e multa'.
    """
    col_infracao = _resolver_coluna_data(df, "Data Infração")
    col_notif_autuacao = _resolver_coluna_data(df, "Data Primeira Notificação Autuação")
    col_notif_multa = _resolver_coluna_data(df, "Data Primeira Notificação Multa")
    if not col_infracao or (not col_notif_autuacao and not col_notif_multa):
        return pd.Series([''] * len(df), index=df.index)

    data_infracao = pd.to_datetime(df[col_infracao], errors='coerce', dayfirst=True)
    data_notif_autuacao = pd.to_datetime(df[col_notif_autuacao], errors='coerce', dayfirst=True) if col_notif_autuacao else pd.Series([pd.NaT] * len(df), index=df.index)
    data_notif_multa = pd.to_datetime(df[col_notif_multa], errors='coerce', dayfirst=True) if col_notif_multa else pd.Series([pd.NaT] * len(df), index=df.index)

    series_todas_datas = pd.concat([
        data_infracao.dropna(),
        data_notif_autuacao.dropna(),
        data_notif_multa.dropna(),
    ]) if (data_notif_autuacao is not None and data_notif_multa is not None) else data_infracao.dropna()

    anos = set()
    for v in series_todas_datas:
        try:
            t = pd.Timestamp(v)
            if pd.notna(t):
                anos.add(t.year)
        except Exception:
            pass
    feriados = set()
    for a in anos:
        feriados |= _feriados_nacionais_brasil(a)
        feriados |= _feriados_nacionais_brasil(a + 1)

    def primeiro_dia_util_a_partir(dt):
        if pd.isna(dt):
            return pd.NaT
        d = dt.date()
        while True:
            if d.weekday() < 5 and d not in feriados:
                return pd.Timestamp(d)
            d += timedelta(days=1)

    inicio_prazo = data_infracao.apply(primeiro_dia_util_a_partir)

    data_notif_autuacao_ajustada = data_notif_autuacao + pd.to_timedelta(AJUSTE_DIAS_AUTUACAO, unit="D")
    data_notif_multa_ajustada = data_notif_multa + pd.to_timedelta(AJUSTE_DIAS_MULTA, unit="D")

    dias_corridos_autuacao = (data_notif_autuacao_ajustada - inicio_prazo).dt.days
    dias_corridos_multa = (data_notif_multa_ajustada - inicio_prazo).dt.days

    decadente_autuacao = dias_corridos_autuacao > PRAZO_DIAS_AUTUACAO
    decadente_multa = dias_corridos_multa > PRAZO_DIAS_MULTA

    data_corte_multa = pd.Timestamp('2021-04-11')
    mask_permite_multa = data_infracao >= data_corte_multa
    decadente_multa = decadente_multa & mask_permite_multa.fillna(False)

    situacao = pd.Series([''] * len(df), index=df.index, dtype=object)
    both_ = decadente_autuacao & decadente_multa
    only_aut = decadente_autuacao & ~decadente_multa
    only_multa = ~decadente_autuacao & decadente_multa
    situacao = situacao.mask(both_, 'Decadente autuação e multa').mask(only_aut, 'Decadente autuação').mask(only_multa, 'Decadente multa')

    col_modal_efetivo = coluna_modal if (coluna_modal and coluna_modal in df.columns) else ('Modais' if 'Modais' in df.columns else None)
    if col_modal_efetivo:
        mask_modal = df[col_modal_efetivo].apply(_modal_tem_decadencia)
        situacao = situacao.where(mask_modal, '')

    return situacao
