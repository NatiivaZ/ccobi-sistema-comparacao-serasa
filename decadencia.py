"""PORTFOLIO — regras de decadencia/prescricao omitidas."""
from portfolio_omitted import omit


def calcular(*args, **kwargs):
    omit("calculo de decadencia")


def __getattr__(name):
    def _m(*a, **k):
        omit(f"decadencia.{name}")

    return _m
