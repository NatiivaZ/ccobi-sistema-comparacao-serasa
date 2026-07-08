"""PORTFOLIO — regras de classificacao de autuados omitidas."""
from portfolio_omitted import omit


def classificar(*args, **kwargs):
    omit("classificacao de autuados")


def __getattr__(name):
    def _m(*a, **k):
        omit(f"classificacao.{name}")

    return _m
