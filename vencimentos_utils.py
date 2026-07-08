"""PORTFOLIO — utilitarios de vencimento omitidos."""
from portfolio_omitted import omit


def comparar_bases(*args, **kwargs):
    omit("comparar_bases por vencimento")


def __getattr__(name):
    def _m(*a, **k):
        omit(f"vencimentos_utils.{name}")

    return _m
