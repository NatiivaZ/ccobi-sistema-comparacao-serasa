"""Regras de classificação do nome do autuado.

Aqui fica a separação entre quem pode cobrar normalmente e os casos como órgão,
banco ou leasing.
"""

import pandas as pd
import re
import unicodedata
import json
from pathlib import Path

_session_config_getter = None

def set_session_config_getter(fn):
    """Guarda uma função para ler a configuração atual direto do session_state."""
    global _session_config_getter
    _session_config_getter = fn

CONFIG_CLASSIFICACAO_PATH = Path(__file__).with_name("config_classificacao_autuados.json")

DEFAULT_CLASSIFICACAO_CONFIG = {
    "extras_orgao": [],
    "extras_banco": [],
    "extras_leasing": [],
    "excecoes_pode_cobrar": [],
}

EXCECOES_FIXAS_PODE_COBRAR = ["SAFRA"]


def carregar_config_classificacao():
    """Carrega o arquivo de configuração da classificação, se ele existir."""
    try:
        if CONFIG_CLASSIFICACAO_PATH.exists():
            data = json.loads(CONFIG_CLASSIFICACAO_PATH.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                config = DEFAULT_CLASSIFICACAO_CONFIG.copy()
                for chave in config.keys():
                    valor = data.get(chave, [])
                    config[chave] = valor if isinstance(valor, list) else []
                return config
    except (json.JSONDecodeError, OSError, UnicodeDecodeError):
        pass
    return DEFAULT_CLASSIFICACAO_CONFIG.copy()


def salvar_config_classificacao(config):
    """Salva a configuração já limpa para a próxima execução do app."""
    config_limpo = {}
    for chave, valor in DEFAULT_CLASSIFICACAO_CONFIG.items():
        itens = config.get(chave, valor)
        if not isinstance(itens, list):
            itens = []
        itens = [str(v).strip() for v in itens if str(v).strip()]
        config_limpo[chave] = itens
    CONFIG_CLASSIFICACAO_PATH.write_text(
        json.dumps(config_limpo, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )


def parse_lista_multilinha(texto):
    """Transforma o texto da interface em lista simples de termos."""
    if not texto:
        return []
    return [linha.strip() for linha in str(texto).splitlines() if linha.strip()]


def _normalizar_texto_para_busca(texto):
    """Normaliza o texto para comparação sem depender de acento ou pontuação."""
    if not texto or (isinstance(texto, float) and pd.isna(texto)):
        return ""
    s = unicodedata.normalize("NFD", str(texto).strip().upper())
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = re.sub(r"[^A-Z0-9]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return f" {s} " if s else ""


def _contem_expressao(texto_norm, expressao):
    """Procura a expressão inteira para evitar falso positivo por pedaço de palavra."""
    expr_norm = _normalizar_texto_para_busca(expressao)
    return bool(expr_norm) and expr_norm in texto_norm


def _contem_alguma_expressao(texto_norm, expressoes):
    return any(_contem_expressao(texto_norm, expr) for expr in expressoes)


_EXPRESSOES_LEASING = [
    "LEASING", "ARRENDAMENTO MERCANTIL", "ARREND MERCANTIL",
    "ARREND. MERCANTIL", "LOCACAO FINANCEIRA", "LOCACAO MERCANTIL",
]

_EXPRESSOES_ORGAO = [
    "CORPO DE BOMBEIROS", "BOMBEIRO MILITAR", "POLICIA MILITAR",
    "POLICIA CIVIL", "POLICIA FEDERAL", "POLICIA RODOVIARIA FEDERAL",
    "RODOVIARIA FEDERAL", "PRF", "EXERCITO BRASILEIRO", "EXERCITO",
    "MARINHA DO BRASIL", "MARINHA", "FORCA AEREA BRASILEIRA", "AERONAUTICA",
    "PREFEITURA MUNICIPAL", "MUNICIPIO DE", "GOVERNO DO ESTADO",
    "GOVERNO FEDERAL", "RECEITA FEDERAL", "MINISTERIO PUBLICO",
    "DEFENSORIA PUBLICA", "TRIBUNAL DE", "TRIBUNAL REGIONAL",
    "CAMARA MUNICIPAL", "SECRETARIA DE ESTADO", "SECRETARIA MUNICIPAL",
    "MINISTERIO DA", "MINISTERIO DO", "AUTARQUIA", "FUNDACAO PUBLICA",
    "INSS", "DATAPREV", "IBAMA", "INCRA", "DETRAN",
    "DEPARTAMENTO ESTADUAL DE TRANSITO", "CORREIOS",
    "EMPRESA BRASILEIRA DE CORREIOS E TELEGRAFOS",
    "AGENCIA NACIONAL DE AGUAS", "AGENCIA NACIONAL DE TRANSPORTES TERRESTRES",
]

_MARCADORES_BANCO = [
    "BANCO", "BANK", "FINANCEIRA", "INSTITUICAO FINANCEIRA",
    "SERVICOS FINANCEIROS", "FINANCIAL SERVICES",
    "CREDITO FINANCIAMENTO E INVESTIMENTO", "COOPERATIVA DE CREDITO",
]

_NOMES_BANCOS_ESPECIFICOS = [
    "BANCO DO BRASIL", "CAIXA ECONOMICA FEDERAL", "BANCO DA AMAZONIA",
    "BANCO DO NORDESTE", "BANCO DE BRASILIA", "BNDES", "BANRISUL",
    "BANESTES", "BANPARA", "BANESE", "BRADESCO", "SANTANDER",
    "ITAU UNIBANCO", "ITAU", "BTG PACTUAL", "SAFRA", "CITIBANK",
    "C6 BANK", "NUBANK", "BANCO INTER", "PAGBANK", "BANCO ORIGINAL",
    "BANCO PAN", "DAYCOVAL", "SICREDI", "SICOOB", "CRESOL", "AGIBANK",
    "OMNI BANCO", "PARANA BANCO", "MERCANTIL DO BRASIL", "BANCO SOFISA",
    "BANCO FIBRA", "BANCO GENIAL", "BANCO MODAL", "BANCO BS2", "DIGIO",
    "OURIBANK", "TRIBANCO", "ABC BRASIL", "ABN AMRO", "BNP PARIBAS",
    "BANK OF AMERICA", "SCOTIABANK", "DEUTSCHE BANK", "GOLDMAN SACHS",
    "JP MORGAN", "MORGAN STANLEY", "HSBC", "RABOBANK", "BMG", "BV",
]


def obter_config_classificacao_ativa():
    """Usa a configuração da sessão quando existir; senão, volta para o arquivo."""
    if _session_config_getter:
        config = _session_config_getter()
        if isinstance(config, dict):
            return config
    return carregar_config_classificacao()


def _lista_regras_classificacao(config):
    """Junta as regras padrão com os termos extras cadastrados na interface."""
    config = config or DEFAULT_CLASSIFICACAO_CONFIG
    return {
        "orgao": _EXPRESSOES_ORGAO + list(config.get("extras_orgao", [])),
        "banco": _MARCADORES_BANCO + _NOMES_BANCOS_ESPECIFICOS + list(config.get("extras_banco", [])),
        "leasing": _EXPRESSOES_LEASING + list(config.get("extras_leasing", [])),
        "excecoes": EXCECOES_FIXAS_PODE_COBRAR + list(config.get("excecoes_pode_cobrar", [])),
    }


def classificar_autuado_detalhado(nome, config=None):
    """Classifica um nome e devolve também o motivo e o termo que bateu."""
    texto_norm = _normalizar_texto_para_busca(nome)
    if not texto_norm:
        return "Pode cobrar", "Nome vazio ou não informado", ""

    regras = _lista_regras_classificacao(config)

    for termo in regras["excecoes"]:
        if _contem_expressao(texto_norm, termo):
            return "Pode cobrar", "Exceção cadastrada", termo

    for termo in regras["leasing"]:
        if _contem_expressao(texto_norm, termo):
            return "Não pode cobrar - Leasing", "Correspondência com regra de leasing", termo

    for termo in regras["orgao"]:
        if _contem_expressao(texto_norm, termo):
            return "Não pode cobrar - Órgão", "Correspondência com regra de órgão público", termo

    for termo in regras["banco"]:
        if _contem_expressao(texto_norm, termo):
            return "Não pode cobrar - Banco", "Correspondência com regra de instituição financeira", termo

    return "Pode cobrar", "Nenhuma regra impeditiva encontrada", ""


def classificar_autuado(nome):
    """Mantém compatibilidade com chamadas antigas que esperam só a classificação."""
    config = obter_config_classificacao_ativa()
    classificacao, _, _ = classificar_autuado_detalhado(nome, config=config)
    return classificacao


def filtrar_autuados_cobraveis(df_base, coluna_nome_autuado):
    """Filtra a base para deixar só quem segue como "Pode cobrar"."""
    if df_base is None or df_base.empty or not coluna_nome_autuado or coluna_nome_autuado not in df_base.columns:
        return df_base

    config = obter_config_classificacao_ativa()
    detalhes = df_base[coluna_nome_autuado].apply(
        lambda nome: classificar_autuado_detalhado(nome, config=config)
    )
    detalhes_df = pd.DataFrame(
        detalhes.tolist(),
        index=df_base.index,
        columns=['_CLASSIFICACAO_AUTUADO', '_MOTIVO_CLASSIFICACAO', '_TERMO_IDENTIFICADO']
    )
    df_filtrado = df_base.copy()
    df_filtrado['_CLASSIFICACAO_AUTUADO'] = detalhes_df['_CLASSIFICACAO_AUTUADO']
    df_filtrado['_MOTIVO_CLASSIFICACAO'] = detalhes_df['_MOTIVO_CLASSIFICACAO']
    df_filtrado['_TERMO_IDENTIFICADO'] = detalhes_df['_TERMO_IDENTIFICADO']
    return df_filtrado[df_filtrado['_CLASSIFICACAO_AUTUADO'] == 'Pode cobrar'].copy()
