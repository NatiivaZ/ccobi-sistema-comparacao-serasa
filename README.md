# Sistema de análise e comparação — SERASA × Dívida Ativa

Aplicação **web local** construída com **Streamlit** para cruzar bases **SERASA** e **Dívida Ativa** (autos de infração / débitos ANTT), gerar métricas, gráficos interativos (**Plotly**) e **exportação Excel** formatada.  
Projeto **CCOBI – SERASA**.

---

## Visão geral

O sistema permite:

- Carregar **duas planilhas** (SERASA e Dívida Ativa) em **Excel** (`.xlsx`, `.xls`) ou **CSV**.  
- Configurar **nomes exatos das colunas** na barra lateral (identificador do débito/auto, CPF/CNPJ, valores, vencimento, processo, modais, etc.).  
- Executar uma **análise completa** que normaliza documentos, cruza registros, aplica filtros (ex.: por **ano de vencimento**), agrupa por CPF/CNPJ, separa faixas de valor (SERASA e critérios CADIN), identifica **divergências** e exibe **dashboard** com indicadores e gráficos.  
- **Exportar** planilhas processadas com timestamp no nome do arquivo.  
- **Classificar autuados** (órgãos, bancos, leasing, exceções) via configuração persistida em `config_classificacao_autuados.json` (listas editáveis na interface).

Há também um aplicativo complementar **`app_vencimentos.py`** focado em **análise por ano de vencimento** a partir de **uma única** base (filtros, deduplicação mantendo vencimento mais antigo, visualizações).

---

## Requisitos

- **Python** 3.8+  
- **pip**

### Dependências

```bash
pip install -r requirements.txt
```

Principais pacotes: `streamlit`, `pandas`, `numpy`, `plotly`, `openpyxl`, `xlrd`.

---

## Instalação e execução

```bash
cd "Sistema de Comparação SERASA"
pip install -r requirements.txt
```

### Aplicativo principal (duas bases — SERASA × Dívida Ativa)

```bash
streamlit run app.py
```

O navegador abrirá a URL local (por padrão `http://localhost:8501`).

### Aplicativo por ano de vencimento (uma base)

```bash
streamlit run app_vencimentos.py
```

Ou use os arquivos `.bat` fornecidos (`iniciar.bat`, `iniciar_vencimentos.bat`) se estiverem configurados no seu ambiente.

---

## Uso do `app.py` (comparação completa)

1. Na **barra lateral**, faça upload da planilha **SERASA** e da **Dívida Ativa**.  
2. Ajuste os campos de texto com os **nomes das colunas** exatamente como aparecem no arquivo:  
   - **Obrigatório:** coluna do **identificador do débito / auto de infração**.  
   - Demais: CPF/CNPJ, valor SERASA, valor Dívida Ativa, data de vencimento, nº do processo, modais (tipo modal / subtipo de débito), nome do autuado (para classificação na exportação).  
3. Revise seções opcionais de **classificação de autuados** (listas de termos para órgão, banco, leasing, exceções).  
4. Execute **“Executar análise completa”** (ou botão equivalente na interface).  
5. Navegue pelas **abas** (resumo, cruzamentos, divergências, gráficos) e use **download** das planilhas geradas.

### Normalização

CPF/CNPJ são **normalizados** (remoção de caracteres não numéricos) para comparação entre bases.

### Lógica de negócio (alto nível)

Conforme implementado em `app.py` (milhares de linhas — resumo conceitual):

- Cruzamento por chaves definidas (auto + documentos e datas conforme regras da aplicação).  
- Filtros temporais (ex.: foco em vencimentos de determinado ano quando aplicável).  
- Quebras por faixas de valor (ex.: SERASA ≤ R$ 1.000 e > R$ 1.000; CADIN > R$ 100).  
- Relatórios de registros só em uma das bases.  
- Visualizações dinâmicas com Plotly.

> Para detalhes finos (ordem das etapas, nomes de abas exportadas), consulte o código-fonte de `app.py` e comentários internos.

---

## Uso do `app_vencimentos.py`

1. Upload de **uma** planilha.  
2. Informe colunas **obrigatórias:** auto de infração e data de vencimento.  
3. Opcional: processo, subtipo/modal, CPF/CNPJ.  
4. O app remove duplicados por auto **mantendo a data de vencimento mais antiga**, extrai **ano** de vencimento e oferece visualizações agregadas por ano.

Útil para **curadoria** de base antes de cruzamentos ou para relatórios só de calendário de vencimento.

---

## Estrutura de arquivos

| Arquivo | Função |
|---------|--------|
| `app.py` | App principal SERASA × Dívida Ativa |
| `app_vencimentos.py` | App análise por ano (uma base) |
| `requirements.txt` | Dependências |
| `config_classificacao_autuados.json` | Persistência de listas de classificação (gerado/alterado pela UI) |
| `GUIA_RAPIDO.txt` | Dicas rápidas de uso |
| `iniciar.bat` / `iniciar_vencimentos.bat` | Atalhos Windows |

---

## Segurança e privacidade

- **Processamento local:** arquivos enviados ao Streamlit rodam na sua máquina; não há upload automático para nuvem pelo próprio código listado.  
- Evite subir planilhas reais para repositórios públicos.  
- Em ambientes corporativos, avalie política de dados antes de compartilhar exports.

---

## Desempenho

- `st.cache_data` é usado para recarregar DataFrames de forma eficiente ao repetir interações.  
- Planilhas muito grandes podem exigir **mais RAM** e tempo de processamento; considere pré-filtrar no Excel.

---

## Solução de problemas

| Sintoma | Ação |
|---------|------|
| Coluna não encontrada | Conferir nome exato (espaços, acentos). |
| CSV com separador errado | O código usa `;` e `decimal=','` para CSV em partes do fluxo — alinhar formato ou converter para XLSX. |
| Gráfico vazio | Verificar se filtros removem todos os registros. |
| Erro ao ler `.xls` | `xlrd` instalado; preferir `.xlsx` quando possível. |

---

## Contexto

Ferramenta de **análise inteligente** para apoio à gestão de débitos e conformidade entre bases **SERASA** e **Dívida Ativa**, no âmbito **CCOBI – SERASA**.
