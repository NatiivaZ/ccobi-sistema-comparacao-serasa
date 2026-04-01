# 📊 Sistema de Análise e Comparação de Bases de Dados
## SERASA × Dívida Ativa

Sistema profissional desenvolvido para análise, cruzamento e comparação de bases de dados SERASA e Dívida Ativa, com interface moderna e funcionalidades avançadas.

## 🚀 Funcionalidades

- ✅ **Upload e Visualização**: Carregamento de planilhas Excel/CSV com preview
- ✅ **Cruzamento Inteligente**: Comparação automática entre bases de dados
- ✅ **Filtros Automáticos**: Filtro por ano de vencimento (2025)
- ✅ **Agrupamento**: Organização por CPF/CNPJ (maior → menor quantidade)
- ✅ **Separação por Valores**: 
  - SERASA: ≤ R$ 1.000 e > R$ 1.000 (individual e acumulativo)
  - CADIN: > R$ 100
- ✅ **Análise de Divergências**: Identificação de registros presentes apenas em uma base
- ✅ **Dashboard Visual**: Gráficos interativos e métricas em tempo real
- ✅ **Exportação**: Download de planilhas processadas em Excel

## 📋 Pré-requisitos

- Python 3.8 ou superior
- pip (gerenciador de pacotes Python)

## 🔧 Instalação

1. **Clone ou baixe este repositório**

2. **Instale as dependências:**
```bash
pip install -r requirements.txt
```

## 🎯 Como Usar

1. **Inicie o sistema:**
```bash
streamlit run app.py
```

2. **No navegador que abrir automaticamente:**
   - Faça upload da planilha SERASA na barra lateral
   - Faça upload da planilha Dívida Ativa na barra lateral
   - Configure os nomes das colunas (CPF/CNPJ, Valor, Vencimento)
   - Clique em "Executar Análise Completa"

3. **Explore os resultados:**
   - Visualize o dashboard com métricas e gráficos
   - Navegue pelas abas para análises detalhadas
   - Exporte as planilhas processadas

## 📊 Estrutura de Análise

O sistema realiza as seguintes operações automaticamente:

1. **Extração e Normalização**: Carrega e normaliza CPF/CNPJ das bases
2. **Cruzamento**: Identifica registros presentes em ambas as bases
3. **Filtragem**: Mantém apenas registros com correspondência completa
4. **Filtro Temporal**: Seleciona apenas vencimentos de 2025
5. **Agrupamento**: Organiza por CPF/CNPJ ordenado por quantidade
6. **Separação**: Divide por critérios de valor (SERASA e CADIN)
7. **Relatório de Divergências**: Lista registros não correspondentes

## 🎨 Interface

- Design moderno e profissional
- Gráficos interativos com Plotly
- Visualização responsiva
- Métricas em tempo real
- Exportação facilitada

## 📝 Notas Importantes

- O sistema prioriza a base SERASA em caso de divergências
- Valores podem ser considerados individual ou acumulativamente
- Todas as planilhas exportadas incluem data no nome do arquivo
- O sistema normaliza automaticamente CPF/CNPJ removendo caracteres especiais

## 🔒 Segurança

- Dados processados localmente (não enviados para servidores externos)
- Cache otimizado para performance
- Validação de dados em todas as etapas

## 📞 Suporte

Para dúvidas ou problemas, verifique:
- Os nomes das colunas estão corretos?
- As planilhas estão no formato correto (Excel/CSV)?
- As datas de vencimento estão no formato reconhecível?

---

**Desenvolvido com ❤️ para otimizar processos de análise de dados**

