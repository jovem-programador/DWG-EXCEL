# 🔴 Projeto Raio

### Extração Automatizada de Carimbos (.DWG/.DXF) → Excel

Sistema interno desenvolvido para automatizar a extração de informações
técnicas contidas nos carimbos de desenhos em formato DWG/DXF,
consolidando os dados em planilha Excel.

------------------------------------------------------------------------

## 📌 Objetivo

Reduzir tempo operacional do setor de Planejamento e Engenharia na
consolidação de dados de desenhos técnicos, eliminando retrabalho manual
e aumentando confiabilidade das informações.

------------------------------------------------------------------------

## ⚙️ Funcionalidades

-   Conversão automática DWG → DXF via ODA File Converter
-   Extração inteligente de:
    -   Classificação
    -   Projeto
    -   Número SE
    -   Número de contrato
    -   Fase do projeto
    -   Título e subtítulos
    -   Área/Subárea
    -   Revisões dinâmicas
-   Sistema de fallback por coordenadas relativas
-   Ajuste de tolerância espacial (X/Y)
-   Processamento em lote
-   Exportação automática para Excel
-   Interface gráfica via Streamlit

------------------------------------------------------------------------

## 🗂 Estrutura do Projeto

6 - Tela Raio/ │ ├── app.py ├── core_extracao.py ├── scriptTela.py ├──
Logo/ ├── Projetos_DWG/ ├── venv/ └── run_raio.bat

------------------------------------------------------------------------

## 🖥 Requisitos

-   Windows 10 ou 11
-   Python 3.11+ (recomendado 3.11)
-   ODA File Converter instalado (caso utilize DWG)

Download ODA: https://www.opendesign.com/guestfiles/oda_file_converter

------------------------------------------------------------------------

## 🚀 Instalação (Primeira vez)

Dentro da pasta do projeto:

py -m venv venv\
venv`\Scripts`{=tex}`\activate  `{=tex} pip install streamlit pandas
ezdxf openpyxl\
deactivate

------------------------------------------------------------------------

## ▶ Execução

### Método recomendado (via .bat)

Clique duas vezes em:

run_raio.bat

Ou execute manualmente:

py -m streamlit run app.py

------------------------------------------------------------------------

## 📐 Parâmetros Técnicos

### Janela X (fallback)

Tolerância horizontal de busca de texto em relação à posição esperada do
campo.

### Janela Y (fallback)

Tolerância vertical de busca.

Esses parâmetros permitem adaptar o sistema a pequenas variações de
coordenadas entre desenhos.

------------------------------------------------------------------------

## 📊 Fluxo Operacional

1.  Selecionar modo (DWG ou DXF)
2.  Definir pasta de origem
3.  Ajustar parâmetros se necessário
4.  Processar
5.  Baixar Excel consolidado

------------------------------------------------------------------------

## ⚠ Observações Técnicas

-   Tolerâncias muito altas podem capturar texto incorreto.
-   Caso campo não seja encontrado, revisar posição do carimbo no DWG.
-   Recomenda-se padronização de templates de desenho.

------------------------------------------------------------------------

## 🏢 Aplicação Interna

Sistema desenvolvido para uso interno da Projeta --- Engenharia &
Planejamento.

Versão: 1.0\
Ano: 2026

------------------------------------------------------------------------

## 🔮 Evoluções Futuras

-   Barra de progresso detalhada por arquivo
-   Log técnico de extração
-   Validador de campos obrigatórios
-   Exportação de relatório de inconsistências
-   Empacotamento como executável
-   Deploy interno em servidor corporativo

------------------------------------------------------------------------
