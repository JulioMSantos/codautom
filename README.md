# ⚡ Raichu Pro - Automação de Documentos Institucionais

O **Raichu Pro** é um sistema inteligente desenvolvido para automatizar a extração de dados e a geração de documentação de projetos institucionais, como Acordos de Cooperação Técnica (ACT), Contratos Globais (CG) e Acordos de Parceria (AP) no âmbito da Universidade Federal de Santa Maria (UFSM).

## 🎯 Objetivo
Reduzir drasticamente o tempo gasto em trabalhos manuais e repetitivos na elaboração de documentos. O sistema realiza a leitura autônoma de relatórios em formato PDF, processa os dados com base em expressões regulares avançadas e gera pacotes completos `.zip` contendo formulários formatados em Word (`.docx`) e Excel (`.xlsx`).

## 🏗️ Arquitetura do Sistema
O projeto foi desenhado com uma arquitetura *Decoupled* (desacoplada) para garantir escalabilidade e fácil manutenção:

*   **Front-End (Interface UI):** Desenvolvido em **Streamlit**, proporcionando uma experiência de usuário fluida, responsiva e sem necessidade de instalação de software na máquina local.
*   **Back-End (Motor de Processamento):** API construída em **FastAPI**, responsável por receber o arquivo, executar as lógicas de extração (via `pdfplumber`), manipular os templates (`docxtpl` e `openpyxl`) e devolver o pacote empacotado para o cliente.

## 📦 Implantação e Infraestrutura (Pronto para o CPD)
O sistema foi projetado pensando em fácil alocação em máquinas virtuais. 

*   O repositório conta com um arquivo `Dockerfile` nativo, permitindo a conteinerização imediata do motor FastAPI.
*   Pode ser hospedado de forma independente na infraestrutura de TI da instituição, garantindo segurança dos dados e alta disponibilidade.

## 🛠️ Tecnologias Utilizadas
*   Python 3.10+
*   FastAPI & Uvicorn (API e Servidor Web)
*   Streamlit (Interface Gráfica)
*   Pdfplumber (Extração e leitura de PDFs)
*   Docxtpl / Openpyxl (Manipulação de arquivos Office)
*   Docker (Conteinerização)

---
*Desenvolvido por Julio Maia para otimização de rotinas administrativas e acadêmicas.*