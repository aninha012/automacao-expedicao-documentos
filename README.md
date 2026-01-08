# Automacão De Expedição De Documentos

# PROJCORREIOS

## Visão Geral

Projeto desenvolvido em Python para automatizar atividades relacionadas à expedição de documentos no SEI. A solução acessa processos, identifica e baixa arquivos vinculados, extrai informações relevantes e consolida os dados em planilhas Excel, permitindo a geração de etiquetas e apoio à etapa de impressão.

O objetivo principal é reduzir trabalho manual, padronizar a execução das tarefas e minimizar erros operacionais, utilizando automações baseadas em Selenium e PyAutoGUI.

---

## Funcionalidades Principais

* Acesso automatizado a processos no SEI
* Download de documentos e arquivos vinculados aos processos
* Extração e organização de dados relevantes
* Preenchimento automático de planilhas Excel
* Geração de informações para etiquetas
* Apoio ao processo de impressão de documentos

---

## Estrutura do Projeto

```
PROJCORREIOS/
│
├── .vscode/
│   └── Configurações do ambiente de desenvolvimento
│
├── exp-bloco/
│   └── bloco-de-assinatura.py
│      Script responsável pelo tratamento e manipulação de blocos de assinatura
│
├── exp-documento/
│   ├── excel.py
│   │   Manipulação, leitura e escrita de dados em planilhas Excel
│   │
│   ├── processo-especifico.py
│   │   Automação voltada para um processo específico
│   │
│   └── processo-inteiro.py
│       Automação para tratamento completo de processos
│
├── .env
│   Arquivo de variáveis de ambiente (credenciais e configurações sensíveis)
│   *Este arquivo é ignorado pelo Git via `.gitignore`*
│
├── imprimir.py
│   Script de apoio à impressão de documentos
│
├── ordem.txt
│   Arquivo auxiliar para controle de ordem ou sequenciamento
│
├── posicao.py
│   Script auxiliar que utiliza PyAutoGUI para capturar e identificar a posição do mouse na tela,
│   servindo de apoio para ajustes finos na automação gráfica
│
│
└── teste.py
    Arquivo destinado a testes e validações
```

---

## Tecnologias Utilizadas

* Python
* PyAutoGUI (automação gráfica e captura de posição do mouse)
* Selenium WebDriver (automação de navegação web)
* Manipulação de arquivos e PDFs (pikepdf, base64)
* Excel (openpyxl / pandas, conforme uso)
* Variáveis de ambiente (.env)

---

## Pré-requisitos

* Python 3.10 ou superior
* Bibliotecas Python necessárias:

  * pyautogui
  * selenium
  * python-dotenv
  * pikepdf
* Acesso autorizado ao sistema SEI
* Navegador compatível e WebDriver configurado

---

## Configuração do Ambiente

1. Clonar o repositório do projeto
2. Criar o arquivo `.env` com as credenciais e parâmetros necessários
3. Garantir que o `.env` esteja listado no `.gitignore`
4. Instalar as dependências do projeto
5. Ajustar caminhos e parâmetros conforme o ambiente local

---

## Observações Importantes

* O arquivo `.env` não é versionado e deve permanecer fora do repositório

* O projeto depende de sistemas externos, podendo exigir ajustes conforme mudanças no SEI

* Recomenda-se executar inicialmente os scripts de teste

* Uso indicado em ambiente controlado

---

## Status do Projeto

Em desenvolvimento contínuo, com melhorias e ajustes conforme necessidade operacional.

---

## Autoria

Projeto desenvolvido para fins de automação administrativa e aprendizado prático em Python e automações.
