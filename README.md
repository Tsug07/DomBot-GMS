<p align="center">
  <img src="assets/DomBot_New.png" alt="DomBot GMS Logo" width="150">
</p>

<h1 align="center">DomBot - Taxa GMS</h1>

<p align="center">
  Automação inteligente para geração de relatórios de Taxa GMS no sistema Domínio Folha
</p>

<p align="center">
  <img src="https://img.shields.io/badge/python-3.8+-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python">
  <img src="https://img.shields.io/badge/platform-Windows-0078D6?style=for-the-badge&logo=windows&logoColor=white" alt="Windows">
  <img src="https://img.shields.io/badge/GUI-CustomTkinter-1ABC9C?style=for-the-badge" alt="CustomTkinter">
  <img src="https://img.shields.io/badge/automation-PyWinAuto-E74C3C?style=for-the-badge" alt="PyWinAuto">
</p>

<p align="center">
  <img src="https://img.shields.io/github/last-commit/Tsug07/DomBot-GMS?style=flat-square&color=2ECC71" alt="Last Commit">
  <img src="https://img.shields.io/github/repo-size/Tsug07/DomBot-GMS?style=flat-square&color=3498DB" alt="Repo Size">
  <img src="https://img.shields.io/badge/status-em%20desenvolvimento-F39C12?style=flat-square" alt="Status">
</p>

---

## Sobre

O **DomBot GMS** automatiza o processo de geração de relatórios de Taxa GMS no sistema **Domínio Folha**, eliminando o trabalho manual repetitivo de:

- Trocar entre empresas
- Navegar até o Gerenciador de Relatórios
- Preencher parâmetros do relatório
- Gerar e salvar PDFs com nomes padronizados

Tudo controlado por uma interface gráfica moderna com logs em tempo real, estatísticas e controle total da execução.

## Funcionalidades

| Funcionalidade | Descrição |
|---|---|
| **Processamento em lote** | Processa múltiplas empresas a partir de uma planilha Excel |
| **Interface moderna** | GUI dark theme com paleta de cores profissional |
| **Logs coloridos** | Logs em tempo real com cores por tipo (sucesso, erro, aviso) |
| **Preview do Excel** | Visualização dos dados antes de iniciar |
| **Controle de execução** | Iniciar, pausar, retomar e parar a qualquer momento |
| **Estatísticas em tempo real** | Cards com total, sucesso, erros, empresa atual e tempo |
| **Exportação de logs** | Salvar logs da sessão em arquivo texto |
| **Tratamento de erros** | Detecção e tratamento automático de diálogos de erro |
| **Timer** | Cronômetro mostrando tempo decorrido da execução |

## Screenshot

```
┌──────────────────────────────────────────────────────┐
│  🤖 DomBot - GMS                  ● Aguardando...   │
├──────────────────────────────────────────────────────┤
│  📁 [arquivo.xlsx]  [Procurar]  Linha: [2]           │
│  [▶ Iniciar]  [⏸ Pausar]  [⏹ Parar]                │
├──────────────────────────────────────────────────────┤
│  📊 Total  ✅ Sucesso  ❌ Erros  🏢 Empresa  ⏱ Tempo │
│  ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓░░░░░░░░░░░░░░░░░░░░  45.2%       │
├──────────────────────────────────────────────────────┤
│  📋 Logs  │  📊 Preview                              │
│  [10:30:15] ✅ Linha 2 processada com sucesso        │
│  [10:31:02] ⏳ Processando linha 3 - Empresa 105     │
│  [10:31:45] ❌ Erro na linha 3                        │
└──────────────────────────────────────────────────────┘
```

## Pré-requisitos

- **Windows** (obrigatório - utiliza Win32 API)
- **Python 3.8+**
- **Domínio Folha** instalado e aberto

## Instalação

```bash
# Clonar o repositório
git clone https://github.com/Tsug07/DomBot-GMS.git
cd DomBot-GMS

# Instalar dependências
pip install customtkinter pandas pywinauto pywin32 pillow openpyxl
```

## Uso

### 1. Preparar a planilha Excel

A planilha deve conter as seguintes colunas obrigatórias:

| Coluna | Descrição |
|---|---|
| `Nº` | Número da empresa no Domínio |
| `Periodo` | Período do relatório |
| `Salvar Como` | Nome do arquivo PDF a ser gerado |

### 2. Executar

```bash
python DomBot_GMS.py
```

### 3. Na interface

1. Clique em **Procurar** e selecione a planilha Excel
2. Verifique o preview na aba **📊 Preview**
3. Ajuste a **linha inicial** se necessário
4. Certifique-se que o **Domínio Folha** está aberto
5. Clique em **▶ Iniciar**

## Estrutura do Projeto

```
DomBot-GMS/
├── DomBot_GMS.py           # Aplicação principal
├── Old_Version.py          # Versão anterior
├── assets/
│   ├── DomBot_New.png      # Logo do aplicativo
│   ├── favicon.ico         # Ícone da janela
│   └── ...
├── logs/                   # Logs de execução (gerado automaticamente)
│   ├── success_YYYY-MM-DD.log
│   └── error_YYYY-MM-DD.log
├── DomBot_Publicar/        # Módulo de publicação
│   └── DomBot_Pub.py
└── README.md
```

## Dependências

| Pacote | Uso |
|---|---|
| `customtkinter` | Interface gráfica moderna |
| `pandas` | Leitura e manipulação do Excel |
| `pywinauto` | Automação da interface do Domínio |
| `pywin32` | Interação com janelas do Windows |
| `Pillow` | Processamento da logo/ícones |
| `openpyxl` | Engine para leitura de arquivos .xlsx |

## Fluxo da Automação

```
Início
  │
  ├─ Carregar planilha Excel
  ├─ Conectar ao Domínio Folha
  │
  └─ Para cada linha:
       ├─ Trocar empresa (F8)
       ├─ Fechar avisos de vencimento
       ├─ Abrir Relatórios Integrados (ALT+R → I → I)
       ├─ Navegar até Taxa GMS
       ├─ Preencher parâmetros (código, período)
       ├─ Executar relatório
       ├─ Salvar como PDF (Ctrl+D)
       │   ├─ Navegar até pasta GMS
       │   └─ Definir nome do arquivo
       ├─ Fechar janelas
       └─ Próxima linha
  │
  Fim → Resumo da execução
```

---

<p align="center">
  Desenvolvido por <a href="https://github.com/Tsug07">Tsug07</a>
</p>
