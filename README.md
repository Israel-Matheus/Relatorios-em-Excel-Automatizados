# Automação de Relatórios com Excel

Este projeto realiza uma sequência de operações em arquivos Excel com o objetivo de gerar relatórios de vendas de veículos de forma automatizada e enviar por e-mail. As etapas envolvem leitura de dados, criação de tabelas, gráficos e aplicação de fórmulas em planilhas.

---

## 📁 Estrutura do Projeto

- `1-start_here/`  
  Pasta inicial (vazia). Para testar o projeto, mantenha apenas o arquivo `CarSales.xlsx` dentro da pasta `data/` e remova as demais saídas geradas.

- `data/`  
  Pasta onde ficam o arquivo de entrada `CarSales.xlsx` e os arquivos gerados (`pivot_table.xlsx`, `barchart.xlsx`, `test.xlsx`).

- `requirements.txt`  
  Lista das bibliotecas necessárias para rodar o projeto.

---

## ⚙️ Requisitos

- Python 3.8+
- VSCode (opcional, mas recomendado)
- Extensão do VSCode: `Excel Viewer` (ver `extensions.json`)
- Conexão com internet (para envio de e-mails)

---

## 📦 Instalação

1. Clone o repositório:
   ```bash
   git clone https://github.com/Israel-Matheus/Relatorios-em-Excel-Automatizados
   cd Relatorios-em-Excel-Automatizados
