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

### 1. Clone o repositório

```bash
git clone https://github.com/Israel-Matheus/Relatorios-em-Excel-Automatizados
cd Relatorios-em-Excel-Automatizados
```

### 2. Crie e ative um ambiente virtual

**Windows:**

```bash
python -m venv venv
venv\Scripts\activate
```

**Linux/macOS:**

```bash
python3 -m venv venv
source venv/bin/activate
```

### 3. Instale as dependências

```bash
pip install -r requirements.txt
```

---

## 🚀 Execução passo a passo

### 1. Importar dados

```bash
python import_data.py
```

Leitura e visualização básica do arquivo `CarSales.xlsx`.

### 2. Gerar tabela pivô

```bash
python pivot_table.py
```

Cria o arquivo `pivot_table.xlsx` com a tabela dinâmica.

### 3. Ler valores da planilha

```bash
python sheet_read.py
```

Lê e exibe valores específicos da planilha gerada.

### 4. Adicionar gráfico

```bash
python add_chart.py
```

Adiciona um gráfico de barras com os dados de vendas por fabricante e salva como `barchart.xlsx`.

### 5. Aplicar fórmulas

```bash
python forms.py
```

Aplica fórmulas de soma por fabricante na última linha da planilha. Salva como `test.xlsx`.

### 6. Enviar e-mail com anexo

```bash
python email.py
```

Envia o arquivo final `test.xlsx` por e-mail.

---
### 📜 Licença

Este projeto foi desenvolvido para fins educacionais e não possui fins comerciais.  
Programa feito junto com a trilha Start em Python da OneBitCode.

## ✉️ Sobre o envio de e-mail

- Preencha seu e-mail e o destinatário diretamente no script `email.py`.
- Crie um arquivo chamado `senha` (sem extensão) contendo sua [senha de app do Gmail](https://support.google.com/accounts/answer/185833).
