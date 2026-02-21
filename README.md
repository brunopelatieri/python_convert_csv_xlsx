# 📋 Chatwoot CSV → Excel Converter

> Converte exportações de contatos do **Chatwoot** de CSV para planilha Excel (`.xlsx`), corrigindo automaticamente caracteres especiais quebrados como acentos e cedilha.

---

## 🔍 O Problema

Ao exportar contatos pelo Chatwoot, o arquivo CSV gerado frequentemente apresenta caracteres especiais corrompidos:

| ❌ Com problema | ✅ Corrigido |
|---|---|
| `AndrÃ© Camily` | `André Camily` |
| `ElisÃ¢ngela` | `Elisângela` |
| `JoÃ£o da Silva` | `João da Silva` |
| `ApareciÃ§a` | `Aparecição` |

Isso acontece por um conflito de **encoding** — o arquivo é salvo em UTF-8 pelo Chatwoot, mas muitos programas (como o Excel) o abrem assumindo Latin-1/ISO-8859-1, ou vice-versa.

---

## ✅ O que o Script Faz

- 🔎 **Detecta automaticamente** o encoding correto do CSV (testa utf-8, utf-8-sig, latin-1, iso-8859-1, cp1252)
- 🔧 **Corrige** caracteres especiais corrompidos (acentos, ç, ã, ê, etc.)
- 📊 **Converte** o CSV para `.xlsx` com formatação profissional:
  - Cabeçalho com fundo azul e texto branco em negrito
  - Largura de colunas ajustada automaticamente ao conteúdo
  - Fonte Arial padronizada
- 📁 **Nomeia** o arquivo de saída automaticamente (mesmo nome do CSV, extensão `.xlsx`)

---

## 🚀 Como Usar

### 1. Pré-requisitos

Certifique-se de ter o **Python 3.7+** instalado:

```bash
python --version
# ou
python3 --version
```

> Não tem Python? Baixe em [python.org](https://www.python.org/downloads/)

### 2. Instalar dependências

```bash
pip install pandas openpyxl
```

### 3. Baixar o script

Clone o repositório ou baixe o arquivo `converter_chatwoot.py` diretamente.

```bash
git clone https://github.com/seu-usuario/chatwoot-csv-excel.git
cd chatwoot-csv-excel
```

### 4. Exportar os contatos no Chatwoot

No Chatwoot, vá em: **Contatos → Importar/Exportar → Exportar Contatos**

O arquivo será baixado como `.csv`.

### 5. Executar o script

**Uso básico** (gera o `.xlsx` com o mesmo nome do CSV):
```bash
python converter_chatwoot.py contatos.csv
```

**Especificando o nome do arquivo de saída:**
```bash
python converter_chatwoot.py contatos.csv minha_planilha.xlsx
```

---

## 💻 Exemplos por Sistema Operacional

### Windows (Prompt de Comando)

```cmd
# Navegue até a pasta onde estão o script e o CSV
cd C:\Users\SeuNome\Downloads

# Execute
python converter_chatwoot.py contatos.csv
```

### macOS / Linux (Terminal)

```bash
# Navegue até a pasta onde estão o script e o CSV
cd ~/Downloads

# Execute
python3 converter_chatwoot.py contatos.csv
```

> 💡 **Dica:** coloque o script na mesma pasta que o arquivo CSV para facilitar a execução.

---

## 📂 Estrutura do Repositório

```
chatwoot-csv-excel/
│
├── converter_chatwoot.py   # Script principal
├── README.md               # Esta documentação
└── exemplo/
    └── contatos_exemplo.csv  # Arquivo de exemplo para teste
```

---

## ⚙️ Como Funciona (Detalhes Técnicos)

O script segue três etapas principais:

**1. Detecção de Encoding**
Testa uma lista de encodings em ordem de prioridade até conseguir ler o arquivo sem erros:
```
utf-8 → utf-8-sig → latin-1 → iso-8859-1 → cp1252
```

**2. Correção de Caracteres**
Aplica a conversão `latin-1 → utf-8` coluna a coluna para reverter o mojibake (nome técnico para o problema de caracteres trocados). Colunas que já estão corretas são ignoradas automaticamente.

**3. Exportação para Excel**
Usa `pandas` + `openpyxl` para gerar o `.xlsx` com:
- Aba nomeada `Contatos`
- Cabeçalho formatado (fundo azul `#2B5ED6`, texto branco, negrito)
- Colunas com largura proporcional ao conteúdo (máximo 40 caracteres)

---

## 🐛 Solução de Problemas

**`python: command not found`**
> Use `python3` no lugar de `python`, ou verifique se o Python está instalado e no PATH.

**`ModuleNotFoundError: No module named 'pandas'`**
> Execute: `pip install pandas openpyxl`

**Caracteres ainda aparecem errados após a conversão**
> O arquivo pode ter um encoding menos comum. Abra o CSV em um editor de texto (como VS Code ou Notepad++) e verifique o encoding exibido no rodapé. Ajuste a lista `encodings` no script conforme necessário.

**Erro ao abrir o `.xlsx` no Excel**
> Certifique-se de que o arquivo não está aberto em outro programa durante a conversão.

---

## 🛠️ Tecnologias Utilizadas

| Biblioteca | Versão mínima | Uso |
|---|---|---|
| Python | 3.7+ | Linguagem base |
| pandas | 1.3+ | Leitura e manipulação do CSV |
| openpyxl | 3.0+ | Geração e formatação do Excel |

---

## 🤝 Contribuindo

Contribuições são bem-vindas! Sinta-se à vontade para:

1. Fazer um **fork** do repositório
2. Criar uma **branch** para sua feature: `git checkout -b minha-melhoria`
3. Fazer o **commit**: `git commit -m 'Adiciona suporte a X'`
4. Fazer o **push**: `git push origin minha-melhoria`
5. Abrir um **Pull Request**

---

## 📄 Licença

Distribuído sob a licença MIT. Veja `LICENSE` para mais informações.

---

## 📬 Contato

Dúvidas ou sugestões? Abra uma [issue](../../issues) no repositório.

---

*Feito para facilitar a vida de quem usa o Chatwoot no dia a dia* 🚀
