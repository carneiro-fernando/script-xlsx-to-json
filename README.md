# Excel to CSV/JSON Converter 📊

Ferramenta Python para converter arquivos Excel (.xlsx) em múltiplos formatos com controle de linhas.

## 🎯 Funcionalidades

- ✅ Conversão automática de múltiplos arquivos .xlsx
- ✅ Exportação para CSV com limite de 20.000 linhas por arquivo
- ✅ Exportação para JSON
- ✅ Processamento em lote de todos os arquivos de uma pasta
- ✅ Tratamento de erros robusto

## 🚀 Como Usar

### Instalação
```bash
pip install -r requirements.txt
```

### Uso Básico

1. Coloque seus arquivos .xlsx na pasta `files/`
2. Execute o notebook ou o script:
```bash
jupyter notebook excel-json.ipynb
```

ou
```bash
python converter.py
```

3. Os arquivos convertidos estarão em:
   - `output/csv/` - Arquivos CSV (divididos em partes se necessário)
   - `output/json/` - Arquivos JSON

## 📋 Requisitos

- Python 3.7+
- pandas
- openpyxl

## 💡 Exemplo

**Input:** `file_1.xlsx` com 45.000 linhas

**Output:**
- `file_1(pt1).csv` - linhas 0-19.999
- `file_1(pt2).csv` - linhas 20.000-39.999
- `file_1(pt3).csv` - linhas 40.000-44.999
- `file_1.json` - arquivo completo

## 🛠️ Tecnologias

- **pandas** - Manipulação de dados
- **openpyxl** - Leitura de arquivos Excel
- **Python 3** - Linguagem base

## 📝 Contexto

Este projeto foi desenvolvido como parte de uma proposta de freelance na Upwork, 
demonstrando habilidades em manipulação de dados e automação com Python.

## 📄 Licença

MIT License - sinta-se livre para usar e modificar.

## 👤 Autor

[Fernando Carneiro] - [GitHub](https://github.com/carneiro-fernando) | [LinkedIn](https://www.linkedin.com/in/fernandohcarneiro)

![Python](https://img.shields.io/badge/python-3.7+-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)
![Status](https://img.shields.io/badge/status-active-success.svg)