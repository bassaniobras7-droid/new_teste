# Estrutura de Arquivos da Aplicação

## 📄 Arquivos Principais

| Arquivo | Tamanho | Linhas | Descrição |
|---------|---------|--------|-----------|
| `gerar_relatorios_com_formulas_1.3.py` | 104 KB | 1.746 | **Principal** - Gerador de relatórios em Excel com lógica de negócio |
| `main.py` | 1.7 KB | 49 | Executor/ponto de entrada alternativo |
| `QTD.py` | 1.8 KB | 42 | Script auxiliar para quantidade |

## 📊 Dados de Configuração

| Arquivo | Descrição |
|---------|-----------|
| `Subclasse.csv` | Tabela de subclasses para lógica de derivados |
| `Valores Ctba.csv` | Tabela de valores contábeis |
| `Valores Ctba_.csv` | Versão alternativa de valores |
| `OBSERVACOES_GERAIS.txt` | Observações genéricas incluídas nos relatórios |
| `requirements.txt` | Dependências Python (openpyxl, etc) |

## 🔧 Scripts de Diagnóstico/Debug (helpers)

| Arquivo | Descrição |
|---------|-----------|
| `check_aditivos_decrypt.py` | Verifica descriptografia de aditivos |
| `check_aditivos_structure.py` | Inspeciona estrutura de aditivos |
| `check_area.py` | Verifica áreas |
| `check_sheet.py` | Valida abas |
| `inspect_sheets.py` | Inspeciona estrutura de abas |
| `decrypt_excel.py` | Desencripta Excel |
| `decrypt_excel2.py` | Variante de desencriptação |
| `compare_excel.py` | Compara arquivos Excel |
| `compare_excel2.py` | Variante de comparação |

## 🎯 Scripts de Substituição (utilities)

| Arquivo | Descrição |
|---------|-----------|
| `replace_aditivos_func.py` | Substitui funções de aditivos |
| `replace_aditivos_func_1_3.py` | Versão 1.3 de substituição |

## 🧪 Scripts Temporários/Teste

| Arquivo | Descrição |
|---------|-----------|
| `tmp_check_sheet.py` | Teste temporário 1 |
| `tmp_check_sheet2.py` | Teste temporário 2 |
| `tmp_print.py` | Teste de impressão |
| `tmp_test_run.py` | Teste de execução |

## 📦 Módulos do Pacote `src/`

| Arquivo | Descrição |
|---------|-----------|
| `src/src/__init__.py` | Inicializador do pacote |
| `src/src/aspg_logic.py` | Lógica ASPG |
| `src/src/data_processing.py` | Processamento de dados |
| `src/src/excel_writer.py` | Escritor Excel |
| `src/src/lix_j_logic.py` | Lógica LIX-J |
| `src/src/lp_tub_logic.py` | Lógica LP-TUB |
| `src/src/utils.py` | Utilidades |

## 🚫 Arquivos Ignorados pelo Git (.gitignore)

```
.venv/                  # Ambiente virtual Python
env/                    # Alternativa de ambiente virtual
__pycache__/            # Cache Python
*.xlsx                  # Arquivos Excel (saídas geradas)
*.7z                    # Arquivos compactados
~$*                     # Arquivos abertos do Excel
BKP_projeto/            # Backup do projeto
Relatorios_OLD/         # Relatórios antigos
Valores_Ctba/           # Pasta de valores
```

## 🎯 Núcleo Ativo da Aplicação

### Fluxo Principal

A aplicação funciona principalmente através de:

1. **`gerar_relatorios_com_formulas_1.3.py`** ← Arquivo principal
   - Processa dados de planilhas `ADD_*.xlsx` e `DIS_*.xlsx`
   - Gera relatório Excel com 5 abas:
     - Cliente
     - Resumo
     - Aditivos x Distrato
     - Relação Média Material
   - Implementa lógica de negócio (cálculos, fórmulas, formatação)

2. **Arquivos CSV de entrada** (configuração):
   - `Subclasse.csv` - Define tipos derivados
   - `Valores Ctba.csv` - Preços e valores unitários

3. **Arquivos de entrada** (não rastreados, entrada do usuário):
   - `ADD_*.xlsx` - Planilhas de aditivos
   - `DIS_*.xlsx` - Planilhas de distratos
   - `OBSERVACOES_GERAIS.txt` - Texto customizado

### Saída Gerada

- `Relatorios_Com_Formulas.xlsx` - Arquivo Excel final com relatório completo

## 📝 Histórico de Versões

- **gerar_relatorios_com_formulas_1.3.py** - Versão atual em uso
- **gerar_relatorios_com_formulas.py** - Versão anterior (legado)

## 🔄 Scripts de Suporte

Alguns scripts são utilitários de suporte:
- **Scripts de check** (`check_*.py`) - Para validação de dados
- **Scripts de decrypt** (`decrypt_*.py`) - Para desencriptação de Excel
- **Scripts de comparação** (`compare_*.py`) - Para comparar estruturas
- **Scripts de replace** (`replace_*.py`) - Para substituições em massa
- **Scripts temporários** (`tmp_*.py`) - Para testes rápidos

Estes não fazem parte do core da aplicação, mas ajudam no desenvolvimento e debug.

## 📚 Dependências

Ver `requirements.txt` para lista completa. Principais:
- **openpyxl** - Leitura e escrita de Excel
- Outras dependências conforme necessário
