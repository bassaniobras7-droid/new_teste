# 📋 Arquivos Obrigatórios e Opcionais da Aplicação

## ✅ ABSOLUTAMENTE OBRIGATÓRIOS

### 1. **ADD_*.xlsx** (Aditivos)

- **Padrão:** `ADD_[projeto].xlsx`
- **Exemplo:** `ADD_Projeto RB 10837.xlsx`
- **Conteúdo esperado:** Planilha com dados de aditivos
- **Abas esperadas:**
  - `1.__Tabela de Forro`
  - `2.__Tabela de Modelo Genérico`
  - `3.__Tabela de Paredes`
- **Se faltando:** ❌ **ERRO** - Script aborta com mensagem:
  ```
  "Erro: Não foi possível encontrar todos os arquivos Excel necessários. Abortando."
  ```
- **Localização:** Mesmo diretório que `gerar_relatorios_com_formulas_1.3.py`
- **Código relevante:** Linha 1899
  ```python
  excel_file_add = find_latest_excel_file('ADD_')
  ```

---

### 2. **DIS_*.xlsx** (Distratos)

- **Padrão:** `DIS_[projeto]_DISTRATO.xlsx` ou similar
- **Exemplo:** `DIS_Projeto RB 10837 DISTRATO.xlsx`
- **Conteúdo esperado:** Mesma estrutura do `ADD_*.xlsx`
- **Abas esperadas:**
  - `1.__Tabela de Forro`
  - `2.__Tabela de Modelo Genérico`
  - `3.__Tabela de Paredes`
- **Se faltando:** ❌ **ERRO** - Script aborta imediatamente
- **Localização:** Mesmo diretório que o script Python
- **Código relevante:** Linha 1900-1904
  ```python
  excel_file_dis = find_latest_excel_file('DIS_')

  if not excel_file_add or not excel_file_dis:
      print("Erro: Não foi possível encontrar todos os arquivos Excel necessários. Abortando.")
      exit()
  ```

---

### 3. **Valores Ctba.csv** (Tabela de Preços)

- **Nome exato:** `Valores Ctba.csv` (com espaço)
- **Colunas obrigatórias:**
  - `Tipo R. Bassani` (chave primária)
  - `Un` (unidade de medida)
  - `Valor do Material + MO` (preço total)
  - `Custo MO à Pagar` (custo de mão de obra)
  - `Forros` (descrição do item)
- **Separador:** `;` (ponto e vírgula)
- **Encoding:** UTF-8
- **Se faltando:** ⚠️ **AVISO** - Script continua mas com preços vazios
  ```
  "AVISO: Arquivo de preços 'Valores Ctba.csv' não encontrado."
  ```
- **Impacto:** Relatório pode estar incompleto ou com valores em branco
- **Localização:** Mesmo diretório que o script Python
- **Código relevante:** Linha 237-256
  ```python
  @lru_cache(maxsize=1)
  def load_price_data(filename='Valores Ctba.csv'):
      try:
          df_prices = pd.read_csv(filename, sep=';', header=0)
          # ...
      except FileNotFoundError:
          print(f"AVISO: Arquivo de preços '{filename}' não encontrado.")
          return {}
  ```

---

### 4. **Subclasse.csv** (Tabela de Subclasses)

- **Nome exato:** `Subclasse.csv`
- **Colunas obrigatórias:**
  - `Tipo` (código da subclasse)
- **Separador:** `;` (ponto e vírgula)
- **Encoding:** UTF-8
- **Se faltando:** ⚠️ **AVISO** - Script continua mas sem subclasses
  ```
  "AVISO: Arquivo de subclasses 'Subclasse.csv' não encontrado."
  ```
- **Impacto:** Lógica de negócio relacionada a subclasses pode não funcionar
- **Localização:** Mesmo diretório que o script Python
- **Código relevante:** Linha 259-269
  ```python
  @lru_cache(maxsize=1)
  def load_subclass_data(filename='Subclasse.csv'):
      try:
          df_subclass = pd.read_csv(filename, sep=';', header=0)
          # ...
      except FileNotFoundError:
          print(f"AVISO: Arquivo de subclasses '{filename}' não encontrado.")
          return set()
  ```

---

## ⚠️ OPCIONAIS (Funcionam com fallback)

### 1. **OBSERVACOES_GERAIS.txt** (Observações Genéricas)

- **Nome exato:** `OBSERVACOES_GERAIS.txt`
- **Conteúdo:** Texto livre (observações genéricas incluídas no relatório)
- **Encoding:** UTF-8
- **Se faltando:** ⚠️ **AVISO** - Script continua normalmente
  ```
  "AVISO: Arquivo 'OBSERVACOES_GERAIS.txt' não encontrado.
          As observações gerais não serão incluídas."
  ```
- **Impacto:** Relatório é gerado sem observações (funciona 100%)
- **Fallback:** Conteúdo vazio (`""`)
- **Localização:** Mesmo diretório que o script Python
- **Código relevante:** Linha 1884-1891
  ```python
  try:
      with open('OBSERVACOES_GERAIS.txt', 'r', encoding='utf-8') as f:
          observacoes_gerais_content = f.read()
  except FileNotFoundError:
      print("AVISO: Arquivo 'OBSERVACOES_GERAIS.txt' não encontrado...")
      observacoes_gerais_content = ""  # Fallback: vazio
  except Exception as e:
      print(f"ERRO ao ler 'OBSERVACOES_GERAIS.txt': {e}")
      observacoes_gerais_content = ""
  ```

---

## 📊 Tabela de Status

| Arquivo | Status | Tipo | Ação se faltar | Impacto |
|---------|--------|------|---|---|
| `ADD_*.xlsx` | **CRÍTICA** | Entrada | Halt (exit) | Script aborta |
| `DIS_*.xlsx` | **CRÍTICA** | Entrada | Halt (exit) | Script aborta |
| `Valores Ctba.csv` | **CRÍTICA** | Configuração | Warn + continue | Preços vazios |
| `Subclasse.csv` | **CRÍTICA** | Configuração | Warn + continue | Lógica incompleta |
| `OBSERVACOES_GERAIS.txt` | **Opcional** | Texto | Warn + continue | Sem observações |

---

## ✓ Checklist Pré-Execução

Antes de rodar a aplicação, verifique:

- [ ] ✅ `ADD_*.xlsx` existe no diretório
- [ ] ✅ `DIS_*.xlsx` existe no diretório
- [ ] ✅ `Valores Ctba.csv` existe no diretório
- [ ] ✅ `Subclasse.csv` existe no diretório
- [ ] ⚠️ (Opcional) `OBSERVACOES_GERAIS.txt` existe no diretório
- [ ] 📁 Todos estão no **MESMO diretório** que `gerar_relatorios_com_formulas_1.3.py`

---

## 🔍 Fluxo de Carregamento de Arquivos

```
Início da Aplicação
│
├─→ 1. load_price_data('Valores Ctba.csv')
│        ├─ Se existe: Carrega tabela de preços
│        └─ Se não existe: AVISO + continua com {}
│
├─→ 2. load_subclass_data('Subclasse.csv')
│        ├─ Se existe: Carrega lista de subclasses
│        └─ Se não existe: AVISO + continua com set()
│
├─→ 3. Lê OBSERVACOES_GERAIS.txt (opcional)
│        ├─ Se existe: Carrega observações
│        └─ Se não existe: AVISO + continua com ""
│
├─→ 4. find_latest_excel_file('ADD_')
│        ├─ Se existe: Encontra ADD_*.xlsx
│        └─ Se não existe: ❌ ERRO + exit()
│
├─→ 5. find_latest_excel_file('DIS_')
│        ├─ Se existe: Encontra DIS_*.xlsx
│        └─ Se não existe: ❌ ERRO + exit()
│
└─→ 6. Processa dados e gera Relatorios_Com_Formulas.xlsx
         (Continua só se ADD_ e DIS_ foram encontrados)
```

---

## 📝 Notas Importantes

### Função `find_latest_excel_file()`

Localiza o arquivo Excel mais recente com o prefixo especificado:

```python
def find_latest_excel_file(prefix):
    """
    Procura pelo arquivo Excel mais recente com o prefixo dado
    no diretório atual. Retorna o caminho do arquivo ou None
    se nenhum arquivo for encontrado.
    """
    search_pattern = f"{prefix}*.xlsx"
    files = glob.glob(search_pattern)
    if not files:
        return None
    return max(files, key=os.path.getmtime)  # Retorna o mais recente
```

**Comportamento:**
- Procura por todos os arquivos matching `ADD_*.xlsx` ou `DIS_*.xlsx`
- Retorna o **mais recente** por data de modificação
- Se múltiplos `ADD_*.xlsx` existem, usa o último modificado

---

## 🎯 Exemplo de Estrutura Mínima para Execução

```
/projeto_bassani/
├── gerar_relatorios_com_formulas_1.3.py    ← Script principal
│
├── ADD_Projeto RB 10837.xlsx                ← ✅ Obrigatório
├── DIS_Projeto RB 10837 DISTRATO.xlsx       ← ✅ Obrigatório
│
├── Valores Ctba.csv                         ← ✅ Obrigatório
├── Subclasse.csv                           ← ✅ Obrigatório
│
├── OBSERVACOES_GERAIS.txt                  ← ⚠️ Opcional
│
└── [Saída gerada]
    └── Relatorios_Com_Formulas.xlsx        ← Output do script
```

---

## ⚡ Comando para Executar

```bash
# No diretório com todos os arquivos acima:
python gerar_relatorios_com_formulas_1.3.py
```

Será solicitado ao usuário:
```
Orçamento com Guias e Montantes? (s/n):
```

---

Última atualização: 2026-03-25
