import pandas as pd
fn = 'ADD_RVT_2021_BASE_RB 10885 .xlsx'

try:
    df = pd.read_excel(fn, sheet_name='Paredes', header=1)
except Exception as e:
    print('Erro ao abrir arquivo:', e)
    raise

print('Colunas:', df.columns.tolist())

# Encontrar linhas onde a área não é numérica após limpeza

def clean(v):
    try:
        s = str(v).strip().replace(',', '.')
        if s in ['nan', 'None', '', 'nan']:
            return None
        return float(s)
    except:
        return None

problem = []
for idx, row in df.iterrows():
    raw = row.get('Área')
    if pd.isna(raw):
        continue
    val = clean(raw)
    if val is None:
        problem.append((idx, raw, row.get('Sistema Construtivo R. Bassani'), row.get('ID. Bloco/Torre')))

print('Total problemas:', len(problem))
for i, raw, tipo, bloco in problem[:50]:
    print(i, bloco, tipo, repr(raw))

mask = df['Sistema Construtivo R. Bassani'].astype(str).str.contains('TP395', na=False)
print('TP395 count', mask.sum())
print(df.loc[mask, ['ID. Bloco/Torre','Sistema Construtivo R. Bassani','Área']].head(20))
