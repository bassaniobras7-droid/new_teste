from pathlib import Path
p = Path('gerar_relatorios_com_formulas_1.3.py')
lines = p.read_text(encoding='utf8').splitlines()
for i in range(1025, 1046):
    print(i, lines[i-1])
