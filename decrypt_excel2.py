import win32com.client
import os

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

base = "C:\\Users\\Wesllen.Santana\\Downloads\\bassani\\"

files = [
    ("Relatorios_Com_Formulas.xlsx", "Relatorios_Com_Formulas_dec2.xlsx"),
    ("Relatorios_Com_Formulas. novo layoutxlsx.xlsx", "Relatorios_Com_Formulas_novo_dec2.xlsx"),
]

for src, dst in files:
    src_path = base + src
    dst_path = base + dst
    print(f"Abrindo: {src}")
    try:
        # Remover arquivo destino se existir
        if os.path.exists(dst_path):
            os.remove(dst_path)

        wb = excel.Workbooks.Open(src_path, UpdateLinks=False, ReadOnly=True)
        sheets = []
        for i in range(1, wb.Sheets.Count + 1):
            sheets.append(wb.Sheets(i).Name)
        print(f"  Abas: {sheets}")

        # 51 = xlOpenXMLWorkbook (xlsx sem macros)
        # Precisamos passar o caminho completo Windows style
        wb.SaveAs(dst_path, 51, '', '', False, False)
        wb.Close(False)
        print(f"  Salvo: {dst}")

        # Verificar header
        with open(dst_path, 'rb') as f:
            h = f.read(8)
            print(f"  Header: {h.hex()} ({'ZIP/xlsx' if h[:4] == b'PK\x03\x04' else 'OLE2 ainda'})")
    except Exception as e:
        print(f"  ERRO: {e}")
        try:
            wb.Close(False)
        except:
            pass

excel.Quit()
print("Concluido!")
