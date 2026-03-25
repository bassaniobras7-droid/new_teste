import win32com.client
import os

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

base = "C:\\Users\\Wesllen.Santana\\Downloads\\bassani\\"

files = [
    ("Relatorios_Com_Formulas.xlsx", "Relatorios_Com_Formulas_decrypted.xlsx"),
    ("Relatorios_Com_Formulas. novo layoutxlsx.xlsx", "Relatorios_Com_Formulas_novo_decrypted.xlsx"),
]

for src, dst in files:
    src_path = base + src
    dst_path = base + dst
    print(f"Abrindo: {src}")
    try:
        wb = excel.Workbooks.Open(src_path, UpdateLinks=False, ReadOnly=True)
        sheets = []
        for i in range(1, wb.Sheets.Count + 1):
            sheets.append(wb.Sheets(i).Name)
        print(f"  Abas: {sheets}")
        # Salvar como xlsx sem senha
        wb.SaveAs(dst_path, FileFormat=51)  # 51 = xlOpenXMLWorkbook
        wb.Close(False)
        print(f"  Salvo: {dst}")
    except Exception as e:
        print(f"  ERRO: {e}")
        try:
            wb.Close(False)
        except:
            pass

excel.Quit()
print("Concluido!")
