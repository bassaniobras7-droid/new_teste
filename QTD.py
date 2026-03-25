from openpyxl import Workbook

# Criar uma nova planilha
wb = Workbook()

# =====================
# Aba 1 - Entrada de dados
# =====================
entrada = wb.active
entrada.title = "Entrada"

entrada.append(["Tipo de Parede", "Comprimento (m)", "Altura (m)", "Área (m²)", "Espessura do Perfil (mm)", "Espaçamento (m)"])
entrada.append(["MS90/600 Std/Std", 202, 3.2, "=B2*C2", 90, 0.6])

# =====================
# Aba 2 - Quantitativos
# =====================
quant = wb.create_sheet(title="Quantitativos")
quant.append(["Item", "Tipo", "Unidade", "Quantidade"])

# Itens e fórmulas
itens = [
    ("Montante 90 mm", "Perfil", "UD", "=(ARREDONDAR.PARA.CIMA(((Entrada!B2/Entrada!F2)+1)*(Entrada!C2/3),1))"),
    ("Guia 90 mm", "Perfil", "UD", "=(ARREDONDAR.PARA.CIMA((2*Entrada!B2)/3,1))"),
    ("Placa STD 12,5", "Placa", "UD", "=(ARREDONDAR.PARA.CIMA((Entrada!D2*2)/2,16,1))"),
    ("Parafuso T25", "Fixador", "CX", "=(ARREDONDAR.PARA.CIMA(Entrada!D2*35/100,1))"),
    ("Parafuso 4,2x13", "Fixador", "CX", "=(ARREDONDAR.PARA.CIMA(((Entrada!B2/Entrada!F2)+1)*10/100,1))"),
    ("Fixador Estrutura", "Fixador", "CX", "=(ARREDONDAR.PARA.CIMA(((2*Entrada!B2)/3)/10,1))"),
    ("Fita Juntas 50 mm", "Fita", "RL", "=(ARREDONDAR.PARA.CIMA(((Entrada!D2*2)/2)/150,1))"),
    ("Fita Canto Metálica 30 m", "Fita", "RL", "=(ARREDONDAR.PARA.CIMA((Entrada!C2*4)/30,1))"),
    ("Fita Isolamento 90 mm", "Fita", "RL", "=(ARREDONDAR.PARA.CIMA((Entrada!B2*2)/10,1))"),
    ("Massa para Juntas 20 kg", "Massa", "BD", "=(ARREDONDAR.PARA.CIMA((Entrada!D2*2)/4/20,1))")
]

for item in itens:
    quant.append(item)

# =====================
# Salvar arquivo
# =====================
wb.save("Estimativa_Knauf_MS90_600.xlsx")
print("Arquivo Excel criado com sucesso!")
