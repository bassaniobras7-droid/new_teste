from src.data_processing import load_price_data, load_subclass_data, process_client_data, apply_wool_logic, calculate_and_add_derived_items, process_summary_data
from src.excel_writer import write_excel_with_formulas

precos = load_price_data()
subclasses = load_subclass_data()

client_normal = process_client_data(precos, subclasses, 'Tabela de Forro.csv', 'Tabela de Modelo Genérico.csv', 'Tabela de Paredes.csv')
client_normal = apply_wool_logic(client_normal)
client_normal = calculate_and_add_derived_items(client_normal, precos)
summary_normal = process_summary_data(client_normal)

client_distratado = process_client_data(precos, subclasses, '__Tabela de Forro.csv', '__Tabela de Modelo Genérico.csv', '__Tabela de Paredes.csv')
client_distratado = apply_wool_logic(client_distratado)
client_distratado = calculate_and_add_derived_items(client_distratado, precos)
summary_distratado = process_summary_data(client_distratado)

write_excel_with_formulas(summary_normal, summary_distratado, client_normal, client_distratado, precos, filename='Relatorios_Com_Formulas_test.xlsx')
print('done')
