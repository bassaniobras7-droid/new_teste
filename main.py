from src.data_processing import (
    load_price_data,
    load_subclass_data,
    process_client_data,
    apply_wool_logic,
    calculate_and_add_derived_items,
    process_summary_data
)
from src.excel_writer import write_excel_with_formulas

def main():
    print("Carregando dados de preços e subclasses...")
    precos = load_price_data()
    subclasses = load_subclass_data()

    dados_cliente_normal = {}
    dados_resumo_normal = {}
    dados_cliente_distratado = {}
    dados_resumo_distratado = {}

    for scope in ['normal', 'distratado']:
        print(f"Processando dados para: {scope.upper()}...")
        prefix = '__' if scope == 'distratado' else ''
        client_data = process_client_data(
            precos, subclasses,
            f'{prefix}Tabela de Forro.csv',
            f'{prefix}Tabela de Modelo Genérico.csv',
            f'{prefix}Tabela de Paredes.csv'
        )
        print("  Aplicando lógicas de negócio...")
        client_data = apply_wool_logic(client_data)
        client_data = calculate_and_add_derived_items(client_data, precos)
        summary_data = process_summary_data(client_data)
        
        if scope == 'normal':
            dados_cliente_normal, dados_resumo_normal = client_data, summary_data
        else:
            dados_cliente_distratado, dados_resumo_distratado = client_data, summary_data

    print("Gerando arquivo Excel...")
    write_excel_with_formulas(
        dados_resumo_normal, dados_resumo_distratado,
        dados_cliente_normal, dados_cliente_distratado,
        precos,
        filename='Relatorios_Com_Formulas.xlsx'
    )

if __name__ == '__main__':
    main()
