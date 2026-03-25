import pandas as pd
import math
import re
from src.utils import clean_numeric_column

# ==============================================================================
# FUNÇÕES DE CARREGAMENTO DE DADOS
# ==============================================================================

def load_price_data(filename='Valores Ctba.csv'):
    try:
        df_prices = pd.read_csv(filename, sep=';', header=0)
        df_prices.dropna(subset=['Tipo R. Bassani'], inplace=True)
        price_data = {}
        for _, row in df_prices.iterrows():
            tipo_code = row['Tipo R. Bassani']
            price_data[tipo_code] = {
                'Un': row.get('Un'),
                'Valor': clean_numeric_column(pd.Series([row.get('Valor do Material + MO')]))[0],
                'Custo MO': clean_numeric_column(pd.Series([row.get('Custo MO à Pagar')]))[0],
                'Descricao': row.get('Forros')
            }
        return price_data
    except FileNotFoundError:
        print(f"AVISO: Arquivo de preços '{filename}' não encontrado.")
        return {}
    except Exception as e:
        print(f"ERRO ao carregar o arquivo de preços: {e}")
        return {}

def load_subclass_data(filename='Subclasse.csv'):
    try:
        df_subclass = pd.read_csv(filename, sep=';', header=0)
        df_subclass.dropna(subset=['Tipo'], inplace=True)
        return set(df_subclass['Tipo'].astype(str).str.strip())
    except FileNotFoundError:
        print(f"AVISO: Arquivo de subclasses '{filename}' não encontrado.")
        return set()
    except Exception as e:
        print(f"ERRO ao carregar o arquivo de subclasses: {e}")
        return set()

# ==============================================================================
# PROCESSAMENTO DE DADOS
# ==============================================================================

def process_summary_data(client_data):
    summary_agg = {}
    for client_id, client_items in client_data.items():
        subclass_total_qty = sum(item_data.get('BaseQuantity', 0) for item_data in client_items.values() if item_data.get('is_subclass', False))

        for item_data in client_items.values():
            cat = item_data['Categoria']
            summary_cat = 'Paredes' if cat in ['Parede', 'Revestimento'] else ('Forros' if cat == 'Forro' else cat)

            for carenagem_item in item_data.get('carenagem_items', {}).values():
                carenagem_tipo_code = carenagem_item['Tipo Code']
                if carenagem_tipo_code not in summary_agg:
                    summary_agg[carenagem_tipo_code] = {'Quantidade': 0, 'Descricao': carenagem_item['Descricao'], 'Categoria': summary_cat}
                summary_agg[carenagem_tipo_code]['Quantidade'] += carenagem_item['Quantidade']

            for insulation_item in item_data.get('insulation_items', {}).values():
                insul_tipo_code = insulation_item['Tipo Code']
                if insul_tipo_code not in summary_agg:
                    summary_agg[insul_tipo_code] = {'Quantidade': 0, 'Descricao': insulation_item['Descricao'], 'Categoria': 'Isolamento'}
                summary_agg[insul_tipo_code]['Quantidade'] += insulation_item['Quantidade']

            main_tipo_code = item_data['Tipo Code']
            qty = item_data['BaseQuantity']
            if main_tipo_code == 'FT16' and subclass_total_qty > 0:
                qty -= subclass_total_qty
            
            if main_tipo_code not in summary_agg:
                summary_agg[main_tipo_code] = {'Quantidade': 0, 'Descricao': item_data['Descricao'], 'Categoria': summary_cat}
            summary_agg[main_tipo_code]['Quantidade'] += qty
    return summary_agg

def process_client_data(price_data, subclass_types, forro_file, generico_file, paredes_file):
    data_by_client = {}

    def add_or_aggregate_item(client_id, item_data, has_insulation=False, insulation_data=None, carenagem_data=None):
        if client_id not in data_by_client: data_by_client[client_id] = {}
        key = (item_data['Tipo Code'], item_data['Categoria'], has_insulation)
        if key not in data_by_client[client_id]:
            data_by_client[client_id][key] = {
                'Tipo Code': item_data['Tipo Code'], 'Descricao': item_data.get('Descricao'), 'BaseQuantity': 0,
                'Categoria': item_data['Categoria'], 'has_insulation': has_insulation, 'insulation_items': {},
                'carenagem_items': {}, 'is_subclass': item_data.get('is_subclass', False)
            }
        if item_data.get('Descricao') and not data_by_client[client_id][key].get('Descricao'):
            data_by_client[client_id][key]['Descricao'] = item_data.get('Descricao')
        data_by_client[client_id][key]['BaseQuantity'] += item_data.get('Quantidade', 0)
        if item_data.get('is_subclass', False): data_by_client[client_id][key]['is_subclass'] = True
        if insulation_data:
            insul_key = insulation_data['Tipo Code']
            if insul_key not in data_by_client[client_id][key]['insulation_items']:
                data_by_client[client_id][key]['insulation_items'][insul_key] = {'Tipo Code': insul_key, 'Descricao': insulation_data['Descricao'], 'Quantidade': 0, 'is_la_dupla': insulation_data.get('is_la_dupla', False)}
            data_by_client[client_id][key]['insulation_items'][insul_key]['Quantidade'] += insulation_data['Quantidade']
        if carenagem_data:
            carenagem_key = carenagem_data['Descricao']
            if carenagem_key not in data_by_client[client_id][key]['carenagem_items']:
                data_by_client[client_id][key]['carenagem_items'][carenagem_key] = {'Tipo Code': carenagem_data['Tipo Code'], 'Descricao': carenagem_data['Descricao'], 'Quantidade': 0}
            data_by_client[client_id][key]['carenagem_items'][carenagem_key]['Quantidade'] += carenagem_data['Quantidade']

    for file_path, process_func in [(forro_file, 'forro'), (generico_file, 'generico'), (paredes_file, 'paredes')]:
        try:
            df = pd.read_csv(file_path, sep=';', header=1)
            df.dropna(subset=['ID. Bloco/Torre'], inplace=True)
            if process_func == 'forro':
                df['Área'] = clean_numeric_column(df['Área'])
                df['Perímetro'] = clean_numeric_column(df['Perímetro'])
                for _, row in df.iterrows():
                    tipo_code = str(row['Sistema Construtivo R. Bassani']).strip()
                    unit = price_data.get(tipo_code, {}).get('Un', '')
                    qty = row['Área'] if unit == 'm²' else (row['Perímetro'] if unit == 'm' else row['Área'])
                    add_or_aggregate_item(row['ID. Bloco/Torre'], {'Tipo Code': tipo_code, 'Descricao': price_data.get(tipo_code, {}).get('Descricao', row['Tipo']), 'Quantidade': qty, 'Categoria': 'Forro', 'is_subclass': tipo_code in subclass_types})
            elif process_func == 'generico':
                df['Contador'] = clean_numeric_column(df['Contador'])
                for _, row in df.iterrows():
                    tipo_code = str(row['Sistema Construtivo R. Bassani']).strip()
                    add_or_aggregate_item(row['ID. Bloco/Torre'], {'Tipo Code': tipo_code, 'Descricao': price_data.get(tipo_code, {}).get('Descricao', row['Tipo']), 'Quantidade': row['Contador'], 'Categoria': 'Parede' if row['Classe'] == 'Parede' else 'Forro', 'is_subclass': tipo_code in subclass_types})
            elif process_func == 'paredes':
                df['Área'] = clean_numeric_column(df['Área'])
                df['Altura desconectada'] = clean_numeric_column(df['Altura desconectada'])
                df['Comprimento'] = clean_numeric_column(df['Comprimento'])
                for _, row in df.iterrows():
                    client_id, tipo_code, categoria = row['ID. Bloco/Torre'], str(row['Sistema Construtivo R. Bassani']).strip(), row['Classe']
                    is_subclass, osb_perfil = tipo_code in subclass_types, str(row.get('OSB/Perfil')).strip()
                    is_carenagem, is_la_dupla = osb_perfil == 'Carenagem', osb_perfil == 'Lã dupla'
                    unit, row_quantity = price_data.get(tipo_code, {}).get('Un', ''), 0
                    if categoria in ['Parede', 'Revestimento']:
                        if unit == 'm²':
                            row_quantity = row['Área']
                        elif unit == 'm':
                            row_quantity = row['Altura desconectada']
                        else:
                            row_quantity = 0
                    elif categoria == 'Forro':
                        if unit == 'm²':
                            row_quantity = row['Área']
                        elif unit == 'm':
                            row_quantity = row['Comprimento']
                        else:
                            row_quantity = 0
                    item_data, carenagem_data, insulation_data = None, None, None
                    has_insulation = pd.notna(row['Sistema de Isolamento']) and str(row['Sistema de Isolamento']).strip() != ''
                    if is_carenagem:
                        tipo_code_carenagem = f"{tipo_code}-CAR"
                        desc = price_data.get(tipo_code_carenagem, {}).get('Descricao', row['Tipo'])
                        # Check if 'desc' already ends with '(Carenagem)'
                        if not desc.strip().endswith('(Carenagem)'):
                            desc = f"{desc} (Carenagem)"
                        carenagem_data = {'Tipo Code': tipo_code_carenagem, 'Descricao': desc, 'Quantidade': row_quantity}
                        item_data = {'Tipo Code': tipo_code, 'Categoria': categoria, 'Quantidade': 0, 'is_subclass': is_subclass}
                    else:
                        item_data = {'Tipo Code': tipo_code, 'Descricao': price_data.get(tipo_code, {}).get('Descricao', row['Tipo']), 'Quantidade': row_quantity, 'Categoria': categoria, 'is_subclass': is_subclass}
                        if has_insulation:
                            tipo_code_insul = row['Sistema de Isolamento']
                            insulation_data = {'Tipo Code': tipo_code_insul, 'Descricao': price_data.get(tipo_code_insul, {}).get('Descricao', f"Isolamento {tipo_code_insul}"), 'Quantidade': row['Área'] * (2 if is_la_dupla else 1), 'is_la_dupla': is_la_dupla}
                    add_or_aggregate_item(client_id, item_data, has_insulation=has_insulation, insulation_data=insulation_data, carenagem_data=carenagem_data)
        except FileNotFoundError:
            print(f"AVISO: Arquivo '{file_path}' não encontrado.")
    return data_by_client

# ==============================================================================
# AJUSTES E REGRAS DE NEGÓCIO
# ==============================================================================

def apply_wool_logic(data_by_client):
    for client_id, client_items in data_by_client.items():
        grouped_keys = {}
        for key in list(client_items.keys()):
            tipo_code, categoria, has_insulation = key
            base_key = (tipo_code, categoria)
            if base_key not in grouped_keys: grouped_keys[base_key] = []
            grouped_keys[base_key].append(key)

        for keys in grouped_keys.values():
            if len(keys) == 2:
                key_with_wool = next((k for k in keys if k[2]), None)
                key_without_wool = next((k for k in keys if not k[2]), None)
                
                if key_with_wool and key_without_wool:
                    item_com_la, item_sem_la = client_items[key_with_wool], client_items[key_without_wool]
                    original_qty_com_la, qty_sem_la = item_com_la['BaseQuantity'], item_sem_la['BaseQuantity']
                    item_com_la['FormulaBase'] = qty_sem_la
                    item_com_la['BaseQuantity'] = qty_sem_la + original_qty_com_la
                    if item_com_la['insulation_items']:
                        insul_key = next(iter(item_com_la['insulation_items']))
                        if not item_com_la['insulation_items'][insul_key].get('is_la_dupla', False):
                            item_com_la['insulation_items'][insul_key]['Quantidade'] = original_qty_com_la
                    item_com_la['logic_applied'] = True
                    item_com_la['carenagem_items'].update(item_sem_la.get('carenagem_items', {}))
                    del client_items[key_without_wool]
    return data_by_client

def calculate_and_add_derived_items(data_by_client, price_data):
    known_profiles = ['MS48', 'MS70', 'MS90', 'F47', 'MS140']
    rules = { 'SHF-40': {'m': 0.35, 'f': ['Ru', '/400'], 'c': True}, 'SH48-40': {'m': 0.45, 'f': ['MS48', '/400'], 'c': True}, 'SH48-60': {'m': 0.45, 'f': ['MS48', '/600'], 'c': True},
              'SH70-40': {'m': 0.45, 'f': ['MS70', '/400'], 'c': True}, 'SH70-60': {'m': 0.45, 'f': ['MS70', '/600'], 'c': True}, 'SH90-40': {'m': 0.45, 'f': ['MS90', '/400'], 'c': True},
              'SH90-60': {'m': 0.45, 'f': ['MS90', '/600'], 'c': True}, 'SH140-40': {'m': 0.45, 'f': ['MS140', '/400'], 'c': True}, 'SH140-60': {'m': 0.45, 'f': ['MS140', '/600'], 'c': True} }

    for client_id, client_items in data_by_client.items():
        contributions = {code: [] for code in rules}
        for item_key, item in client_items.items():
            if item.get('Categoria') in ['Parede', 'Revestimento'] and item.get('Tipo Code', '').startswith('TP') and item.get('Descricao') and item.get('BaseQuantity', 0) > 0:
                for code, rule in rules.items():
                    if all(f in item['Descricao'] for f in rule['f']):
                        count = item['Descricao'].count('Ru') if code == 'SHF-40' else (item['Descricao'].count(next((p for p in known_profiles if re.search(r'\d+', code).group(0) in p), None)) if rule['c'] and next((p for p in known_profiles if re.search(r'\d+', code).group(0) in p), None) else 1)
                        count = 1 if count == 0 else count
                        if count > 0: contributions[code].append({'item_key': item_key, 'count': count})
        
        for code, contribs in contributions.items():
            if contribs:
                total_contrib_val = sum(client_items[c['item_key']].get('BaseQuantity', 0) * c['count'] for c in contribs)
                final_qty = math.ceil(total_contrib_val * rules[code]['m'])
                if final_qty > 0:
                    data_by_client[client_id][(code, 'Guias e Montantes', False)] = {
                        'Tipo Code': code, 'Descricao': price_data.get(code, {}).get('Descricao', f'Montante {code}'), 'BaseQuantity': final_qty, 'Categoria': 'Guias e Montantes',
                        'has_insulation': False, 'insulation_items': {}, 'carenagem_items': {}, 'is_subclass': False,
                        'formula_contributors': contribs, 'formula_multiplier': rules[code]['m']
                    }
    return data_by_client
