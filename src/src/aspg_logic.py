import re

def build_aspg_formula_parts(summary_data, parede_metragem_cells_map):
    """
    Constrói as partes da fórmula para o item ASP-G.

    Args:
        summary_data (dict): O dicionário com os dados do resumo.
        parede_metragem_cells_map (dict): Um mapa de tipo_code para a coordenada da célula de metragem.

    Returns:
        list: Uma lista de strings, cada uma sendo uma parte da fórmula do Excel.
    """
    aspg_formula_parts = []
    pattern = re.compile(r'(?:MS|MD)(?:48|70|90|140)', re.IGNORECASE)

    for tipo_code, data in summary_data.items():
        # A lógica original depende das 'parede_metragem_cells', então vamos filtrar por elas.
        if tipo_code not in parede_metragem_cells_map:
            continue

        desc = data.get('Descricao')
        if not desc:
            continue

        matches = pattern.findall(str(desc))
        if matches:
            count = len(matches)
            coord = parede_metragem_cells_map[tipo_code]
            if count > 0:
                aspg_formula_parts.append(f"({coord}*{count})")

    return aspg_formula_parts
