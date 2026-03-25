def build_lptub_formula(aspg_metragem_coord):
    """
    Constrói a fórmula para o item LP-TUB.

    Args:
        aspg_metragem_coord (str): A coordenada da célula de metragem do ASP-G.

    Returns:
        str: A fórmula do Excel para LP-TUB.
    """
    if not aspg_metragem_coord:
        return 0
    return f"=CEILING({aspg_metragem_coord}*1504/17262.3655,1)"
