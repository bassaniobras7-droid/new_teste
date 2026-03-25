import re

# Constantes extraídas de modelo.py para a lógica LIX-J' #
LIX_J_INCLUSION_WORDS = [
    "forro", "drymold", "sanca", "cortineiro", "faixa de acabamento",
    "fechamento vertical", "revestimento",
]

LIX_J_EXCLUSION_WORDS = [
    "cim", "cimentícia", "f127", "tabica", "modular", "amf", "knauf",
    "star", "nrc", "cleaneo", "osb", "suporte", "madeira", "estrutural",
    "alçapão", "guia", "lsf"
]

def _lix_j_extract_factor_cm(desc):
    """Extrai um fator de soma de dimensões em cm."""
    m = re.search(r"(\d+(?:[.,]?\d+)?(?:\s*[xX]\s*\d+(?:[.,]?\d+)?)+)\s*cm", desc)
    if not m: return None
    parts = re.split(r"[xX]", m.group(1))
    values = [float(p.strip().replace(",", ".")) for p in parts]
    return sum(values) / 100 if values else None

def _lix_j_extract_factor_m(desc):
    """Extrai um fator de multiplicação em m."""
    m = re.search(r"(\d+[.,]?\d*)\s*m\b", desc)
    if not m: return None
    return float(m.group(1).replace(",", "."))

def _lix_j_should_include(desc):
    """Verifica se uma descrição deve ser incluída no cálculo de LIX-J'."""
    if not isinstance(desc, str): return False
    d = desc.lower().strip()
    if not any(re.search(r'\b' + re.escape(p) + r'\b', d) for p in LIX_J_INCLUSION_WORDS):
        return False
    exclusion_list = LIX_J_EXCLUSION_WORDS
    if 'parede' in d:
        exclusion_list = [p for p in LIX_J_EXCLUSION_WORDS if p not in ["cim", "lsf"]]
    if any(re.search(r'\b' + re.escape(p) + r'\b', d) for p in exclusion_list):
        return False
    return True
