import pandas as pd
import re

def clean_numeric_column(series):
    return pd.to_numeric(series.astype(str).str.replace(',', '.'), errors='coerce').fillna(0)

def natural_sort_key(s):
    return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', str(s))]
