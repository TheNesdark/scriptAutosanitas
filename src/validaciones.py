import pandas as pd
import re

def validar_telefono(telefono):
    if pd.isna(telefono):
        return ""
    
    telefono_str = str(telefono)
    digitos = re.sub(r'\D', '', telefono_str)
    
    if not digitos:
        return ""
    
    if len(digitos) > 10:
        digitos = digitos[:10]
    
    if len(digitos) < 10:
        digitos = digitos.zfill(10)
    
    return digitos
