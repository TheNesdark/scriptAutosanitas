import pandas as pd
import datetime as date

NIT = "812000344"
CSBH = "11033"
NPR = "ESE HOSPITAL LOCAL DE MONTELIBANO"

datos = []
ex = pd.read_excel("base5.xlsx", engine="openpyxl", header=1)

for _, fila in ex.iterrows():

    datos.append({
        "FECHA_ENVIO": date.datetime.now(),
        "NIT_PRESTADOR_REMITENTE": NIT,
        "CODIGO_SUCURSAL_BH": CSBH,
        "NOMBRE_PRESTADOR_REMITENTE": NPR,
        "TIPO_IDENTIFICACION_DEL_AFILIADO": fila["Tipo_Identificación_del_Afiliado"],
        "NUMERO_IDENTIFICACION_DEL_AFILIADO": fila["Numero_de_Identificación_del_Afiliado"],
        "NOMBRE_PACIENTE": fila["Nombre_Paciente"],
        "TELEFONO_CELULAR_1": fila["Telefono_Celular_1"],
        "TELEFONO_CELULAR_2": fila["Telefono_Celular_2"],
        "FECHA_ATENCION": fila["Fecha_Atencion"],
        "CIE10": fila["cie10"],
        "NOMBRE_MEDICO": fila["Nombre_Medico"],
        "CODIGO_ESPECIALIDAD_REMITENTE": fila["Codigo_Especialidad_Remitente"],
        "ESPECIALIDAD_REMITENTE": fila["Especialidad_Remitente"],
        "CODIGO_CUPS_PRESTACION": fila["Codigo_CUPS_Prestacion"],
        "DESCRIPCION_PRESTACION": fila["Descripcion_Prestacion"],
        "CANTIDAD": fila["Cantidad"],
        "JUSTIFICACION_CLINICA": fila["Justificacion_Clinica"],
        "EDAD_GESTACIONAL_SEMANAS": fila["Edad_Gestacional"],
        "ANESTESIA": fila["Anestesia_A"],
        "SEDACION": fila["Sedación_S"],
        "CONTRASTE": fila["Contraste_T"],
        "COMPARATIVO": fila["Comparativo_C"],
        "BILATERAL": fila["Bilateral_B"],
        "NOAPLICA": "X",
    })

nv = pd.DataFrame(datos)

nv.to_excel("NV.xlsx", index=False, engine="openpyxl")