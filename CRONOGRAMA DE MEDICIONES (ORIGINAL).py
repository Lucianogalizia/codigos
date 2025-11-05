#!/usr/bin/env python
# coding: utf-8

# In[1]:


#EXPORTA LA MATRIZ DE AIB - PRIMER CODIGO A EJECUTAR

# -*- coding: utf-8 -*-
# Matriz de Criticidad de AIBs (fuera de Power BI) - cx_Oracle
# Hoja 1: columnas solicitadas de FDP_DINA transformada
# Hoja 2: consulta ACCIONES + columna ETIQUETA (clasificación por reglas + re-etiquetado)

from __future__ import annotations
from pathlib import Path
import re
import unicodedata
import numpy as np
import pandas as pd
import cx_Oracle

# ================== CREDENCIALES / RUTAS ==================
ORA_USER    = "RY33872"
ORA_PASS    = "Contraseña_0725"
ORA_HOST    = "slplpgmoora03"
ORA_PORT    = 1527
ORA_SERVICE = "psfu"

EXPO_PATH   = r"C:\Users\ry16123\Downloads\EXPOSICION-new 1.xlsx"
EXPO_SHEET  = None  # auto-detección de hoja

# ================== QUERIES ==================
SQL_FDP_DINA = """
SELECT DBU_FIC_ORG_ESTRUCTURAL.NOMBRE_CORTO,DBU_FIC_ORG_ESTRUCTURAL.NOMBRE_CORTO_POZO,
       DBU_FIC_ORG_ESTRUCTURAL.NIVEL_1,
       DBU_FIC_ORG_ESTRUCTURAL.NIVEL_2,
       DBU_FIC_ORG_ESTRUCTURAL.NIVEL_3,
       DBU_FIC_ORG_ESTRUCTURAL.NIVEL_4,
       DBU_FIC_ORG_ESTRUCTURAL.NIVEL_5,
       FIC_ULTIMO_DINAMOMETRO.FECHA_HORA,
       FIC_ULTIMO_DINAMOMETRO.REALIZADO_POR,
       FIC_ULTIMO_DINAMOMETRO.AIB_MARCA_Y_DESC_API,
       FIC_ULTIMO_DINAMOMETRO.AIBRR_TORQUE_MAXIMO_REDUCTOR,
       FIC_ULTIMO_DINAMOMETRO.AIBEB_TORQUE_MAXIMO_REDUCTOR,
       FIC_ULTIMO_DINAMOMETRO.AIBRR_TORQUE_DISPONIBLE,
       FIC_ULTIMO_DINAMOMETRO.BBA_LLENADO_DE_BOMBA/100 AS LLENAD_BOMBA_PCT,
       FIC_ULTIMO_DINAMOMETRO.AIBRE_SOLICITACION_DE_ESTRUCT/100 AS ESTRUCTURA_PCT,
       VTOW_WELL_LAST_CONTROL_DET.PROD_OIL,
       VTOW_WELL_LAST_CONTROL_DET.PROD_GAS,
       VTOW_WELL_LAST_CONTROL_DET.PROD_WAT,
       FIC_ULTIMO_DINAMOMETRO.AIB_GPM,
       FIC_ULTIMO_DINAMOMETRO.AIB_DIAMETRO_POLEA_REDUCTOR,
       FIC_ULTIMO_DINAMOMETRO.MOTOR_DIAMETRO_POLEA
FROM DISC_ADMINS.DBU_FIC_ORG_ESTRUCTURAL DBU_FIC_ORG_ESTRUCTURAL,
     DISC_ADMINS.VTOW_WELL_LAST_CONTROL_DET VTOW_WELL_LAST_CONTROL_DET,
     DISC_ADMINS.FIC_ULTIMO_DINAMOMETRO FIC_ULTIMO_DINAMOMETRO
WHERE ((VTOW_WELL_LAST_CONTROL_DET.COMP_SK = DBU_FIC_ORG_ESTRUCTURAL.COMP_SK)
       AND (FIC_ULTIMO_DINAMOMETRO.COMP_SK = DBU_FIC_ORG_ESTRUCTURAL.COMP_SK))
  AND (DBU_FIC_ORG_ESTRUCTURAL.ESTADO = 'Produciendo')
  AND ((FIC_ULTIMO_DINAMOMETRO.AIBRE_SOLICITACION_DE_ESTRUCT/100) <> 0)
  AND ((FIC_ULTIMO_DINAMOMETRO.BBA_LLENADO_DE_BOMBA/100) <> 0)
  AND (DBU_FIC_ORG_ESTRUCTURAL.MET_PROD = 'Bombeo Mecánico')
"""

SQL_PRODUCCION = """
SELECT UPPER(DBU_FIC_ORG_ESTRUCTURAL.NIVEL_4) AS NIVEL_4,
       SUM(VTOW_WELL_LAST_CONTROL_DET.PROD_OIL) AS PROD_OIL_1
FROM DISC_ADMINS.DBU_FIC_ORG_ESTRUCTURAL DBU_FIC_ORG_ESTRUCTURAL,
     DISC_ADMINS.VTOW_WELL_LAST_CONTROL_DET VTOW_WELL_LAST_CONTROL_DET,
     DISC_ADMINS.FIC_ULTIMO_DINAMOMETRO FIC_ULTIMO_DINAMOMETRO
WHERE ((VTOW_WELL_LAST_CONTROL_DET.COMP_SK = DBU_FIC_ORG_ESTRUCTURAL.COMP_SK)
       AND (FIC_ULTIMO_DINAMOMETRO.COMP_SK = DBU_FIC_ORG_ESTRUCTURAL.COMP_SK))
  AND (DBU_FIC_ORG_ESTRUCTURAL.ESTADO = 'Produciendo')
  AND ((FIC_ULTIMO_DINAMOMETRO.AIBRE_SOLICITACION_DE_ESTRUCT/100) <> 0)
  AND ((FIC_ULTIMO_DINAMOMETRO.BBA_LLENADO_DE_BOMBA/100) <> 0)
  AND (DBU_FIC_ORG_ESTRUCTURAL.MET_PROD = 'Bombeo Mecánico')
GROUP BY UPPER(DBU_FIC_ORG_ESTRUCTURAL.NIVEL_4)
"""

SQL_ACCIONES = """
SELECT DISTINCT
  DBU_FIC_ORG_ESTRUCTURAL.NIVEL_3,
  FIC_ACCIONES.ACTIVIDAD,
  FIC_ACCIONES.ESTADO,
  FIC_ACCIONES."ESTADO AUTORIZACIÓN",
  FIC_ACCIONES.FECHAACCION,
  FIC_ACCIONES."FECHA AUTORIZACION",
  FIC_ACCIONES.FECHAREALIZACION,
  FIC_ACCIONES.JUSTIFICACION,
  FIC_ACCIONES.OBJETIVO,
  FIC_ACCIONES.OBSERVACION,
  FIC_ACCIONES.ORIGEN,
  FIC_ACCIONES."INCREMENTO BRUTA",
  FIC_ACCIONES."INCREMENTO PETRÓLEO",
  FIC_ACCIONES.RECURSO,
  FIC_ACCIONES.SUBACTIVIDAD,
  FIC_ACCIONES."USUARIO AUTORIZANTE",
  FIC_ACCIONES."USUARIO CREADOR",
  DBU_FIC_ORG_ESTRUCTURAL.NOMBRE_POZO,DBU_FIC_ORG_ESTRUCTURAL.NOMBRE_CORTO,
  DBU_FIC_ORG_ESTRUCTURAL.NOMBRE_CORTO_POZO
FROM DISC_ADMINS.DBU_FIC_ORG_ESTRUCTURAL DBU_FIC_ORG_ESTRUCTURAL,
     DISC_ADMINS.FIC_ACCIONES FIC_ACCIONES
WHERE (FIC_ACCIONES.CLAVEPOZO = DBU_FIC_ORG_ESTRUCTURAL.COMP_SK)
  AND (DBU_FIC_ORG_ESTRUCTURAL.COMP_SK = FIC_ACCIONES.CLAVEPOZO)
  AND (FIC_ACCIONES.FECHAACCION >= TO_DATE('20230101000000','YYYYMMDDHH24MISS'))
  AND (DBU_FIC_ORG_ESTRUCTURAL.NIVEL_3 IN ('Las Heras CG - Canadon Escondida','Los Perales','El Guadal','Seco Leon - Pico Truncado'))
"""

# ================== CONEXIÓN ORACLE (cx_Oracle) ==================
def get_connection():
    dsn = cx_Oracle.makedsn(ORA_HOST, ORA_PORT, service_name=ORA_SERVICE)
    return cx_Oracle.connect(ORA_USER, ORA_PASS, dsn, encoding="UTF-8")

def read_sql_df(conn, sql: str) -> pd.DataFrame:
    return pd.read_sql(sql, con=conn)

# ================== EXPO (auto-detección) ==================
def load_exposicion(path: str, sheet: str | None) -> pd.DataFrame:
    xls = pd.ExcelFile(path, engine="openpyxl")
    print(f"Hojas encontradas en '{path}': {xls.sheet_names}")

    if sheet and sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet)
    else:
        candidatos = ["EXPOSICION", "EXPOSICIÓN", "Exposicion", "Hoja1", "Sheet1"]
        df = None
        for s in candidatos:
            if s in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=s)
                break
        if df is None:
            for s in xls.sheet_names:
                tmp = pd.read_excel(xls, sheet_name=s)
                cols_norm = [str(c).strip().lower() for c in tmp.columns]
                if "denominacion api" in cols_norm:
                    df = tmp
                    print(f"Usando hoja detectada por columna: {s}")
                    break
        if df is None:
            raise ValueError(f"No pude encontrar la hoja de EXPOSICIÓN. Hojas: {xls.sheet_names}")

    if "Cantidad" in df.columns:
        df = df.drop(columns=["Cantidad"])

    # Normalizar columna clave
    col_api = None
    for c in df.columns:
        if str(c).strip().lower() == "denominacion api":
            col_api = c; break
    if col_api is None:
        raise ValueError("No encuentro la columna 'Denominacion API' en la hoja seleccionada.")

    df[col_api] = df[col_api].astype(str).str.upper().str.strip()
    df = df.rename(columns={col_api: "Denominacion API"})

    for c in ["Marca","Tipo","Torque max","SENSIBILIDAD CR","SENSIBILIDAD EST"]:
        if c not in df.columns:
            df[c] = np.nan
    return df

# ================== Utilidades NLP/Reglas ==================
_num = r"(\d+(?:[.,]\d+)?)"

MODEL_WORDS = (
    r"(lufkin|vulcan|siam|pump ?jack|wuel?fel|darco|maxii?|weatherford|"
    r"mel ?altium|altium|vc\d{3}|lm[-\s]?\d+|m\d{3}|c\d{3}|rm[-\s]?\d+)"
)

def _normalize(s):
    if s is None or (isinstance(s, float) and np.isnan(s)): return ""
    s = str(s)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.lower()

def _to_float(x):
    try: return float(str(x).replace(",", "."))
    except: return None

def _any(txt, words):
    t = _normalize(txt)
    return any(w in t for w in words)

def _has_up(txt):   return _any(txt, ["increment", "aument", "subir", "sube", "suba", "+"])
def _has_down(txt): return _any(txt, ["bajar", "disminu", "reducir", "baja", "-"])

def _numbers_trend_up(txt):
    t = _normalize(txt)

    m = re.search(r"actual[^0-9]*"+_num+r".*final[^0-9]*"+_num, t)
    if m:
        a,f = _to_float(m.group(1)), _to_float(m.group(2))
        return None if a is None or f is None else f>a

    m = re.search(r"llevar[^0-9]*a[^0-9]*"+_num+r".*actual[^0-9]*"+_num, t)
    if m:
        target,actual = _to_float(m.group(1)), _to_float(m.group(2))
        return None if target is None or actual is None else target>actual

    m = re.search(r"sale[^0-9]*"+_num+r".*(entra|ingresa)[^0-9]*"+_num, t)
    if m:
        s,e = _to_float(m.group(2)), _to_float(m.group(3))
        return None if s is None or e is None else e>s

    m = re.search(r"\bs[^0-9:]*[: ]\s*"+_num+r".*\be[^0-9:]*[: ]\s*"+_num, t)
    if m:
        s,e = _to_float(m.group(1)), _to_float(m.group(2))
        return None if s is None or e is None else e>s

    m = re.search(_num+r"\s*(rpm|gpm|hz|spm)\D{0,15}(?:a|@|→|-?>)\D{0,15}"+_num, t)
    if m:
        a,f = _to_float(m.group(1)), _to_float(m.group(3))
        return None if a is None or f is None else f>a

    return None

def _polea_direction(txt):
    trend = _numbers_trend_up(txt)
    if trend is None: return None
    return "SUBE REGIMEN" if trend else "BAJA REGIMEN"

def classify_accion(objetivo, observacion) -> str:
    o   = _normalize(objetivo)
    c   = _normalize(observacion)
    txt = f"{o} {c}"

    # --- A) CAMBIO CARRERA (prioritario sobre contrapesar)
    if "carrera" in txt and re.search(r"\b(minim|maxim|intermed|bajar|subir|aumentar|disminuir|cambio de carrera|54|64|74|84|96|118|130|149|168|192)\b", txt):
        return "CAMBIO CARRERA"

    # --- B) CAMBIO AIB: verbos + AIB/modelo; “proveniente/locación/trasladar/tomar AIB”
    if re.search(r"\b(instalar|montar|desmontar|transportar|intercambio|enroque|cambio|tomar)\b.*\b(aib|"+MODEL_WORDS+r")\b", txt)        or re.search(r"\b(proveniente de|locaci[oó]n|trasladar|mover|llevar)\b.*\b(aib|"+MODEL_WORDS+r")\b", txt):
        return "CAMBIO AIB"

    # --- C) CONTRAPESAR “puro”
    if re.search(r"\b(mover|colocar|ubicar|retirar|balancear|balanceo|contrapesar)\b.*\b(contrapesos?|placas?)\b", txt):
        return "CONTRAPESAR"

    # --- D) SUBE/BAJA por evidencia numérica o de polea
    trend = _numbers_trend_up(txt)
    if trend is True:  base = "SUBE REGIMEN"
    elif trend is False: base = "BAJA REGIMEN"
    else: base = ""

    if base == "" and ("polea" in txt or "cambio de polea" in txt):
        polea = _polea_direction(txt)
        if polea: base = polea
        else:
            base = "SUBE REGIMEN" if ("optimizacion" in o or "optimización" in objetivo.lower()) else ("BAJA REGIMEN" if "operativa" in o else "SUBE REGIMEN")

    # verbos de subir/bajar (sin números claros)
    if base == "":
        if _has_up(txt)  and re.search(r"\b(regimen|gpm|rpm|hz|spm|mci|vsd|pid|pip)\b", txt):  base = "SUBE REGIMEN"
        if _has_down(txt) and re.search(r"\b(regimen|gpm|rpm|hz|spm|mci)\b", txt):             base = "BAJA REGIMEN"

    # --- E) Ajustes sin evidencia (llevar/estabilizar/ajustar/adecuar/setear X)
    if base == "" and re.search(r"\b(llevar|estabilizar|ajustar|adecuar|sete(ar|o))\b", txt) and re.search(r"\b(gpm|rpm|hz|spm|mci|vsd)\b", txt):
        base = "BAJA REGIMEN" if re.search(r"\b(golpe de fluido|gdf|gdf)\b", txt) else "SUBE REGIMEN"

    # --- F) ACONDICIONAR (superficie)
    if base == "":
        if re.search(r"\b(hot ?oil|ho\b|bache|batch|quimic|químic|dispersante|acido|ácido|solvente|freno|reparar freno|leuter|dispositivo|cubre ?polea|cubrepoleas?|rotador|medicion|medición|nivel|mf\b|nf\b|muestra|analisis|análisis|hermeticidad)\b", txt):
            base = "VER EQUIP SUPERFICIE"

    # Guardrails
    if "optimizacion" in o or "optimización" in objetivo.lower():
        if base == "BAJA REGIMEN" or base == "": base = "SUBE REGIMEN"
    if "operativa" in o:
        strong_up = _has_up(txt) or trend is True or (_polea_direction(txt) == "SUBE REGIMEN")
        if base == "SUBE REGIMEN" and not strong_up:
            base = "BAJA REGIMEN"

    return base or "VER EQUIP SUPERFICIE"

# ---- Re-etiquetado (capa post) ----
def _is_contrapesar(txt):
    return _any(txt, ["contrapes", "placa", "balancear", "balanceo", "manivela"])

def _is_carrera_action(txt):
    t = _normalize(txt)
    return ("carrera" in t) and _any(t, [
        "cambio", "cambiar", "llevar", "pasar", "disminuir", "aumentar",
        "minima", "mínima", "maxima", "máxima", "intermedia", "subir", "bajar"
    ])

def _is_surface_conditioning(txt):
    return _any(txt, [
        "hot oil"," ho","bache","batch","quimic","químic","dispersante",
        "acido","ácido","solvente","freno","leuter","dispositivo",
        "cubre","cubrepolea","cubre-polea","rotador","medicion","medición",
        "nivel"," mf"," nf","muestra","analisis","análisis","hermeticidad"
    ])

def _is_polea_change(txt):
    t = _normalize(txt)
    if "polea" not in t: return False
    return bool(re.search(r"(sale|entra|ingresa|cambio de polea|cambiar polea|s:|e:|montar polea)", t))

def _has_model_movement(txt):
    t = _normalize(txt)
    movement = re.search(r"\b(mover|trasladar|transportar|llevar)\b", t)
    model = re.search(r"\b(aib|"+MODEL_WORDS+r")\b", t)
    return bool(movement and model)

def reclassify_label(etiqueta_inicial: str, objetivo: str, observacion: str) -> str:
    """Aplica reglas de re-etiquetado sobre la etiqueta ya calculada (no cambia la lógica base)."""
    etq = (etiqueta_inicial or "").strip().upper()
    obj = _normalize(objetivo)
    obs = _normalize(observacion)
    txt = f"{obj} {obs}"

    is_opt = ("optimizacion" in obj or "optimización" in objetivo.lower())
    is_opr = ("operativa" in obj)

    # 1) Si es ACONDICIONAR EQUIP SUPERFICIE
    if etq == "VER EQUIP SUPERFICIE":
        # 1.a) contrapesar siempre gana
        if _is_contrapesar(txt):
            return "CONTRAPESAR"
        # 1.b) cambio de polea: etiqueta por objetivo
        if _is_polea_change(txt):
            return "SUBE REGIMEN" if is_opt else "BAJA REGIMEN"
        # 1.c) movimiento + modelo → CAMBIO AIB
        if _has_model_movement(txt):
            return "CAMBIO AIB"
        # 1.d) GPM/RPM/Hz/SPM/MCI/VSD con verbos de ajuste (incluye setear)
        if _any(txt, ["gpm","rpm","hz","spm","mci","vsd"]) and _any(txt, ["llevar","estabilizar","dejar","ajustar","adecuar","setear","seteo"]):
            return "SUBE REGIMEN" if is_opt else "BAJA REGIMEN"
        return etq

    # 2) Si es BAJA REGIMEN
    if etq == "BAJA REGIMEN":
        if _is_contrapesar(txt):
            return "CONTRAPESAR"
        if _is_carrera_action(txt):
            return "CAMBIO CARRERA"
        if _is_surface_conditioning(txt) and not _is_polea_change(txt):
            return "VER EQUIP SUPERFICIE"
        # cambio de polea domina: BAJA si Operativa, SUBE si Opti
        if _is_polea_change(txt):
            return "SUBE REGIMEN" if is_opt else "BAJA REGIMEN"
        return etq

    # 3) Si es CAMBIO AIB
    if etq == "CAMBIO AIB":
        # 3.a) “montar polea” (no AIB) y objetivo Optimización → SUBE
        if "montar polea" in txt and is_opt:
            return "SUBE REGIMEN"
        # 3.b) si NO hay verbos de AIB pero sí cambio de polea → por objetivo
        verbos_aib = re.search(r"\b(instalar|montar|desmontar|transportar|intercambio|enroque|tomar)\b", txt)
        if (not verbos_aib) and _is_polea_change(txt):
            dir_ = _polea_direction(txt)
            if dir_:
                return dir_
            return "SUBE REGIMEN" if is_opt else "BAJA REGIMEN"
        return etq

    # 4) Si es CAMBIO CARRERA
    if etq == "CAMBIO CARRERA":
        if re.search(r"\b(montar|instalar|transportar|trasladar|intercambio|proveniente|locaci[oó]n|enroque|tomar)\b", txt) and            _any(txt, ["aib","lufkin","vulcan","siam","pump jack","darco","weatherford","maxii","wuel","mel altium","altium"]):
            return "CAMBIO AIB"
        return etq

    # 5) Si es CONTRAPESAR
    if etq == "CONTRAPESAR":
        if _is_carrera_action(txt):
            return "CAMBIO CARRERA"
        return etq

    # 6) Si es SUBE REGIMEN
    if etq == "SUBE REGIMEN":
        if _is_carrera_action(txt):
            return "CAMBIO CARRERA"
        return etq

    return etq

# ================== Cargas base ==================
def load_fdp(conn) -> pd.DataFrame:
    df = read_sql_df(conn, SQL_FDP_DINA)
    df["NIVEL_4"] = df["NIVEL_4"].astype(str).str.upper()
    return df

def load_produccion(conn) -> pd.DataFrame:
    return read_sql_df(conn, SQL_PRODUCCION)

def load_acciones(conn) -> pd.DataFrame:
    df = read_sql_df(conn, SQL_ACCIONES)
    for c in ["FECHAACCION","FECHAREALIZACION","FECHA AUTORIZACION"]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")
    return df

# ================== Transformaciones ==================
def join_expo(df_fdp: pd.DataFrame, expo: pd.DataFrame) -> pd.DataFrame:
    df = df_fdp.copy()
    df["__API_KEY__"] = df["AIB_MARCA_Y_DESC_API"].astype(str).str.upper().str.strip()
    ex = expo.copy()
    ex["__API_KEY__"] = ex["Denominacion API"]

    df = df.merge(
        ex[["__API_KEY__","Denominacion API","Marca","Tipo","Torque max","SENSIBILIDAD CR","SENSIBILIDAD EST"]],
        how="left", on="__API_KEY__"
    ).drop(columns=["__API_KEY__"])

    df = df.rename(columns={
        "Denominacion API": "EXPOSICIÓN.Denominacion API",
        "Marca": "EXPOSICIÓN.Marca",
        "Tipo": "EXPOSICIÓN.Tipo",
        "Torque max": "EXPOSICIÓN.Torque max",
        "SENSIBILIDAD CR": "EXPOSICIÓN.SENSIBILIDAD CR",
        "SENSIBILIDAD EST": "EXPOSICIÓN.SENSIBILIDAD EST",
    })

    df["ID_tmp"] = (df["NOMBRE_CORTO"].fillna("").astype(str) + df["NIVEL_5"].fillna("").astype(str))
    df = df[(df["ID_tmp"]!="") & df["NOMBRE_CORTO"].notna()]
    df = df.drop_duplicates(subset=["ID_tmp"]).drop(columns=["ID_tmp"])

    for c in ["AIBRR_TORQUE_MAXIMO_REDUCTOR","AIBEB_TORQUE_MAXIMO_REDUCTOR"]:
        df[c] = df[c].fillna(0)

    df["NIVEL_2"] = df["NIVEL_2"].replace({"Tierra del Fuego": "Chubut"})
    return df

def attach_produccion(df: pd.DataFrame, prod: pd.DataFrame) -> pd.DataFrame:
    return df.merge(prod, how="left", on="NIVEL_4")

def add_features(df: pd.DataFrame) -> pd.DataFrame:
    d = df.copy()
    d["PROD_OIL_1"] = d["PROD_OIL_1"].replace({0: np.nan})
    d["Ratio_Produccion"] = d["PROD_OIL"] / d["PROD_OIL_1"]

    def seg_consecuencia(x):
        if pd.isna(x): return "0 - 0,0025"
        if x >= 0.0065: return "Mayor a 0,0065"
        if x >= 0.0025: return "0,0025 - 0,0065"
        return "0 - 0,0025"
    d["f_CONSECUENCIA"] = d["Ratio_Produccion"].apply(seg_consecuencia)
    d["CONSECUENCIA"] = d["f_CONSECUENCIA"]

    d["TORQUE_BAL[%]"] = d["AIBEB_TORQUE_MAXIMO_REDUCTOR"] / d["AIBRR_TORQUE_DISPONIBLE"]
    d["TORQUE[%]"]     = d["AIBRR_TORQUE_MAXIMO_REDUCTOR"] / d["AIBRR_TORQUE_DISPONIBLE"]

    d["f_CR"]  = np.select([d["TORQUE[%]"] > 1, d["TORQUE[%]"] >= 0.85], [10, 3.7], default=0)
    d["f_CR3"] = d["EXPOSICIÓN.SENSIBILIDAD CR"].map({"ALTA":3, "MEDIA":1.5}).fillna(1.0)
    d["CAJA REDUCTORA"]  = d["f_CR"] * d["f_CR3"]
    d["CAJA REDUCTORA_"] = np.where(d["TORQUE_BAL[%]"]>1, 1.5*d["CAJA REDUCTORA"], d["CAJA REDUCTORA"])

    d["f_estr"]  = np.select([d["ESTRUCTURA_PCT"] >= 0.95, d["ESTRUCTURA_PCT"] > 0.85], [20,5], default=1)
    d["f_estr3"] = d["EXPOSICIÓN.SENSIBILIDAD EST"].map({"ALTA":3,"MEDIA":1.5}).fillna(1.0)
    d["ESTRUCTURA"] = d["f_estr"] * d["f_estr3"]

    d["f_GPM"]  = np.select([d["AIB_GPM"] > 9, d["AIB_GPM"] > 7], [3, 1.5], default=1)
    d["f_GPM2"] = np.where(d["LLENAD_BOMBA_PCT"] < 0.75, d["f_GPM"]*2, d["f_GPM"])

    d["EXIGENCIA"]   = d["f_GPM2"] * d["ESTRUCTURA"] + d["CAJA REDUCTORA_"]
    d["F_EXIGENCIA"] = np.select([d["EXIGENCIA"]>10, d["EXIGENCIA"]>4], ["A","M"], default="B")

    def crit(row):
        rp = row["Ratio_Produccion"]; fe = row["F_EXIGENCIA"]
        if pd.isna(rp) or pd.isna(fe): return "NORMAL"
        if rp >= 0.0065 and fe == "A": return "CRITICO"
        if rp >= 0.0025 and fe == "A": return "ALERTA"
        if rp <  0.0025 and fe == "A": return "ALERTA"
        if rp >= 0.0065 and fe == "M": return "ALERTA"
        if rp >= 0.0025 and fe == "M": return "ALERTA"
        if rp <  0.0025 and fe == "M": return "ALERTA"
        return "NORMAL"
    d["CRITICIDAD"] = d.apply(crit, axis=1).astype(str)
    d["ref_criticidad"] = d["CRITICIDAD"].map({"CRITICO":1, "ALERTA":2, "NORMAL":3}).astype("Int64")

    def detalle_crit(rp, fe):
        if pd.isna(rp) or pd.isna(fe): return "NORMAL"
        if rp >= 0.0065 and fe == "A": return "AA"
        if rp >= 0.0025 and fe == "A": return "AM"
        if rp <  0.0025 and fe == "A": return "AB"
        if rp >= 0.0065 and fe == "M": return "MA"
        if rp >= 0.0025 and fe == "M": return "MM"
        if rp <  0.0025 and fe == "M": return "MB"
        if rp >= 0.0065 and fe == "B": return "BA"
        if rp >= 0.0025 and fe == "B": return "BM"
        if rp <  0.0025 and fe == "B": return "BB"
        return "NORMAL"
    d["DETALLE CRITICIDAD"] = np.vectorize(detalle_crit)(d["Ratio_Produccion"], d["F_EXIGENCIA"])

    d["FECHA_HORA"] = pd.to_datetime(d["FECHA_HORA"], errors="coerce")
    max_por_pozo = d.groupby("NOMBRE_CORTO", dropna=False)["FECHA_HORA"].max().rename("FECHA_MAX").reset_index()
    d = d.merge(max_por_pozo, on="NOMBRE_CORTO", how="left")
    d["Días sin dina"] = (pd.Timestamp.today().normalize() - d["FECHA_MAX"].dt.normalize()).dt.days

    d["Correlativo Exigencia"] = d["F_EXIGENCIA"].map({"A":"10 +", "M":"04-10", "B":"0-04"}).fillna("")
    d["ID"] = d["f_CONSECUENCIA"].fillna("").astype(str) + d["Correlativo Exigencia"].fillna("").astype(str)

    d = d[(d["NOMBRE_CORTO"].notna()) & (d["NOMBRE_CORTO"].astype(str)!="")]

    return d

# ================== PIPELINE ==================
def main():
    print("Conectando a Oracle (cx_Oracle)…")
    with get_connection() as conn:
        print("Leyendo FDP_DINA…")
        fdp = load_fdp(conn)

        print("Leyendo Producción…")
        prod = load_produccion(conn)

        print("Leyendo ACCIONES…")
        acciones = load_acciones(conn)

    print("Leyendo Excel EXPOSICIÓN…")
    expo = load_exposicion(EXPO_PATH, EXPO_SHEET)

    print("Aplicando joins/limpieza (EXPOSICIÓN)…")
    fdp = join_expo(fdp, expo)

    print("Adjuntando Producción…")
    fdp = attach_produccion(fdp, prod)

    print("Calculando métricas / criticidad…")
    final = add_features(fdp)

    # -------- (1) Subset de columnas solicitadas --------
    columnas_matriz = [
        "NOMBRE_CORTO","NOMBRE_CORTO_POZO",
        "FECHA_HORA",
        "AIB_MARCA_Y_DESC_API",
        "PROD_OIL",
        "PROD_GAS",
        "PROD_WAT",
        "AIB_GPM",
        "MOTOR_DIAMETRO_POLEA",
        "EXPOSICIÓN.Denominacion API",
        "CRITICIDAD",
        "Días sin dina",
    ]
    faltantes = [c for c in columnas_matriz if c not in final.columns]
    if faltantes:
        raise KeyError(f"Faltan columnas en la matriz final: {faltantes}")
    matriz_subset = final[columnas_matriz].copy()

    # -------- (2) Etiquetado de ACCIONES --------
    for col in ["OBJETIVO", "OBSERVACION"]:
        if col not in acciones.columns:
            acciones[col] = ""

    acciones["ETIQUETA_BASE"] = acciones.apply(
        lambda r: classify_accion(r.get("OBJETIVO", ""), r.get("OBSERVACION", "")),
        axis=1
    )
    acciones["ETIQUETA"] = acciones.apply(
        lambda r: reclassify_label(r.get("ETIQUETA_BASE", ""), r.get("OBJETIVO", ""), r.get("OBSERVACION", "")),
        axis=1
    )

    
        # ================== ACCIONES EN PROCESO / FINALIZADAS ==================
    # ================== ACCIONES EN PROCESO / FINALIZADAS ==================
    # Utilidades comunes
    def _plural_dia(n):
        try:
            n = int(n)
        except Exception:
            return f"{n} días"
        return f"{n} día" if n == 1 else f"{n} días"

    def _fila_texto(obs, etq, dias, msg="sin hacerse la acción"):
        obs = ("" if pd.isna(obs) else str(obs)).strip()
        etq = ("" if pd.isna(etq) else str(etq)).strip()
        dias_txt = _plural_dia(dias if dias is not pd.NA else 0)
        # sin punto final
        return f"{obs} -{etq} - Lleva {dias_txt} {msg}"

    def _acciones_texto_por_pozo(acciones_df: pd.DataFrame,
                                 matriz_df: pd.DataFrame,
                                 estados_validos: set,
                                 col_salida: str,
                                 msg_final: str = "sin hacerse la acción",
                                 add_finalizada: bool = False,
                                 date_tail: tuple[str, str, str] | None = None,
                                 use_realizacion_only: bool = False):  # <-- NUEVO
        """
        Devuelve DF [NOMBRE_CORTO, col_salida] enumerado por pozo.
        ...
        """

        acc = acciones_df.copy()

        # Traer FECHA_HORA y CRITICIDAD desde MATRIZ
        mat_cols_needed = ["NOMBRE_CORTO", "FECHA_HORA", "CRITICIDAD"]
        mat_min = matriz_df[mat_cols_needed].drop_duplicates("NOMBRE_CORTO")
        acc = acc.merge(mat_min, on="NOMBRE_CORTO", how="left", suffixes=("", "_MATRIZ"))

        # Filtros base
        crit_mask  = acc["CRITICIDAD"].isin(["CRITICO", "ALERTA"])

        # >>>>>> ÚNICO CAMBIO DE LÓGICA EN FECHAS <<<<<<
        if use_realizacion_only:
            # Solo para ACCIONES FINALIZADAS: FECHAREALIZACION > FECHA_HORA
            fecha_mask = (
                pd.to_datetime(acc["FECHA_HORA"], errors="coerce")
                < pd.to_datetime(acc["FECHAREALIZACION"], errors="coerce")
            )
        else:
            # Resto (EN PROCESO, etc.): se mantiene como estaba
            fecha_mask = (
                pd.to_datetime(acc["FECHA_HORA"], errors="coerce")
                < pd.to_datetime(acc["FECHAACCION"], errors="coerce")
            )

        fecha_mask = fecha_mask.fillna(False)

        act_mask   = ~acc["ACTIVIDAD"].fillna("").str.contains("Intervención de Fondo", case=False, na=False)
        obj_mask   = (
            acc["OBJETIVO"].fillna("").str.contains("Operativa", case=False, na=False) |
            acc["OBJETIVO"].fillna("").str.contains("Optimización de producción", case=False, na=False)
        )

        acc_fil = acc[crit_mask & fecha_mask & act_mask & obj_mask].copy()
        

        # Estado
        acc_fil = acc_fil[acc_fil["ESTADO"].fillna("").str.upper().isin(estados_validos)].copy()
        if acc_fil.empty:
            return pd.DataFrame(columns=["NOMBRE_CORTO", col_salida])

        # Texto + días desde FECHAACCION
        today = pd.Timestamp.today().normalize()
        acc_fil["FECHAACCION"] = pd.to_datetime(acc_fil["FECHAACCION"], errors="coerce")
        acc_fil["dias_sin"] = (today - acc_fil["FECHAACCION"].dt.normalize()).dt.days.clip(lower=0).astype("Int64")

        # Para “y lleva XXXX días finalizada.” usamos FECHAREALIZACION
        if add_finalizada:
            acc_fil["FECHAREALIZACION"] = pd.to_datetime(acc_fil["FECHAREALIZACION"], errors="coerce")
            acc_fil["dias_fin"] = (today - acc_fil["FECHAREALIZACION"].dt.normalize()).dt.days.clip(lower=0).astype("Int64")
            acc_fil["__extra__"] = acc_fil["dias_fin"].apply(
                lambda d: f" y lleva {_plural_dia(0 if pd.isna(d) else d)} finalizada."
            )
        else:
            acc_fil["__extra__"] = ""

        # Sufijo de fecha (opcional) — p.ej. " FECHA ACCION: 13/08/25"
        if date_tail:
            col_fecha, etiqueta, fmt = date_tail
            col_fecha_norm = pd.to_datetime(acc_fil[col_fecha], errors="coerce")
            acc_fil["__fecha_tail__"] = col_fecha_norm.dt.strftime(fmt).fillna("")
            acc_fil["__fecha_tail__"] = acc_fil["__fecha_tail__"].apply(
                lambda s: f" {etiqueta}: {s}" if s else ""
            )
        else:
            acc_fil["__fecha_tail__"] = ""

        def _fila_texto(obs, etq, dias, extra="", fecha_tail=""):
            obs = ("" if pd.isna(obs) else str(obs)).strip()
            etq = ("" if pd.isna(etq) else str(etq)).strip()
            dias_txt = _plural_dia(dias if dias is not pd.NA else 0)
            return f"{obs} -{etq} - Lleva {dias_txt} {msg_final}.{extra}{fecha_tail}"

        acc_fil["__texto__"] = acc_fil.apply(
            lambda r: _fila_texto(r["OBSERVACION"], r["ETIQUETA"], r["dias_sin"],
                                  r["__extra__"], r["__fecha_tail__"]),
            axis=1
        )

        # Ordenar y enumerar
        acc_fil = acc_fil.sort_values(["NOMBRE_CORTO", "FECHAACCION"], ascending=[True, True])

        def _enumerar_series(s):
            lst = list(s)
            return "\n".join(f"{i}) {t}" for i, t in enumerate(lst, start=1))

        out = (
            acc_fil.groupby("NOMBRE_CORTO")["__texto__"]
            .apply(_enumerar_series)
            .reset_index(name=col_salida)
        )
        return out

    
    # ---- Construir y mergear columnas en MATRIZ ----
    
    # 1) ACCIONES EN PROCESO  -> agrega " FECHA ACCION: dd/mm/yy"
    estados_proceso = {"EN PROCESO", "BORRADOR", "NO INICIADA", "PROPUESTO"}
    agg_proceso = _acciones_texto_por_pozo(
        acciones, matriz_subset, estados_proceso, "ACCIONES EN PROCESO",
        msg_final="sin hacerse la acción",
        add_finalizada=False,
        date_tail=("FECHAACCION", "FECHA ACCION", "%d/%m/%y")
    )
    if agg_proceso.empty:
        agg_proceso = pd.DataFrame(columns=["NOMBRE_CORTO", "ACCIONES EN PROCESO"])
    matriz_subset = matriz_subset.merge(agg_proceso, on="NOMBRE_CORTO", how="left")
    matriz_subset["ACCIONES EN PROCESO"] = matriz_subset["ACCIONES EN PROCESO"].fillna("")

    # 2) ACCIONES FINALIZADAS -> agrega “y lleva XXXX días finalizada.” y “ FECHA REALIZACION dd/mm/YYYY”
    estados_finalizadas = {"FINALIZADA"}
    agg_final = _acciones_texto_por_pozo(
        acciones, matriz_subset, estados_finalizadas, "ACCIONES FINALIZADAS",
        msg_final="iniciada la acción",   # o “finalizada la acción” si preferís
        add_finalizada=True,
        date_tail=("FECHAREALIZACION", "FECHA REALIZACION", "%d/%m/%Y"),
        use_realizacion_only=True  # <-- NUEVO: usa solo FECHAREALIZACION > FECHA_HORA
    )

    if agg_final.empty:
        agg_final = pd.DataFrame(columns=["NOMBRE_CORTO", "ACCIONES FINALIZADAS"])
    matriz_subset = matriz_subset.merge(agg_final, on="NOMBRE_CORTO", how="left")
    matriz_subset["ACCIONES FINALIZADAS"] = matriz_subset["ACCIONES FINALIZADAS"].fillna("")

    # ================== Columna VER ==================
    # Normalizamos textos (mayúsculas, NaN -> "")
    proc = matriz_subset["ACCIONES EN PROCESO"].fillna("").str.upper()

    # Soporta nombre singular o plural para finalizadas
    if "ACCIONES FINALIZADAS" in matriz_subset.columns:
        fin = matriz_subset["ACCIONES FINALIZADAS"].fillna("").str.upper()
    elif "ACCIONES FINALIZADA" in matriz_subset.columns:
        fin = matriz_subset["ACCIONES FINALIZADA"].fillna("").str.upper()
    else:
        fin = pd.Series("", index=matriz_subset.index)

    # 1) Proceso contiene: MATRIZ | CAMBIO AIB | CAMBIO CARRERA | CONTRAPESAR
    m1 = proc.str.contains(r"\bMATRIZ\b|CAMBIO AIB|CAMBIO CARRERA|CONTRAPESAR", regex=True)

    # 2) Proceso contiene: ACONDICIONAR EQUIP SUPERFICIE
    m2 = proc.str.contains("VER EQUIP SUPERFICIE", regex=False)

    # 3) Finalizadas contiene: MATRIZ | CAMBIO AIB | CAMBIO CARRERA | CONTRAPESAR | ACONDICIONAR EQUIP SUPERFICIE
    m3 = fin.str.contains(r"\bMATRIZ\b|CAMBIO AIB|CAMBIO CARRERA|CONTRAPESAR|ACONDICIONAR EQUIP SUPERFICIE", regex=True)

    matriz_subset["VER"] = np.select(
        [m1, m2, m3],
        [
            "ACCION EN PROCESO PARA REGULARIZAR MATRIZ AIB",
            "ACONDICIONAR SUPERFICIE PARA TOMAR MEDICION",
            "ACTUALIZAR MEDICION DINAMOMETRICA",
        ],
        default=""
    )

 
    
    # -------- Exportar a Excel con 2 hojas --------
    out = Path("Matriz_AIB_export.xlsx")
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        matriz_subset.to_excel(writer, sheet_name="MATRIZ", index=False)
        acciones.to_excel(writer, sheet_name="ACCIONES", index=False)

    print(f"✅ Exportado {len(matriz_subset):,} filas (MATRIZ) + {len(acciones):,} filas (ACCIONES) → {out.absolute()}")


if __name__ == "__main__":
    main()


# In[2]:


#ESTA CONSULTA ES LA SEGUNDA EN EJECUTAR.
# --- Outlook: toma el último "Mediciones totales - LP", combina por Fecha, agrega POZO NORMALIZADO y CARGAR MEDICION ---
import pandas as pd
from pathlib import Path
import unicodedata, sys, re
from datetime import datetime
from difflib import SequenceMatcher

SAVE_DIR   = Path(r"C:\MedicionesInbox\compilados")
TMP_DIR    = SAVE_DIR / "_tmp"
COORD_PATH = Path(r"C:\Users\ry16123\OneDrive - YPF\Escritorio\power BI\GUADAL- POWER BI\Inteligencia Artificial\coordenadas1.xlsx")
MATRIZ_PATH = Path(r"C:\Users\ry16123\Matriz_AIB_export.xlsx")  # <-- AJUSTAR SI HACE FALTA
TMP_DIR.mkdir(parents=True, exist_ok=True)

# ---------------------- helpers ----------------------
def _normalize_for_match(s: str) -> str:
    s = s.replace("\xa0", " ")
    s = " ".join(s.split())
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    return s.lower()

def _match_filename(name: str) -> bool:
    n = _normalize_for_match(name)
    if not re.search(r"\.(xls|xlsx|xlsm|xltx|xltm)$", n):
        return False
    return re.search(r"mediciones\s+totales\s*-\s*lp", n) is not None

def get_latest_lp_excel_from_outlook(save_to_dir: Path) -> Path:
    try:
        import win32com.client as win32
    except ImportError:
        sys.exit("Falta pywin32. Instalá: pip install pywin32")

    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # Inbox
    items = inbox.Items
    items.Sort("[ReceivedTime]", True)

    for item in items:
        try:
            atts = item.Attachments
        except Exception:
            continue
        for i in range(1, atts.Count + 1):
            att = atts.Item(i)
            fn = att.FileName or ""
            if _match_filename(fn):
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                out_path = save_to_dir / f"{Path(fn).stem}__{ts}{Path(fn).suffix}"
                att.SaveAsFile(str(out_path))
                return out_path
    sys.exit("No encontré adjuntos 'Mediciones totales - LP' en Outlook.")

def norm_col(s):
    s = str(s).replace("\xa0", " ").strip()
    s = " ".join(s.split())
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    return s.lower()

def fix_fecha(df):
    if df is None or df.empty:
        return None
    df = df.rename(columns=lambda c: str(c).replace("\xa0", " ").strip())
    fecha_col = None
    for c in df.columns:
        if "fecha" in norm_col(c):
            fecha_col = c
            break
    if fecha_col is None:
        print("Omito hoja sin columna 'Fecha'. Columnas:", list(df.columns))
        return None
    if fecha_col != "Fecha":
        df = df.rename(columns={fecha_col: "Fecha"})
    df["Fecha"] = pd.to_datetime(df["Fecha"], dayfirst=True, errors="coerce")
    df = df.dropna(subset=["Fecha"]).dropna(how="all")
    return df

# ------ utilidades normalización de POZO ------
import re as _re
def extract_number(s):
    m = _re.search(r'(\d+)', str(s))
    return m.group(1) if m else ""

def extract_letters(s):
    letters = _re.findall(r'[A-Z]+', str(s).upper())
    return "".join(letters)

from difflib import SequenceMatcher
def custom_normalize_pozo(user_pozo, coord_list, letter_threshold=0.5):
    if not isinstance(user_pozo, str):
        return user_pozo
    user_pozo = user_pozo.strip()
    user_number  = extract_number(user_pozo)
    user_letters = extract_letters(user_pozo)

    candidates = []
    for cand in coord_list:
        cand = str(cand).strip()
        if extract_number(cand) == user_number and user_number != "":
            candidates.append(cand)

    if not candidates:
        return user_pozo
    if len(candidates) == 1:
        return candidates[0]

    best_candidate, best_ratio = None, 0
    for cand in candidates:
        ratio = SequenceMatcher(None, user_letters, extract_letters(cand)).ratio()
        if ratio > best_ratio:
            best_ratio, best_candidate = ratio, cand
    return best_candidate if best_ratio >= letter_threshold else user_pozo

def find_pozo_column(df):
    for c in df.columns:
        if "pozo" in norm_col(c):
            return c
    return None

# ---------------------- main ----------------------
def main():
    # 1) Traer adjunto desde Outlook
    src = get_latest_lp_excel_from_outlook(TMP_DIR)
    print("Usando adjunto de Outlook:", src.name)

    # 2) Combinar hojas por Fecha (misma lógica)
    engine = "openpyxl" if src.suffix.lower() in (".xlsx", ".xlsm", ".xltx", ".xltm") else None
    sheets = pd.read_excel(src, sheet_name=None, engine=engine)

    comb = []
    for name, df in sheets.items():
        fixed = fix_fecha(df)
        if fixed is not None and not fixed.empty:
            fixed["__hoja__"] = name
            comb.append(fixed)

    if not comb:
        sys.exit("No hubo datos combinables con 'Fecha' válida.")

    final = pd.concat(comb, ignore_index=True).sort_values("Fecha", ascending=True)

    # 3) Agregar POZO NORMALIZADO usando 'coordenadas1.xlsx'
    pozo_col = find_pozo_column(final)
    if pozo_col is None:
        print("⚠️ No encontré columna 'Pozo'. No se agregará POZO NORMALIZADO.")
        final["POZO NORMALIZADO"] = None
    else:
        if not COORD_PATH.exists():
            print(f"⚠️ No encuentro {COORD_PATH}. Dejo POZO NORMALIZADO = Pozo original.")
            final["POZO NORMALIZADO"] = final[pozo_col]
        else:
            df_coords = pd.read_excel(COORD_PATH, engine="openpyxl")
            coord_pozo_col = find_pozo_column(df_coords)
            if coord_pozo_col is None:
                print("⚠️ En coordenadas1.xlsx no hay columna 'POZO'. Dejo POZO NORMALIZADO = Pozo original.")
                final["POZO NORMALIZADO"] = final[pozo_col]
            else:
                final["POZO_TMP"]  = final[pozo_col].astype(str).str.strip().str.upper()
                df_coords["POZO_TMP"] = df_coords[coord_pozo_col].astype(str).str.strip().str.upper()

                df_merged = final.merge(
                    df_coords[[coord_pozo_col, "POZO_TMP"]],
                    on="POZO_TMP",
                    how="left",
                    suffixes=("", "_coords"),
                )

                coord_pozo_list = df_coords[coord_pozo_col].dropna().astype(str).tolist()
                df_merged["POZO NORMALIZADO"] = df_merged.apply(
                    lambda row: row.get(coord_pozo_col) if pd.notnull(row.get(coord_pozo_col))
                    else custom_normalize_pozo(row[pozo_col], coord_pozo_list, letter_threshold=0.5),
                    axis=1
                )
                df_merged.drop(columns=["POZO_TMP", coord_pozo_col], inplace=True)
                final = df_merged

    # 4) Agregar columna "CARGAR MEDICION" en base a Matriz_AIB_export
    #    Merge por POZO NORMALIZADO ↔ NOMBRE_CORTO_POZO y comparar FECHA_HORA < Fecha y Cancelada != 1
    if "POZO NORMALIZADO" not in final.columns:
        print("⚠️ No hay 'POZO NORMALIZADO'; no se puede calcular 'CARGAR MEDICION'.")
        final["CARGAR MEDICION"] = None
    else:
        if not MATRIZ_PATH.exists():
            print(f"⚠️ No encuentro {MATRIZ_PATH}. 'CARGAR MEDICION' quedará vacío.")
            final["CARGAR MEDICION"] = None
        else:
            df_matriz = pd.read_excel(MATRIZ_PATH, engine="openpyxl")
            # Detectar columna NOMBRE_CORTO_POZO (o similar)
            col_ncpozo = None
            for c in df_matriz.columns:
                if "nombre_corto_pozo" in norm_col(c):
                    col_ncpozo = c; break
            if col_ncpozo is None:
                # fallback por "nombre corto" / "nombre_corto"
                for c in df_matriz.columns:
                    if "nombre_corto" in norm_col(c):
                        col_ncpozo = c; break

            col_fh = None
            for c in df_matriz.columns:
                if "fecha_hora" in norm_col(c):
                    col_fh = c; break
                if norm_col(c) == "fecha hora":  # por si viene con espacio
                    col_fh = c; break

            if col_ncpozo is None or col_fh is None:
                print("⚠️ En Matriz_AIB_export no encuentro 'NOMBRE_CORTO_POZO' y/o 'FECHA_HORA'.")
                final["CARGAR MEDICION"] = None
            else:
                # Nos quedamos con la FECHA_HORA más reciente por pozo (por seguridad)
                df_m = df_matriz[[col_ncpozo, col_fh]].copy()
                df_m[col_ncpozo] = df_m[col_ncpozo].astype(str).str.strip().str.upper()
                df_m[col_fh] = pd.to_datetime(df_m[col_fh], errors="coerce", dayfirst=True)
                df_m = df_m.dropna(subset=[col_ncpozo, col_fh])
                df_m = df_m.sort_values(col_fh).groupby(col_ncpozo, as_index=False).last()

                # Preparar final
                final["_POZO_NORM_UP"] = final["POZO NORMALIZADO"].astype(str).str.strip().str.upper()
                final["_Fecha_dia"] = pd.to_datetime(final["Fecha"], errors="coerce", dayfirst=True).dt.floor("D")

                df_join = final.merge(
                    df_m.rename(columns={col_ncpozo: "_POZO_NORM_UP", col_fh: "_FECHA_HORA"}),
                    on="_POZO_NORM_UP",
                    how="left"
                )

                # Cancelada puede venir como 1, "1", 1.0; la normalizo a flag
                def _cancelada_flag(x):
                    try:
                        return float(str(x).strip()) == 1.0
                    except Exception:
                        return False

                df_join["_cancelada1"] = df_join.get("Cancelada").apply(_cancelada_flag)
                # Condición: FECHA_HORA < Fecha (día) y Cancelada != 1  -> "CARGAR MEDICION EN SISTEMA"
                cond = (pd.notnull(df_join["_FECHA_HORA"])) & (df_join["_FECHA_HORA"].dt.date < df_join["_Fecha_dia"].dt.date) & (~df_join["_cancelada1"])
                df_join["CARGAR MEDICION"] = df_join.apply(
                    lambda r: "CARGAR MEDICION EN SISTEMA" if cond.loc[r.name] else None,
                    axis=1
                )

                # Limpiar auxiliares
                final = df_join.drop(columns=["_POZO_NORM_UP", "_Fecha_dia", "_cancelada1"])

    # 5) Exportar
    out_path = SAVE_DIR / "Mediciones_totales_LP__COMBINADO_1.xlsx"
    with pd.ExcelWriter(out_path, engine="openpyxl") as xw:
        final.to_excel(xw, sheet_name="MEDICIONES", index=False)

    print(f"✅ Combinado + POZO NORMALIZADO + CARGAR MEDICION. Filas: {len(final)}  Columnas: {len(final.columns)}")
    print("Salida:", out_path)

if __name__ == "__main__":
    main()


# In[5]:


#TERCER CONSULTA A EJECUTAR
# -*- coding: utf-8 -*-
"""
Agrega 4 pestañas al Excel COMBINADO (sin incremental, sin append, sin dedup):
 A) "Potenciales y controles"  (PSFU)  -> filtra ESTADO = 'Produciendo'
 B) "Pérdidas"                 (PCCT)  -> siempre día anterior a hoy (por SQL)
 C) "Último evento"            (PSFU)  -> último evento por pozo
 D) "Eventos 2024+"            (PSFU)  -> eventos desde 2024-01-01 (INT/RES/TER/ISQ/REP/INW)

Comportamiento:
- No guarda ni lee estado previo.
- No apendea ni deduplica: sobreescribe cada hoja en cada corrida.
"""
from pathlib import Path
import pandas as pd
from sqlalchemy import create_engine, text
from typing import Optional

SAVE_DIR   = Path(r"C:\MedicionesInbox\compilados")
OUT_XLSX   = SAVE_DIR / "Mediciones_totales_LP__COMBINADO_2.xlsx"
SAVE_DIR.mkdir(parents=True, exist_ok=True)

# ========= Conexiones Oracle (SQLAlchemy + cx_Oracle) =========
def engine_psfu(user, password, host, port, service):
    dsn = f"(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST={host})(PORT={port}))"           f"(CONNECT_DATA=(SERVICE_NAME={service})))"
    url = f"oracle+cx_oracle://{user}:{password}@/?dsn={dsn}"
    return create_engine(url)

def engine_pcct(user, password, host, port, service):
    dsn = f"(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST={host})(PORT={port}))"           f"(CONNECT_DATA=(SERVICE_NAME={service})))"
    url = f"oracle+cx_oracle://{user}:{password}@/?dsn={dsn}"
    return create_engine(url)

# ========= Helper de escritura simple =========
def write_sheet(df: pd.DataFrame, sheet_name: str, date_col: Optional[str] = None):
    """Sobrescribe la pestaña indicada. Si date_col existe, ordena por esa columna."""
    if date_col and (date_col in df.columns):
        df = df.copy()
        df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
        df = df.sort_values(by=date_col, ascending=True, kind="mergesort")

    if OUT_XLSX.exists():
        # Archivo ya existe: abrimos en append y reemplazamos la hoja
        with pd.ExcelWriter(
            OUT_XLSX, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as xw:
            df.to_excel(xw, sheet_name=sheet_name, index=False)
    else:
        # Archivo no existe: creamos nuevo
        with pd.ExcelWriter(
            OUT_XLSX, engine="openpyxl", mode="w"
        ) as xw:
            df.to_excel(xw, sheet_name=sheet_name, index=False)

# ========= Consultas =========
def run_potenciales_controles():
    """
    A) PSFU — Potenciales y controles (solo ESTADO='Produciendo').
    Incluye columna LAST_DT (por si querés ordenarla o inspeccionarla).
    """
    sheet_name = "Potenciales y controles"
    user, pwd, host, port, svc = ("RY33872", "Contraseña_0725", "slplpgmoora03", 1527, "psfu")
    eng = engine_psfu(user, pwd, host, port, svc)

    sql = """
    SELECT 
        MONTHS_BETWEEN(SYSDATE, v.MAX_EFF_DT)                         AS MESES_DESDE_POTENCIAL,
        MONTHS_BETWEEN(SYSDATE, f.FECHA_DESDE) * 30                   AS DIAS_DESDE_DIAG,
        SYSDATE - v.MAX_EFF_DT                                        AS DIAS_HASTA_POT,
        o.NOMBRE_CORTO,
        o.NOMBRE_CORTO_POZO,
        o.NOMBRE_POZO,
        o.ESTADO,
        o.MET_PROD,
        o.NIVEL_1,
        o.NIVEL_3,
        o.NIVEL_5,
        t.FECHA_NVL,
        c.MAX_EFF_DT AS FECHA_ULT_CONTROL,
        c.TEST_PURP_CD,
        c.TEST_REASON_CD,
        v.MAX_EFF_DT AS FECHA_ULT_POT,
        t.TIPONIVEL,
        o.CLAS_ABC,
        f.FECHA_DESDE,
        f.DIAGNOSTICO,
        f.MOTIVO,
        (v.POT_OIL + v.POT_WAT) - c.TOTAL_VOL                       AS DIFF_POT_VS_CONT,
        v.POT_OIL - c.PROD_OIL                                      AS OIL_DIFF,
        o.COMP_SK,
        c.DAYS_FROM_TODAY,
        c.PROD_OIL,
        c.PROD_GAS,
        c.FREE_WAT_PCT,
        c.TOTAL_VOL,
        v.POT_OIL,
        v.POT_GAS,
        (v.POT_WAT * 100) / DECODE(v.POT_OIL+v.POT_WAT, 0, 1, NULL, 1, v.POT_OIL+v.POT_WAT) AS PORC_AGUA,
        v.POT_OIL + v.POT_WAT                                      AS TOTAL_POT,
        t.SUMERGENCIA,
        t.REGIMEN_PRIM,
        t.PROF_ENT_BOMBA,
        t.REGIMEN_SEC,
        /* fecha compuesta para referencia/orden */
        GREATEST(
            NVL(c.MAX_EFF_DT,    DATE '1900-01-01'),
            NVL(v.MAX_EFF_DT,    DATE '1900-01-01'),
            NVL(f.FECHA_DESDE,   DATE '1900-01-01'),
            NVL(t.FECHA_NVL,     DATE '1900-01-01')
        ) AS LAST_DT
    FROM DISC_ADMINS.DBU_FIC_ORG_ESTRUCTURAL o
    LEFT JOIN DISC_ADMINS.TEST_FIC_ULT_NIVELES t 
        ON t.COMP_SK = o.COMP_SK
    JOIN DISC_ADMINS.VTOW_WELL_LAST_CONTROL_DET c 
        ON c.COMP_SK = o.COMP_SK
    LEFT JOIN DISC_ADMINS.VTOW_WELL_LAST_POTENTIAL_DET v 
        ON v.COMP_SK = o.COMP_SK
    JOIN DISC_ADMINS.FIC_ULT_DIAGNOSTICO f 
        ON f.COMP_SK = o.COMP_SK
    WHERE 
        o.NIVEL_3 IN (
            'Las Heras CG - Canadon Escondida',
            'Los Perales',
            'El Guadal',
            'Seco Leon - Pico Truncado'
        )
        AND o.MET_PROD IN (
            'Auto Gas Lift','Bombeo Hidráulico','Bombeo Mecánico',
            'Cavidad Progresiva','Electro Sumergible','Pistoneo',
            'Plunger Lift','Recoil','Surgente'
        )
        AND o.ESTADO = 'Produciendo'
    ORDER BY o.NIVEL_5 ASC, o.NOMBRE_CORTO ASC, o.NOMBRE_POZO ASC
    """
    df = pd.read_sql_query(text(sql), eng)

    # Por si el driver no materializa el alias LAST_DT:
    if "LAST_DT" not in df.columns:
        cand = [c for c in ["FECHA_ULT_CONTROL", "FECHA_ULT_POT", "FECHA_DESDE", "FECHA_NVL"] if c in df.columns]
        for c in cand:
            df[c] = pd.to_datetime(df[c], errors="coerce")
        df["LAST_DT"] = df[cand].max(axis=1) if cand else pd.NaT

    write_sheet(df, sheet_name, date_col="LAST_DT")

def run_perdidas():
    """
    B) PCCT — Pérdidas:
       Siempre trae el día anterior completo mediante TRUNC(SYSDATE).
    """
    sheet_name = "Pérdidas"
    user, pwd, host, port, svc = ("RY16123", "Luciano285", "suarbultowp01", 1521, "pcct")
    eng = engine_pcct(user, pwd, host, port, svc)

    sql = """
    SELECT 
        c.COMP_S_NAME,
        p.NET_LOSE,
        j.ORG_ENT_DS3,
        j.ORG_ENT_DS4,
        j.ORG_ENT_DS5,
        p.PROD_DT,
        r.PROD_STATUS_CD,
        r.PROD_STATUS_DS,
        g.REF_DS,
        p.WAT_LOSE,
        SUM(p.WAT_LOSE) OVER () AS TOTAL_WAT_LOSE,
        SUM(p.NET_LOSE)  OVER () AS TOTAL_NET_LOSE,
        SUM(p.GAS_LOSE)  OVER () AS TOTAL_GAS_LOSE
    FROM DISC_ADMIN_TOW.TOW_COMPLETACIONES c
    JOIN DISC_ADMIN_TOW.TOW_PERDIDAS p 
        ON p.COMP_SK = c.COMP_SK
    JOIN DISC_ADMIN_TOW.TOW_JERARQUIA j 
        ON j.ASSGN_SK = c.COMP_SK
    JOIN DISC_ADMIN_TOW.TOW_RUBROS_DE_PARO r 
        ON r.PROD_STATUS_CD = p.CODIGO_RUBRO
    JOIN DISC_ADMIN_TOW.TOW_REF_GRAN_RUBRO g 
        ON g.REF_ID = r.USER_A1
    WHERE 
        j.ASSGN_SK_TYPE = 'CC'
        AND j.ORG_SK = 1
        AND j.ORG_ENT_DS3 IN (
            'Las Heras CG - Canadon Escondida',
            'Los Perales',
            'El Guadal',
            'Seco Leon - Pico Truncado'
        )
        /* Día anterior completo: [TRUNC(SYSDATE) - 1, TRUNC(SYSDATE)) */
        AND p.PROD_DT >= TRUNC(SYSDATE) - 1
        AND p.PROD_DT <  TRUNC(SYSDATE)
    """
    df = pd.read_sql_query(text(sql), eng)
    write_sheet(df, sheet_name, date_col="PROD_DT")

def run_ultimo_evento():
    """
    C) PSFU — Último evento por pozo (INT/RES/REP/TER) con estado Produciendo.
       Se queda con el más reciente por pozo por medio de ROW_NUMBER en SQL.
    """
    sheet_name = "Último evento"
    user, pwd, host, port, svc = ("RY33872", "Contraseña_0725", "slplpgmoora03", 1527, "psfu")
    eng = engine_psfu(user, pwd, host, port, svc)

    sql = """
    WITH E AS (
        SELECT 
            o.NOMBRE                   AS NOMBRE,
            o.NOMBRE_POZO              AS NOMBRE_POZO,
            o.ESTADO,
            o.NIVEL_2,
            o.NIVEL_3,
            ev.I_KEY,
            ev.E_KEY,
            ev.END_STATUS,
            ev.START_DATE,
            ev.END_DATE,
            ev.EQUIP_TYPE,
            ev.EVNTCODE,
            ev.JOB_TYPE,
            ev.OBJECTIVE,
            ev.OBJECTIVE2,
            ev.CONTRATISTA,
            ROW_NUMBER() OVER (PARTITION BY o.NOMBRE_POZO ORDER BY ev.START_DATE DESC NULLS LAST) AS RN
        FROM DISC_ADMINS.DBU_FIC_ORG_ESTRUCTURAL o
        JOIN DISC_ADMINS.DBU_FIC_EVENTOS ev
            ON ev.I_KEY = o.I_KEY
        WHERE 
            o.NIVEL_2 = 'Negocio Santa Cruz'
            AND o.ESTADO = 'Produciendo'
            AND ev.EVNTCODE IN ('INT','RES','REP','TER')
    )
    SELECT 
        NOMBRE, ESTADO, NIVEL_2, NIVEL_3, I_KEY, E_KEY, END_STATUS, 
        START_DATE, END_DATE, EQUIP_TYPE, EVNTCODE, JOB_TYPE, OBJECTIVE, OBJECTIVE2, CONTRATISTA, NOMBRE_POZO
    FROM E
    WHERE RN = 1
    """
    df = pd.read_sql_query(text(sql), eng)
    write_sheet(df, sheet_name, date_col="START_DATE")

def run_eventos_2024_psfu():
    """
    D) PSFU — Eventos desde 2024-01-01 para INT/RES/TER/ISQ/REP/INW,
       en áreas seleccionadas y para pozos Productores (Oil/Gas/GyC).
       Conexión: suarbuworap09:1527/psfu
    """
    sheet_name = "Eventos 2024+"
    user, pwd, host, port, svc = ("RY33872", "Contraseña_0725", "slplpgmoora03", 1527, "psfu")
    eng = engine_psfu(user, pwd, host, port, svc)

    sql = """
    SELECT DISTINCT
        o.NOMBRE_CORTO,
        o.TIPO,
        o.ESTADO,
        o.MET_PROD,
        o.BATERIA,
        o.NIVEL_3,
        e.END_STATUS,
        e.START_DATE,
        e.END_DATE,
        e.EVNTCODE,
        e.JOB_TYPE,
        e.OBJECTIVE,
        e.OBJECTIVE2,
        e.CONTRATISTA,
        o.NOMBRE_POZO,
        o.NOMBRE_CORTO_POZO
    FROM DISC_ADMINS.DBU_FIC_ORG_ESTRUCTURAL o
    JOIN DISC_ADMINS.DBU_FIC_EVENTOS e
        ON e.I_KEY = o.I_KEY
    WHERE
        e.EVNTCODE IN ('INT','RES','TER','ISQ','REP','INW')
        AND e.START_DATE >= TO_DATE('20240601000000','YYYYMMDDHH24MISS')
        AND o.TIPO IN ('Productor - Gas','Productor - Gas Y Condensado','Productor - Petróleo')
        AND o.NIVEL_3 IN (
            'Las Heras CG - Canadon Escondida',
            'Los Perales',
            'El Guadal',
            'Seco Leon - Pico Truncado'
        )
    ORDER BY o.NIVEL_3, o.NOMBRE_CORTO, e.START_DATE
    """
    df = pd.read_sql_query(text(sql), eng)
    write_sheet(df, sheet_name, date_col="START_DATE")

# ========= Entrypoint =========
if __name__ == "__main__":
    # A)
    run_potenciales_controles()
    # B)
    run_perdidas()
    # C)
    run_ultimo_evento()
    # D) NUEVA
    run_eventos_2024_psfu()

    print(f"✅ Listo. Actualizado (sobrescribiendo hojas): {OUT_XLSX}")


# In[6]:


# -*- coding: utf-8 -*-CUARTA CONSULTA A EJECUTAR
#CONSOLIDA LOS DOS EXCEL - USAR ESTE - ( antes hay que ejecutar dos más) ESTA DEVUELVE SIN LAS CATEGORIAS. 
"""
Consolida dos Excels en uno con cuatro pestañas y completa "CARGAR MEDICION" en MEDICIONES
usando la hoja "Potenciales y controles" (POZO NORMALIZADO ↔ nombre_corto_pozo, fecha_nvl < Fecha).
Además:
- Agrega en "Potenciales y controles" la columna "TIPO DE NIVEL QUE FALTA CARGAR".
- Trae del correo el Excel más reciente "Relevamiento de campo - Novedades LP-LM"
  y mergea por Pozo+Fecha (día) para inyectar en MEDICIONES:
  Observaciones (Relev), SOLICITUD (Relev), OPERACIÓN REALIZADA (Relev), ACONDICIONAR (Relev).
- En "Potenciales y controles", agrega (por merge con MEDICIONES)
  Observaciones_rel, SOLICITUD (Relev), OPERACIÓN REALIZADA (Relev), ACONDICIONAR (Relev),
  y "Fecha (más reciente)" por pozo.
- NUEVA REGLA: En "Potenciales y controles", si met_prod ≠ "Bombeo Mecánico",
  Observaciones_rel está vacío, "Fecha (más reciente)" > fecha_nvl y
  "TIPO DE NIVEL QUE FALTA CARGAR" está vacío → poner "CARGAR MEDICION DE NIVEL".
"""
from pathlib import Path
import pandas as pd
import unicodedata
from datetime import datetime
import re
from typing import Optional
from __future__ import annotations

# ====== CONFIGURACIÓN ======
BASE_DIR = Path(r"C:\MedicionesInbox\compilados")
SRC1 = BASE_DIR / "Mediciones_totales_LP__COMBINADO_1.xlsx"   # tiene 'MEDICIONES'
SRC2 = BASE_DIR / "Mediciones_totales_LP__COMBINADO_2.xlsx"   # tiene 3 pestañas
OUT  = BASE_DIR / "Mediciones_totales_LP__CONSOLIDADO_FINAL.xlsx"
TMP_DIR = BASE_DIR / "_tmp"
TMP_DIR.mkdir(parents=True, exist_ok=True)

# ====== Helpers ======
def _norm_text(s: str) -> str:
    s = str(s).replace("\xa0", " ")
    s = " ".join(s.split())
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    return s.lower()

def _find_col(df: pd.DataFrame, targets: list[str]):
    """Busca columna por nombres objetivo (case/espacios/acentos-insensible)."""
    normmap = {c: _norm_text(c) for c in df.columns}
    for t in targets:
        tnorm = _norm_text(t)
        # exact
        for c, nc in normmap.items():
            if nc == tnorm:
                return c
        # contiene
        for c, nc in normmap.items():
            if tnorm in nc:
                return c
    return None

def read_sheet_if_exists(path: Path, sheet: str):
    try:
        df = pd.read_excel(path, sheet_name=sheet, engine="openpyxl")
        if df is None or df.empty:
            print(f"⚠️ Hoja vacía: '{sheet}' en {path.name}")
        return df
    except ValueError:
        print(f"⚠️ No encontré la hoja '{sheet}' en {path.name}")
        return None

# ====== Outlook: bajar último "Relevamiento de campo - Novedades LP-LM" ======
def _match_relevamiento_filename(name: str) -> bool:
    n = _norm_text(name)
    if not re.search(r"\.(xls|xlsx|xlsm|xltx|xltm)$", n):
        return False
    return "relevamiento de campo - novedades lp-lm" in n

def get_latest_relevamiento_from_outlook(save_to_dir: Path) -> Optional[Path]:
    try:
        import win32com.client as win32
    except ImportError:
        print("⚠️ Falta pywin32. Instalá: pip install pywin32")
        return None

    ns = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = ns.GetDefaultFolder(6)  # 6 = Inbox
    items = inbox.Items
    items.Sort("[ReceivedTime]", True)

    for item in items:
        try:
            atts = item.Attachments
        except Exception:
            continue
        for i in range(1, atts.Count + 1):
            att = atts.Item(i)
            fn = att.FileName or ""
            if _match_relevamiento_filename(fn):
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                out_path = save_to_dir / f"{Path(fn).stem}__{ts}{Path(fn).suffix}"
                att.SaveAsFile(str(out_path))
                print("📥 Relevamiento descargado:", out_path.name)
                return out_path
    print("⚠️ No encontré adjuntos 'Relevamiento de campo - Novedades LP-LM' en Outlook.")
    return None

# ====== Leer relevamiento (detección robusta de encabezados) ======
def read_relevamiento_table(path: Path) -> Optional[pd.DataFrame]:
    """
    Auto-detecta encabezados. Recorre hojas y busca fila que contenga 'pozo' y 'fecha'.
    """
    try:
        xls = pd.ExcelFile(path, engine="openpyxl")
    except Exception as e:
        print(f"⚠️ No pude abrir '{path.name}': {e}")
        return None

    must_have = {"pozo", "fecha"}
    nice_to_have = {"observaciones", "solicitud", "operación", "operacion", "acondicionar"}

    def _norm_cell(v) -> str:
        from unicodedata import normalize, category
        s = str(v) if v is not None else ""
        s = s.replace("\xa0", " ")
        s = " ".join(s.split())
        s = "".join(c for c in normalize("NFD", s) if category(c) != "Mn")
        return s.strip().lower()

    best_df, best_score = None, -1

    for sheet in xls.sheet_names:
        try:
            preview = pd.read_excel(xls, sheet_name=sheet, nrows=250, header=None, engine="openpyxl")
        except Exception as e:
            print(f"ℹ️ Salteo hoja '{sheet}' por error de lectura preliminar: {e}")
            continue

        header_row_idx, header_score = None, -1
        max_rows_to_scan = min(200, len(preview))
        for i in range(max_rows_to_scan):
            row_vals = [_norm_cell(v) for v in preview.iloc[i].tolist()]
            has_must = all(any(m in cell for cell in row_vals) for m in must_have)
            if not has_must:
                continue
            score = sum(any(n in cell for cell in row_vals) for n in nice_to_have)
            if score > header_score:
                header_score = score
                header_row_idx = i

        if header_row_idx is None:
            try:
                df_try = pd.read_excel(xls, sheet_name=sheet, header=0, engine="openpyxl")
                cols_norm = {_norm_cell(c) for c in df_try.columns}
                if {"pozo"} & cols_norm and {"fecha"} & cols_norm:
                    df_try = df_try.dropna(how="all").dropna(axis=1, how="all")
                    if not df_try.empty and len(df_try.columns) >= 2 and best_df is None:
                        print(f"ℹ️ Usé header=0 en hoja '{sheet}' (fallback).")
                        best_df, best_score = df_try, 0
            except Exception:
                pass
            continue

        try:
            df = pd.read_excel(xls, sheet_name=sheet, header=header_row_idx, engine="openpyxl")
            df = df.dropna(how="all").dropna(axis=1, how="all")
            if df.empty:
                continue

            cols_norm = {_norm_cell(c) for c in df.columns}
            has_core = ("pozo" in " ".join(cols_norm)) and ("fecha" in " ".join(cols_norm))
            if not has_core:
                continue

            score = sum(1 for n in nice_to_have if any(n in c for c in cols_norm))
            if score > best_score:
                best_df, best_score = df, score
                print(f"✅ Detecté encabezados en hoja '{sheet}' fila {header_row_idx+1} (1-based). Score={score}")
        except Exception as e:
            print(f"ℹ️ No pude releer hoja '{sheet}' con header={header_row_idx}: {e}")

    if best_df is None:
        print("⚠️ No se detectó una fila de encabezados confiable en ninguna hoja.")
        return None

    return best_df

# ====== Completar CARGAR MEDICION en MEDICIONES ======
def completar_cargar_medicion(df_med: pd.DataFrame, df_pot: pd.DataFrame) -> pd.DataFrame:
    # MEDICIONES
    col_fecha_med   = _find_col(df_med, ["Fecha"])
    col_pozo_norm   = _find_col(df_med, ["POZO NORMALIZADO", "Pozo normalizado"])
    col_fh_med      = _find_col(df_med, ["_FECHA_HORA", "FECHA_HORA", "fecha_hora"])
    col_cargar      = _find_col(df_med, ["CARGAR MEDICION", "Cargar medicion"])
    col_cancelada   = _find_col(df_med, ["Cancelada"])
    col_sumerg      = _find_col(df_med, ["Sumergencia"])
    # Potenciales
    col_ncpozo_pot  = _find_col(df_pot, ["nombre_corto_pozo", "nombre corto pozo"])
    col_fnvl_pot    = _find_col(df_pot, ["fecha_nvl", "fecha nvl"])

    missing = []
    for n, c in [
        ("Fecha (MEDICIONES)", col_fecha_med),
        ("POZO NORMALIZADO (MEDICIONES)", col_pozo_norm),
        ("_FECHA_HORA (MEDICIONES)", col_fh_med),
        ("nombre_corto_pozo (Potenciales)", col_ncpozo_pot),
        ("fecha_nvl (Potenciales)", col_fnvl_pot),
    ]:
        if c is None:
            missing.append(n)
    if missing:
        print("⚠️ No pude ubicar estas columnas para completar 'CARGAR MEDICION':")
        for m in missing:
            print("   -", m)
        return df_med

    med = df_med.copy()
    pot = df_pot.copy()

    med[col_fecha_med] = pd.to_datetime(med[col_fecha_med], errors="coerce", dayfirst=True)
    med[col_fh_med]    = pd.to_datetime(med[col_fh_med],    errors="coerce", dayfirst=True)
    pot[col_fnvl_pot]  = pd.to_datetime(pot[col_fnvl_pot],  errors="coerce", dayfirst=True)

    if col_cargar not in med.columns:
        med[col_cargar] = None

    med["_POZO_KEY_"] = med[col_pozo_norm].astype(str).str.strip().str.upper()
    pot["_POZO_KEY_"] = pot[col_ncpozo_pot].astype(str).str.strip().str.upper()

    med2 = med.merge(
        pot[["_POZO_KEY_", col_fnvl_pot]].rename(columns={col_fnvl_pot: "__FECHA_NVL__"}),
        on="_POZO_KEY_",
        how="left"
    )

    cond_fh_vacio  = med2[col_fh_med].isna()
    cond_match_nvl = med2["__FECHA_NVL__"].notna()
    cond_nvl_menor = med2["__FECHA_NVL__"] < med2[col_fecha_med]

    if col_cancelada is not None:
        def _is_one(x):
            try:
                return float(str(x).strip()) == 1.0
            except Exception:
                return False
        cancelada_es1 = med2[col_cancelada].apply(_is_one)
    else:
        cancelada_es1 = pd.Series(False, index=med2.index)

    if col_sumerg is not None:
        sumerg_no_indef = ~med2[col_sumerg].astype(str).str.lower().str.contains("indef", na=False)
    else:
        sumerg_no_indef = pd.Series(True, index=med2.index)

    cond_extra = (~cancelada_es1) & (sumerg_no_indef)
    cond_final = cond_fh_vacio & cond_match_nvl & cond_nvl_menor & cond_extra

    def _is_empty(x):
        if pd.isna(x):
            return True
        if isinstance(x, str) and x.strip() == "":
            return True
        return False

    mask_empty_cargar = med2[col_cargar].apply(_is_empty)
    fill_mask = cond_final & mask_empty_cargar
    med2.loc[fill_mask, col_cargar] = "CARGAR MEDICION EN SISTEMA"

    med2 = med2.drop(columns=["_POZO_KEY_", "__FECHA_NVL__"])
    print(f"ℹ️ Filas marcadas (FECHA_HORA vacío + match + fecha_nvl<Fecha + Cancelada!=1 AND Sumergencia no 'indef'): {int(fill_mask.sum())}")
    return med2

# ====== Agregar TIPO DE NIVEL QUE FALTA CARGAR en Potenciales ======
def agregar_tipo_nivel_que_falta(df_med: pd.DataFrame, df_pot: pd.DataFrame) -> pd.DataFrame:
    # --- columnas MEDICIONES ---
    c_med_pozo   = _find_col(df_med, ["POZO NORMALIZADO", "Pozo normalizado"])
    c_med_fecha  = _find_col(df_med, ["Fecha"])
    c_med_cargar = _find_col(df_med, ["CARGAR MEDICION", "Cargar medicion"])
    c_med_nivel  = _find_col(df_med, ["Nivel"])
    c_med_dina   = _find_col(df_med, ["Dina"])
    c_med_comb   = _find_col(df_med, ["Comb"])
    # --- columnas Potenciales ---
    c_pot_pozo   = _find_col(df_pot, ["nombre_corto_pozo", "nombre corto pozo"])
    c_pot_met    = _find_col(df_pot, ["met_prod", "met prod", "metodo", "método"])

    missing = []
    for n, c in [
        ("POZO NORMALIZADO (MEDICIONES)", c_med_pozo),
        ("Fecha (MEDICIONES)", c_med_fecha),
        ("CARGAR MEDICION (MEDICIONES)", c_med_cargar),
        ("Nivel (MEDICIONES)", c_med_nivel),
        ("Dina (MEDICIONES)", c_med_dina),
        ("Comb (MEDICIONES)", c_med_comb),
        ("nombre_corto_pozo (Potenciales)", c_pot_pozo),
        ("met_prod (Potenciales)", c_pot_met),
    ]:
        if c is None:
            missing.append(n)

    pot = df_pot.copy()
    if missing:
        print("⚠️ No pude ubicar columnas para 'TIPO DE NIVEL QUE FALTA CARGAR':")
        for m in missing:
            print("   -", m)
        if "TIPO DE NIVEL QUE FALTA CARGAR" not in pot.columns:
            pot["TIPO DE NIVEL QUE FALTA CARGAR"] = None
        return pot

    med = df_med.copy()
    med[c_med_fecha] = pd.to_datetime(med[c_med_fecha], errors="coerce", dayfirst=True)
    med["_POZO_KEY_"] = med[c_med_pozo].astype(str).str.strip().str.upper()
    pot["_POZO_KEY_"] = pot[c_pot_pozo].astype(str).str.strip().str.upper()

    med_sel = med[med[c_med_cargar].astype(str).str.strip().str.upper() == "CARGAR MEDICION EN SISTEMA"].copy()
    med_sel = med_sel.sort_values(c_med_fecha).groupby("_POZO_KEY_", as_index=False).tail(1)

    cols_traer = ["_POZO_KEY_", c_med_fecha, c_med_nivel, c_med_dina, c_med_comb]
    med_min = med_sel[cols_traer].rename(columns={
        c_med_fecha: "__FECHA_MED_REC__",
        c_med_nivel: "__NIVEL__",
        c_med_dina:  "__DINA__",
        c_med_comb:  "__COMB__",
    })

    pot2 = pot.merge(med_min, on="_POZO_KEY_", how="left")

    def _is_one(x):
        try:
            return float(str(x).strip()) == 1.0
        except Exception:
            return False
    def _is_empty(x):
        if pd.isna(x):
            return True
        if isinstance(x, str) and x.strip() == "":
            return True
        return False

    # Comparar método con acento y sin acento
    es_bm = pot2[c_pot_met].apply(lambda v: _norm_text(v) == "bombeo mecanico")
    fstr = pot2["__FECHA_MED_REC__"].dt.strftime("%Y-%m-%d %H:%M:%S").where(pot2["__FECHA_MED_REC__"].notna(), "")

    nivel_1   = pot2["__NIVEL__"].apply(_is_one)
    nivel_vac = pot2["__NIVEL__"].apply(_is_empty)
    dina_1    = pot2["__DINA__"].apply(_is_one)
    dina_vac  = pot2["__DINA__"].apply(_is_empty)
    comb_1    = pot2["__COMB__"].apply(_is_one)
    comb_vac  = pot2["__COMB__"].apply(_is_empty)

    col_out = "TIPO DE NIVEL QUE FALTA CARGAR"
    if col_out not in pot2.columns:
        pot2[col_out] = None

    m1 = es_bm & nivel_1 & dina_vac & comb_vac & pot2["__FECHA_MED_REC__"].notna()
    pot2.loc[m1, col_out] = "CARGAR NIVEL – NO SE PUDO REALIZAR DINA " + fstr[m1]
    m2 = es_bm & nivel_vac & dina_1 & comb_vac & pot2["__FECHA_MED_REC__"].notna()
    pot2.loc[m2, col_out] = "CARGAR DINAMOMETRIA – NO SE PUDO REALIZAR NIVEL " + fstr[m2]
    m3 = es_bm & nivel_vac & dina_vac & comb_1 & pot2["__FECHA_MED_REC__"].notna()
    pot2.loc[m3, col_out] = "CARGAR DINA Y NIVEL " + fstr[m3]

    pot2 = pot2.drop(columns=["_POZO_KEY_", "__FECHA_MED_REC__", "__NIVEL__", "__DINA__", "__COMB__"])
    return pot2

# ====== Inyectar Observaciones / Solicitud / Operación Realizada / Acondicionar -> MEDICIONES ======
def merge_relevamiento_into_mediciones(df_med: pd.DataFrame, relevamiento_df: pd.DataFrame) -> pd.DataFrame:
    med = df_med.copy()
    rel = relevamiento_df.copy()

    c_med_pozo  = _find_col(med, ["Pozo"])
    c_med_fecha = _find_col(med, ["Fecha"])

    c_rel_pozo  = _find_col(rel, ["Pozo"])
    c_rel_fecha = _find_col(rel, ["Fecha", "fecha relevamiento", "fecha_relevamiento"])

    c_rel_obs   = _find_col(rel, ["Observaciones"])
    c_rel_solic = _find_col(rel, ["Solicitud", "SOLICITUD"])
    c_rel_oper  = _find_col(rel, ["Operación realizada", "OPERACIÓN REALIZADA", "Operacion realizada"])
    c_rel_acond = _find_col(rel, ["Acondicionar", "ACONDICIONAR"])

    missing = []
    for n, c in [("Pozo (MEDICIONES)", c_med_pozo), ("Fecha (MEDICIONES)", c_med_fecha),
                 ("Pozo (Relevamiento)", c_rel_pozo), ("Fecha (Relevamiento)", c_rel_fecha)]:
        if c is None:
            missing.append(n)
    if missing:
        print("⚠️ No se puede mergear relevamiento. Faltan columnas:")
        for m in missing:
            print("   -", m)
        return med

    med[c_med_fecha] = pd.to_datetime(med[c_med_fecha], errors="coerce", dayfirst=True)
    rel[c_rel_fecha] = pd.to_datetime(rel[c_rel_fecha], errors="coerce", dayfirst=True)

    med["_POZO_KEY_"] = med[c_med_pozo].astype(str).str.strip().str.upper()
    rel["_POZO_KEY_"] = rel[c_rel_pozo].astype(str).str.strip().str.upper()
    med["_FECHA_DIA_"] = med[c_med_fecha].dt.date
    rel["_FECHA_DIA_"] = rel[c_rel_fecha].dt.date

    traer_cols = [c for c in [c_rel_obs, c_rel_solic, c_rel_oper, c_rel_acond] if c is not None]
    rel_sub = rel[["_POZO_KEY_", "_FECHA_DIA_"] + traer_cols].copy()
    if not rel_sub.empty:
        rel_sub = rel_sub.groupby(["_POZO_KEY_", "_FECHA_DIA_"], as_index=False).last()

    med2 = med.merge(rel_sub, on=["_POZO_KEY_", "_FECHA_DIA_"], how="left", suffixes=("", "_rel"))

    def _set_or_create(src_col, dst_name):
        if src_col is None:
            if dst_name not in med2.columns:
                med2[dst_name] = pd.NA
            return
        if dst_name in med2.columns:
            mask = med2[src_col].notna()
            med2.loc[mask, dst_name] = med2.loc[mask, src_col]
            med2.drop(columns=[src_col], inplace=True)
        else:
            med2.rename(columns={src_col: dst_name}, inplace=True)

    _set_or_create(c_rel_obs,   "Observaciones (Relev)")
    _set_or_create(c_rel_solic, "SOLICITUD (Relev)")
    _set_or_create(c_rel_oper,  "OPERACIÓN REALIZADA (Relev)")
    _set_or_create(c_rel_acond, "ACONDICIONAR (Relev)")

    med2 = med2.drop(columns=[col for col in ["_POZO_KEY_", "_FECHA_DIA_"] if col in med2.columns])
    return med2

# ====== Enriquecer "Potenciales y controles" con columnas desde MEDICIONES ======
def enriquecer_potenciales_con_relev_desde_med(df_pot: pd.DataFrame, df_med: pd.DataFrame) -> pd.DataFrame:
    pot = df_pot.copy()
    med = df_med.copy()

    c_pot_pozo  = _find_col(pot, ["nombre_corto_pozo", "nombre corto pozo"])
    c_med_pozo  = _find_col(med, ["POZO NORMALIZADO", "Pozo normalizado"])
    c_med_fecha = _find_col(med, ["Fecha"])

    if any(x is None for x in [c_pot_pozo, c_med_pozo, c_med_fecha]):
        print("⚠️ No se puede enriquecer Potenciales: faltan columnas clave (pozo/fecha).")
        return pot

    c_obs   = _find_col(med, ["Observaciones (Relev)", "Observaciones"])
    c_solic = _find_col(med, ["SOLICITUD (Relev)", "Solicitud (Relev)"])
    c_oper  = _find_col(med, ["OPERACIÓN REALIZADA (Relev)", "Operacion realizada (Relev)", "Operación realizada (Relev)"])
    c_acond = _find_col(med, ["ACONDICIONAR (Relev)", "Acondicionar (Relev)"])

    med[c_med_fecha] = pd.to_datetime(med[c_med_fecha], errors="coerce", dayfirst=True)
    med["_POZO_KEY_"] = med[c_med_pozo].astype(str).str.strip().str.upper()

    if c_obs is None:   med["__OBS_TMP__"]   = pd.NA; c_obs   = "__OBS_TMP__"
    if c_solic is None: med["__SOL_TMP__"]   = pd.NA; c_solic = "__SOL_TMP__"
    if c_oper is None:  med["__OPER_TMP__"]  = pd.NA; c_oper  = "__OPER_TMP__"
    if c_acond is None: med["__ACOND_TMP__"] = pd.NA; c_acond = "__ACOND_TMP__"

    med_sorted = med.sort_values(c_med_fecha)
    med_last = med_sorted.groupby("_POZO_KEY_", as_index=False).tail(1)[
        ["_POZO_KEY_", c_med_fecha, c_obs, c_solic, c_oper, c_acond]
    ].rename(columns={
        c_med_fecha: "Fecha (más reciente)",
        c_obs:       "Observaciones_rel",
        c_solic:     "SOLICITUD (Relev)",
        c_oper:      "OPERACIÓN REALIZADA (Relev)",
        c_acond:     "ACONDICIONAR (Relev)"
    })

    pot["_POZO_KEY_"] = pot[c_pot_pozo].astype(str).str.strip().str.upper()
    pot_enr = pot.merge(med_last, on="_POZO_KEY_", how="left").drop(columns=["_POZO_KEY_"])

    if "Fecha (más reciente)" in pot_enr.columns:
        pot_enr["Fecha (más reciente)"] = pd.to_datetime(pot_enr["Fecha (más reciente)"], errors="coerce")

    for tmp in ["__OBS_TMP__", "__SOL_TMP__", "__OPER_TMP__", "__ACOND_TMP__"]:
        if tmp in pot_enr.columns:
            pot_enr.drop(columns=[tmp], inplace=True, errors="ignore")

    return pot_enr

# ====== NUEVA REGLA sobre "Potenciales y controles" ======
def regla_no_bm_observaciones_fecha(df_pot: pd.DataFrame) -> pd.DataFrame:
    """
    Si met_prod ≠ 'Bombeo Mecánico' Y Observaciones_rel vacío
    Y 'Fecha (más reciente)' > fecha_nvl
    Y 'TIPO DE NIVEL QUE FALTA CARGAR' vacío
    => set 'CARGAR MEDICION DE NIVEL'
    """
    pot = df_pot.copy()

    c_met     = _find_col(pot, ["met_prod", "met prod", "metodo", "método"])
    c_obs_rel = _find_col(pot, ["Observaciones_rel"])
    c_frec    = _find_col(pot, ["Fecha (más reciente)"])
    c_fnvl    = _find_col(pot, ["fecha_nvl", "fecha nvl"])
    c_out     = _find_col(pot, ["TIPO DE NIVEL QUE FALTA CARGAR"])

    # Si faltan columnas clave, no aplicamos
    if any(x is None for x in [c_met, c_obs_rel, c_frec, c_fnvl]):
        print("ℹ️ Regla no-BM no aplicada (faltan columnas clave en Potenciales).")
        return pot

    if c_out is None:
        c_out = "TIPO DE NIVEL QUE FALTA CARGAR"
        pot[c_out] = None

    # Fechas a datetime
    pot[c_frec] = pd.to_datetime(pot[c_frec], errors="coerce", dayfirst=True)
    pot[c_fnvl] = pd.to_datetime(pot[c_fnvl], errors="coerce", dayfirst=True)

    # Helpers
    def _is_empty(x):
        if pd.isna(x):
            return True
        if isinstance(x, str) and x.strip() == "":
            return True
        return False

    not_bm   = pot[c_met].apply(lambda v: _norm_text(v) != "bombeo mecanico")
    obs_vac  = pot[c_obs_rel].apply(_is_empty)
    fecha_ok = pot[c_frec] > pot[c_fnvl]
    out_vac  = pot[c_out].apply(_is_empty)

    mask = not_bm & obs_vac & fecha_ok & out_vac
    pot.loc[mask, c_out] = "CARGAR MEDICION DE NIVEL"

    if mask.any():
        print(f"ℹ️ Regla no-BM aplicada en {int(mask.sum())} fila(s).")

    return pot

def limpiar_tipo_nivel_si_fechas_iguales(df_pot: pd.DataFrame) -> pd.DataFrame:
    """
    Si fecha_nvl (día) == Fecha (más reciente) (día) -> borrar 'TIPO DE NIVEL QUE FALTA CARGAR'.
    Ignora horas/minutos; compara solo la parte de fecha.
    """
    pot = df_pot.copy()

    c_out  = _find_col(pot, ["TIPO DE NIVEL QUE FALTA CARGAR"])
    c_fnvl = _find_col(pot, ["fecha_nvl", "fecha nvl"])
    c_frec = _find_col(pot, ["Fecha (más reciente)"])

    if any(x is None for x in [c_out, c_fnvl, c_frec]):
        print("ℹ️ Limpieza fechas iguales no aplicada (faltan columnas en Potenciales).")
        return pot

    # Asegurar datetime
    pot[c_fnvl] = pd.to_datetime(pot[c_fnvl], errors="coerce", dayfirst=True)
    pot[c_frec] = pd.to_datetime(pot[c_frec], errors="coerce", dayfirst=True)

    # Comparar solo por fecha (día)
    mask = pot[c_fnvl].notna() & pot[c_frec].notna() & (pot[c_fnvl].dt.date == pot[c_frec].dt.date)

    # Borrar valores donde coinciden las fechas
    num = int(mask.sum())
    if num > 0:
        pot.loc[mask, c_out] = None
        print(f"ℹ️ Limpieza: borrado 'TIPO DE NIVEL...' en {num} fila(s) por fecha_nvl == Fecha (más reciente).")

    return pot

def agregar_dias_sin_medicion(df_pot: pd.DataFrame) -> pd.DataFrame:
    """
    Agrega la columna 'DIAS SIN MEDICION' = (hoy - fecha_nvl) en días.
    Compara solo la parte de fecha (ignora horas).
    """
    pot = df_pot.copy()
    c_fnvl = _find_col(pot, ["fecha_nvl", "fecha nvl"])
    if c_fnvl is None:
        print("ℹ️ No se pudo calcular 'DIAS SIN MEDICION' (no está fecha_nvl).")
        return pot

    pot[c_fnvl] = pd.to_datetime(pot[c_fnvl], errors="coerce", dayfirst=True)

    hoy = pd.Timestamp.today().normalize()  # fecha de hoy (00:00)
    # Normalizamos fecha_nvl al día para ignorar horas
    dias = (hoy - pot[c_fnvl].dt.normalize()).dt.days

    # Si querés enteros con soporte para NaN:
    pot["DIAS SIN MEDICION"] = dias.astype("Int64")
    return pot

# ====== MERGE MATRIZ -> Potenciales y controles ======
def merge_matriz_into_potenciales(df_pot: pd.DataFrame, df_matriz: pd.DataFrame) -> pd.DataFrame:
    """
    Mergea por nombre_corto_pozo (Potenciales) ↔ NOMBRE_CORTO_POZO (MATRIZ)
    y trae las columnas: CRITICIDAD, ACCIONES EN PROCESO, ACCIONES FINALIZADAS, VER.

    - No borra columnas existentes en Potenciales.
    - Si la columna ya existe en Potenciales, sobrescribe solo donde viene valor no nulo desde MATRIZ.
    """
    pot = df_pot.copy()
    mat = df_matriz.copy()

    # Claves
    c_pot_pozo = _find_col(pot, ["nombre_corto_pozo", "nombre corto pozo"])
    c_mat_pozo = _find_col(mat, ["NOMBRE_CORTO_POZO", "nombre_corto_pozo", "nombre corto pozo"])

    if any(x is None for x in [c_pot_pozo, c_mat_pozo]):
        print("⚠️ No se puede mergear MATRIZ: faltan columnas clave de pozo.")
        return pot

    # Columnas a traer
    c_crit = _find_col(mat, ["CRITICIDAD"])
    c_proc = _find_col(mat, ["ACCIONES EN PROCESO", "acciones en proceso"])
    c_fin  = _find_col(mat, ["ACCIONES FINALIZADAS", "acciones finalizadas"])
    c_ver  = _find_col(mat, ["VER", "ver"])

    traer = [c for c in [c_crit, c_proc, c_fin, c_ver] if c is not None]
    if not traer:
        print("ℹ️ MATRIZ no tiene columnas CRITICIDAD/ACCIONES/VER detectables.")
        return pot

    # Normalizar claves
    pot["_POZO_KEY_"] = pot[c_pot_pozo].astype(str).str.strip().str.upper()
    mat["_POZO_KEY_"] = mat[c_mat_pozo].astype(str).str.strip().str.upper()

    # Dedup en MATRIZ por pozo (última fila por si hay repetidos)
    mat_sub = mat[["_POZO_KEY_"] + traer].copy()
    mat_sub = mat_sub.groupby("_POZO_KEY_", as_index=False).last()

    # Merge
    merged = pot.merge(mat_sub, on="_POZO_KEY_", how="left", suffixes=("", "_mat"))

    # Helper para mover/mezclar columnas sin perder lo ya existente
    def _bring(src_col_name: str, dst_name: str):
        if src_col_name is None:
            # crear vacía si no existe
            if dst_name not in merged.columns:
                merged[dst_name] = pd.NA
            return
        # Si ya existe en Potenciales, sobreescribimos solo donde MATRIZ trae valor
        if dst_name in merged.columns:
            mask = merged[src_col_name].notna()
            merged.loc[mask, dst_name] = merged.loc[mask, src_col_name]
            if src_col_name != dst_name:
                merged.drop(columns=[src_col_name], inplace=True)
        else:
            merged.rename(columns={src_col_name: dst_name}, inplace=True)

    _bring(c_crit, "CRITICIDAD")
    _bring(c_proc, "ACCIONES EN PROCESO")
    _bring(c_fin,  "ACCIONES FINALIZADAS")
    _bring(c_ver,  "VER")

    # limpiar clave
    merged.drop(columns=["_POZO_KEY_"], inplace=True)

    return merged

# ====== POST PULLING: merge con "Último evento" y marcar condición ======
def _safe_get_col(df: pd.DataFrame, preferred_name: str, aliases: list[str]) -> str | None:
    """
    Devuelve el nombre REAL de la columna en df buscando por:
    1) match exacto con preferred_name
    2) match por normalizado contra cualquiera de 'aliases'
    Si no encuentra, devuelve None.
    """
    if preferred_name in df.columns:
        return preferred_name
    normmap = {c: _norm_text(c) for c in df.columns}
    alias_norms = {_norm_text(a) for a in aliases}
    for real_name, norm_name in normmap.items():
        if norm_name in alias_norms:
            return real_name
    return None


def agregar_post_pulling(df_pot: pd.DataFrame, df_ult: pd.DataFrame) -> pd.DataFrame:
    """
    POST PULLING:
    Merge nombre_pozo (Potenciales) ↔ nombre (Último evento).
    Si END_STATUS = 'COMPLETADO' y END_DATE > fecha_nvl  => 'HACER MEDICION LUEGO DEL PULLING'.
    """
    pot = df_pot.copy()
    ult = df_ult.copy()

    # Claves
    c_pot_nombrepozo = _find_col(pot, ["nombre_pozo", "nombre pozo"])
    c_ult_nombre     = _find_col(ult, ["nombre"])  # en tu SQL es 'NOMBRE'
    c_fnvl           = _find_col(pot, ["fecha_nvl", "fecha nvl"])

    # Campos en “Último evento”
    c_end_status_src = _find_col(ult, ["END_STATUS", "end_status", "end status"])
    c_end_date_src   = _find_col(ult, ["END_DATE", "end_date", "end date"])

    missing = [label for label, col in [
        ("nombre_pozo (Potenciales)", c_pot_nombrepozo),
        ("nombre (Último evento)",    c_ult_nombre),
        ("fecha_nvl (Potenciales)",   c_fnvl),
        ("END_STATUS (Último evento)",c_end_status_src),
        ("END_DATE (Último evento)",  c_end_date_src),
    ] if col is None]

    if missing:
        print("ℹ️ POST PULLING no aplicado; faltan columnas:", ", ".join(missing))
        return pot

    # Normalizar claves
    pot["_POZO_KEY_"] = pot[c_pot_nombrepozo].astype(str).str.strip().str.upper()
    ult["_POZO_KEY_"] = ult[c_ult_nombre].astype(str).str.strip().str.upper()

    # Tipos fecha
    pot[c_fnvl] = pd.to_datetime(pot[c_fnvl], errors="coerce", dayfirst=True)
    ult[c_end_date_src] = pd.to_datetime(ult[c_end_date_src], errors="coerce", dayfirst=True)

    # Último evento por pozo (mayor END_DATE)
    ult_sorted = ult.sort_values(c_end_date_src)
    ult_last = ult_sorted.groupby("_POZO_KEY_", as_index=False).tail(1)[["_POZO_KEY_", c_end_status_src, c_end_date_src]]

    # Merge
    pot2 = pot.merge(ult_last, on="_POZO_KEY_", how="left")

    # Resolver nombres REALES tras el merge (por si vinieron con variantes)
    c_end_status = _safe_get_col(
        pot2, c_end_status_src, ["END_STATUS", "end_status", "end status"]
    )
    c_end_date = _safe_get_col(
        pot2, c_end_date_src, ["END_DATE", "end_date", "end date"]
    )

    if any(x is None for x in [c_end_status, c_end_date]):
        print("ℹ️ POST PULLING no aplicado; no pude resolver columnas END_STATUS/END_DATE tras el merge.")
        # print("Cols en pot2:", list(pot2.columns))
        pot2.drop(columns=["_POZO_KEY_"], inplace=True, errors="ignore")
        return pot2

    # Columna destino
    col_out = "POST PULLING"
    if col_out not in pot2.columns:
        pot2[col_out] = None

    status_compl = pot2[c_end_status].astype(str).str.strip().str.upper() == "COMPLETADO"
    end_mayor_nvl = pot2[c_end_date] > pot2[c_fnvl]

    mask = status_compl & end_mayor_nvl
    pot2.loc[mask, col_out] = "HACER MEDICION LUEGO DEL PULLING"

    pot2.drop(columns=["_POZO_KEY_"], inplace=True, errors="ignore")
    return pot2

# ====== POZO NUEVO/REPARADO (último año, REP/TER, COMPLETADO) ======
# ====== POZO NUEVO/REPARADO (último año, REP/TER, COMPLETADO) ======
def marcar_pozo_nuevo_reparado(
    df_pot: pd.DataFrame,
    df_ult: pd.DataFrame,
    overwrite_post_pulling: bool = True
) -> pd.DataFrame:
    pot = df_pot.copy()
    ult = df_ult.copy()

    # ----- columnas clave
    c_pot_nombrepozo = _find_col(pot, ["nombre_pozo", "nombre pozo"])
    c_ult_nombre     = _find_col(ult, ["nombre"])  # en tu SQL es 'NOMBRE'
    c_end_status     = _find_col(ult, ["END_STATUS", "end_status", "end status"])
    c_end_date       = _find_col(ult, ["END_DATE", "end_date", "end date"])
    c_evntcode       = _find_col(ult, ["evntcode", "EVNTCODE", "event_code", "EVENT_CODE"])

    missing = [lab for lab, col in [
        ("nombre_pozo (Potenciales)", c_pot_nombrepozo),
        ("nombre (Último evento)",    c_ult_nombre),
        ("END_STATUS (Último evento)",c_end_status),
        ("END_DATE (Último evento)",  c_end_date),
        ("EVNTCODE (Último evento)",  c_evntcode),
    ] if col is None]
    if missing:
        print("ℹ️ 'POZO NUEVO/REPARADO' no aplicado; faltan columnas: " + ", ".join(missing))
        return pot

    # ----- normalizar claves
    pot["_POZO_KEY_"] = pot[c_pot_nombrepozo].astype(str).str.strip().str.upper()
    ult["_POZO_KEY_"] = ult[c_ult_nombre].astype(str).str.strip().str.upper()

    # ----- tipos y ventana temporal
    ult[c_end_date] = pd.to_datetime(ult[c_end_date], errors="coerce", dayfirst=True)
    hoy = pd.Timestamp.today().normalize()
    desde = hoy - pd.Timedelta(days=365)

    # ----- condición en Último evento
    status_ok = ult[c_end_status].astype(str).str.strip().str.upper().eq("COMPLETADO")
    evnt_ok   = ult[c_evntcode].astype(str).str.upper().str.contains(r"(REP|TER)", na=False)
    fecha_ok  = ult[c_end_date].between(desde, hoy, inclusive="both")

    ult_ok = ult[status_ok & evnt_ok & fecha_ok][["_POZO_KEY_"]].drop_duplicates()

    # ----- preparar columna destino
    col_flag = "POZO NUEVO/REPARADO"
    if col_flag not in pot.columns:
        pot[col_flag] = pd.NA

    # ----- marcar en Potenciales
    mask = pot["_POZO_KEY_"].isin(ult_ok["_POZO_KEY_"])
    n = int(mask.sum())
    if n > 0:
        pot.loc[mask, col_flag] = "POZO NUEVO/REPARADO"
        if overwrite_post_pulling:
            if "POST PULLING" not in pot.columns:
                pot["POST PULLING"] = pd.NA
            pot.loc[mask, "POST PULLING"] = "POZO NUEVO/REPARADO"
        print(f"ℹ️ 'POZO NUEVO/REPARADO' marcado en {n} pozo(s).")

    pot.drop(columns=["_POZO_KEY_"], inplace=True, errors="ignore")
    return pot

def marcar_pozo_frecuente(
    df_pot: pd.DataFrame,
    df_evt: pd.DataFrame,
    window_days: int = 365,
    min_events: int = 2,
    status_ok: str = "COMPLETADO"
) -> pd.DataFrame:
    """
    Agrega columna 'Pozo frecuente' a Potenciales y controles, marcando 'POZO FRECUENTE'
    si en Eventos 2024+ el mismo nombre_corto tiene >= min_events end_date distintos,
    con end_status == COMPLETADO, dentro de los últimos 'window_days' días.
    """
    pot = df_pot.copy()
    evt = df_evt.copy()

    # --- columnas necesarias
    c_pot_ncorto = _find_col(pot, ["nombre_corto", "nombre corto"])
    c_evt_ncorto = _find_col(evt, ["nombre_corto", "nombre corto"])
    c_evt_status = _find_col(evt, ["END_STATUS", "end_status", "end status", "status"])
    c_evt_enddt  = _find_col(evt, ["END_DATE", "end_date", "end date"])

    faltan = [lab for lab, col in [
        ("nombre_corto (Potenciales)", c_pot_ncorto),
        ("nombre_corto (Eventos 2024+)", c_evt_ncorto),
        ("END_STATUS (Eventos 2024+)", c_evt_status),
        ("END_DATE (Eventos 2024+)", c_evt_enddt),
    ] if col is None]
    if faltan:
        print("ℹ️ 'Pozo frecuente' no aplicado; faltan columnas: " + ", ".join(faltan))
        if "Pozo frecuente" not in pot.columns:
            pot["Pozo frecuente"] = pd.NA
        return pot

    # --- normalizar claves
    pot["_POZO_KEY_"] = pot[c_pot_ncorto].astype(str).str.strip().str.upper()
    evt["_POZO_KEY_"] = evt[c_evt_ncorto].astype(str).str.strip().str.upper()

    # --- fechas y filtros
    evt[c_evt_enddt] = pd.to_datetime(evt[c_evt_enddt], errors="coerce", dayfirst=True)
    hoy = pd.Timestamp.today().normalize()
    desde = hoy - pd.Timedelta(days=window_days)

    status_ok_up = status_ok.strip().upper()
    cond_status  = evt[c_evt_status].astype(str).str.strip().str.upper().eq(status_ok_up)
    cond_fecha   = evt[c_evt_enddt].between(desde, hoy, inclusive="both")
    cond_vivas   = evt[c_evt_enddt].notna()

    evt_ok = evt[cond_status & cond_fecha & cond_vivas].copy()

    # --- contar end_date DISTINTOS por pozo
    counts = (evt_ok.groupby("_POZO_KEY_")[c_evt_enddt]
              .nunique(dropna=True)
              .reset_index(name="n_end_dates"))

    frecuentes = set(counts.loc[counts["n_end_dates"] >= min_events, "_POZO_KEY_"])

    # --- marcar en Potenciales
    if "Pozo frecuente" not in pot.columns:
        pot["Pozo frecuente"] = pd.NA

    mask = pot["_POZO_KEY_"].isin(frecuentes)
    pot.loc[mask, "Pozo frecuente"] = "POZO FRECUENTE"

    pot.drop(columns=["_POZO_KEY_"], inplace=True, errors="ignore")
    return pot

from typing import Optional, List
import pandas as pd

def agregar_estado_actual(df_pot: pd.DataFrame, df_perd: pd.DataFrame) -> pd.DataFrame:
    """
    ESTADO ACTUAL:
    - Merge lógico por nombre_corto (Potenciales) ↔ comp_s_name (Pérdidas)
    - Si aparece en Pérdidas -> 'PARADO', si no -> 'EN MARCHA'
    """
    pot = df_pot.copy()

    # Buscar columnas con tolerancia de nombre
    c_pot_ncorto = _find_col(pot, ["nombre_corto", "nombre corto"])
    c_perd_comp  = _find_col(df_perd, ["comp_s_name", "comp s name", "comp_s", "comp"])

    if c_pot_ncorto is None:
        print("⚠️ 'ESTADO ACTUAL' no aplicado: falta nombre_corto en Potenciales.")
        if "ESTADO ACTUAL" not in pot.columns:
            pot["ESTADO ACTUAL"] = pd.NA
        return pot

    if df_perd is None or df_perd.empty or c_perd_comp is None:
        # Si no hay pérdidas, consideramos todos EN MARCHA (no hay evidencia de paro)
        pot["ESTADO ACTUAL"] = "EN MARCHA"
        return pot

    # Normalizar claves
    pot["_POZO_KEY_"] = pot[c_pot_ncorto].astype(str).str.strip().str.upper()
    perd = df_perd.copy()
    perd["_POZO_KEY_"] = perd[c_perd_comp].astype(str).str.strip().str.upper()

    # Conjunto de pozos que aparecen en Pérdidas
    pozos_parados = set(perd["_POZO_KEY_"].dropna().unique())

    # Marcar
    pot["ESTADO ACTUAL"] = "EN MARCHA"
    mask = pot["_POZO_KEY_"].isin(pozos_parados)
    pot.loc[mask, "ESTADO ACTUAL"] = "PARADO"

    pot.drop(columns=["_POZO_KEY_"], inplace=True, errors="ignore")
    return pot


def filtrar_columnas_potenciales(df_pot: pd.DataFrame) -> pd.DataFrame:
    """
    Mantiene SOLO las columnas indicadas por el usuario y en ese orden.
    Si alguna no existe, la crea vacía (NaN).
    También renombra columnas detectadas a los nombres objetivo.
    """
    pot = df_pot.copy()

    objetivos: List[str] = [
        "nombre_corto","nombre_corto_pozo","nombre_pozo","estado","met_prod","nivel_3","nivel_5",
        "fecha_nvl","fecha_ult_control","tiponivel","prod_oil","prod_gas","total_vol","sumergencia",
        "regimen_prim","prof_ent_bomba","regimen_sec","Fecha (más reciente)","Observaciones_rel",
        "SOLICITUD (Relev)","OPERACIÓN REALIZADA (Relev)","ACONDICIONAR (Relev)",
        "TIPO DE NIVEL QUE FALTA CARGAR","DIAS SIN MEDICION","CRITICIDAD","ACCIONES EN PROCESO",
        "ACCIONES FINALIZADAS","VER","POST PULLING","POZO NUEVO/REPARADO","Pozo frecuente","ESTADO ACTUAL","CATEGORIAS","ACCIONES PARA ACONDICIONAR","CARGAR MEDICION"
    ]

    # Aliases mínimos para mapear nombres reales -> objetivo
    aliases = {
        "nombre_corto": ["nombre_corto", "nombre corto"],
        "nombre_corto_pozo": ["nombre_corto_pozo", "nombre corto pozo"],
        "nombre_pozo": ["nombre_pozo", "nombre pozo"],
        "estado": ["estado"],
        "met_prod": ["met_prod", "met prod", "metodo", "método"],
        "nivel_3": ["nivel_3", "nivel 3"],
        "nivel_5": ["nivel_5", "nivel 5"],
        "fecha_nvl": ["fecha_nvl", "fecha nvl"],
        "fecha_ult_control": ["fecha_ult_control", "fecha ult control", "fecha_ult_control", "FECHA_ULT_CONTROL"],
        "tiponivel": ["tiponivel", "tipo nivel", "tipo_nivel"],
        "prod_oil": ["prod_oil", "produccion aceite", "prod oil"],
        "prod_gas": ["prod_gas", "prod gas"],
        "total_vol": ["total_vol", "total vol"],
        "sumergencia": ["sumergencia"],
        "regimen_prim": ["regimen_prim", "regimen prim"],
        "prof_ent_bomba": ["prof_ent_bomba", "prof ent bomba"],
        "regimen_sec": ["regimen_sec", "regimen sec"],
        "Fecha (más reciente)": ["Fecha (más reciente)", "fecha mas reciente", "fecha (mas reciente)"],
        "Observaciones_rel": ["Observaciones_rel", "observaciones_rel"],
        "SOLICITUD (Relev)": ["SOLICITUD (Relev)", "Solicitud (Relev)"],
        "OPERACIÓN REALIZADA (Relev)": ["OPERACIÓN REALIZADA (Relev)", "Operacion realizada (Relev)", "Operación realizada (Relev)"],
        "ACONDICIONAR (Relev)": ["ACONDICIONAR (Relev)", "Acondicionar (Relev)"],
        "TIPO DE NIVEL QUE FALTA CARGAR": ["TIPO DE NIVEL QUE FALTA CARGAR"],
        "DIAS SIN MEDICION": ["DIAS SIN MEDICION", "dias sin medicion"],
        "CRITICIDAD": ["CRITICIDAD"],
        "ACCIONES EN PROCESO": ["ACCIONES EN PROCESO"],
        "ACCIONES FINALIZADAS": ["ACCIONES FINALIZADAS"],
        "VER": ["VER"],
        "POST PULLING": ["POST PULLING"],
        "POZO NUEVO/REPARADO": ["POZO NUEVO/REPARADO"],
        "Pozo frecuente": ["Pozo frecuente", "pozo frecuente"],
        "ESTADO ACTUAL": ["ESTADO ACTUAL", "estado actual"],
        "ACCIONES PARA ACONDICIONAR": ["ACCIONES PARA ACONDICIONAR", "acciones para acondicionar"],
        "CARGAR MEDICION": ["CARGAR MEDICION", "Cargar medicion"],
    }

    # Renombrar las que existan con alias
    for objetivo in objetivos:
        real = _find_col(pot, aliases.get(objetivo, [objetivo]))
        if real is not None and real != objetivo:
            pot.rename(columns={real: objetivo}, inplace=True)

    # Crear faltantes
    for objetivo in objetivos:
        if objetivo not in pot.columns:
            pot[objetivo] = pd.NA

    # Orden final y recorte
    pot = pot[objetivos]
    return pot

def marcar_categoria_A(df_pot: pd.DataFrame, overwrite: bool = True) -> pd.DataFrame:
    """
    Regla A (sin 'VER'):
    - met_prod contiene: Bombeo Mecánico / Cavidad Progresiva / Electro Sumergible
    - fecha_nvl <= Fecha (más reciente)  (comparando solo la FECHA, sin horas)
    - ACONDICIONAR (Relev) vacío o contiene: DERRAME / DESMONTADO / DETENIDO / EMPRESA OPERANDO
    - CRITICIDAD = CRITICO
    - ESTADO ACTUAL = EN MARCHA
    => CATEGORIAS = "A"
    """
    pot = df_pot.copy()

    # localizar columnas
    c_met    = _find_col(pot, ["met_prod", "met prod", "metodo", "método"])
    c_fnvl   = _find_col(pot, ["fecha_nvl", "fecha nvl"])
    c_frec   = _find_col(pot, ["Fecha (más reciente)", "fecha mas reciente"])
    c_acond  = _find_col(pot, ["ACONDICIONAR (Relev)", "Acondicionar (Relev)"])
    c_crit   = _find_col(pot, ["CRITICIDAD"])
    c_estado = _find_col(pot, ["ESTADO ACTUAL", "estado actual"])

    faltan = [name for name, col in [
        ("met_prod", c_met), ("fecha_nvl", c_fnvl), ("Fecha (más reciente)", c_frec),
        ("ACONDICIONAR (Relev)", c_acond), ("CRITICIDAD", c_crit), ("ESTADO ACTUAL", c_estado)
    ] if col is None]
    if faltan:
        print("ℹ️ 'CATEGORIAS=A' no aplicado (faltan columnas): " + ", ".join(faltan))
        if "CATEGORIAS" not in pot.columns:
            pot["CATEGORIAS"] = pd.NA
        return pot

    # helpers
    def _is_empty(x):
        if x is None or pd.isna(x): return True
        return str(x).strip() == ""

    # fechas -> comparar solo la fecha (día)
    pot[c_fnvl] = pd.to_datetime(pot[c_fnvl], errors="coerce", dayfirst=True)
    pot[c_frec] = pd.to_datetime(pot[c_frec], errors="coerce", dayfirst=True)
    fnvl_day = pot[c_fnvl].dt.date
    frec_day = pot[c_frec].dt.date
    # <= como pediste
    cond_fecha = pot[c_frec].isna() | (pot[c_frec].notna() & (pot[c_fnvl].isna() | (fnvl_day <= frec_day)))


    # met_prod permitido
    met_ok = pot[c_met].astype(str).apply(
        lambda v: any(k in _norm_text(v) for k in ["bombeo mecanico", "cavidad progresiva", "electro sumergible"])
    )

    # ACONDICIONAR vacío o contiene palabras clave
    acond_ok = pot[c_acond].apply(
        lambda v: _is_empty(v) or any(k in _norm_text(str(v)) for k in
                                      ["derrame", "desmontado", "detenido", "empresa operando"])
    )

    # CRITICIDAD = CRITICO
    crit_ok = pot[c_crit].astype(str).map(_norm_text).eq("critico")

    # ESTADO ACTUAL = EN MARCHA
    estado_ok = pot[c_estado].astype(str).str.strip().str.upper().eq("EN MARCHA")

    mask_A = met_ok & cond_fecha & acond_ok & crit_ok & estado_ok

    if "CATEGORIAS" not in pot.columns:
        pot["CATEGORIAS"] = pd.NA

    if overwrite:
        pot.loc[mask_A, "CATEGORIAS"] = "A"
    else:
        pot.loc[mask_A & pot["CATEGORIAS"].isna(), "CATEGORIAS"] = "A"

    return pot

def marcar_categoria_B(df_pot: pd.DataFrame, overwrite: bool = False) -> pd.DataFrame:
    """
    Regla B:
    - met_prod contiene: Bombeo Mecánico / Cavidad Progresiva / Electro Sumergible
    - fecha_nvl <= Fecha (más reciente)  (comparando solo la FECHA, sin horas)
    - ACONDICIONAR (Relev) vacío o contiene: DERRAME / DESMONTADO / DETENIDO / EMPRESA OPERANDO
    - CRITICIDAD = ALERTA
    - ESTADO ACTUAL = EN MARCHA
    => CATEGORIAS = "B"

    overwrite=False mantiene la prioridad de reglas previas (p.ej. A).
    """
    pot = df_pot.copy()

    # localizar columnas
    c_met    = _find_col(pot, ["met_prod", "met prod", "metodo", "método"])
    c_fnvl   = _find_col(pot, ["fecha_nvl", "fecha nvl"])
    c_frec   = _find_col(pot, ["Fecha (más reciente)", "fecha mas reciente"])
    c_acond  = _find_col(pot, ["ACONDICIONAR (Relev)", "Acondicionar (Relev)"])
    c_crit   = _find_col(pot, ["CRITICIDAD"])
    c_estado = _find_col(pot, ["ESTADO ACTUAL", "estado actual"])

    faltan = [name for name, col in [
        ("met_prod", c_met), ("fecha_nvl", c_fnvl), ("Fecha (más reciente)", c_frec),
        ("ACONDICIONAR (Relev)", c_acond), ("CRITICIDAD", c_crit), ("ESTADO ACTUAL", c_estado)
    ] if col is None]
    if faltan:
        print("ℹ️ 'CATEGORIAS=B' no aplicado (faltan columnas): " + ", ".join(faltan))
        if "CATEGORIAS" not in pot.columns:
            pot["CATEGORIAS"] = pd.NA
        return pot

    # helpers
    def _is_empty(x):
        if x is None or pd.isna(x): 
            return True
        return str(x).strip() == ""

    # fechas -> comparar solo la fecha (día)
    pot[c_fnvl] = pd.to_datetime(pot[c_fnvl], errors="coerce", dayfirst=True)
    pot[c_frec] = pd.to_datetime(pot[c_frec], errors="coerce", dayfirst=True)
    fnvl_day = pot[c_fnvl].dt.date
    frec_day = pot[c_frec].dt.date
    cond_fecha = pot[c_frec].isna() | (pot[c_frec].notna() & (pot[c_fnvl].isna() | (fnvl_day <= frec_day)))

    # condiciones
    met_ok = pot[c_met].astype(str).apply(
        lambda v: any(k in _norm_text(v) for k in ["bombeo mecanico", "cavidad progresiva", "electro sumergible"])
    )
    acond_ok = pot[c_acond].apply(
        lambda v: _is_empty(v) or any(k in _norm_text(str(v)) for k in
                                      ["derrame", "desmontado", "detenido", "empresa operando"])
    )
    crit_ok = pot[c_crit].astype(str).map(_norm_text).eq("alerta")
    estado_ok = pot[c_estado].astype(str).str.strip().str.upper().eq("EN MARCHA")

    mask_B = met_ok & cond_fecha & acond_ok & crit_ok & estado_ok

    if "CATEGORIAS" not in pot.columns:
        pot["CATEGORIAS"] = pd.NA

    if overwrite:
        pot.loc[mask_B, "CATEGORIAS"] = "B"
    else:
        pot.loc[mask_B & pot["CATEGORIAS"].isna(), "CATEGORIAS"] = "B"

    return pot

def marcar_categoria_C(df_pot: pd.DataFrame, overwrite: bool = False) -> pd.DataFrame:
    """
    Regla C:
    - met_prod contiene: Bombeo Mecánico / Cavidad Progresiva / Electro Sumergible
    - fecha_nvl <= Fecha (más reciente)  (comparando solo la FECHA, sin horas)
    - ACONDICIONAR (Relev) vacío o contiene: DERRAME / DESMONTADO / DETENIDO / EMPRESA OPERANDO
    - CRITICIDAD = NORMAL o vacía
    - ESTADO ACTUAL = EN MARCHA
    - POZO NUEVO/REPARADO = "POZO NUEVO/REPARADO"
    => CATEGORIAS = "C"
    """
    pot = df_pot.copy()

    # columnas necesarias
    c_met       = _find_col(pot, ["met_prod", "met prod", "metodo", "método"])
    c_fnvl      = _find_col(pot, ["fecha_nvl", "fecha nvl"])
    c_frec      = _find_col(pot, ["Fecha (más reciente)", "fecha mas reciente"])
    c_acond     = _find_col(pot, ["ACONDICIONAR (Relev)", "Acondicionar (Relev)"])
    c_crit      = _find_col(pot, ["CRITICIDAD"])
    c_estado    = _find_col(pot, ["ESTADO ACTUAL", "estado actual"])
    c_pozo_n_r  = _find_col(pot, ["POZO NUEVO/REPARADO"])

    faltan = [name for name, col in [
        ("met_prod", c_met), ("fecha_nvl", c_fnvl), ("Fecha (más reciente)", c_frec),
        ("ACONDICIONAR (Relev)", c_acond), ("CRITICIDAD", c_crit),
        ("ESTADO ACTUAL", c_estado), ("POZO NUEVO/REPARADO", c_pozo_n_r),
    ] if col is None]
    if faltan:
        print("ℹ️ 'CATEGORIAS=C' no aplicado (faltan columnas): " + ", ".join(faltan))
        if "CATEGORIAS" not in pot.columns:
            pot["CATEGORIAS"] = pd.NA
        return pot

    # helpers
    def _is_empty(x):
        if x is None or pd.isna(x):
            return True
        return str(x).strip() == ""

    # fechas (solo día)
    pot[c_fnvl] = pd.to_datetime(pot[c_fnvl], errors="coerce", dayfirst=True)
    pot[c_frec] = pd.to_datetime(pot[c_frec], errors="coerce", dayfirst=True)
    fnvl_day = pot[c_fnvl].dt.date
    frec_day = pot[c_frec].dt.date
    cond_fecha = pot[c_frec].isna() | (pot[c_frec].notna() & (pot[c_fnvl].isna() | (fnvl_day <= frec_day)))

    # condiciones base
    met_ok = pot[c_met].astype(str).apply(
        lambda v: any(k in _norm_text(v) for k in ["bombeo mecanico", "cavidad progresiva", "electro sumergible"])
    )
    acond_ok = pot[c_acond].apply(
        lambda v: _is_empty(v) or any(k in _norm_text(str(v)) for k in
                                      ["derrame", "desmontado", "detenido", "empresa operando"])
    )
    crit_norm = pot[c_crit].astype(str).map(_norm_text)
    crit_ok = crit_norm.eq("normal") | pot[c_crit].apply(_is_empty)

    estado_ok = pot[c_estado].astype(str).str.strip().str.upper().eq("EN MARCHA")
    pozo_nuevo_rep_ok = pot[c_pozo_n_r].astype(str).str.upper().eq("POZO NUEVO/REPARADO")

    mask_C = met_ok & cond_fecha & acond_ok & crit_ok & estado_ok & pozo_nuevo_rep_ok

    # asegurar columna destino
    if "CATEGORIAS" not in pot.columns:
        pot["CATEGORIAS"] = pd.NA

    if overwrite:
        pot.loc[mask_C, "CATEGORIAS"] = "C"
    else:
        pot.loc[mask_C & pot["CATEGORIAS"].isna(), "CATEGORIAS"] = "C"

    return pot
def marcar_categoria_D(df_pot: pd.DataFrame, overwrite: bool = False) -> pd.DataFrame:
    """
    Regla D:
    - met_prod contiene: Bombeo Mecánico / Cavidad Progresiva / Electro Sumergible
    - fecha_nvl <= Fecha (más reciente)  (comparando solo la FECHA, sin horas)
    - ACONDICIONAR (Relev) vacío o contiene: DERRAME / DESMONTADO / DETENIDO / EMPRESA OPERANDO
    - CRITICIDAD = NORMAL o vacía
    - ESTADO ACTUAL = EN MARCHA
    - Pozo frecuente = "POZO FRECUENTE"
    => CATEGORIAS = "D"
    """
    pot = df_pot.copy()

    # columnas necesarias
    c_met      = _find_col(pot, ["met_prod", "met prod", "metodo", "método"])
    c_fnvl     = _find_col(pot, ["fecha_nvl", "fecha nvl"])
    c_frec     = _find_col(pot, ["Fecha (más reciente)", "fecha mas reciente"])
    c_acond    = _find_col(pot, ["ACONDICIONAR (Relev)", "Acondicionar (Relev)"])
    c_crit     = _find_col(pot, ["CRITICIDAD"])
    c_estado   = _find_col(pot, ["ESTADO ACTUAL", "estado actual"])
    c_frecpozo = _find_col(pot, ["Pozo frecuente", "pozo frecuente"])

    faltan = [name for name, col in [
        ("met_prod", c_met), ("fecha_nvl", c_fnvl), ("Fecha (más reciente)", c_frec),
        ("ACONDICIONAR (Relev)", c_acond), ("CRITICIDAD", c_crit),
        ("ESTADO ACTUAL", c_estado), ("Pozo frecuente", c_frecpozo),
    ] if col is None]
    if faltan:
        print("ℹ️ 'CATEGORIAS=D' no aplicado (faltan columnas): " + ", ".join(faltan))
        if "CATEGORIAS" not in pot.columns:
            pot["CATEGORIAS"] = pd.NA
        return pot

    # helper
    def _is_empty(x):
        if x is None or pd.isna(x): return True
        return str(x).strip() == ""

    # fechas (solo día)
    pot[c_fnvl] = pd.to_datetime(pot[c_fnvl], errors="coerce", dayfirst=True)
    pot[c_frec] = pd.to_datetime(pot[c_frec], errors="coerce", dayfirst=True)
    fnvl_day = pot[c_fnvl].dt.date
    frec_day = pot[c_frec].dt.date
    cond_fecha = pot[c_frec].isna() | (pot[c_frec].notna() & (pot[c_fnvl].isna() | (fnvl_day <= frec_day)))

    # condiciones
    met_ok = pot[c_met].astype(str).apply(
        lambda v: any(k in _norm_text(v) for k in ["bombeo mecanico", "cavidad progresiva", "electro sumergible"])
    )
    acond_ok = pot[c_acond].apply(
        lambda v: _is_empty(v) or any(k in _norm_text(str(v)) for k in
                                      ["derrame", "desmontado", "detenido", "empresa operando"])
    )
    crit_norm = pot[c_crit].astype(str).map(_norm_text)
    crit_ok = crit_norm.eq("normal") | pot[c_crit].apply(_is_empty)

    estado_ok = pot[c_estado].astype(str).str.strip().str.upper().eq("EN MARCHA")
    pozo_frec_ok = pot[c_frecpozo].astype(str).str.upper().eq("POZO FRECUENTE")

    mask_D = met_ok & cond_fecha & acond_ok & crit_ok & estado_ok & pozo_frec_ok

    if "CATEGORIAS" not in pot.columns:
        pot["CATEGORIAS"] = pd.NA

    if overwrite:
        pot.loc[mask_D, "CATEGORIAS"] = "D"
    else:
        pot.loc[mask_D & pot["CATEGORIAS"].isna(), "CATEGORIAS"] = "D"

    return pot

def marcar_categoria_E(df_pot: pd.DataFrame, overwrite: bool = False) -> pd.DataFrame:
    """
    Regla E:
    - met_prod contiene: Bombeo Mecánico / Cavidad Progresiva / Electro Sumergible
    - fecha_nvl <= Fecha (más reciente)  (comparando solo la FECHA, sin horas)
    - ACONDICIONAR (Relev) vacío o contiene: DERRAME / DESMONTADO / DETENIDO / EMPRESA OPERANDO
    - CRITICIDAD = NORMAL o vacía
    - ESTADO ACTUAL = EN MARCHA
    - POST PULLING = "HACER MEDICION LUEGO DEL PULLING"
    => CATEGORIAS = "E"
    """
    pot = df_pot.copy()

    # columnas necesarias
    c_met       = _find_col(pot, ["met_prod", "met prod", "metodo", "método"])
    c_fnvl      = _find_col(pot, ["fecha_nvl", "fecha nvl"])
    c_frec      = _find_col(pot, ["Fecha (más reciente)", "fecha mas reciente"])
    c_acond     = _find_col(pot, ["ACONDICIONAR (Relev)", "Acondicionar (Relev)"])
    c_crit      = _find_col(pot, ["CRITICIDAD"])
    c_estado    = _find_col(pot, ["ESTADO ACTUAL", "estado actual"])
    c_postpull  = _find_col(pot, ["POST PULLING"])

    faltan = [name for name, col in [
        ("met_prod", c_met), ("fecha_nvl", c_fnvl), ("Fecha (más reciente)", c_frec),
        ("ACONDICIONAR (Relev)", c_acond), ("CRITICIDAD", c_crit),
        ("ESTADO ACTUAL", c_estado), ("POST PULLING", c_postpull),
    ] if col is None]
    if faltan:
        print("ℹ️ 'CATEGORIAS=E' no aplicado (faltan columnas): " + ", ".join(faltan))
        if "CATEGORIAS" not in pot.columns:
            pot["CATEGORIAS"] = pd.NA
        return pot

    # helpers
    def _is_empty(x):
        if x is None or pd.isna(x): 
            return True
        return str(x).strip() == ""

    # fechas (solo día)
    pot[c_fnvl] = pd.to_datetime(pot[c_fnvl], errors="coerce", dayfirst=True)
    pot[c_frec] = pd.to_datetime(pot[c_frec], errors="coerce", dayfirst=True)
    fnvl_day = pot[c_fnvl].dt.date
    frec_day = pot[c_frec].dt.date
    cond_fecha = pot[c_frec].isna() | (pot[c_frec].notna() & (pot[c_fnvl].isna() | (fnvl_day <= frec_day)))

    # condiciones
    met_ok = pot[c_met].astype(str).apply(
        lambda v: any(k in _norm_text(v) for k in ["bombeo mecanico", "cavidad progresiva", "electro sumergible"])
    )
    acond_ok = pot[c_acond].apply(
        lambda v: _is_empty(v) or any(k in _norm_text(str(v)) for k in
                                      ["derrame", "desmontado", "detenido", "empresa operando"])
    )
    crit_norm = pot[c_crit].astype(str).map(_norm_text)
    crit_ok = crit_norm.eq("normal") | pot[c_crit].apply(_is_empty)

    estado_ok = pot[c_estado].astype(str).str.strip().str.upper().eq("EN MARCHA")
    postpull_ok = pot[c_postpull].astype(str).str.upper().eq("HACER MEDICION LUEGO DEL PULLING")

    mask_E = met_ok & cond_fecha & acond_ok & crit_ok & estado_ok & postpull_ok

    # asegurar columna destino
    if "CATEGORIAS" not in pot.columns:
        pot["CATEGORIAS"] = pd.NA

    if overwrite:
        pot.loc[mask_E, "CATEGORIAS"] = "E"
    else:
        pot.loc[mask_E & pot["CATEGORIAS"].isna(), "CATEGORIAS"] = "E"

    return pot

def marcar_categoria_F(df_pot: pd.DataFrame, overwrite: bool = False) -> pd.DataFrame:
    """
    Regla F:
    - met_prod contiene: Bombeo Mecánico / Cavidad Progresiva / Electro Sumergible
    - fecha_nvl <= Fecha (más reciente)  (comparando solo la FECHA, sin horas)
    - ACONDICIONAR (Relev) vacío o contiene: DERRAME / DESMONTADO / DETENIDO / EMPRESA OPERANDO
    - CRITICIDAD = NORMAL o vacía
    - ESTADO ACTUAL = EN MARCHA
    - DIAS SIN MEDICION > 30
    => CATEGORIAS = "F"
    """
    pot = df_pot.copy()

    # columnas necesarias
    c_met    = _find_col(pot, ["met_prod", "met prod", "metodo", "método"])
    c_fnvl   = _find_col(pot, ["fecha_nvl", "fecha nvl"])
    c_frec   = _find_col(pot, ["Fecha (más reciente)", "fecha mas reciente"])
    c_acond  = _find_col(pot, ["ACONDICIONAR (Relev)", "Acondicionar (Relev)"])
    c_crit   = _find_col(pot, ["CRITICIDAD"])
    c_estado = _find_col(pot, ["ESTADO ACTUAL", "estado actual"])
    c_dias   = _find_col(pot, ["DIAS SIN MEDICION", "dias sin medicion"])

    faltan = [name for name, col in [
        ("met_prod", c_met), ("fecha_nvl", c_fnvl), ("Fecha (más reciente)", c_frec),
        ("ACONDICIONAR (Relev)", c_acond), ("CRITICIDAD", c_crit),
        ("ESTADO ACTUAL", c_estado), ("DIAS SIN MEDICION", c_dias),
    ] if col is None]
    if faltan:
        print("ℹ️ 'CATEGORIAS=F' no aplicado (faltan columnas): " + ", ".join(faltan))
        if "CATEGORIAS" not in pot.columns:
            pot["CATEGORIAS"] = pd.NA
        return pot

    # helpers
    def _is_empty(x):
        if x is None or pd.isna(x):
            return True
        return str(x).strip() == ""

    # fechas (solo día)
    pot[c_fnvl] = pd.to_datetime(pot[c_fnvl], errors="coerce", dayfirst=True)
    pot[c_frec] = pd.to_datetime(pot[c_frec], errors="coerce", dayfirst=True)
    fnvl_day = pot[c_fnvl].dt.date
    frec_day = pot[c_frec].dt.date
    cond_fecha = pot[c_frec].isna() | (pot[c_frec].notna() & (pot[c_fnvl].isna() | (fnvl_day <= frec_day)))

    # condiciones base
    met_ok = pot[c_met].astype(str).apply(
        lambda v: any(k in _norm_text(v) for k in ["bombeo mecanico", "cavidad progresiva", "electro sumergible"])
    )
    acond_ok = pot[c_acond].apply(
        lambda v: _is_empty(v) or any(k in _norm_text(str(v)) for k in
                                      ["derrame", "desmontado", "detenido", "empresa operando"])
    )
    crit_norm = pot[c_crit].astype(str).map(_norm_text)
    crit_ok = crit_norm.eq("normal") | pot[c_crit].apply(_is_empty)

    estado_ok = pot[c_estado].astype(str).str.strip().str.upper().eq("EN MARCHA")

    # DIAS SIN MEDICION > 30
    dias_num = pd.to_numeric(pot[c_dias], errors="coerce")
    dias_ok = dias_num > 30

    mask_F = met_ok & cond_fecha & acond_ok & crit_ok & estado_ok & dias_ok

    # asegurar columna destino
    if "CATEGORIAS" not in pot.columns:
        pot["CATEGORIAS"] = pd.NA

    if overwrite:
        pot.loc[mask_F, "CATEGORIAS"] = "F"
    else:
        pot.loc[mask_F & pot["CATEGORIAS"].isna(), "CATEGORIAS"] = "F"

    return pot

def marcar_categoria_G(df_pot: pd.DataFrame, overwrite: bool = False) -> pd.DataFrame:
    """
    Regla G:
    - met_prod contiene: Bombeo Mecánico / Cavidad Progresiva / Electro Sumergible
    - fecha_nvl <= Fecha (más reciente)  (comparando solo la FECHA, sin horas)
    - ACONDICIONAR (Relev) vacío o contiene: DERRAME / DESMONTADO / DETENIDO / EMPRESA OPERANDO
    - CRITICIDAD = NORMAL o vacía
    - ESTADO ACTUAL = EN MARCHA
    - DIAS SIN MEDICION < 30
    => CATEGORIAS = "G"
    """
    pot = df_pot.copy()

    # columnas necesarias
    c_met    = _find_col(pot, ["met_prod", "met prod", "metodo", "método"])
    c_fnvl   = _find_col(pot, ["fecha_nvl", "fecha nvl"])
    c_frec   = _find_col(pot, ["Fecha (más reciente)", "fecha mas reciente"])
    c_acond  = _find_col(pot, ["ACONDICIONAR (Relev)", "Acondicionar (Relev)"])
    c_crit   = _find_col(pot, ["CRITICIDAD"])
    c_estado = _find_col(pot, ["ESTADO ACTUAL", "estado actual"])
    c_dias   = _find_col(pot, ["DIAS SIN MEDICION", "dias sin medicion"])

    faltan = [name for name, col in [
        ("met_prod", c_met), ("fecha_nvl", c_fnvl), ("Fecha (más reciente)", c_frec),
        ("ACONDICIONAR (Relev)", c_acond), ("CRITICIDAD", c_crit),
        ("ESTADO ACTUAL", c_estado), ("DIAS SIN MEDICION", c_dias),
    ] if col is None]
    if faltan:
        print("ℹ️ 'CATEGORIAS=G' no aplicado (faltan columnas): " + ", ".join(faltan))
        if "CATEGORIAS" not in pot.columns:
            pot["CATEGORIAS"] = pd.NA
        return pot

    # helper
    def _is_empty(x):
        if x is None or pd.isna(x):
            return True
        return str(x).strip() == ""

    # fechas -> comparar solo día
    pot[c_fnvl] = pd.to_datetime(pot[c_fnvl], errors="coerce", dayfirst=True)
    pot[c_frec] = pd.to_datetime(pot[c_frec], errors="coerce", dayfirst=True)
    fnvl_day = pot[c_fnvl].dt.date
    frec_day = pot[c_frec].dt.date
    cond_fecha = pot[c_frec].isna() | (pot[c_frec].notna() & (pot[c_fnvl].isna() | (fnvl_day <= frec_day)))

    # condiciones base
    met_ok = pot[c_met].astype(str).apply(
        lambda v: any(k in _norm_text(v) for k in ["bombeo mecanico", "cavidad progresiva", "electro sumergible"])
    )
    acond_ok = pot[c_acond].apply(
        lambda v: _is_empty(v) or any(k in _norm_text(str(v)) for k in
                                      ["derrame", "desmontado", "detenido", "empresa operando"])
    )
    crit_norm = pot[c_crit].astype(str).map(_norm_text)
    crit_ok = crit_norm.eq("normal") | pot[c_crit].apply(_is_empty)

    estado_ok = pot[c_estado].astype(str).str.strip().str.upper().eq("EN MARCHA")

    # DIAS SIN MEDICION < 30
    dias_num = pd.to_numeric(pot[c_dias], errors="coerce")
    dias_ok = dias_num < 30

    mask_G = met_ok & cond_fecha & acond_ok & crit_ok & estado_ok & dias_ok

    # asegurar columna destino
    if "CATEGORIAS" not in pot.columns:
        pot["CATEGORIAS"] = pd.NA

    if overwrite:
        pot.loc[mask_G, "CATEGORIAS"] = "G"
    else:
        pot.loc[mask_G & pot["CATEGORIAS"].isna(), "CATEGORIAS"] = "G"

    return pot

def marcar_categoria_F1(df_pot: pd.DataFrame, overwrite: bool = False) -> pd.DataFrame:
    """
    Regla F1:
    - CATEGORIAS vacío
    - ESTADO ACTUAL = EN MARCHA
    - ACONDICIONAR (Relev) contiene {DERRAME, DESMONTADO, DETENIDO, EMPRESA OPERANDO} o está vacío
    => CATEGORIAS = "F1"
    """
    pot = df_pot.copy()

    # columnas
    c_estado = _find_col(pot, ["ESTADO ACTUAL", "estado actual"])
    c_acond  = _find_col(pot, ["ACONDICIONAR (Relev)", "Acondicionar (Relev)"])

    if "CATEGORIAS" not in pot.columns:
        pot["CATEGORIAS"] = pd.NA

    faltan = [n for n,c in [("ESTADO ACTUAL", c_estado), ("ACONDICIONAR (Relev)", c_acond)] if c is None]
    if faltan:
        print("ℹ️ 'CATEGORIAS=F1' no aplicado (faltan columnas): " + ", ".join(faltan))
        return pot

    def _is_empty(x): 
        return x is None or pd.isna(x) or str(x).strip() == ""

    cat_vacia = pot["CATEGORIAS"].isna() | (pot["CATEGORIAS"].astype(str).str.strip() == "")
    estado_ok = pot[c_estado].astype(str).str.strip().str.upper().eq("EN MARCHA")

    palabras = ["derrame", "desmontado", "detenido", "empresa operando"]
    acond_ok = pot[c_acond].apply(
        lambda v: _is_empty(v) or any(p in _norm_text(str(v)) for p in palabras)
    )

    mask = cat_vacia & estado_ok & acond_ok

    if overwrite:
        pot.loc[mask, "CATEGORIAS"] = "F1"
    else:
        pot.loc[mask & pot["CATEGORIAS"].isna(), "CATEGORIAS"] = "F1"

    return pot

def marcar_categoria_ACONDICIONAR(df_pot: pd.DataFrame, overwrite: bool = False) -> pd.DataFrame:
    """
    Regla ACONDICIONAR:
    - CATEGORIAS vacío
    - Fecha (más reciente) > fecha_nvl  (comparar SOLO fecha, sin horas)
    - ESTADO ACTUAL = EN MARCHA
    - ACONDICIONAR (Relev) NO contiene {DERRAME, DESMONTADO, DETENIDO, EMPRESA OPERANDO} y NO está vacío
    => CATEGORIAS = "ACONDICIONAR"
    """
    pot = df_pot.copy()

    # columnas
    c_fnvl   = _find_col(pot, ["fecha_nvl", "fecha nvl"])
    c_frec   = _find_col(pot, ["Fecha (más reciente)", "fecha mas reciente"])
    c_estado = _find_col(pot, ["ESTADO ACTUAL", "estado actual"])
    c_acond  = _find_col(pot, ["ACONDICIONAR (Relev)", "Acondicionar (Relev)"])

    if "CATEGORIAS" not in pot.columns:
        pot["CATEGORIAS"] = pd.NA

    faltan = [n for n,c in [
        ("fecha_nvl", c_fnvl), ("Fecha (más reciente)", c_frec),
        ("ESTADO ACTUAL", c_estado), ("ACONDICIONAR (Relev)", c_acond)
    ] if c is None]
    if faltan:
        print("ℹ️ 'CATEGORIAS=ACONDICIONAR' no aplicado (faltan columnas): " + ", ".join(faltan))
        return pot

    def _is_empty(x): 
        return x is None or pd.isna(x) or str(x).strip() == ""

    cat_vacia = pot["CATEGORIAS"].isna() | (pot["CATEGORIAS"].astype(str).str.strip() == "")
    estado_ok = pot[c_estado].astype(str).str.strip().str.upper().eq("EN MARCHA")

    # fecha (solo día) y comparación estricta >
    pot[c_fnvl] = pd.to_datetime(pot[c_fnvl], errors="coerce", dayfirst=True)
    pot[c_frec] = pd.to_datetime(pot[c_frec], errors="coerce", dayfirst=True)
    fnvl_day = pot[c_fnvl].dt.normalize()
    frec_day = pot[c_frec].dt.normalize()
    cond_fecha = pot[c_fnvl].notna() & pot[c_frec].notna() & (frec_day > fnvl_day)

    palabras = ["derrame", "desmontado", "detenido", "empresa operando"]
    acond_no = pot[c_acond].apply(
        lambda v: (not _is_empty(v)) and all(p not in _norm_text(str(v)) for p in palabras)
    )

    mask = cat_vacia & cond_fecha & estado_ok & acond_no

    if overwrite:
        pot.loc[mask, "CATEGORIAS"] = "ACONDICIONAR"
    else:
        pot.loc[mask & pot["CATEGORIAS"].isna(), "CATEGORIAS"] = "ACONDICIONAR"

    return pot

def marcar_categoria_ACONDICIONAR_VER(df_pot: pd.DataFrame, overwrite: bool = False) -> pd.DataFrame:
    """
    Regla ACONDICIONAR- VER:
    - CATEGORIAS vacío
    - Fecha (más reciente) < fecha_nvl  (comparando SOLO la fecha, sin horas/minutos)
    - ESTADO ACTUAL = EN MARCHA
    - ACONDICIONAR (Relev) NO contiene {DERRAME, DESMONTADO, DETENIDO, EMPRESA OPERANDO} y NO está vacío
    => CATEGORIAS = "ACONDICIONAR- VER"
    """
    pot = df_pot.copy()

    # columnas
    c_fnvl   = _find_col(pot, ["fecha_nvl", "fecha nvl"])
    c_frec   = _find_col(pot, ["Fecha (más reciente)", "fecha mas reciente"])
    c_estado = _find_col(pot, ["ESTADO ACTUAL", "estado actual"])
    c_acond  = _find_col(pot, ["ACONDICIONAR (Relev)", "Acondicionar (Relev)"])

    if "CATEGORIAS" not in pot.columns:
        pot["CATEGORIAS"] = pd.NA

    faltan = [n for n,c in [
        ("fecha_nvl", c_fnvl), ("Fecha (más reciente)", c_frec),
        ("ESTADO ACTUAL", c_estado), ("ACONDICIONAR (Relev)", c_acond)
    ] if c is None]
    if faltan:
        print("ℹ️ 'CATEGORIAS=ACONDICIONAR- VER' no aplicado (faltan columnas): " + ", ".join(faltan))
        return pot

    def _is_empty(x):
        return (x is None) or pd.isna(x) or (str(x).strip() == "")

    # Solo marcar si CATEGORIAS está vacío
    cat_vacia = pot["CATEGORIAS"].isna() | (pot["CATEGORIAS"].astype(str).str.strip() == "")

    # Estado actual EN MARCHA
    estado_ok = pot[c_estado].astype(str).str.strip().str.upper().eq("EN MARCHA")

    # Fechas: comparar solo la parte de fecha (día)
    pot[c_fnvl] = pd.to_datetime(pot[c_fnvl], errors="coerce", dayfirst=True)
    pot[c_frec] = pd.to_datetime(pot[c_frec], errors="coerce", dayfirst=True)
    fnvl_day = pot[c_fnvl].dt.normalize()
    frec_day = pot[c_frec].dt.normalize()
    cond_fecha = pot[c_fnvl].notna() & pot[c_frec].notna() & (frec_day <= fnvl_day)

    # ACONDICIONAR NO contiene palabras clave y NO está vacío
    palabras = ["derrame", "desmontado", "detenido", "empresa operando"]
    acond_no = pot[c_acond].apply(
        lambda v: (not _is_empty(v)) and all(p not in _norm_text(str(v)) for p in palabras)
    )

    mask = cat_vacia & cond_fecha & estado_ok & acond_no

    if overwrite:
        pot.loc[mask, "CATEGORIAS"] = "ACONDICIONAR- VER"
    else:
        pot.loc[mask & pot["CATEGORIAS"].isna(), "CATEGORIAS"] = "ACONDICIONAR- VER"

    return pot

def marcar_acciones_para_acondicionar(df_pot: pd.DataFrame) -> pd.DataFrame:
    pot = df_pot.copy()

    c_cat   = _find_col(pot, ["CATEGORIAS", "categorias"])
    c_fnvl  = _find_col(pot, ["fecha_nvl", "fecha nvl"])
    c_frec  = _find_col(pot, ["Fecha (más reciente)", "fecha mas reciente"])
    c_fin   = _find_col(pot, ["ACCIONES FINALIZADAS", "acciones finalizadas"])
    c_proc  = _find_col(pot, ["ACCIONES EN PROCESO", "acciones en proceso"])
    dest    = "ACCIONES PARA ACONDICIONAR"

    if dest not in pot.columns:
        pot[dest] = pd.NA

    faltan = [n for n,c in [
        ("CATEGORIAS", c_cat), ("fecha_nvl", c_fnvl), ("Fecha (más reciente)", c_frec),
        ("ACCIONES FINALIZADAS", c_fin), ("ACCIONES EN PROCESO", c_proc)
    ] if c is None]
    if faltan:
        print("ℹ️ 'ACCIONES PARA ACONDICIONAR' no aplicado (faltan columnas): " + ", ".join(faltan))
        return pot

    pot[c_fnvl] = pd.to_datetime(pot[c_fnvl], errors="coerce", dayfirst=True)
    pot[c_frec] = pd.to_datetime(pot[c_frec], errors="coerce", dayfirst=True)
    fnvl_day = pot[c_fnvl].dt.normalize()
    frec_day = pot[c_frec].dt.normalize()
    cond_fecha = pot[c_fnvl].notna() & pot[c_frec].notna() & (frec_day >= fnvl_day)

    # CATEGORIAS contiene ACONDICIONAR (cubre 'ACONDICIONAR' y 'ACONDICIONAR- VER')
    cat_acond = pot[c_cat].astype(str).apply(lambda v: "acondicionar" in _norm_text(v))

    # Palabras a buscar
    def _contains_any(s: str, words: list[str]) -> bool:
        t = _norm_text(s)
        return any(w in t for w in words)

    words_finalizadas = ["acondicionar"]
    words_en_proceso  = ["acondicionar", "reparar", "torcido", "danado"]  # 'dañado' -> 'danado' por normalización

    # 1) FINALIZADAS primero (tiene prioridad)
    mask_fin = cat_acond & cond_fecha & pot[c_fin].astype(str).apply(lambda s: _contains_any(s, words_finalizadas))
    pot.loc[mask_fin, dest] = "ya se finalizó acción para acondicionar"

    # 2) EN PROCESO (solo donde aún no se escribió nada)
    mask_proc = cat_acond & cond_fecha & pot[dest].isna() & pot[c_proc].astype(str).apply(lambda s: _contains_any(s, words_en_proceso))
    pot.loc[mask_proc, dest] = "acción en proceso para acondicionar"

    return pot

def marcar_cargar_medicion_en_pot(df_pot: pd.DataFrame) -> pd.DataFrame:
    pot = df_pot.copy()

    c_fnvl  = _find_col(pot, ["fecha_nvl", "fecha nvl"])
    c_frec  = _find_col(pot, ["Fecha (más reciente)", "fecha mas reciente"])
    c_oper  = _find_col(pot, ["OPERACIÓN REALIZADA (Relev)", "Operacion realizada (Relev)", "Operación realizada (Relev)"])
    dest    = "CARGAR MEDICION"

    if dest not in pot.columns:
        pot[dest] = pd.NA

    faltan = [n for n,c in [("fecha_nvl", c_fnvl), ("Fecha (más reciente)", c_frec), ("OPERACIÓN REALIZADA (Relev)", c_oper)] if c is None]
    if faltan:
        print("ℹ️ 'CARGAR MEDICION' no aplicado (faltan columnas): " + ", ".join(faltan))
        return pot

    pot[c_fnvl] = pd.to_datetime(pot[c_fnvl], errors="coerce", dayfirst=True)
    pot[c_frec] = pd.to_datetime(pot[c_frec], errors="coerce", dayfirst=True)

    fnvl_day = pot[c_fnvl].dt.normalize()
    frec_day = pot[c_frec].dt.normalize()

    cond_fecha = pot[c_fnvl].notna() & pot[c_frec].notna() & (frec_day > fnvl_day)

    # NO contiene "CANCELADA"
    oper_ok = ~pot[c_oper].astype(str).map(_norm_text).str.contains("cancelad", na=False)

    # hoy - Fecha (más reciente) < 30 días
    hoy = pd.Timestamp.today().normalize()
    delta_dias = (hoy - frec_day).dt.days
    cond_ventana = delta_dias < 30

    mask = cond_fecha & oper_ok & cond_ventana

    pot.loc[mask, dest] = "CARGAR MEDICION EN SISTEMA"
    return pot

# ====== LÓGICA PRINCIPAL ======

# ====== CONFIG extra para MATRIZ ======
MAT_PATH = Path(r"C:\Users\ry16123\Matriz_AIB_export.xlsx")
MAT_SHEET = "MATRIZ"

def main():
    if not SRC1.exists():
        raise FileNotFoundError(f"No existe SRC1: {SRC1}")
    if not SRC2.exists():
        raise FileNotFoundError(f"No existe SRC2: {SRC2}")

    # 0) Leer hojas base
    df_med = read_sheet_if_exists(SRC1, "MEDICIONES")
    targets = ["Potenciales y controles", "Pérdidas", "Último evento", "Eventos 2024+"]
    sheets_2 = {s: read_sheet_if_exists(SRC2, s) for s in targets}

    # Leer MATRIZ (opcional)
    if MAT_PATH.exists():
        df_matriz = read_sheet_if_exists(MAT_PATH, MAT_SHEET)
    else:
        df_matriz = None
        print(f"ℹ️ No existe el archivo de MATRIZ en {MAT_PATH}")

    # 1) Completar "CARGAR MEDICION" en MEDICIONES usando Potenciales
    if df_med is not None and sheets_2.get("Potenciales y controles") is not None:
        df_med = completar_cargar_medicion(df_med, sheets_2["Potenciales y controles"])
    else:
        print("⚠️ No se completó 'CARGAR MEDICION' (faltó MEDICIONES o Potenciales y controles).")

    # 2) Post-filtro global en MEDICIONES (cancelada==1 o sumergencia 'indef' apaga la marca)
    if df_med is not None:
        col_cargar    = _find_col(df_med, ["CARGAR MEDICION", "Cargar medicion"])
        col_cancelada = _find_col(df_med, ["Cancelada"])
        col_sumerg    = _find_col(df_med, ["Sumergencia"])

        if col_cargar is not None:
            if col_cancelada is not None:
                def _is_one(x):
                    try:
                        return float(str(x).strip()) == 1.0
                    except Exception:
                        return False
                mask_cancel1 = df_med[col_cancelada].apply(_is_one)
            else:
                mask_cancel1 = pd.Series(False, index=df_med.index)

            if col_sumerg is not None:
                mask_indef = df_med[col_sumerg].astype(str).str.lower().str.contains("indef", na=False)
            else:
                mask_indef = pd.Series(False, index=df_med.index)

            bad_mask = (mask_cancel1 | mask_indef) & (df_med[col_cargar].astype(str).str.strip() == "CARGAR MEDICION EN SISTEMA")
            if bad_mask.any():
                df_med.loc[bad_mask, col_cargar] = None
                print(f"ℹ️ Post-filtro: filas desmarcadas por Cancelada==1 o Sumergencia 'indef': {int(bad_mask.sum())}")

    # 3) Descargar relevamiento desde Outlook y mergear en MEDICIONES (Pozo+Fecha(día))
    rel_path = get_latest_relevamiento_from_outlook(TMP_DIR)
    if rel_path is not None and df_med is not None:
        rel_df = read_relevamiento_table(rel_path)
        if rel_df is not None:
            df_med = merge_relevamiento_into_mediciones(df_med, rel_df)
        else:
            print("⚠️ No se pudo incorporar el relevamiento (DF vacío).")
    else:
        print("ℹ️ Salteo merge de relevamiento (no hay archivo o no hay MEDICIONES).")

    # 4) Potenciales y controles: enriquecer + reglas en ORDEN correcto
    # 4) Potenciales y controles: enriquecer + reglas en ORDEN correcto
        # 4) Potenciales y controles: enriquecer + reglas en ORDEN correcto
    if df_med is not None and sheets_2.get("Potenciales y controles") is not None:
        pot = sheets_2["Potenciales y controles"]

        # a) ...
        pot = enriquecer_potenciales_con_relev_desde_med(pot, df_med)
        # b)
        pot = agregar_tipo_nivel_que_falta(df_med, pot)
        # c)
        pot = regla_no_bm_observaciones_fecha(pot)
        # d)
        pot = limpiar_tipo_nivel_si_fechas_iguales(pot)
        # e)
        pot = agregar_dias_sin_medicion(pot)
        # f)
        if df_matriz is not None:
            pot = merge_matriz_into_potenciales(pot, df_matriz)
        else:
            print("ℹ️ Salteo merge de MATRIZ en Potenciales (no se cargó la hoja MATRIZ).")

        # g) POST PULLING existente (merge con 'Último evento')
        if sheets_2.get("Último evento") is not None:
            pot = agregar_post_pulling(pot, sheets_2["Último evento"])
        else:
            print("ℹ️ Salteo 'POST PULLING' (no está la hoja 'Último evento').")

        # h) NUEVO: POZO NUEVO/REPARADO (último año, REP/TER, COMPLETADO)
        if sheets_2.get("Último evento") is not None:
            pot = marcar_pozo_nuevo_reparado(
                pot, sheets_2["Último evento"], overwrite_post_pulling=True
            )
        else:
            print("ℹ️ Salteo 'POZO NUEVO/REPARADO' (no está la hoja 'Último evento').")

        
        # i) POZO FRECUENTE (últimos 365 días, >=2 end_date COMPLETADO)
        evt_df = sheets_2.get("Eventos 2024+") 
        if evt_df is not None:
            pot = marcar_pozo_frecuente(pot, evt_df, window_days=365, min_events=2, status_ok="COMPLETADO")
        else:
            print("ℹ️ Salteo 'Pozo frecuente' (no está la hoja 'Eventos 2024+' / 'Eventos 2024').")
            
        # j) ESTADO ACTUAL (Pérdidas)
        perd_df = sheets_2.get("Pérdidas")
        if perd_df is not None:
            pot = agregar_estado_actual(pot, perd_df)
        else:
            print("ℹ️ Salteo 'ESTADO ACTUAL' (no está la hoja 'Pérdidas'); marco EN MARCHA por defecto.")
            pot["ESTADO ACTUAL"] = "EN MARCHA"

        # A primero (puede overwrite=True si querés que A tape lo que hubiera)
        pot = marcar_categoria_A(pot, overwrite=True)

        # B después (overwrite=False para NO pisar A)
        pot = marcar_categoria_B(pot, overwrite=False)
        
        # C después (no pisa A/B)
        pot = marcar_categoria_C(pot, overwrite=False)
        
        # D después (no pisa A/B/C)
        pot = marcar_categoria_D(pot, overwrite=False)
        
        # E después (no pisa A/B/C/D)
        pot = marcar_categoria_E(pot, overwrite=False)
        
        # F después (no pisa A/B/C/D/E)
        pot = marcar_categoria_F(pot, overwrite=False) 
        
        pot = marcar_categoria_G(pot, overwrite=False)
        
        # 1) F1 (no pisa categorías previas)
        pot = marcar_categoria_F1(pot, overwrite=False)
        
        # 2) ACONDICIONAR (no pisa lo ya asignado, incluido F1)
        pot = marcar_categoria_ACONDICIONAR(pot, overwrite=False)
        
        pot = marcar_categoria_ACONDICIONAR_VER(pot, overwrite=False)
        
        pot = marcar_acciones_para_acondicionar(pot)
        pot = marcar_cargar_medicion_en_pot(pot) 
        
        # K) columnas finales
        pot = filtrar_columnas_potenciales(pot)
        sheets_2["Potenciales y controles"] = pot

       

    # 5) Escribir salida (incluye escribir también la hoja MATRIZ si se encontró)
    with pd.ExcelWriter(OUT, engine="openpyxl", mode="w") as xw:
        wrote_any = False

        if df_med is not None:
            df_med.to_excel(xw, sheet_name="MEDICIONES", index=False)
            wrote_any = True

        for s in targets:
            df = sheets_2.get(s)
            if df is not None:
                df.to_excel(xw, sheet_name=s, index=False)
                wrote_any = True

        # Escribir MATRIZ como pestaña adicional si existe
        if df_matriz is not None and not df_matriz.empty:
            df_matriz.to_excel(xw, sheet_name="MATRIZ", index=False)
            wrote_any = True

    if wrote_any:
        print(f"✅ Consolidado creado: {OUT}")
    else:
        print("⚠️ No se escribió ninguna hoja (todas faltaban o estaban vacías).")




if __name__ == "__main__":
    main()


# In[7]:


#QUINTA Y ULTIMA CODIGO A EJECUTAR
# -*- coding: utf-8 -*-
# Cronograma por proximidad geográfica + pestañas extra

from pathlib import Path
import pandas as pd
import unicodedata
from datetime import datetime, timedelta, date
from typing import Optional, List
import math
import numpy as np

# ==================== CONFIG ====================
IN_XLSX  = Path(r"C:\MedicionesInbox\compilados\Mediciones_totales_LP__CONSOLIDADO_FINAL.xlsx")
SHEET_IN = "Potenciales y controles"

COORDS_XLSX = Path(r"C:\Users\ry16123\OneDrive - YPF\Escritorio\power BI\GUADAL- POWER BI\Inteligencia Artificial\coordenadas1.xlsx")
OUT_XLSX = Path(r"C:\MedicionesInbox\compilados\Cronograma_mediciones_por_proximidad.xlsx")

CAPACIDAD_POR_EQUIPO_POR_DIA = 20  # pozos por equipo por día
CATS_PRIORIDAD = ["A","B","C","D","ACONDICIONAR- VER","ACONDICIONAR","E","F","F1","G"]

# ==================== HELPERS BÁSICOS ====================
def _norm_text(s: str) -> str:
    s = "" if s is None else str(s)
    s = s.replace("\xa0", " ").strip()
    s = " ".join(s.split())
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    return s.lower()

def _find_col(df: pd.DataFrame, targets: List[str]):
    normmap = {c: _norm_text(c) for c in df.columns}
    for t in targets:
        tnorm = _norm_text(t)
        for c, nc in normmap.items():
            if nc == tnorm:
                return c
        for c, nc in normmap.items():
            if tnorm in nc:
                return c
    return None

def _is_empty(x) -> bool:
    if x is None or pd.isna(x):
        return True
    return str(x).strip() == ""

def _next_monday_from(today: date) -> date:
    wd = today.weekday()  # 0=Mon
    return today + timedelta(days=(7 - wd) % 7 or 7)

def _weekday_only(d: date) -> date:
    while d.weekday() >= 5:
        d += timedelta(days=1)
    return d

def haversine_km(lat1, lon1, lat2, lon2) -> float:
    R = 6371.0
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlmb = math.radians(lon2 - lon1)
    a = math.sin(dphi/2.0)**2 + math.cos(phi1)*math.cos(phi2)*math.sin(dlmb/2.0)**2
    c = 2*math.atan2(math.sqrt(a), math.sqrt(1 - a))
    return R*c

# ==================== NORMALIZAR POTENCIALES (SIN FILTRO) ====================
def _normalizar_potenciales(df: pd.DataFrame) -> pd.DataFrame:
    """Solo normaliza nombres/formatos; NO filtra filas."""
    d = df.copy()

    # Detectar columnas
    c_nombre   = _find_col(d, ["nombre_corto","nombre corto"])
    c_ncpozo   = _find_col(d, ["nombre_corto_pozo","nombre corto pozo"])
    c_npozo    = _find_col(d, ["nombre_pozo","nombre pozo"])
    c_estado   = _find_col(d, ["estado"])
    c_met      = _find_col(d, ["met_prod","met prod","metodo","método"])
    c_lvl3     = _find_col(d, ["nivel_3","nivel 3"])
    c_lvl5     = _find_col(d, ["nivel_5","nivel 5"])
    c_fnvl     = _find_col(d, ["fecha_nvl","fecha nvl"])
    c_fctrl    = _find_col(d, ["fecha_ult_control","fecha ult control","FECHA_ULT_CONTROL"])
    c_tipon    = _find_col(d, ["tiponivel","tipo nivel","tipo_nivel"])
    c_oil      = _find_col(d, ["prod_oil","prod oil","produccion aceite"])
    c_gas      = _find_col(d, ["prod_gas","prod gas"])
    c_tvol     = _find_col(d, ["total_vol","total vol"])
    c_sum      = _find_col(d, ["sumergencia"])
    c_rprim    = _find_col(d, ["regimen_prim","regimen prim"])
    c_pbomba   = _find_col(d, ["prof_ent_bomba","prof ent bomba"])
    c_rsec     = _find_col(d, ["regimen_sec","regimen sec"])
    c_frec     = _find_col(d, ["Fecha (más reciente)","fecha mas reciente"])
    c_obs      = _find_col(d, ["Observaciones_rel","observaciones_rel"])
    c_sol      = _find_col(d, ["SOLICITUD (Relev)","Solicitud (Relev)"])
    c_oper     = _find_col(d, ["OPERACIÓN REALIZADA (Relev)","Operacion realizada (Relev)","Operación realizada (Relev)"])
    c_acond    = _find_col(d, ["ACONDICIONAR (Relev)","Acondicionar (Relev)"])
    c_tipo_fal = _find_col(d, ["TIPO DE NIVEL QUE FALTA CARGAR"])
    c_dias     = _find_col(d, ["DIAS SIN MEDICION","dias sin medicion"])
    c_crit     = _find_col(d, ["CRITICIDAD"])
    c_accp     = _find_col(d, ["ACCIONES EN PROCESO"])
    c_accf     = _find_col(d, ["ACCIONES FINALIZADAS"])
    c_ver      = _find_col(d, ["VER"])
    c_post     = _find_col(d, ["POST PULLING"])
    c_pnr      = _find_col(d, ["POZO NUEVO/REPARADO"])
    c_frecpozo = _find_col(d, ["Pozo frecuente","pozo frecuente"])
    c_estado2  = _find_col(d, ["ESTADO ACTUAL","estado actual"])
    c_cat      = _find_col(d, ["CATEGORIAS","categorias"])
    c_accpara  = _find_col(d, ["ACCIONES PARA ACONDICIONAR","acciones para acondicionar"])
    c_cargar   = _find_col(d, ["CARGAR MEDICION","Cargar medicion"])

    rename_map = {}
    if c_nombre: rename_map[c_nombre] = "nombre_corto"
    if c_ncpozo: rename_map[c_ncpozo] = "nombre_corto_pozo"
    if c_npozo:  rename_map[c_npozo]  = "nombre_pozo"
    if c_estado: rename_map[c_estado] = "estado"
    if c_met:    rename_map[c_met]    = "met_prod"
    if c_lvl3:   rename_map[c_lvl3]   = "nivel_3"
    if c_lvl5:   rename_map[c_lvl5]   = "nivel_5"
    if c_fnvl:   rename_map[c_fnvl]   = "fecha_nvl"
    if c_fctrl:  rename_map[c_fctrl]  = "fecha_ult_control"
    if c_tipon:  rename_map[c_tipon]  = "tiponivel"
    if c_oil:    rename_map[c_oil]    = "prod_oil"
    if c_gas:    rename_map[c_gas]    = "prod_gas"
    if c_tvol:   rename_map[c_tvol]   = "total_vol"
    if c_sum:    rename_map[c_sum]    = "sumergencia"
    if c_rprim:  rename_map[c_rprim]  = "regimen_prim"
    if c_pbomba: rename_map[c_pbomba] = "prof_ent_bomba"
    if c_rsec:   rename_map[c_rsec]   = "regimen_sec"
    if c_frec:   rename_map[c_frec]   = "Fecha (más reciente)"
    if c_obs:    rename_map[c_obs]    = "Observaciones_rel"
    if c_sol:    rename_map[c_sol]    = "SOLICITUD (Relev)"
    if c_oper:   rename_map[c_oper]   = "OPERACIÓN REALIZADA (Relev)"
    if c_acond:  rename_map[c_acond]  = "ACONDICIONAR (Relev)"
    if c_tipo_fal: rename_map[c_tipo_fal] = "TIPO DE NIVEL QUE FALTA CARGAR"
    if c_dias:   rename_map[c_dias]   = "DIAS SIN MEDICION"
    if c_crit:   rename_map[c_crit]   = "CRITICIDAD"
    if c_accp:   rename_map[c_accp]   = "ACCIONES EN PROCESO"
    if c_accf:   rename_map[c_accf]   = "ACCIONES FINALIZADAS"
    if c_ver:    rename_map[c_ver]    = "VER"
    if c_post:   rename_map[c_post]   = "POST PULLING"
    if c_pnr:    rename_map[c_pnr]    = "POZO NUEVO/REPARADO"
    if c_frecpozo: rename_map[c_frecpozo] = "Pozo frecuente"
    if c_estado2: rename_map[c_estado2]   = "ESTADO ACTUAL"
    if c_cat:      rename_map[c_cat]      = "CATEGORIAS"
    if c_accpara:  rename_map[c_accpara]  = "ACCIONES PARA ACONDICIONAR"
    if c_cargar:   rename_map[c_cargar]   = "CARGAR MEDICION"

    d.rename(columns=rename_map, inplace=True)

    # tipos
    for c in ["fecha_nvl","fecha_ult_control","Fecha (más reciente)"]:
        if c in d.columns:
            d[c] = pd.to_datetime(d[c], errors="coerce", dayfirst=True)
    if "DIAS SIN MEDICION" in d.columns:
        d["DIAS SIN MEDICION"] = pd.to_numeric(d["DIAS SIN MEDICION"], errors="coerce")

    return d

# ==================== PRESELECCIÓN (PARA CRONOGRAMA) ====================
def _preseleccionar_para_cronograma(pot_norm: pd.DataFrame) -> pd.DataFrame:
    d = pot_norm.copy()

    for col in ["met_prod","nivel_3","ESTADO ACTUAL","CATEGORIAS"]:
        if col in d.columns:
            d[col] = d[col].astype(str)

    # filtros base
    met_ok = d["met_prod"].apply(lambda v: _norm_text(v) in {"bombeo mecanico","cavidad progresiva","electro sumergible"})
    lvl3_ok = d["nivel_3"].apply(lambda v: _norm_text(v) == "los perales")
    est_ok = d["ESTADO ACTUAL"].apply(lambda v: _norm_text(v) == "en marcha")
    dias_ok = d["DIAS SIN MEDICION"] > 30

    base = d[met_ok & lvl3_ok & est_ok & dias_ok].copy()

    # asegurar columnas que usamos luego
    for c in ["nombre_corto_pozo","ACCIONES PARA ACONDICIONAR","CARGAR MEDICION",
              "Observaciones_rel","SOLICITUD (Relev)","OPERACIÓN REALIZADA (Relev)","ACONDICIONAR (Relev)"]:
        if c not in base.columns:
            base[c] = pd.NA

    return base

# ==================== COORDENADAS ====================
def _leer_coordenadas(path: Path) -> pd.DataFrame:
    dfc = pd.read_excel(path, engine="openpyxl")

    c_pozo = _find_col(dfc, ["POZO"])
    c_lat  = _find_col(dfc, ["GEO_LATITUDE","latitude","lat","LATITUD"])
    c_lon  = _find_col(dfc, ["GEO_LONGITUDE","longitude","lon","long","LONGITUD"])
    if any(x is None for x in [c_pozo,c_lat,c_lon]):
        raise RuntimeError("El archivo coordenadas no tiene columnas detectables: POZO / GEO_LATITUDE / GEO_LONGITUDE.")

    dfc = dfc[[c_pozo, c_lat, c_lon]].rename(
        columns={c_pozo:"POZO_RAW", c_lat:"GEO_LATITUDE", c_lon:"GEO_LONGITUDE"}
    )

    # Clave simple compatible con 'nombre_corto_pozo'
    dfc["POZO_KEY"] = dfc["POZO_RAW"].astype(str).str.upper().str.strip()

    # ⇩⇩ convertir coma decimal a punto antes de to_numeric
    def _to_float_col(s: pd.Series) -> pd.Series:
        return pd.to_numeric(
            s.astype(str).str.replace("\u00a0"," ").str.strip().str.replace(",", ".", regex=False),
            errors="coerce"
        )

    dfc["GEO_LATITUDE"]  = _to_float_col(dfc["GEO_LATITUDE"])
    dfc["GEO_LONGITUDE"] = _to_float_col(dfc["GEO_LONGITUDE"])

    dfc = dfc.dropna(subset=["GEO_LATITUDE","GEO_LONGITUDE"])
    dfc = dfc.groupby("POZO_KEY", as_index=False).last()  # por si hay duplicados
    return dfc[["POZO_KEY","GEO_LATITUDE","GEO_LONGITUDE"]]


def _merge_coords(base: pd.DataFrame, coords: pd.DataFrame) -> pd.DataFrame:
    b = base.copy()
    if "nombre_corto_pozo" not in b.columns:
        print("⚠️ _merge_coords: falta 'nombre_corto_pozo' en base.")
        return b

    b["POZO_KEY"] = b["nombre_corto_pozo"].astype(str).str.upper().str.strip()
    b = b.merge(coords, on="POZO_KEY", how="left")
    b.drop(columns=["POZO_KEY"], inplace=True)

    # Diagnóstico
    ok = b["GEO_LATITUDE"].notna() & b["GEO_LONGITUDE"].notna()
    print(f"🔎 _merge_coords: coords presentes en {int(ok.sum())}/{len(b)} pozos "
          f"({100.0*float(ok.sum())/max(1,len(b)):.1f}% con geo).")
    return b


# ==================== PRIORIDAD & MOTIVO ====================
def _orden_prioridad(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # Para A,B,C,D,E,F,F1,G se respeta "CARGAR MEDICION" vacío
    if "CARGAR MEDICION" in df.columns:
        carga_vacia = df["CARGAR MEDICION"].apply(_is_empty)
    else:
        carga_vacia = pd.Series([True]*len(df), index=df.index)

    def _is_cat(series, etiqueta: str):
        t = series.astype(str).map(_norm_text)
        tgt = _norm_text(etiqueta)
        if tgt in {"acondicionar- ver", "acondicionar-ver"}:
            # aceptar variantes con/ sin guion / espacio
            return t.apply(lambda s: ("acondicionar" in s) and ("ver" in s))
        if tgt == "acondicionar":
            # aceptar "acondicionar" pero no la variante "- ver"
            return t.apply(lambda s: ("acondicionar" in s) and ("ver" not in s))
        # match exacto para categorías simples
        return t == tgt

    parts = []

    # A-B-C-D (con CARGAR MEDICION vacío)
    for lab in ["A","B","C","D"]:
        m = _is_cat(df["CATEGORIAS"], lab) & carga_vacia
        parts.append(df[m])

    # ACONDICIONAR- VER (SIN condicionar por CARGAR MEDICION y SIN exigir "ya se finalizó...")
    parts.append(df[_is_cat(df["CATEGORIAS"], "ACONDICIONAR- VER")])

    # ACONDICIONAR (solo si ACCIONES PARA ACONDICIONAR = "ya se finalizó acción para acondicionar")
    acc_col = df.get("ACCIONES PARA ACONDICIONAR", pd.Series([None]*len(df)))
    fin_ok = acc_col.astype(str).map(_norm_text).eq("ya se finalizo accion para acondicionar")
    parts.append(df[_is_cat(df["CATEGORIAS"], "ACONDICIONAR") & fin_ok])

    # E-F-F1-G (con CARGAR MEDICION vacío)
    for lab in ["E","F","F1","G"]:
        m = _is_cat(df["CATEGORIAS"], lab) & carga_vacia
        parts.append(df[m])

    # Unificamos y orden secundario por categoría y días sin medición
    pri = pd.concat(parts, axis=0).drop_duplicates(subset=["nombre_corto"])
    cat_order = {lab:i for i,lab in enumerate(CATS_PRIORIDAD)}
    pri = pri.sort_values(
        ["CATEGORIAS","DIAS SIN MEDICION"],
        ascending=[True, False],
        key=lambda s: s if s.name!="CATEGORIAS" else s.map(cat_order)
    )
    return pri


def _motivo_base_por_categoria(cat: str) -> str:
    c = (cat or "").strip().upper()
    if c == "A":
        return "AIB en estado Crítico, no hay medición hace 30 días, no hay inconvenientes en superficie para medirlo"
    if c == "B":
        return "AIB en estado Alerta, no hay medición hace 30 días, no hay inconvenientes en superficie para medirlo"
    if c == "C":
        return "POZO NUEVO/REPARADO, no hay medición hace 30 días, no hay inconvenientes en superficie para medirlo"
    if c == "D":
        return "POZO con intervención FRECUENTE, no hay medición hace 30 días, no hay inconvenientes en superficie para medirlo"
    if c == "E":
        return "Pozo NUEVO/REP/TER, no hay medición hace 30 días, no hay inconvenientes en superficie para medirlo"
    return ""

# ==================== CLUSTER CERCANOS ====================
def _cluster_cercanos(df_ord: pd.DataFrame, assigned: np.ndarray, capacidad: int, seed_cursor: int) -> List[int]:
    n = len(df_ord)
    if n == 0:
        return []
    idxs = list(range(n))

    # elegir seed
    seed = None
    for k in range(n):
        i = (seed_cursor + k) % n
        if not assigned[i]:
            seed = i
            break
    if seed is None:
        return []

    elegidos = [seed]
    assigned[seed] = True

    seed_lat = df_ord.iloc[seed].get("GEO_LATITUDE")
    seed_lon = df_ord.iloc[seed].get("GEO_LONGITUDE")

    # rellenar por cercanía si tenemos coords
    if pd.notna(seed_lat) and pd.notna(seed_lon):
        dists = []
        for j in idxs:
            if assigned[j] or j == seed: 
                continue
            lat = df_ord.iloc[j].get("GEO_LATITUDE")
            lon = df_ord.iloc[j].get("GEO_LONGITUDE")
            if pd.notna(lat) and pd.notna(lon):
                d = haversine_km(float(seed_lat), float(seed_lon), float(lat), float(lon))
                dists.append((d, j))
        dists.sort(key=lambda x: x[0])
        for _, j in dists:
            if len(elegidos) >= capacidad:
                break
            elegidos.append(j)
            assigned[j] = True

    # si faltan, tomar siguientes por orden
    if len(elegidos) < capacidad:
        for j in idxs:
            if len(elegidos) >= capacidad:
                break
            if not assigned[j]:
                elegidos.append(j)
                assigned[j] = True

    return elegidos

# ==================== EXTRA SHEETS ====================
def _select_cols(df: pd.DataFrame, cols: List[str]) -> pd.DataFrame:
    out = df.copy()
    for c in cols:
        if c not in out.columns:
            out[c] = pd.NA
    return out[cols]

def _write_extra_sheets(xw, pot_full: pd.DataFrame):
    # Falta acondicionar
    def _is_acond_cat(v):
        t = _norm_text(v)
        return t in {"acondicionar","acondicionar- ver","acondicionar-ver"}
    falta_acond = pot_full[pot_full["CATEGORIAS"].apply(_is_acond_cat)] if "CATEGORIAS" in pot_full.columns else pot_full.iloc[0:0]
    cols_falta = ["nombre_corto","nombre_corto_pozo","met_prod","nivel_3","nivel_5","fecha_nvl",
                  "Fecha (más reciente)","Observaciones_rel","SOLICITUD (Relev)",
                  "OPERACIÓN REALIZADA (Relev)","ACONDICIONAR (Relev)"]
    _select_cols(falta_acond, cols_falta).to_excel(xw, sheet_name="Falta acondicionar", index=False)

    # Consolidado (las 33+ columnas)
    cols_conso = [
        "nombre_corto","nombre_corto_pozo","nombre_pozo","estado","met_prod","nivel_3","nivel_5",
        "fecha_nvl","fecha_ult_control","tiponivel","prod_oil","prod_gas","total_vol","sumergencia",
        "regimen_prim","prof_ent_bomba","regimen_sec","Fecha (más reciente)","Observaciones_rel",
        "SOLICITUD (Relev)","OPERACIÓN REALIZADA (Relev)","ACONDICIONAR (Relev)",
        "TIPO DE NIVEL QUE FALTA CARGAR","DIAS SIN MEDICION","CRITICIDAD","ACCIONES EN PROCESO",
        "ACCIONES FINALIZADAS","VER","POST PULLING","POZO NUEVO/REPARADO","Pozo frecuente",
        "ESTADO ACTUAL","CATEGORIAS","ACCIONES PARA ACONDICIONAR","CARGAR MEDICION"
    ]
    _select_cols(pot_full, cols_conso).to_excel(xw, sheet_name="Consolidado", index=False)

    # Mediciones a cargar
    if "CARGAR MEDICION" in pot_full.columns:
        m_cargar = pot_full["CARGAR MEDICION"].astype(str).str.upper().eq("CARGAR MEDICION EN SISTEMA")
        med_cargar = pot_full[m_cargar].copy()
    else:
        med_cargar = pot_full.iloc[0:0]
    cols_med = [
        "nombre_corto","nombre_corto_pozo","nombre_pozo","met_prod","nivel_3","nivel_5","fecha_nvl",
        "Fecha (más reciente)","SOLICITUD (Relev)","OPERACIÓN REALIZADA (Relev)","ACONDICIONAR (Relev)",
        "TIPO DE NIVEL QUE FALTA CARGAR","DIAS SIN MEDICION","CRITICIDAD","ACCIONES EN PROCESO",
        "ACCIONES FINALIZADAS","VER","POST PULLING","POZO NUEVO/REPARADO","Pozo frecuente"
    ]
    _select_cols(med_cargar, cols_med).to_excel(xw, sheet_name="Mediciones a cargar", index=False)
    
    
# ==> NUEVO: distancia al siguiente pozo del mismo día
def _dist_next_same_day(df):
    df = df.copy()
    for c in ["GEO_LATITUDE","GEO_LONGITUDE"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    # siguiente pozo dentro del mismo día
    df["lat_next"] = df.groupby("Fecha")["GEO_LATITUDE"].shift(-1)
    df["lon_next"] = df.groupby("Fecha")["GEO_LONGITUDE"].shift(-1)

    def _dist_row(r):
        try:
            if pd.notna(r["GEO_LATITUDE"]) and pd.notna(r["GEO_LONGITUDE"]) and                pd.notna(r["lat_next"]) and pd.notna(r["lon_next"]):
                return haversine_km(float(r["GEO_LATITUDE"]), float(r["GEO_LONGITUDE"]),
                                    float(r["lat_next"]), float(r["lon_next"]))
        except Exception:
            pass
        return np.nan

    df["DIST_KM_SIG_POZO"] = df.apply(_dist_row, axis=1)
    df.drop(columns=["lat_next","lon_next"], inplace=True)
    return df

def fetch_datos_instalacion_oracle() -> Optional[pd.DataFrame]:
    """
    Ejecuta la consulta de Oracle y devuelve un DataFrame ya transformado
    y con las columnas reordenadas como en Power Query.
    """
    # Intentar usar oracledb (recomendado) y si no está, cx_Oracle
    dbmod = None
    try:
        import oracledb as dbmod  # pip install oracledb
    except Exception:
        try:
            import cx_Oracle as dbmod  # pip install cx_Oracle
        except Exception:
            print("⚠️ No se encontró ni 'oracledb' ni 'cx_Oracle'. Instalá uno de los dos para consultar Oracle.")
            return None

    user = "RY33872"
    password = "Contraseña_0725"
    host = "slplpgmoora03"
    port = 1527
    service = "psfu"

    try:
        dsn = dbmod.makedsn(host=host, port=port, service_name=service)
        conn = dbmod.connect(user=user, password=password, dsn=dsn)
    except Exception as e:
        print(f"⚠️ No se pudo conectar a Oracle: {e}")
        return None

    query = """
    SELECT
      FIC_PROF_SIST_EXTRAC_INSTAL."Fecha de Instalacion",
      DBU_FIC_ORG_ESTRUCTURAL.NOMBRE_CORTO_POZO,
      VARILLAS_BM.T1_GRADO, VARILLAS_BM.T2_GRADO, VARILLAS_BM.T3_GRADO, VARILLAS_BM.T4_GRADO,
      FIC_PROF_SIST_EXTRAC_INSTAL."Profundidad de la bomba",
      VARILLAS_BM.PROF_INSTAL_ANCLA, VARILLAS_BM.PROF_BBA, VARILLAS_BM.D_PISTON,
      VARILLAS_BM.T1_CANT, VARILLAS_BM.T1_DIAM,
      VARILLAS_BM.T2_CANT, VARILLAS_BM.T2_DIAM,
      VARILLAS_BM.T3_CANT, VARILLAS_BM.T3_DIAM,
      VARILLAS_BM.T4_CANT, VARILLAS_BM.T4_DIAM
    FROM
      DISC_ADMINS.DBU_FIC_ORG_ESTRUCTURAL DBU_FIC_ORG_ESTRUCTURAL,
      DISC_ADMINS.VMOSE_FIC_MONITOREOPOZOS VMOSE_FIC_MONITOREOPOZOS,
      DISC_ADMINS.FIC_PROF_SIST_EXTRAC_INSTAL FIC_PROF_SIST_EXTRAC_INSTAL,
      DISC_ADMINS.VARILLAS_BM VARILLAS_BM
    WHERE
      ( FIC_PROF_SIST_EXTRAC_INSTAL.COMP_SK(+) = DBU_FIC_ORG_ESTRUCTURAL.COMP_SK )
      AND ( VMOSE_FIC_MONITOREOPOZOS.COMP_SK = DBU_FIC_ORG_ESTRUCTURAL.COMP_SK )
      AND ( VARILLAS_BM.I_KEY = DBU_FIC_ORG_ESTRUCTURAL.I_KEY )
      AND ( VMOSE_FIC_MONITOREOPOZOS.ESTADO = 'Produciendo' )
      AND ( DBU_FIC_ORG_ESTRUCTURAL.NIVEL_3 IN ( 'Los Perales' ) )
    ORDER BY DBU_FIC_ORG_ESTRUCTURAL.NOMBRE_CORTO_POZO ASC
    """

    try:
        df = pd.read_sql(query, conn)
    except Exception as e:
        print(f"⚠️ Error ejecutando la consulta Oracle: {e}")
        try:
            conn.close()
        except Exception:
            pass
        return None
    finally:
        try:
            conn.close()
        except Exception:
            pass

    # --- Transformaciones equivalentes al Power Query ---
    # 1) Asegurar que los nombres lleguen como esperamos (Oracle suele devolver mayúsculas)
    #    Si vienen con mayúsculas, renombramos exactamente a lo esperado.
    rename_map = {
        'NOMBRE_CORTO_POZO': 'NOMBRE_CORTO_POZO',
        'Fecha de Instalacion': 'Fecha de Instalacion',
        'Profundidad de la bomba': 'Profundidad de la bomba',
        'PROF_INSTAL_ANCLA': 'PROF_INSTAL_ANCLA',
        'PROF_BBA': 'PROF_BBA',
        'D_PISTON': 'D_PISTON',
        'T1_CANT': 'T1_CANT', 'T1_DIAM': 'T1_DIAM', 'T1_GRADO': 'T1_GRADO',
        'T2_CANT': 'T2_CANT', 'T2_DIAM': 'T2_DIAM', 'T2_GRADO': 'T2_GRADO',
        'T3_CANT': 'T3_CANT', 'T3_DIAM': 'T3_DIAM', 'T3_GRADO': 'T3_GRADO',
        'T4_CANT': 'T4_CANT', 'T4_DIAM': 'T4_DIAM', 'T4_GRADO': 'T4_GRADO',
    }
    # Alinear claves si Oracle devolvió comillas o espacios raros
    df.columns = [c.replace('"','') for c in df.columns]
    # No cambiamos nombres si ya coinciden, solo nos aseguramos de que existan
    # Convertir fecha
    if "Fecha de Instalacion" in df.columns:
        df["Fecha de Instalacion"] = pd.to_datetime(df["Fecha de Instalacion"], errors="coerce").dt.date

    # Orden final (como en tu step #"Columnas reordenadas1")
    final_cols = [
        "NOMBRE_CORTO_POZO", "Fecha de Instalacion", "Profundidad de la bomba",
        "PROF_INSTAL_ANCLA", "PROF_BBA", "D_PISTON",
        "T1_CANT", "T1_DIAM", "T1_GRADO",
        "T2_CANT", "T2_DIAM", "T2_GRADO",
        "T3_CANT", "T3_DIAM", "T3_GRADO",
        "T4_CANT", "T4_DIAM", "T4_GRADO",
    ]
    # Agregar columnas faltantes como vacías para no romper el reorden
    for c in final_cols:
        if c not in df.columns:
            df[c] = pd.NA
    df = df[final_cols]

    return df

def fetch_datos_produ_oracle() -> Optional[pd.DataFrame]:
    """
    Ejecuta la query 'Datos produ' en Oracle y devuelve un DataFrame
    con las mismas transformaciones que definiste en Power Query.
    """
    # Cargar driver (oracledb o cx_Oracle)
    try:
        import oracledb as dbmod
    except Exception:
        try:
            import cx_Oracle as dbmod
        except Exception:
            print("⚠️ No se encontró ni 'oracledb' ni 'cx_Oracle'. Instalá uno para consultar Oracle.")
            return None

    user = "RY33872"
    password = "Contraseña_0725"
    host = "slplpgmoora03"
    port = 1527
    service = "psfu"

    # DSN y conexión
    try:
        dsn = dbmod.makedsn(host=host, port=port, service_name=service)
        conn = dbmod.connect(user=user, password=password, dsn=dsn)
    except Exception as e:
        print(f"⚠️ No se pudo conectar a Oracle (Datos produ): {e}")
        return None

    query = """
    SELECT
        in442 as E_442,in445 as E_445,in447 as E_447,in464 as E_464,in484 as E_484,
        i106762 as E106762,in449 as E_449,in451 as E_451,in469 as E_469,in471 as E_471,
        in474 as E_474,in476 as E_476,in478 as E_478,in480 as E_480,i106744 as E106744,
        i106782 as E106782,in487 as E_487,i148551 as E148551,i293499 as E293499,
        i293502 as E293502,i293503 as E293503,i293505 as E293505
    FROM (
        SELECT i100019 AS in455, i100462 AS in464, i101270 AS in469, i101271 AS in471, i101272 AS in474,
               i101273 AS in476, i101285 AS in478, i101287 AS in480, i100996 AS in484, i107231 AS in487
        FROM (
            SELECT MET_PROD_DESDE AS i275948, CLAS_ABC AS i275949, NOMBRE_POZO AS i293225, NOMBRE_CORTO_POZO AS i293226,
                   CONTROLADOR_MODELO AS i390184, CALIDAD_SCADA AS i390185, TELESUP AS i390186, COMP_SK AS i100019,
                   API_CD AS i100020, I_KEY AS i100021, UWI AS i100022, NOMBRE AS i100023, NOMBRE_CORTO AS i100024,
                   COD_TIPO AS i100025, TIPO AS i100026, ESTADO AS i100027, COD_MET_PROD AS i100028, MET_PROD AS i100029,
                   BATERIA AS i100030, NIVEL_1 AS i100031, SK_NIVEL_1 AS i100032, NIVEL_2 AS i100033, SK_NIVEL_2 AS i100034,
                   NIVEL_3 AS i100035, SK_NIVEL_3 AS i100036, NIVEL_4 AS i100037, SK_NIVEL_4 AS i100038,
                   NIVEL_5 AS i100039, SK_NIVEL_5 AS i100040, COORD_X AS i100041, COORD_Y AS i100042,
                   TRAYECTORIA AS i408270, COTA AS i529913, BLOQUE_MONITOREO AS i587247
            FROM DISC_ADMINS.DBU_FIC_ORG_ESTRUCTURAL
        ) IO100018,
        (
            SELECT BAT_SK AS i100456, BAT_TYPE AS i100457, COMP_SK AS i100458, COMP_NAME AS i100459,
                   API_CD AS i100460, FIELD_SK AS i100461, BAT_NAME AS i100462, PRD_METH_DS AS i100463
            FROM DISC_ADMINS.VTOW_WELLS_BAT
        ) IO100453,
        (
            SELECT COMP_SK AS i101268, MAX_EFF_DT AS i101269, DAYS_FROM_TODAY AS i101270, PROD_OIL AS i101271,
                   PROD_GAS AS i101272, PROD_WAT AS i101273, TEST_PURP_CD AS i101274, TEST_TYPE AS i101275,
                   T_HRS AS i101276, GRAV_OIL AS i101277, GRAV_GAS AS i101278, AMBIENT_TMP AS i101279, INJ_VOL AS i101280,
                   INJ_PROD_CD AS i101281, INJ_PRESS AS i101282, EMUL_PCT AS i101283, LIFT_INPUT AS i101284,
                   FREE_WAT_PCT AS i101285, TEST_REASON_CD AS i101286, TOTAL_VOL AS i101287, LINE_PRESS AS i101288,
                   TUB_PRESS_LOW AS i101289, TUB_PRESS_HI AS i101290, CHOKE_SIZE AS i101291, USER_N5 AS i101292,
                   USER_N6 AS i101293, SEP_PRESS AS i101294, CAS_PRESS_HI AS i101295, CAS_PRESS_LOW AS i101296,
                   TRIPLEX_PSI AS i101297, WH_TUB_PRESS AS i101298, USER_N7 AS i101299, WH_TMP AS i101300,
                   CHL_PPM AS i101301, USER_N1 AS i280192
            FROM DISC_ADMINS.VTOW_WELL_LAST_CONTROL_DET
        ) IO101267,
        (
            SELECT NIV_CAL_DIN_FECHA AS i145613, NVL_DESCRIPCION AS i145615, FECHA_CARGA_NVL AS i145616,
                   PARAM_METODO_MEDICION AS i145618, DENSIDAD_GAS AS i145619, DENSIDAD_CRUDO AS i145620,
                   DENSIDAD_AGUA AS i145622, LONG_TUBING AS i145623, PROF_ENT_BOMBA AS i145624,
                   FECHA_MEDICION_PRES AS i145625, PRESION_FINAL_CASING AS i145626, COMENTARIO_PRES AS i145627,
                   FECHA_CARGA_PRES AS i145628, PRESION_ENT_BOMBA AS i145630, PEB_ID_ORIGEN_CALCULO AS i145631,
                   ID_REGIMEN AS i145600, NIV_MED_FECHA AS i145602, NIV_COR_FECHA AS i145605, NIV_CAL_SEN_FECHA AS i145608,
                   NIVEL_CALC_SENSOR AS i146172, FECHA_MEDICION_REG AS i145634, REGIMEN_PRIM AS i145635,
                   REGIMEN_SEC AS i145636, ID_SISTEMA_EXTRACCION AS i145637, COMENTARIO_REG AS i145638,
                   FECHA_CARGA_REG AS i145639, PRESION_LINEA AS i107227, PRESION_CASING AS i107228,
                   PRESION_TUBING AS i107229, VAR_PRESION AS i107230, SUMERGENCIA AS i107231,
                   NIVEL_ID AS i100995, FECHA_NVL AS i100996, EMPRESA AS i100997, NVL_SONOLOG AS i100998,
                   NVL_CORREGIDO AS i100999, COMP_SK AS i101000, TIPONIVEL AS i180842, METOMEDIOCION AS i180843,
                   CONEXIONENTRECANIO AS i180844, NIVEL AS i180845, ORIGENIVEL AS i180846, NIVEL_CALC_DINA AS i180847,
                   USUARIOCARGANVL AS i180848, ESTADONIVELMEDIDO AS i180849, NIV_MED_USUARIO AS i180850,
                   ESTADONIVELCORREGIDO AS i180851, NIVCORUSUARIO AS i180852, ESTADONIVELCALCSENSOR AS i180853,
                   NIV_CAL_SEN_USUARIO AS i180854, ESTADONIVELCALCDINA AS i180855, NIVCALDINUSUARIO AS i180856,
                   ESTADOPEB AS i180857, PEB_USUARIO AS i180858, TIEMPO_DE_CIERRE AS i180859,
                   USUARIO_CARGA_PRES AS i180860, USUARIO_CARGA_REG AS i180861
            FROM DISC_ADMINS.TEST_FIC_ULT_NIVELES
        ) IO100994
        WHERE (i100019 = i100458(+)) AND (i101268 = i100458) AND (i101000 = i100458)
    ) on453,
    (
        SELECT i100019 AS in433, i100024 AS in442, i100027 AS in445, i100029 AS in447, i275949 AS in449,
               i293226 AS in451, i100020 AS in503
        FROM (
            SELECT MET_PROD_DESDE AS i275948, CLAS_ABC AS i275949, NOMBRE_POZO AS i293225, NOMBRE_CORTO_POZO AS i293226,
                   CONTROLADOR_MODELO AS i390184, CALIDAD_SCADA AS i390185, TELESUP AS i390186, COMP_SK AS i100019,
                   API_CD AS i100020, I_KEY AS i100021, UWI AS i100022, NOMBRE AS i100023, NOMBRE_CORTO AS i100024,
                   COD_TIPO AS i100025, TIPO AS i100026, ESTADO AS i100027, COD_MET_PROD AS i100028, MET_PROD AS i100029,
                   BATERIA AS i100030, NIVEL_1 AS i100031, SK_NIVEL_1 AS i100032, NIVEL_2 AS i100033, SK_NIVEL_2 AS i100034,
                   NIVEL_3 AS i100035, SK_NIVEL_3 AS i100036, NIVEL_4 AS i100037, SK_NIVEL_4 AS i100038,
                   NIVEL_5 AS i100039, SK_NIVEL_5 AS i100040, COORD_X AS i100041, COORD_Y AS i100042,
                   TRAYECTORIA AS i408270, COTA AS i529913, BLOQUE_MONITOREO AS i587247
            FROM DISC_ADMINS.DBU_FIC_ORG_ESTRUCTURAL
        ) IO100018,
        (
            SELECT COMP_SK AS i290857, STATUS_DT_FR AS i290858, STATUS_DT_TO AS i290859, STATUS_CD AS i290860,
                   DB_OWN_SK AS i290861, USER_A1 AS i290862, USER_A2 AS i290863, REF_DS AS i290864
            FROM DISC_ADMINS.VTOW_ESTADOS_DE_POZO
        ) IO290856
        WHERE (i100019 = i290857(+)) AND (i100035 = 'Los Perales') AND (i290864 IN ('Produciendo'))
    ) on431,
    (
        SELECT SARTA_TENSION_MIN_BOMBA AS i107165, SARTA_TENSION_MAX_BOMBA AS i107166, BBA_PERDIDA_DE_VALV_FIJA AS i107167,
               BBA_PRESION_DE_FONDO AS i107168, BBA_NIVEL_CORREGIDO_MCKOY AS i106757, BBA_NIVEL_CALCULADO AS i106758,
               BBA_SUMERGENCIA AS i106759, BBA_PO_CALCULADO AS i106760, IDNIVEL AS i106761, FECHA_HORA AS i106762,
               REALIZADO_POR AS i106763, AIB_MARCA_Y_DESC_API AS i106764, AIB_CARRERA_DEL_VASTAGO AS i106765,
               AIB_SENTIDO_DE_GIRO AS i106766, AIB_RELACION_DE_TRANSMISION AS i106767, AIB_CARRERA_MEDIDA_SUP AS i106768,
               AIB_DIAMETRO_POLEA_REDUCTOR AS i106769, CP_DISTANCIA_EM AS i106770, CP_TIPO_PRINCIPAL AS i106771,
               CP_CANTIDAD_TIPO_PRINCIPAL AS i106772, CP_TIPO_SECUNDARIO AS i106773, CP_CANTIDAD_TIPO_SECUNDARIO AS i106774,
               CP_TIPO_DE_MANIVELA AS i106775, CP_CONSTANTE_DE_RESORTE AS i106776, CANT_DIAM_GRADO_API_PROF_LONG AS i106777,
               MOTOR_POTENCIA_NOMINAL AS i106778, MOTOR_DIAMETRO_POLEA AS i106779, BBA_MARCA_DESC_API AS i106780,
               BBA_PROF_DE_BBA AS i106782, BBA_DIAMETRO_TUBING AS i106783, BBA_CORTE_DE_AGUA AS i106784,
               BBA_VISCOSIDAD_PO AS i106786, BBA_DENSIDAD_PO AS i106787, BBA_DENSIDAD_PROMEDIO AS i106788,
               BBA_DENSIDAD_DE_COLUMNA AS i106789, BBA_DISP_ESPECIALES AS i106791, AIBRE_CARGA_MAXIMA_ESTRUCTURA AS i106792,
               BBA_TPO_ENTRE_TVI_TVF AS i106838, BBA_PERDIDA_DE_VALVULA_MOVIL AS i106839,
               BBA_CARRERA_CALCULADA AS i106840, BBA_TPO_ENTRE_SVI_SVF AS i106834, BBA_PESO_FLUID_VAR_INI AS i106836,
               BBA_PESO_FLUID_VAR_FIN AS i106837, COMP_SK AS i106739, AIB_GPM AS i106740, MOTOR_MARCA AS i106741,
               MOTOR_TIPO AS i106742, MOTOR_RPM AS i106743, BBA_DIAM_PISTON AS i106744, BBA_PRODUCCION_BRUTA AS i106745,
               BBA_TIPO_EXPLOTACION AS i106746, BBA_CARRERA_EFECTIVA AS i106747, BBA_LLENADO_DE_BOMBA AS i106748,
               BBA_EFICIENCIA_VOLUMETRICA AS i106749, BBA_CAUDAL_BRUTO_CALCULADO AS i106750,
               BBA_CAUDAL_BRUTO_EFECTIVO AS i106751, IDDINAMOMETRO AS i106738, BBA_PRESION_ENTRADA_A_BOMBA AS i106753,
               BBA_PO_EFECTIVO AS i106754, BBA_ESCURRIMIENTO AS i106755, BBA_NIVEL_MEDIDO AS i106756,
               AIBRE_CARGA_MINIMA_ESTRUCTURA AS i106793, AIBRE_PESO_FLUIDO AS i106794, AIBRE_SOLICITACION_DE_ESTRUCT AS i106795,
               AIBEB_TORQUE_MAXIMO_REDUCTOR AS i106796, AIBEB_PORCENTAJE AS i106797, AIBEBDISTANCIA_CONTRAPESO AS i106798,
               AIBEB_EFECTO_DE_CONTRAPESO AS i106799, AIBEBCONTRAPESO_IDEAL AS i106800,
               AIBEB_TORQUE_MAXIMO_CONTRAPESO AS i106801, AIBEB_EFICIENCIA_TORSIONAL AS i106802,
               AIBRR_TORQUE_MAXIMO_REDUCTOR AS i106803, AIBRR_DISTANCIA_CONTRAPESO AS i106804,
               AIBRR_EFICIENCIA_TORSIONAL AS i106805, AIBRR_TORQUE_DISPONIBLE AS i106806,
               AIBRR_REGIMEN_DE_OPERACION AS i106807, AIBRR_PORCENTAJE AS i106808,
               AIBRR_EFECTO_DE_CONTRAPESO AS i106809, AIBRR_TORQUE_MAXIMO_CONTRAPESO AS i106810,
               AIBRR_BALANCEO_EXISTENTE AS i106811, MOTOR_POT_VASTAGO AS i106812, MOTOR_POT_HIDRAULICA AS i106813,
               MOTOR_POT_REGIMEN AS i106814, MOTOR_EFICIENCIA_GLOBAL AS i106815, MOTOR_POT_REQ_EXISTENTE AS i106816,
               MOTOR_PERDIDA_DE_POTENCIA AS i106817, MOTOR_FACT_FORMA_EXIST_POT_REQ AS i106818,
               MOTOR_POT_REQ_EN_BALANCE AS i106819, MOTOR_FACT_FORMA_OPT_POT_REQ AS i106820,
               SARTA_CARGA_MIN_FONDO AS i106821, SARTA_CARGA_MAX_FONDO AS i106822, SARTA_PESO_SARTA_AIRE AS i106825,
               SARTA_PESO_SARTA_SUMERGIDO AS i106826, SARTA_PUNTO_NEUTRO AS i106827, SARTA_EST_VARILLAS AS i106828,
               SARTA_SOBRERRECORRIDO AS i106829, SARTA_CARGA_MINIMA AS i106830, SARTA_CARGA_MAXIMA AS i106831,
               BBA_PESO_VARILLA_FLUID_INI AS i106832, BBA_PESO_VARILLA_FLUID_FIN AS i106833,
               AIB_API_CARGA_MAXIMA_ESTRUC AS i108372, AIB_MARCA AS i109250, AIB_API_TIPO_CONTRAPESO AS i109251,
               AIB_API_CAPACIDAD_TORQUE AS i109252, AIB_API_TIPO_REDUCCION AS i109253, AIB_API_CARRERA_MAX AS i109254,
               "BBA_REL.GAS/PO" AS i282282, "BBA_PROF._ANCLA/PACKER" AS i282283, "BBA_LUZ_PISTON/BARRIL" AS i282281
        FROM DISC_ADMINS.FIC_ULTIMO_DINAMOMETRO
    ) o106737,
    (
        SELECT COMP_SK AS i148549, "Fecha de Instalacion" AS i148550, "Profundidad de la bomba" AS i148551
        FROM DISC_ADMINS.FIC_PROF_SIST_EXTRAC_INSTAL
    ) o148547,
    (
        SELECT API_CD AS i293495, I_KEY AS i293496, COMP_SK AS i293497, COMP_NAME AS i293498, PROF_ANCLA AS i293499,
               PROF_PACKER AS i293500, PROF_ZAP AS i293501, D_PISTON AS i293502, LUZ AS i293503, DIAM_TBG AS i293504,
               DIAM_CSG AS i293505, CURR_METH AS i290182, BBA_PROF AS i403182
        FROM DISC_ADMINS.LOCAL_SREP_ULT_EQUIP
    ) o293494
    WHERE ( (i293495 = in503) AND (i148549(+) = in433) AND (i106739 = in433) AND (in433 = in455) )
    """

    try:
        df = pd.read_sql(query, conn)
    except Exception as e:
        print(f"⚠️ Error ejecutando la consulta Oracle (Datos produ): {e}")
        try: conn.close()
        except Exception: pass
        return None
    finally:
        try: conn.close()
        except Exception: pass

    # --- Transformaciones equivalentes al Power Query ---
    # 0) Limpiar posibles comillas en nombres de columnas
    df.columns = [c.replace('"','') for c in df.columns]

    # 1) Eliminar E_442 si existe
    if "E_442" in df.columns:
        df = df.drop(columns=["E_442"])

    # 2) Renombrar inicial
    df = df.rename(columns={
        "E_445": "Estado",
        "E_447": "SEA",
        "E_464": "Batería",
    })

    # 3) Reordenar (primer orden)
    cols_step1 = ["E_451", "Estado", "SEA", "Batería", "E_484", "E106762", "E_449",
                  "E_469", "E_471", "E_474", "E_476", "E_478", "E_480",
                  "E106744", "E106782", "E_487", "E148551", "E293499", "E293502", "E293503", "E293505"]
    for c in cols_step1:
        if c not in df.columns: df[c] = pd.NA
    df = df[cols_step1]

    # 4) Renombrar (segundo paso)
    df = df.rename(columns={
        "E_451": "Pozo",
        "E_471": "Oil",
        "E_474": "Gas",
        "E_476": "Agua",
        "E_478": "WCut",
        "E_480": "Bruta",
    })

    # 5) Reordenar (segundo orden)
    cols_step2 = ["Pozo", "Estado", "SEA", "Batería", "E_484", "E106762", "E_449", "E_469",
                  "Bruta", "Oil", "Gas", "Agua", "WCut",
                  "E106744", "E106782", "E_487", "E148551", "E293499", "E293502", "E293503", "E293505"]
    for c in cols_step2:
        if c not in df.columns: df[c] = pd.NA
    df = df[cols_step2]

    # 6) Renombrar final
    df = df.rename(columns={
        "E_469": "Dias sin control",
        "E_449": "Clase ABC",
        "E106762": "Ultimo Dina",
        "E_484": "Último Nivel",
        "E106744": "Pistón dina",
        "E106782": "Prof bba dina",
        "E_487": "Sumergencia",
        "E148551": "Prof SEA",
        "E293499": "Prof ancla",
        "E293502": "Diam piston",
        "E293503": "Luz bba",
        "E293505": "Casing",
    })

    return df

# ==================== MAIN ====================
def generar_cronograma_por_proximidad(
    equipos: Optional[int] = None,
    in_xlsx: Path = IN_XLSX,
    sheet_in: str = SHEET_IN,
    coords_xlsx: Path = COORDS_XLSX,
    out_xlsx: Path = OUT_XLSX,
    capacidad: int = CAPACIDAD_POR_EQUIPO_POR_DIA,
    start_date: Optional[date] = None
):
    # Equipos
    if equipos is None:
        while True:
            try:
                equipos = int(input("¿Cuántos equipos disponibles? (cada equipo hace 20 pozos/día): ").strip())
                if equipos >= 1:
                    break
                print("Debe ser un entero >= 1.")
            except Exception:
                print("Entrada inválida. Probá de nuevo.")

    # Cargar Potenciales y normalizar (completo, SIN filtro)
    pot_raw = pd.read_excel(in_xlsx, sheet_name=sheet_in, engine="openpyxl")
    pot_full = _normalizar_potenciales(pot_raw)

    # Base para cronograma (con filtros)
    base = _preseleccionar_para_cronograma(pot_full)

    # Coordenadas
    coords = _leer_coordenadas(coords_xlsx)
    base = _merge_coords(base, coords)
    
    # Diagnóstico: ¿cuántos pozos quedaron con coordenadas?
    if "GEO_LATITUDE" in base.columns and "GEO_LONGITUDE" in base.columns:
        n_total   = len(base)
        n_con_geo = base["GEO_LATITUDE"].notna() & base["GEO_LONGITUDE"].notna()
        print(f"Coordenadas presentes en {int(n_con_geo.sum())}/{n_total} pozos "
              f"({100.0*float(n_con_geo.sum())/max(1,n_total):.1f}% con geo).")
    else:
        print("⚠️ No llegaron columnas GEO_LATITUDE/GEO_LONGITUDE a 'base'. Revisa _leer_coordenadas/_merge_coords.")
    
    # Prioridad
    df_ord = _orden_prioridad(base).reset_index(drop=True)

    # Fechas calendario
    hoy = datetime.now().date()
    d = _next_monday_from(hoy) if start_date is None else _weekday_only(start_date)

    assigned = np.zeros(len(df_ord), dtype=bool)
    sched_por_equipo = [[] for _ in range(equipos)]
    seed_cursor = 0

    # Asignación por día & equipo
    while not assigned.all() and len(df_ord) > 0:
        d = _weekday_only(d)
        for team_idx in range(equipos):
            if assigned.all():
                break
            idxs = _cluster_cercanos(df_ord, assigned, capacidad, seed_cursor)
            if idxs:
                for i in idxs:
                    row = df_ord.iloc[i]
                    cat = str(row.get("CATEGORIAS", "")).strip()

                    # MOTIVO base por categoría
                    motivo = _motivo_base_por_categoria(cat)

                    # A) Si CATEGORIAS = ACONDICIONAR- VER => concatenar campos
                    norm_cat = _norm_text(cat)
                    if ("acondicionar" in norm_cat) and ("ver" in norm_cat):
                        # concateno Fecha (más reciente), Observaciones_rel, ACCIONES PARA ACONDICIONAR
                        f_rec = row.get("Fecha (más reciente)")
                        if pd.notna(f_rec):
                            f_txt = pd.to_datetime(f_rec).date().isoformat()
                        else:
                            f_txt = ""
                        obs   = str(row.get("Observaciones_rel", "") or "")
                        accp  = str(row.get("ACCIONES PARA ACONDICIONAR", "") or "")
                        partes = [p for p in [f_txt, obs, accp] if str(p).strip() != ""]
                        motivo = " | ".join(partes)

                    # armar registro de salida
                    sched_por_equipo[team_idx].append({
                        "Fecha": pd.to_datetime(d),
                        "nombre_corto": row.get("nombre_corto", pd.NA),
                        "nombre_corto_pozo": row.get("nombre_corto_pozo", pd.NA),
                        "met_prod": row.get("met_prod", pd.NA),
                        "nivel_5": row.get("nivel_5", pd.NA),
                        "fecha_nvl": row.get("fecha_nvl", pd.NA),
                        "Fecha (más reciente)": row.get("Fecha (más reciente)", pd.NA),
                        "DIAS SIN MEDICION": row.get("DIAS SIN MEDICION", pd.NA),
                        "CATEGORIAS": cat,
                        "MOTIVO DE LA MEDICION": motivo,
                        "GEO_LATITUDE": row.get("GEO_LATITUDE", pd.NA),
                        "GEO_LONGITUDE": row.get("GEO_LONGITUDE", pd.NA),
                    })
            seed_cursor = (seed_cursor + 1) % max(1, len(df_ord))
        d += timedelta(days=1)

    # Exportar
     # ======= NUEVO: traer Oracle una sola vez (si hay driver/credenciales) =======
    df_inst = fetch_datos_instalacion_oracle()  # Datos Instalacion
    df_produ = fetch_datos_produ_oracle()       # Datos produ

    # Nos quedamos SOLO con las columnas necesarias de cada fuente para no inflar merges
    # Nos quedamos SOLO con las columnas necesarias de cada fuente para no inflar merges
    if df_inst is not None and not df_inst.empty:
        cols_inst_nec = [
            "NOMBRE_CORTO_POZO",
            "T1_CANT","T1_DIAM","T1_GRADO",
            "T2_CANT","T2_DIAM","T2_GRADO",
            "T3_CANT","T3_DIAM","T3_GRADO",
            "T4_CANT","T4_DIAM",
            "PROF_BBA","D_PISTON","PROF_INSTAL_ANCLA"
        ]
        for c in cols_inst_nec:
            if c not in df_inst.columns:
                df_inst[c] = pd.NA
        df_inst = df_inst[cols_inst_nec].copy()

    if df_produ is not None and not df_produ.empty:
        # Si existe "Ultimo Dina", nos quedamos con el registro más reciente por Pozo
        if "Ultimo Dina" in df_produ.columns:
            df_produ["Ultimo Dina"] = pd.to_datetime(df_produ["Ultimo Dina"], errors="coerce")
            df_produ = df_produ.sort_values(["Pozo", "Ultimo Dina"], ascending=[True, True])
            df_produ = df_produ.drop_duplicates(subset=["Pozo"], keep="last")
        else:
            df_produ = df_produ.drop_duplicates(subset=["Pozo"], keep="last")

        cols_produ_nec = ["Pozo","Bruta","Oil","WCut","Gas","Casing"]
        for c in cols_produ_nec:
            if c not in df_produ.columns:
                df_produ[c] = pd.NA
        df_produ = df_produ[cols_produ_nec].copy()


    # ======= Exportar =======
    with pd.ExcelWriter(out_xlsx, engine="openpyxl", mode="w") as xw:
        for idx, rows in enumerate(sched_por_equipo, start=1):
            df_out = pd.DataFrame(rows)

            if not df_out.empty:
                # Orden lógico por día/categoría
                df_out = df_out.sort_values(["Fecha","CATEGORIAS","DIAS SIN MEDICION"],
                                            ascending=[True, True, False])

                # Distancia al siguiente pozo del mismo día
                df_out = _dist_next_same_day(df_out)

                # ================== NUEVO: Tip Medición ==================
                # usamos _norm_text para comparar sin acentos/mayúsculas
                _met_norm = df_out["met_prod"].astype(str).map(_norm_text)
                df_out["Tip Medición"] = np.where(
                    _met_norm.eq("bombeo mecanico"),
                    "COMBINADA",
                    "ECO"
                )

                # ================== NUEVO: Merge con Datos Instalacion ==================
                if df_inst is not None and not df_inst.empty:
                    df_out = df_out.merge(
                        df_inst,
                        left_on="nombre_corto_pozo",
                        right_on="NOMBRE_CORTO_POZO",
                        how="left"
                    )
                    # ya no necesitamos la clave derecha duplicada
                    if "NOMBRE_CORTO_POZO" in df_out.columns:
                        df_out.drop(columns=["NOMBRE_CORTO_POZO"], inplace=True)

                # ================== NUEVO: Merge con Datos produ ==================
                # ================== NUEVO: Merge con Datos produ ==================
                if df_produ is not None and not df_produ.empty:
                    df_out = df_out.merge(
                        df_produ,
                        left_on="nombre_corto_pozo",
                        right_on="Pozo",
                        how="left"
                    )
                    if "Pozo" in df_out.columns:
                        df_out.drop(columns=["Pozo"], inplace=True)


                # ================== NUEVO: Observaciones vacía ==================
                if "Observaciones" not in df_out.columns:
                    df_out["Observaciones"] = ""
                
                # >>> NUEVO: dedup por Fecha + nombre_corto_pozo
                df_out = df_out.drop_duplicates(subset=["Fecha", "nombre_corto_pozo"], keep="first")
                
                # ================== Selección y orden final ==================
                cols_out = [
                    "Fecha","nombre_corto","nombre_corto_pozo","met_prod","nivel_5",
                    "fecha_nvl","Fecha (más reciente)","DIAS SIN MEDICION",
                    "CATEGORIAS","MOTIVO DE LA MEDICION","DIST_KM_SIG_POZO",
                    "T1_CANT","T1_DIAM","T2_CANT","T2_DIAM","T2_GRADO",
                    "T3_CANT","T3_DIAM","T3_GRADO","T4_CANT","T4_DIAM","T1_GRADO",
                    "PROF_BBA","D_PISTON","PROF_INSTAL_ANCLA",
                    "Tip Medición","Bruta","Oil","WCut","Gas","Casing","Observaciones"
                ]
                # Garantizamos todas las columnas
                for c in cols_out:
                    if c not in df_out.columns:
                        df_out[c] = pd.NA

                df_out = df_out[cols_out]

            # Escribimos la hoja del equipo
            df_out.to_excel(xw, sheet_name=f"Cronograma Equipo {idx}", index=False)

        # ============ Pestañas extra e info Oracle (opcionales) ============
        _write_extra_sheets(xw, pot_full)

        # Nota: si además querés **dejar** las hojas con datos crudos de Oracle:
        if df_inst is not None and not df_inst.empty:
            df_inst.to_excel(xw, sheet_name="Datos Instalacion", index=False)
            print(f"📝 Escribí 'Datos Instalacion' (unique) con {len(df_inst)} filas.")
        else:
            print("ℹ️ No se escribió 'Datos Instalacion' (consulta vacía o error de conexión).")

        if df_produ is not None and not df_produ.empty:
            df_produ.to_excel(xw, sheet_name="Datos produ", index=False)
            print(f"📝 Escribí 'Datos produ' (unique) con {len(df_produ)} filas.")
        else:
            print("ℹ️ No se escribió 'Datos produ' (consulta vacía o error de conexión).")

    print(f"✅ Cronograma por proximidad creado: {out_xlsx}")
    return out_xlsx


# Ejecutar directamente
if __name__ == "__main__":
    generar_cronograma_por_proximidad()


# In[ ]:




