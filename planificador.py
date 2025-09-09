#!/usr/bin/env python
# coding: utf-8

# In[ ]:


# -*- coding: utf-8 -*-
"""
Planificador dinámico de recorridos de pozos (recupero de petróleo)
- Normalización avanzada de POZO
- Inyección de metadatos desde Nombres-Pozo
- Cálculo de frecuencias y r_m3_d (K=1 tasa más reciente por pedido)
- Priorización por __v = r_m3_d (SIN multiplicar días)
- Equipos fijos por ZONA (un equipo no cambia de ZONA en todo el cronograma)
- Asignación diaria por clústers dentro de un radio (coordenadas) max 4 pozos/día/equipo
- Filtro: sólo r_m3_d > RM3D_MIN (default 0.1)
- Distancias por fila: a semilla y al centroide del clúster
- Sub-filtro por BATERÍA sólo si se elige "Las Heras CG - Canadon Escondida"
- Excel de salida paralelo: "<original>_CRONOGRAMA_YYYYMMDD.xlsx"
- Hojas extra: Frecuencias, Plan_Equipo_i, Parametros_Usados, Exclusiones_Usadas (si aplica),
  Zonas_Seleccionadas, Baterias_Filtradas, Normalizacion_Pozos, Alertas_Normalizacion,
  Alertas de ABM, Alertas_Coordenadas
"""

import os, re, unicodedata, math
import numpy as np
import pandas as pd
from datetime import date, timedelta, datetime

# ==========================
# CONFIG
# ==========================
INPUT_FILE  = r"DIAGRAMA SW.xlsx"   # Excel base (NO se modifica)
SHEET_HIST  = None                  # None => autodetecta hoja/encabezados
NOMBRES_POZO_FILE = r"C:\Users\ry16123\export_org_estructural\Nombres-Pozo.xlsx"

# Coordenadas (POZO, GEO_LATITUDE, GEO_LONGITUDE)
COORDS_FILE = r"C:\Users\ry16123\OneDrive - YPF\Escritorio\power BI\GUADAL- POWER BI\Inteligencia Artificial\coordenadas1.xlsx"

# Radio en km para agrupar por cercanía (cambiar acá si querés +/- 2 km)
RADIUS_KM = 3.0

# Filtro mínimo de potencial
RM3D_MIN = 0.1

# Umbrales para fuzzy
FUZZY_REPLACE_THRESHOLD = 85
FUZZY_SUGGEST_THRESHOLD = 75
LETTERS_SIMILARITY_MIN  = 80

DEFAULTS = {
    "equipos_activos": 4,                 # 1..4
    "dias_por_semana": 5,                 # 5 o 6
    "semanas_plan": 52,                   # horizonte anual
    "k_visitas": 1,                       # tasas (K=1 por pedido)
    "max_pozos_dia_equipo": 4,            # cupo por día por equipo
    "m3_por_visita_objetivo": 2.0,        # informativo
    "min_dias_freq": 7,                   # 1 semana
    "max_dias_freq": 56,                  # 8 semanas
    "dias_asumidos_una_visita": 7,        # para r si hay 1 sola visita
    "freq_dias_ultimo_cero_valido": 30,
}

# ==========================
# Utils
# ==========================
def _norm(s: str) -> str:
    s = "" if s is None or (isinstance(s, float) and np.isnan(s)) else str(s)
    s = s.replace("³", "3")
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = s.lower().strip().replace("\xa0"," ")
    s = s.replace("_"," ").replace("-"," ").replace("."," ").replace("\n"," ")
    return " ".join(s.split())

def _pozo_key(s: str) -> str:
    s = "" if s is None or (isinstance(s, float) and np.isnan(s)) else str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    return "".join(ch for ch in s if ch.isalnum()).upper()

def _canonical_digits(d: str) -> str:
    d = (d or "").lstrip("0")
    return d if d != "" else "0"

def _letters_digits_from_key_both(k: str):
    raw_digits = "".join(re.findall(r"\d+", k))
    digits_canon = _canonical_digits(raw_digits)
    letters = re.sub(r"\d+", "", k)
    return letters, digits_canon, len(raw_digits)

def _ratio_score(a: str, b: str) -> int:
    try:
        from rapidfuzz import fuzz
        return int(fuzz.ratio(a, b))
    except Exception:
        import difflib
        return int(round(difflib.SequenceMatcher(None, a, b).ratio()*100))

def _fuzzy_score(a: str, b: str) -> int:
    try:
        from rapidfuzz import fuzz
        return int(fuzz.partial_ratio(a, b))
    except Exception:
        import difflib
        return int(round(difflib.SequenceMatcher(None, a, b).ratio()*100))

def _canon_prefix_pozo(s: str) -> str:
    """
    Canoniza prefijos del nombre ORIGINAL:
      - CÑE... -> CNE...
      - CNE... -> intacto
      - CN...  -> CNE...
      - CE + dígitos -> CNE + dígitos (p.ej. CE839 -> CNE839)
    """
    if s is None or (isinstance(s, float) and np.isnan(s)):
        return s
    raw = str(s).strip()
    raw_up = raw.upper()
    if raw_up.startswith("CÑE"):
        return "CNE" + raw_up[3:]
    raw_ascii = unicodedata.normalize("NFKD", raw_up).encode("ascii", "ignore").decode("ascii")
    if raw_ascii.startswith("CNE"):
        return raw_ascii
    if raw_ascii.startswith("CN"):
        return "CNE" + raw_ascii[2:]
    m = re.match(r"^CE(\d+)$", raw_ascii)
    if m:
        return "CNE" + m.group(1)
    return raw_ascii

def next_monday(d=None):
    d = d or date.today()
    return d + timedelta(days=(7 - d.weekday()) % 7)  # 0=Lunes

def unique_output_path(base_input_path: str) -> str:
    folder = os.path.dirname(os.path.abspath(base_input_path))
    stem   = os.path.splitext(os.path.basename(base_input_path))[0]
    today  = datetime.now().strftime("%Y%m%d")
    base   = os.path.join(folder, f"{stem}_CRONOGRAMA_{today}.xlsx")
    if not os.path.exists(base): return base
    i = 2
    while True:
        cand = os.path.join(folder, f"{stem}_CRONOGRAMA_{today}_({i}).xlsx")
        if not os.path.exists(cand): return cand
        i += 1

EXPECTED_KEYS = {
    "fecha":       ["fecha"],
    "pozo":        ["pozo"],
    "zona":        ["zona"],
    "bateria":     ["bateria", "batería"],
    "m3":          ["m3 bruta","m3","m3_bruta","m3bruta","m 3 bruta","m 3","m3 bruto","m3 recuperado","m3 recupero"],
    "carreras":    ["n de carreras","n° de carreras","nº de carreras","no de carreras","nro de carreras","numero de carreras","n° carreras","n de carrera","n carreras"],
    "nivel_final": ["nivel final pozo","nivel final","nivel final del pozo"]
}

def _find_header_row(df_raw):
    for i in range(min(200, len(df_raw))):
        row_norm = [_norm(x) for x in df_raw.iloc[i,:].tolist()]
        if not row_norm:
            continue
        colmap = {v:j for v,j in zip(row_norm, range(len(row_norm)))}
        def has_any(keys): return any(k in colmap for k in keys)
        if has_any(EXPECTED_KEYS["fecha"]) and has_any(EXPECTED_KEYS["pozo"]) and has_any(EXPECTED_KEYS["zona"]) and has_any(EXPECTED_KEYS["bateria"]):
            return i, row_norm
    return None, None

# ---------- Nombres pozo: carga diccionario con metadatos ----------
def load_pozo_dictionary(xlsx_path: str):
    """
    Devuelve (mapping_key2oficial, dict_df) usando 'nombre_corto_pozo' y metadatos.
    dict_df incluye: ['oficial','key','letters','digits_canon','digits_len','met_prod','nivel_3','nivel_5','estado']
    """
    try:
        ref = pd.read_excel(xlsx_path)
    except Exception as e:
        print(f"\n[AVISO] No pude leer diccionario de pozos: {xlsx_path}\n{e}\n")
        return {}, pd.DataFrame(columns=["oficial","key","letters","digits_canon","digits_len","met_prod","nivel_3","nivel_5","estado"])

    cols = {c.lower().strip(): c for c in ref.columns}
    if "nombre_corto_pozo" not in cols:
        print(f"\n[AVISO] El diccionario no tiene la columna 'nombre_corto_pozo'. Columnas: {list(ref.columns)}\n")
        return {}, pd.DataFrame(columns=["oficial","key","letters","digits_canon","digits_len","met_prod","nivel_3","nivel_5","estado"])

    c_pozo = cols["nombre_corto_pozo"]
    c_met  = cols.get("met_prod")
    c_n3   = cols.get("nivel_3")
    c_n5   = cols.get("nivel_5")
    c_est  = cols.get("estado")

    refv = ref.loc[ref[c_pozo].notna()].copy()
    refv[c_pozo] = refv[c_pozo].astype(str).str.strip()

    of_list  = refv[c_pozo].tolist()
    met_vals = refv[c_met].astype(str).str.strip() if c_met else np.nan
    n3_vals  = refv[c_n3].astype(str).str.strip()  if c_n3 else np.nan
    n5_vals  = refv[c_n5].astype(str).str.strip()  if c_n5 else np.nan
    est_vals = refv[c_est].astype(str).str.strip() if c_est else np.nan

    keys, letters_, digits_canon_, digits_len_ = [], [], [], []
    for val in of_list:
        k = _pozo_key(val)
        L, Dcanon, Dlen = _letters_digits_from_key_both(k)
        keys.append(k); letters_.append(L); digits_canon_.append(Dcanon); digits_len_.append(Dlen)

    dict_df = pd.DataFrame({
        "oficial": of_list,
        "key": keys,
        "letters": letters_,
        "digits_canon": digits_canon_,
        "digits_len": digits_len_,
        "met_prod": list(met_vals) if isinstance(met_vals, pd.Series) else [np.nan]*len(of_list),
        "nivel_3":  list(n3_vals)  if isinstance(n3_vals,  pd.Series) else [np.nan]*len(of_list),
        "nivel_5":  list(n5_vals)  if isinstance(n5_vals,  pd.Series) else [np.nan]*len(of_list),
        "estado":   list(est_vals) if isinstance(est_vals, pd.Series) else [np.nan]*len(of_list),
    })

    key2off = {}
    for k, off in zip(dict_df["key"], dict_df["oficial"]):
        if k and k not in key2off:
            key2off[k] = off
    return key2off, dict_df

def apply_pozo_normalization(df: pd.DataFrame, key2off: dict, dict_df: pd.DataFrame):
    """
    Normaliza POZO, hace match exacto/fuzzy y, si matchea:
      - POZO := oficial
      - ZONA := nivel_3 SOLO si hubo match; si NO hubo match => ZONA = ""
      - BATERIA := nivel_5 (si existe) o se deja la original
    """
    df = df.copy()
    df["POZO_ORIG"] = df["POZO"].astype(str).str.strip()
    df["POZO_PreCanon"] = df["POZO_ORIG"].apply(_canon_prefix_pozo)
    df["__POZO_KEY"] = df["POZO_PreCanon"].apply(_pozo_key)

    parts = df["__POZO_KEY"].apply(_letters_digits_from_key_both)
    df["__KEY_LET"], df["__KEY_DIG_CANON"], df["__KEY_DIG_LEN"] = zip(*parts)

    df["POZO_MATCH"]   = None
    df["MATCH_TIPO"]   = "NO"
    df["MATCH_SCORE"]  = np.nan
    df["LETTER_SCORE"] = np.nan
    df["APLICADO"]     = "NO"
    df["ALERTA_NORM"]  = ""
    df["VALIDO_POZO"]  = True

    invalid_mask = (df["__KEY_LET"].str.len()==0) | (df["__KEY_DIG_LEN"]==0)
    if invalid_mask.any():
        df.loc[invalid_mask, "ALERTA_NORM"] = "SIN_LETRAS_O_DIGITOS"
        df.loc[invalid_mask, "VALIDO_POZO"] = False

    valid_mask = ~invalid_mask
    exact_mask = valid_mask & df["__POZO_KEY"].isin(key2off.keys())
    df.loc[exact_mask, "POZO_MATCH"]   = df.loc[exact_mask, "__POZO_KEY"].map(key2off)
    df.loc[exact_mask, "MATCH_TIPO"]   = "EXACTO"
    df.loc[exact_mask, "MATCH_SCORE"]  = 100
    df.loc[exact_mask, "LETTER_SCORE"] = 100
    df.loc[exact_mask, "APLICADO"]     = "SI"

    pending = df[valid_mask & (~exact_mask)].index.tolist()
    if pending and not dict_df.empty:
        dict_by_spec = {}
        for spec, sub in dict_df.groupby(["digits_canon","digits_len"]):
            dict_by_spec[spec] = sub

        for idx in pending:
            key_u   = df.at[idx, "__POZO_KEY"]
            let_u   = df.at[idx, "__KEY_LET"]
            digc_u  = df.at[idx, "__KEY_DIG_CANON"]
            digl_u  = int(df.at[idx, "__KEY_DIG_LEN"])

            cand_df = dict_by_spec.get((digc_u, digl_u), pd.DataFrame())
            best_off, best_score, best_lscore = None, -1, -1

            if cand_df is not None and not cand_df.empty:
                for row in cand_df.itertuples():
                    kk = row.key
                    ll = row.letters
                    sc_key = _fuzzy_score(key_u, kk)
                    sc_let = _ratio_score(let_u, ll)
                    if sc_let < LETTERS_SIMILARITY_MIN:
                        continue
                    if sc_key > best_score or (sc_key == best_score and sc_let > best_lscore):
                        best_score = sc_key
                        best_lscore = sc_let
                        best_off   = row.oficial

            if best_off is not None:
                df.at[idx, "POZO_MATCH"]   = best_off
                df.at[idx, "MATCH_TIPO"]   = "SUGERIDO"
                df.at[idx, "MATCH_SCORE"]  = int(best_score)
                df.at[idx, "LETTER_SCORE"] = int(best_lscore)
            else:
                df.at[idx, "ALERTA_NORM"] = "SIN MATCH EN DICCIONARIO"

    # Reemplazos
    df["POZO"] = df["POZO_MATCH"].where(df["POZO_MATCH"].notna(), df["POZO"])
    meta_first = dict_df.groupby("oficial")[["met_prod","nivel_3","nivel_5"]].first()
    df = df.merge(meta_first, how="left", left_on="POZO", right_index=True)

    # ZONA sólo si hubo match; sino, vacío
    if "nivel_3" in df.columns:
        df.loc[df["POZO_MATCH"].isna(), "nivel_3"] = ""
        df["ZONA"] = np.where(df["POZO_MATCH"].notna(), df["nivel_3"].fillna(""), "")

    # BATERIA si hay nivel_5
    if "nivel_5" in df.columns:
        df["BATERIA"] = np.where(
            df["nivel_5"].notna() & (df["nivel_5"].astype(str).str.strip()!=""),
            df["nivel_5"], df["BATERIA"]
        )

    df["__ZONA_NORM"]    = df["ZONA"].apply(_norm)
    df["__BATERIA_NORM"] = df["BATERIA"].apply(_norm)

    norm_table = (df[["POZO_ORIG","POZO_PreCanon","__POZO_KEY",
                      "__KEY_LET","__KEY_DIG_CANON","__KEY_DIG_LEN",
                      "POZO_MATCH","MATCH_TIPO","MATCH_SCORE","LETTER_SCORE",
                      "APLICADO","ALERTA_NORM","VALIDO_POZO",
                      "met_prod","nivel_3","nivel_5"]]
                  .drop_duplicates()
                  .rename(columns={
                      "POZO_ORIG":"Pozo_Original",
                      "POZO_PreCanon":"Pozo_PreCanon",
                      "__POZO_KEY":"Clave_Normalizada",
                      "__KEY_LET":"Letras",
                      "__KEY_DIG_CANON":"Digitos_Canon",
                      "__KEY_DIG_LEN":"Digitos_Len",
                      "POZO_MATCH":"Match_Oficial",
                      "MATCH_TIPO":"Match_Tipo",
                      "MATCH_SCORE":"Match_Score",
                      "LETTER_SCORE":"Letter_Score",
                      "APLICADO":"Aplicado",
                      "ALERTA_NORM":"Alerta",
                      "VALIDO_POZO":"Valido",
                      "met_prod":"met_prod",
                      "nivel_3":"nivel_3",
                      "nivel_5":"nivel_5"
                  })
                  .sort_values(["Valido","Aplicado","Match_Tipo","Pozo_Original"], ascending=[False, False, True, True]))

    alert_table = norm_table[(norm_table["Valido"]==False) | (norm_table["Aplicado"]=="NO") | (norm_table["Match_Tipo"]=="NO")].copy()

    return df, alert_table, norm_table

def read_historial(xlsx_path, sheet_hist=None):
    xl = pd.ExcelFile(xlsx_path)
    sheets = [sheet_hist] if (sheet_hist and sheet_hist in xl.sheet_names) else xl.sheet_names
    for sh in sheets:
        raw = xl.parse(sh, header=None)
        idx, header_norm = _find_header_row(raw)
        if idx is None:
            continue
        data = raw.iloc[idx:, :].copy()
        true_headers = data.iloc[0,:].astype(str).tolist()
        data = data.iloc[1:,:]
        data.columns = true_headers

        name_map = {c: _norm(c) for c in data.columns}
        def find_col(candidates):
            for c, n in name_map.items():
                if n in candidates:
                    return c
            return None

        c_fecha       = find_col(set(EXPECTED_KEYS["fecha"]))
        c_pozo        = find_col(set(EXPECTED_KEYS["pozo"]))
        c_zona        = find_col(set(EXPECTED_KEYS["zona"]))
        c_bateria     = find_col(set(EXPECTED_KEYS["bateria"]))
        c_m3          = find_col(set(EXPECTED_KEYS["m3"]))
        c_carr        = find_col(set(EXPECTED_KEYS["carreras"]))
        c_nivel_final = find_col(set(EXPECTED_KEYS["nivel_final"]))

        if not (c_fecha and c_pozo and c_zona and c_bateria):
            continue

        use_cols = [c_fecha, c_pozo, c_zona, c_bateria]
        headers  = ["FECHA","POZO","ZONA","BATERIA"]
        if c_m3:            use_cols.append(c_m3);            headers.append("M3")
        if c_carr:          use_cols.append(c_carr);          headers.append("CARRERAS")
        if c_nivel_final:   use_cols.append(c_nivel_final);   headers.append("NIVEL_FINAL")

        df = data[use_cols].copy()
        df.columns = headers

        df["FECHA"] = pd.to_datetime(df["FECHA"], errors="coerce")
        if "M3" not in df.columns: df["M3"] = np.nan
        else: df["M3"] = pd.to_numeric(df["M3"], errors="coerce")

        if "CARRERAS" not in df.columns: df["CARRERAS"] = np.nan
        else: df["CARRERAS"] = pd.to_numeric(df["CARRERAS"], errors="coerce")

        if "NIVEL_FINAL" not in df.columns:
            df["NIVEL_FINAL"] = None

        for col in ["POZO","ZONA","BATERIA","NIVEL_FINAL"]:
            df[col] = df[col].astype(str).str.strip().replace({"nan": np.nan})

        df = df.dropna(subset=["FECHA","POZO"]).sort_values(["POZO","FECHA"])
        return df

    raise ValueError("No pude detectar FECHA/POZO/ZONA/BATERÍA en ninguna hoja del Excel.")

def read_exclusions_from_sheet(xlsx_path):
    excl = set()
    try:
        xl = pd.ExcelFile(xlsx_path)
        if "ExcluirPozos" in xl.sheet_names:
            e = xl.parse("ExcluirPozos")
            e.columns = [str(c).strip().lower() for c in e.columns]
            if "pozo" in e.columns:
                if "excluir" in e.columns:
                    excl = set(e.loc[e["excluir"].astype(str).str.upper().isin(
                        ["SI","SÍ","YES","1","TRUE"]), "pozo"].astype(str).str.strip())
                else:
                    excl = set(e["pozo"].astype(str).str.strip())
    except Exception:
        pass
    return excl

def prompt_int(prompt, default, lo, hi):
    try:
        s = input(prompt).strip()
    except EOFError:
        return default
    if s == "": return default
    try:
        v = int(float(s))
    except:
        return default
    return max(lo, min(hi, v))

def pick_list_checkbox(titulo, items, prechecked=None):
    prechecked = set(prechecked or [])
    try:
        import questionary
        from questionary import Choice
        choices = [Choice(title=x, value=x, checked=(x in prechecked)) for x in sorted(items)]
        print("\nUsando pick list. Controles: ↑/↓, ESPACIO marca, ENTER confirma.\n")
        selected = questionary.checkbox(titulo, choices=choices).ask()
        return set(selected or [])
    except Exception:
        print("\nTip: para casillas, instalá questionary ->  pip install questionary")
        items_sorted = sorted(items)
        index_to_item = {i+1: x for i,x in enumerate(items_sorted)}
        cols, width = 3, 28
        per_col = int(np.ceil(len(items_sorted)/cols)) if items_sorted else 0
        for r in range(per_col):
            row = []
            for c in range(cols):
                idx = r + 1 + c*per_col
                if idx in index_to_item:
                    row.append(f"{idx:>3}) {index_to_item[idx]}".ljust(width))
            if row:
                print("  ".join(row))
        try:
            raw = input("Índices (ej: 1,5,7-10). Enter = ninguno: ").strip()
        except EOFError:
            raw = ""
        selected = set()
        if raw:
            toks = re.split(r"[,\s]+", raw)
            for t in toks:
                if not t: continue
                if re.match(r"^\d+-\d+$", t):
                    a,b = t.split("-")
                    try:
                        a=int(a); b=int(b)
                        for i in range(min(a,b), max(a,b)+1):
                            if i in index_to_item: selected.add(index_to_item[i])
                    except: pass
                elif re.match(r"^\d+$", t):
                    i=int(t)
                    if i in index_to_item: selected.add(index_to_item[i])
        return selected

def pick_zonas_checkbox(zonas_series):
    mapping = {}
    for z in zonas_series.dropna().astype(str):
        zn = _norm(z)
        if zn and zn not in mapping:
            mapping[zn] = z.strip()
    etiquetas = set(mapping.values())
    sel_labels = pick_list_checkbox("Zonas (nivel_3) a INCLUIR (desmarcá lo que NO quieras):", etiquetas, prechecked=etiquetas)
    if not sel_labels:
        sel_labels = etiquetas
    sel_norm = {_norm(lbl) for lbl in sel_labels}
    return sel_labels, sel_norm

def pick_baterias_subfilter(df_norm, zonas_labels, zonas_norm):
    """
    Si el usuario eligió 'Las Heras CG - Canadon Escondida', ofrece sub-filtro por BATERÍA.
    Devuelve: dict zona_norm -> set(bateria_norm) ó None si sin restricción.
    """
    allowed = {}
    target_label = "Las Heras CG - Canadon Escondida"
    target_norm  = _norm(target_label)

    for z_lbl in zonas_labels:
        zn = _norm(z_lbl)
        if zn == target_norm:
            # baterías presentes para esa zona
            bats = (df_norm.loc[df_norm["__ZONA_NORM"]==zn, "BATERIA"]
                    .dropna().astype(str).str.strip())
            # quito vacías
            bats = bats[bats != ""]
            unique_bats = sorted(set(bats))
            if unique_bats:
                sel_bats = pick_list_checkbox(
                    f"BATERÍAS a INCLUIR dentro de '{target_label}' (desmarcá lo que NO quieras):",
                    unique_bats,
                    prechecked=unique_bats
                )
                # normalizadas
                allowed[zn] = {_norm(b) for b in sel_bats}
            else:
                allowed[zn] = None  # no hay dato, no restrinjo
        else:
            allowed[zn] = None  # otras zonas: sin restricción
    return allowed

# ==========================
# Frecuencias / r_m3_d
# ==========================
def _count_trailing_zeros_with_carr(g):
    cnt = 0
    for _, row in g.sort_values("FECHA").iloc[::-1].iterrows():
        m3 = row.get("M3", np.nan)
        car = row.get("CARRERAS", np.nan)
        if pd.notna(m3) and float(m3) == 0.0 and pd.notna(car) and float(car) > 0:
            cnt += 1
        else:
            break
    return cnt

def compute_frecuencias(df, params):
    v_target = params["m3_por_visita_objetivo"]
    min_d    = params["min_dias_freq"]
    max_d    = params["max_dias_freq"]
    k        = int(params["k_visitas"])
    one_days = int(params.get("dias_asumidos_una_visita", 7))
    freq_cero_ultimo = int(params.get("freq_dias_ultimo_cero_valido", 30))

    out = []
    for pozo, g0 in df.groupby("POZO", sort=False):
        g = g0.sort_values("FECHA").copy()

        for col in ["ZONA","BATERIA","NIVEL_FINAL"]:
            if col in g.columns:
                g[col] = g[col].replace({None: np.nan})
                g[col] = g[col].ffill().bfill()

        g["__ZONA_NORM"]    = g["ZONA"].apply(_norm)
        g["__BATERIA_NORM"] = g["BATERIA"].apply(_norm)
        g["__nf_norm"]      = g["NIVEL_FINAL"].apply(_norm) if "NIVEL_FINAL" in g.columns else ""

        med_validas_all = g[g["M3"].notna()].copy()

        m3_eq0 = g["M3"].fillna(0) == 0
        carr   = g.get("CARRERAS", pd.Series(index=g.index, dtype=float)).fillna(np.nan)
        zero_cond_a = m3_eq0 & (carr.fillna(0) >= 1)
        zero_cond_b = m3_eq0 & ((carr.isna()) | (carr.fillna(0) == 0)) & (g["__nf_norm"] == "surge")
        cond_cero_valido = zero_cond_a | zero_cond_b

        validas_rate = g[(g["M3"] > 0) | cond_cero_valido].copy()
        zeros_tail = _count_trailing_zeros_with_carr(g)

        ultima_med = med_validas_all["FECHA"].max() if not med_validas_all.empty else pd.NaT
        ultima_exi = g.loc[g["M3"]>0, "FECHA"].max() if "M3" in g.columns and not g[g["M3"]>0].empty else pd.NaT

        last_zero_valido = False
        if not med_validas_all.empty:
            idx_last = med_validas_all["FECHA"].idxmax()
            m3_last  = g.at[idx_last, "M3"]
            if pd.notna(m3_last) and float(m3_last) == 0.0:
                try:
                    last_zero_valido = bool(cond_cero_valido.loc[idx_last])
                except Exception:
                    last_zero_valido = False

        alerta = ""
        if last_zero_valido:
            alerta = f"ULTIMA_M3_0_VALIDO -> FREQ {freq_cero_ultimo}D"
        elif pd.notna(ultima_med):
            if zeros_tail > 0:
                alerta = f"ALERTA: {zeros_tail} cero(s) consecutivo(s) con Carreras>0"

        # === r_m3_d con K=1 (tasa más reciente). Si hay <K tasas, promedia las que haya.
        r = np.nan
        if not validas_rate.empty:
            v = validas_rate.copy()
            v["delta_d"] = v["FECHA"].diff().dt.days
            v.loc[v["delta_d"] <= 0, "delta_d"] = np.nan
            v["rate"] = v["M3"].fillna(0) / v["delta_d"]
            rates = v["rate"].dropna()
            if len(rates) >= 1:
                # K=1 -> última tasa; si K>1 promedio de últimas K
                r = rates.tail(min(k, len(rates))).mean()
            else:
                row = v.iloc[-1]
                m3 = float(row["M3"]) if pd.notna(row["M3"]) else 0.0
                if m3 > 0:
                    r = m3 / max(1, one_days)
                else:
                    r = np.nan
        else:
            if len(med_validas_all) == 1:
                row = med_validas_all.iloc[-1]
                m3 = float(row["M3"]) if pd.notna(row["M3"]) else 0.0
                if m3 > 0:
                    r = m3 / max(1, one_days)
                else:
                    r = np.nan

        # ===== FRECUENCIA (delta) =====
        if last_zero_valido:
            delta = int(freq_cero_ultimo)
        else:
            if pd.isna(r):      delta = 7
            elif r <= 0:        delta = max_d
            else:
                delta = max(min_d, min(max_d, float(v_target)/float(r)))
                delta = int(7 * round(delta / 7.0))
                if delta < 7:
                    delta = 7

        prox = (ultima_med + pd.Timedelta(days=int(delta))) if pd.notna(ultima_med) else pd.Timestamp(next_monday())

        out.append({
            "POZO": pozo,
            "ZONA": g["ZONA"].iloc[-1],
            "BATERIA": g["BATERIA"].iloc[-1],
            "ZONA_NORM": g["__ZONA_NORM"].iloc[-1],
            "BATERIA_NORM": g["__BATERIA_NORM"].iloc[-1],
            "r_m3_d": r,
            "ultima_medicion": ultima_med,
            "ultima_exitosa": ultima_exi,
            "delta_star_dias": int(delta),
            "proxima_visita_base": prox,
            "ceros_consec": zeros_tail,
            "alerta": alerta
        })
    return pd.DataFrame(out)

# ==========================
# Coordenadas
# ==========================
def _to_float_maybe_comma(x):
    if pd.isna(x):
        return np.nan
    if isinstance(x, (int, float, np.number)):
        return float(x)
    s = str(x).strip()
    if s == "": return np.nan
    s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return np.nan

def read_coords(xlsx_path):
    try:
        cdf = pd.read_excel(xlsx_path)
    except Exception as e:
        print(f"\n[AVISO] No pude leer coordenadas: {xlsx_path}\n{e}\n")
        return pd.DataFrame(columns=["POZO","LAT","LON"])
    cols_map = {c.lower().strip(): c for c in cdf.columns}
    c_pozo = cols_map.get("pozo")
    for k in ["geo_latitude","latitude","lat"]:
        if k in cols_map:
            c_lat = cols_map[k]; break
    else:
        c_lat = None
    for k in ["geo_longitude","longitude","lon","long"]:
        if k in cols_map:
            c_lon = cols_map[k]; break
    else:
        c_lon = None

    if not (c_pozo and c_lat and c_lon):
        print(f"[AVISO] Coordenadas: columnas esperadas 'POZO','GEO_LATITUDE','GEO_LONGITUDE'. Columnas encontradas: {list(cdf.columns)}")
        return pd.DataFrame(columns=["POZO","LAT","LON"])

    out = cdf[[c_pozo, c_lat, c_lon]].copy()
    out.columns = ["POZO","LAT","LON"]
    out["POZO"] = out["POZO"].astype(str).str.strip()
    out["LAT"] = out["LAT"].apply(_to_float_maybe_comma)
    out["LON"] = out["LON"].apply(_to_float_maybe_comma)
    out = out.dropna(subset=["POZO"])
    out = out.drop_duplicates(subset=["POZO"], keep="last")
    return out

# ==========================
# Candidatos y utilidades
# ==========================
def build_candidates_with_coords(freq, week_start, week_end, excl_pozos,
                                 zonas_norm_incluidas, coords_df,
                                 allowed_bats_by_zone_norm=None,
                                 next_due_map=None):
    # Partimos de frecuencias
    F = freq.copy()

    # due_date base (permitimos override con next_due_map)
    F["due_date"] = F["proxima_visita_base"]
    if next_due_map:
        F["due_date"] = F["POZO"].map(next_due_map).fillna(F["due_date"])

    F["overdue_d"] = (pd.Timestamp(week_start) - pd.to_datetime(F["due_date"])).dt.days
    F["is_overdue"] = F["overdue_d"] > 0

    # __v = r_m3_d (prioridad)
    F["__v"] = F["r_m3_d"].astype(float)

    # Filtro por zona elegida (nivel_3 ya normalizada)
    if "ZONA_NORM" in F.columns and zonas_norm_incluidas:
        F = F[F["ZONA_NORM"].isin(zonas_norm_incluidas)].copy()

    # Sub-filtro por BATERÍA (si corresponde)
    if allowed_bats_by_zone_norm:
        # Para cada zona filtrada, aplicar su set de baterías si existe
        mask = pd.Series(True, index=F.index)
        for zn in zonas_norm_incluidas:
            bats = allowed_bats_by_zone_norm.get(zn)
            if bats is not None:
                mask &= ~ (F["ZONA_NORM"] == zn) | (F["BATERIA_NORM"].isin(bats))
        F = F[mask].copy()

    # Excluir pozos
    if excl_pozos:
        F = F[~F["POZO"].isin(excl_pozos)].copy()

    # Filtro de potencial mínimo
    F = F[F["r_m3_d"].fillna(0) > RM3D_MIN].copy()

    # Eliminar baterías vacías (evita filas sin BATERIA)
    F = F[F["BATERIA"].notna() & (F["BATERIA"].astype(str).str.strip()!="")].copy()

    # Merge coordenadas (por POZO oficial)
    coords_df = coords_df if coords_df is not None else pd.DataFrame(columns=["POZO","LAT","LON"])
    F = F.merge(coords_df, how="left", on="POZO")
    F["has_coords"] = F["LAT"].notna() & F["LON"].notna()

    # Orden general (no es el orden final, sólo base)
    F = F.sort_values(by=["is_overdue","__v","due_date"], ascending=[False, False, True]).reset_index(drop=True)
    return F

def _v_est_for_day(row, day_date):
    r = row.get("r_m3_d", np.nan)
    u = row.get("ultima_medicion", pd.NaT)
    if pd.isna(u) or pd.isna(r) or r <= 0:
        return 0.0
    dd = max(0, (pd.Timestamp(day_date) - pd.Timestamp(u)).days)
    return max(0.0, float(r) * float(dd))

def haversine_km(lat1, lon1, lat2, lon2):
    try:
        if pd.isna(lat1) or pd.isna(lon1) or pd.isna(lat2) or pd.isna(lon2):
            return np.nan
        R = 6371.0088
        p1 = math.radians(float(lat1)); p2 = math.radians(float(lat2))
        dphi = math.radians(float(lat2) - float(lat1))
        dlmb = math.radians(float(lon2) - float(lon1))
        a = math.sin(dphi/2)**2 + math.cos(p1)*math.cos(p2)*math.sin(dlmb/2)**2
        return 2*R*math.asin(math.sqrt(a))
    except Exception:
        return np.nan

# ==========================
# Clustering "estrella" por día (semilla + vecinos en radio)
# ==========================
def _fill_day_star_clusters(day_date, avail_df, cap_per_day, radius_km, used_set):
    assigned = []
    remaining_cap = int(cap_per_day)

    # Orden semilla: 1) con coords primero, 2) mayor __v, 3) overdue, 4) due_date
    has_xy = avail_df["has_coords"].fillna(False)
    pool = pd.concat([
        avail_df.loc[has_xy].sort_values(["__v","is_overdue","due_date"], ascending=[False, False, True]),
        avail_df.loc[~has_xy].sort_values(["__v","is_overdue","due_date"], ascending=[False, False, True]),
    ], ignore_index=True)

    def _bb_filter(df, lat0, lon0, rad_km):
        if pd.isna(lat0) or pd.isna(lon0):
            return df.iloc[0:0]
        dlat = rad_km / 110.574
        dlon = rad_km / (111.320 * max(0.1, math.cos(math.radians(float(lat0)))))
        return df[(df["LAT"].between(lat0 - dlat, lat0 + dlat)) &
                  (df["LON"].between(lon0 - dlon, lon0 + dlon))].copy()

    while (remaining_cap > 0) and (not pool.empty):
        seed_row = pool.iloc[0]
        seed_pozo = seed_row["POZO"]
        seed_lat  = seed_row.get("LAT", np.nan)
        seed_lon  = seed_row.get("LON", np.nan)

        cluster_rows = [seed_row]

        if pd.notna(seed_lat) and pd.notna(seed_lon):
            neigh = _bb_filter(pool.iloc[1:], seed_lat, seed_lon, radius_km)
            if not neigh.empty:
                neigh = neigh.copy()
                neigh["__dist_seed"] = neigh.apply(
                    lambda r: haversine_km(seed_lat, seed_lon, r["LAT"], r["LON"]), axis=1
                )
                neigh = neigh[neigh["__dist_seed"] <= radius_km].copy()
                neigh = neigh.sort_values(["__v","__dist_seed"], ascending=[False, True])

                take_n = max(0, remaining_cap - 1)
                if take_n > 0 and not neigh.empty:
                    for _, nr in neigh.head(take_n).iterrows():
                        cluster_rows.append(nr)

        coords_cluster = [(r["LAT"], r["LON"]) for r in cluster_rows
                          if pd.notna(r.get("LAT", np.nan)) and pd.notna(r.get("LON", np.nan))]
        if coords_cluster:
            c_lat = float(np.mean([x for x,_ in coords_cluster]))
            c_lon = float(np.mean([y for _,y in coords_cluster]))
        else:
            c_lat, c_lon = (np.nan, np.nan)

        used_now = set()
        for r in cluster_rows:
            if remaining_cap <= 0:
                break
            pozo = r["POZO"]
            if pozo in used_set or pozo in used_now:
                continue

            lat = r.get("LAT", np.nan); lon = r.get("LON", np.nan)
            d_seed = haversine_km(seed_lat, seed_lon, lat, lon) if pd.notna(seed_lat) and pd.notna(seed_lon) else np.nan
            d_cent = haversine_km(c_lat, c_lon, lat, lon)       if pd.notna(c_lat)  and pd.notna(c_lon)  else np.nan

            assigned.append({
                "Plan_Fecha": day_date.date(),
                "Semana_ISO": day_date.isocalendar()[1],
                "ZONA": r["ZONA"],
                "BATERIA": r["BATERIA"],
                "POZO": pozo,
                "r_m3_d": float(r["__v"]),
                "ultima_medicion": r.get("ultima_medicion", pd.NaT),
                "Seed_POZO": seed_pozo,
                "Dist_km_semilla": None if pd.isna(d_seed) else round(float(d_seed), 3),
                "Dist_km_centroid": None if pd.isna(d_cent) else round(float(d_cent), 3),
            })
            used_now.add(pozo)
            remaining_cap -= 1

        if used_now:
            used_set.update(used_now)
            pool = pool[~pool["POZO"].isin(used_now)].copy()
        else:
            pool = pool.iloc[1:].copy()

    return assigned

def assign_week_zone_locked_star_clustering(cand_zone, params, week_start, week_end, radius_km):
    dias    = int(params["dias_por_semana"])
    cap_pz  = int(params["max_pozos_dia_equipo"])

    used = set()
    rows = []
    for d in range(dias):
        day_date = pd.Timestamp(week_start) + pd.Timedelta(days=d)

        pool = cand_zone[~cand_zone["POZO"].isin(used)].copy()
        if pool.empty: continue
        in_window = (pd.to_datetime(pool["due_date"]) <= pd.Timestamp(week_end)) | pool["is_overdue"]
        pool = pool[in_window].copy()
        if pool.empty: continue

        MAX_POOL_PER_DAY = 80
        pool = pool.sort_values(["__v","is_overdue","due_date"], ascending=[False, False, True]).head(MAX_POOL_PER_DAY)

        assigned_today = _fill_day_star_clusters(day_date, pool, cap_pz, radius_km, used)

        for ord_idx, a in enumerate(assigned_today, start=1):
            v_est = 0.0
            try:
                dummy_row = {"r_m3_d": a["r_m3_d"], "ultima_medicion": a["ultima_medicion"]}
                v_est = _v_est_for_day(dummy_row, day_date)
            except Exception:
                v_est = 0.0
            a.update({
                "Equipo": 0,
                "Dia_Idx": d+1,
                "Orden": ord_idx,
                "Vol_Estimado_m3": round(float(v_est), 2)
            })
        rows.extend(assigned_today)

    out = pd.DataFrame(rows, columns=[
        "Plan_Fecha","Semana_ISO","Equipo","Dia_Idx","Orden",
        "ZONA","BATERIA","POZO","r_m3_d","Vol_Estimado_m3",
        "Seed_POZO","Dist_km_semilla","Dist_km_centroid","ultima_medicion"
    ])
    return out

def ensure_annual_coverage_zone_locked(all_pozos_df, plan, params, start_date, equipo_to_zona,
                                       allowed_bats_by_zone_norm=None, r_by_pozo=None):
    """
    all_pozos_df: DataFrame con columnas ['POZO','ZONA','BATERIA'] YA FILTRADAS
                  (por r_m3_d > RM3D_MIN y baterías permitidas)
    En caso de que igualmente entre algo no deseado, reforzamos validaciones aquí.
    """
    cap_pz = params["max_pozos_dia_equipo"]

    keys = []
    for w in range(params["semanas_plan"]):
        w_start = start_date + timedelta(weeks=w)
        for d in range(params["dias_por_semana"]):
            f = w_start + timedelta(days=d)  # datetime.date
            for e in equipo_to_zona.keys():
                keys.append((e, f))

    if not plan.empty:
        plan["__key"] = plan["Equipo"].astype(int).astype(str) + "|" + plan["Plan_Fecha"].astype(str)
        used_counts = plan.groupby("__key")["POZO"].count().to_dict()
    else:
        used_counts = {}

    planned = set(plan["POZO"].unique()) if not plan.empty else set()
    missing_df = all_pozos_df[~all_pozos_df["POZO"].isin(planned)].copy()

    # Reforzar: quitar baterías vacías
    missing_df = missing_df[missing_df["BATERIA"].notna() & (missing_df["BATERIA"].astype(str).str.strip()!="")].copy()

    add = []
    for _, row in missing_df.iterrows():
        pz = row["POZO"]; z = row["ZONA"]
        bat = row.get("BATERIA", "")

        # BATERÍA obligatoria
        if not isinstance(bat, str) or bat.strip() == "":
            continue

        # Re-chequeo de sub-filtro por batería si aplica
        if allowed_bats_by_zone_norm:
            zn = _norm(z)
            bats_allowed = allowed_bats_by_zone_norm.get(zn)
            if bats_allowed is not None:
                if _norm(bat) not in bats_allowed:
                    continue

        # Re-chequeo r_m3_d si nos pasan el dict
        if r_by_pozo is not None:
            r_val = float(r_by_pozo.get(pz, np.nan))
            if not (r_val > RM3D_MIN):
                continue

        target_teams = [e for e, zona in equipo_to_zona.items() if zona == z]
        if not target_teams:
            continue
        placed = False
        for e in target_teams:
            for (ee, f) in keys:
                if ee != e:
                    continue
                key = f"{e}|{f}"
                cnt = used_counts.get(key, 0)
                if cnt < cap_pz:
                    add.append({
                        "Plan_Fecha": f,
                        "Semana_ISO": f.isocalendar()[1],
                        "Equipo": int(e),
                        "Dia_Idx": f.weekday()+1,
                        "Orden": cnt+1,
                        "ZONA": z,
                        "BATERIA": bat,
                        "POZO": pz,
                        "r_m3_d": np.nan,
                        "Vol_Estimado_m3": 0.0,
                        "Seed_POZO": "",
                        "Dist_km_semilla": None,
                        "Dist_km_centroid": None,
                        "ultima_medicion": pd.NaT,
                    })
                    used_counts[key] = cnt+1
                    placed = True
                    break
            if placed:
                break

    if add:
        plan = pd.concat([plan, pd.DataFrame(add)], ignore_index=True)                 .sort_values(["Plan_Fecha","Equipo","Orden"])
    return plan

# ==========================
# Alertas de ABM
# ==========================
def build_alertas_abm(freq_df: pd.DataFrame, norm_table: pd.DataFrame, dict_df: pd.DataFrame) -> pd.DataFrame:
    base = freq_df[["POZO","ZONA","BATERIA","ultima_medicion","ultima_exitosa"]].copy()
    meta_first = dict_df.groupby("oficial")[["estado","met_prod","nivel_3","nivel_5"]].first()
    base = base.merge(meta_first[["estado","met_prod"]], left_on="POZO", right_index=True, how="left")

    out = base.copy()
    for c in ["ultima_medicion","ultima_exitosa"]:
        out[c] = pd.to_datetime(out[c], errors="coerce").dt.date
    out = out.sort_values(["ZONA","BATERIA","POZO"]).reset_index(drop=True)
    return out

# ==========================
# MAIN
# ==========================
def main():
    # 1) Historial original
    df = read_historial(INPUT_FILE, SHEET_HIST)

    # 2) Normalización + reemplazo de POZO; ZONA solo si hubo match; BATERIA si nivel_5
    key2off, dict_df = load_pozo_dictionary(NOMBRES_POZO_FILE)
    df_norm, alert_table, norm_table = apply_pozo_normalization(df, key2off, dict_df)

    # 2.b) Eliminar del historial los POZOS inválidos (sin letras o sin dígitos)
    df = df_norm[df_norm["VALIDO_POZO"] == True].copy()
    if (df_norm["VALIDO_POZO"] == False).any():
        print("[AVISO] Se filtraron filas con POZO inválido (sin letras o sin dígitos). Ver 'Alertas_Normalizacion'.")

    # 3) ZONAS (pick list) sobre ZONA normalizada (nivel_3) y filtro
    zonas_labels, zonas_norm = pick_zonas_checkbox(df["ZONA"])
    df = df[df["__ZONA_NORM"].isin(zonas_norm)].copy()

    # 3.b) Sub-filtro por BATERÍA si corresponde (solo Las Heras CG - Canadon Escondida)
    allowed_bats_by_zone_norm = pick_baterias_subfilter(df, zonas_labels, zonas_norm)

    # 4) Equipos
    pozos_unicos = sorted(df["POZO"].unique())
    print(f"\nZONAS seleccionadas (nivel_3): {', '.join(sorted(zonas_labels))}")
    print(f"Pozos detectados en esas zonas: {len(pozos_unicos)}")

    equipos_user = prompt_int("\n¿Cuántos equipos querés planificar? [1-4, Enter=4]: ",
                              default=DEFAULTS["equipos_activos"], lo=1, hi=4)
    DEFAULTS["equipos_activos"] = equipos_user

    # 5) Exclusión de pozos (archivo + usuario) — ya con POZO normalizado
    excl_archivo = read_exclusions_from_sheet(INPUT_FILE)
    excl_archivo = set([p for p in excl_archivo if p in pozos_unicos])
    excl_user = pick_list_checkbox("Pozos a EXCLUIR:", pozos_unicos, prechecked=excl_archivo)
    excl_total = set(excl_archivo) | set(excl_user)
    print(f"\nPozos excluidos: {len(excl_total)}")

    # 6) Frecuencias (ya con ZONA/BATERIA normalizadas)
    params = DEFAULTS.copy()
    freq = compute_frecuencias(df, params)

    # 7) Cargar coordenadas y unir
    coords_df = read_coords(COORDS_FILE)

    # 8) Mapas auxiliares
    delta_by_pozo = freq.set_index("POZO")["delta_star_dias"].to_dict()
    r_by_pozo     = freq.set_index("POZO")["r_m3_d"].to_dict()

    # 9) Semanas a planificar
    start = next_monday(date.today())
    weeks = [(start + timedelta(weeks=i), start + timedelta(weeks=i, days=6)) for i in range(params["semanas_plan"])]

    # 10) Mapeo Equipo -> ZONA (fijo)
    zonas_list = sorted(set(zonas_labels))
    equipo_to_zona = {}
    for i in range(1, params["equipos_activos"]+1):
        zona_asignada = zonas_list[min(i-1, len(zonas_list)-1)]
        equipo_to_zona[i] = zona_asignada
    print("\nAsignación fija de equipos por ZONA:")
    for e, z in equipo_to_zona.items():
        print(f"  Equipo {e} -> ZONA: {z}")

    # 11) Plan (por semana y por equipo, bloqueado a su zona) con due dinámico
    plan_all = []
    next_due = {row.POZO: row.proxima_visita_base for row in freq.itertuples()}

    for (w_start, w_end) in weeks:
        for eq, zona_label in equipo_to_zona.items():
            zona_norm_label = _norm(zona_label)
            cand_all = build_candidates_with_coords(
                freq=freq,
                week_start=w_start,
                week_end=w_end,
                excl_pozos=excl_total,
                zonas_norm_incluidas={zona_norm_label},
                coords_df=coords_df,
                allowed_bats_by_zone_norm=allowed_bats_by_zone_norm,
                next_due_map=next_due  # << usa due actualizado
            )
            if cand_all.empty:
                continue

            cand_zone = cand_all[[
                "POZO","ZONA","BATERIA","due_date","is_overdue","__v","LAT","LON","has_coords","r_m3_d","ultima_medicion"
            ]].copy()

            plan_week = assign_week_zone_locked_star_clustering(
                cand_zone=cand_zone,
                params=params,
                week_start=w_start,
                week_end=w_end,
                radius_km=RADIUS_KM
            )

            if not plan_week.empty:
                plan_week["Equipo"] = eq
                plan_all.append(plan_week)

                # Actualizar due_date de los pozos asignados (siguiente visita según delta)
                for pz, fcal in plan_week[["POZO","Plan_Fecha"]].drop_duplicates().itertuples(index=False):
                    dd = int(delta_by_pozo.get(pz, params["min_dias_freq"]))
                    next_due[pz] = pd.Timestamp(fcal) + pd.Timedelta(days=dd)

    plan = pd.concat(plan_all, ignore_index=True) if plan_all else pd.DataFrame(columns=[
        "Plan_Fecha","Semana_ISO","Equipo","Dia_Idx","Orden","ZONA","BATERIA","POZO","r_m3_d","Vol_Estimado_m3","Seed_POZO","Dist_km_semilla","Dist_km_centroid","ultima_medicion"
    ])

    # 12) Cobertura anual (respetando zonas, baterías y umbral de r_m3_d)
    if not freq.empty:
        eligible_mask = (freq["ZONA"].isin(zonas_labels)) & (freq["r_m3_d"].fillna(0) > RM3D_MIN)

        if allowed_bats_by_zone_norm:
            for zn, bats in allowed_bats_by_zone_norm.items():
                if bats is not None:
                    eligible_mask &= (~(freq["ZONA_NORM"] == zn)) | (freq["BATERIA_NORM"].isin(bats))

        all_pozos_in_zonas = freq.loc[eligible_mask, ["POZO","ZONA","BATERIA"]].drop_duplicates().copy()
        all_pozos_in_zonas = all_pozos_in_zonas[
            all_pozos_in_zonas["BATERIA"].notna() & (all_pozos_in_zonas["BATERIA"].astype(str).str.strip() != "")
        ].copy()

        plan = ensure_annual_coverage_zone_locked(
            all_pozos_in_zonas,
            plan,
            params,
            start,
            equipo_to_zona,
            allowed_bats_by_zone_norm=allowed_bats_by_zone_norm,
            r_by_pozo=r_by_pozo  # refuerzo r_m3_d > RM3D_MIN
        )

    # 13) Excel NUEVO (+ Alertas de ABM + Alertas de Coordenadas)
    output_path = unique_output_path(INPUT_FILE)
    with pd.ExcelWriter(output_path, engine="openpyxl", mode="w") as writer:
        # Frecuencias
        freq_out = freq.copy()
        for c in ["proxima_visita_base","ultima_medicion","ultima_exitosa"]:
            freq_out[c] = pd.to_datetime(freq_out[c], errors="coerce").dt.date
        freq_out = freq_out.sort_values(["ZONA","BATERIA","POZO"])
        freq_out.to_excel(writer, "Frecuencias", index=False)

        # Plan por equipo (ordenado y con columnas nuevas)
        cols_plan = ["Plan_Fecha","Semana_ISO","Equipo","Dia_Idx","Orden",
                     "ZONA","BATERIA","POZO","r_m3_d","Vol_Estimado_m3",
                     "Seed_POZO","Dist_km_semilla","Dist_km_centroid"]
        for eq in range(1, params["equipos_activos"]+1):
            pe = plan.loc[plan["Equipo"]==eq].copy()
            if pe.empty:
                pe = pd.DataFrame(columns=cols_plan + ["Ejecutado"])
            else:
                pe["Ejecutado"] = ""
                pe = pe.sort_values(["Plan_Fecha","Dia_Idx","Orden","POZO"])
                for c in cols_plan:
                    if c not in pe.columns:
                        pe[c] = ""
                pe = pe[cols_plan + ["Ejecutado"]]
            pe.to_excel(writer, f"Plan_Equipo_{eq}", index=False)

        # Parámetros / exclusiones / zonas elegidas / baterías filtradas
        pd.DataFrame(list(params.items()), columns=["Parametro","Valor"]).to_excel(writer, "Parametros_Usados", index=False)
        if excl_total:
            pd.DataFrame(sorted(excl_total), columns=["Pozo_Excluido"]).to_excel(writer, "Exclusiones_Usadas", index=False)
        pd.DataFrame(sorted(zonas_labels), columns=["ZONA_SELECCIONADA"]).to_excel(writer, "Zonas_Seleccionadas", index=False)

        # Baterías filtradas (si aplica)
        if allowed_bats_by_zone_norm:
            rows_b = []
            for zn, bats in allowed_bats_by_zone_norm.items():
                rows_b.append({
                    "ZONA_NORM": zn,
                    "BATERIAS_NORM_PERMITIDAS": ", ".join(sorted(bats)) if bats is not None else "(sin restricción)"
                })
            pd.DataFrame(rows_b).to_excel(writer, "Baterias_Filtradas", index=False)

        # Normalización (con metadatos)
        if not norm_table.empty:
            norm_table.to_excel(writer, "Normalizacion_Pozos", index=False)
        if not alert_table.empty:
            alert_table.to_excel(writer, "Alertas_Normalizacion", index=False)

        # Alertas de ABM (con estado y met_prod)
        alertas_abm = build_alertas_abm(freq_out, norm_table, dict_df)
        alertas_abm.to_excel(writer, "Alertas de ABM", index=False)

        # Alertas de Coordenadas (pozos sin lat/lon) vistos en el plan
        if not plan.empty:
            plan_pozos = plan[["POZO"]].drop_duplicates()
            coords_used = plan_pozos.merge(read_coords(COORDS_FILE), how="left", on="POZO")
            faltantes = coords_used[coords_used[["LAT","LON"]].isna().any(axis=1)].copy()
            if not faltantes.empty:
                faltantes.to_excel(writer, "Alertas_Coordenadas", index=False)

    print("\nListo ✅  Cronograma generado en:", output_path)
    print(f"Radio usado para clústers: {RADIUS_KM} km | Filtro r_m3_d > {RM3D_MIN}")

if __name__ == "__main__":
    main()

