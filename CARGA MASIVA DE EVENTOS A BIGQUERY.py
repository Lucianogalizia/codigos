#!/usr/bin/env python
# coding: utf-8

# In[8]:


#---------con una columna adicional que me filtra los pozos y la batería

import os
import cx_Oracle
import pandas as pd

# Configurar las variables de entorno (si usás cliente Oracle tipo "thick")
os.environ['ORACLE_HOME'] = r"C:\app\product\12.1.0\client_1"
os.environ['PATH'] = os.environ['ORACLE_HOME'] + r"\bin;" + os.environ['PATH']

# Configuración de la conexión
dsn = cx_Oracle.makedsn("slplpgmoora03", 1527, service_name="psfu")
connection = None

try:
    connection = cx_Oracle.connect(
        user="RY33872",
        password="Contraseña_112025",
        dsn=dsn,
        encoding="UTF-8"
    )
    print("Conexión exitosa")
    
    query = """
    SELECT 
           o.NOMBRE_CORTO_POZO,
           o.NOMBRE_CORTO,
           o.NOMBRE_POZO,
           o.TIPO,
           o.ESTADO,
           o.MET_PROD,
           o.BATERIA,
           e.ASS_NAME,
           e.COMP_GROUP,
           e.COMP_NAME,
           e.CONDITION,
           e.NO_JOINTS,
           e.INSTL_DATE,
           e.LENGTH,
           e.NOM_SIZE,
           e.TOP_SET
    FROM DISC_ADMINS.DBU_FIC_ORG_ESTRUCTURAL o,
         DISC_ADMINS.DBU_FIC_ULTIMO_EQUIPAM   e
    WHERE e.UWI(+) = o.API_CD
      AND o.ESTADO IN ('Produciendo', 'Reserva Recuperación Secundaria')
      AND o.TIPO NOT IN ('Inyector - Agua', 'Inyector - Butano', 'Inyector - Gas', 'Inyector - Marginal', 'Inyector - Polímero')
      AND o.TIPO NOT IN ('Productor - Agua', 'Productor - Gas Y Condensado')
      AND o.MET_PROD NOT IN ('Plunger Lift Asistido', 'Plunger Lift')
      AND e.ASS_NAME NOT IN ('CAÑERIA AISLACION', 'CAÑERIA GUIA', 'CAÑERIA INTERMEDIA')
    ORDER BY 
      o.BATERIA ASC,
      o.NOMBRE_CORTO_POZO ASC,
      e.ASS_NAME,
      e.COMP_GROUP ASC,
      e.INSTL_DATE ASC,
      e.TOP_SET ASC
    """

    # Ejecutar la consulta y obtener los datos
    df = pd.read_sql(query, connection)

    # Crear una nueva columna eliminando duplicados de Nombre del Pozo
    df['NOMBRE_CORTO_POZO_UNICO'] = df['NOMBRE_CORTO_POZO'].mask(
        df.duplicated(subset='NOMBRE_CORTO_POZO')
    )

    # Inicializar columnas vacías
    df['ASS_NAME_INICIO'] = ''
    df['ASS_NAME_FINAL'] = ''

    # Aplicar la lógica de división condicional
    for i, row in df.iterrows():
        ass_name = row['ASS_NAME']
        if pd.isna(ass_name):
            continue
        if any(x in ass_name for x in ['SARTA TUBING -BM-', 'SARTA TUBING -ES-', 'SARTA TUBING -PCP-']):
            df.at[i, 'ASS_NAME_INICIO'] = ass_name[:12]
            df.at[i, 'ASS_NAME_FINAL'] = ass_name[13:]
        elif any(x in ass_name for x in ['SARTA VARILLAS -BM-', 'SARTA VARILLAS -PCP-']):
            df.at[i, 'ASS_NAME_INICIO'] = ass_name[:14]
            df.at[i, 'ASS_NAME_FINAL'] = ass_name[14:]
    
    # ==========================
    # ORDEN QUE PEDISTE
    # ==========================

    # Normalizar espacios en ASS_NAME_INICIO
    df['ASS_NAME_INICIO_LIMPIO'] = df['ASS_NAME_INICIO'].fillna('').str.strip()

    # Mapear orden deseado
    orden_ass = {
        'SARTA TUBING': 1,
        'SARTA VARILLAS': 2
    }
    # Cualquier otro (incluido vacío) va como 3
    df['ASS_ORDER'] = df['ASS_NAME_INICIO_LIMPIO'].map(orden_ass).fillna(3).astype(int)

    # Ordenar:
    # 1) por NOMBRE_POZO
    # 2) por orden de SARTA (tubing, varillas, vacíos)
    # 3) por TOP_SET de menor a mayor
    df = df.sort_values(
        by=['NOMBRE_POZO', 'ASS_ORDER', 'TOP_SET'],
        ascending=[True, True, True]
    ).reset_index(drop=True)

    # Si no querés que queden las columnas auxiliares en el Excel:
    df = df.drop(columns=['ASS_NAME_INICIO_LIMPIO', 'ASS_ORDER'])

    # Guardar el Excel en tu usuario
    output_file = r"C:\Users\ry16123\Instalacion_pozo1.xlsx"
    df.to_excel(output_file, sheet_name='Instalacion_pozo1', index=False)
    print(f"Excel guardado en: {output_file}")

except cx_Oracle.Error as error:
    print("Error al conectar / ejecutar:", error)
finally:
    if connection:
        connection.close()


# In[6]:


import os
import glob
import pandas as pd
from google.cloud import bigquery

# 🔐 Credenciales
os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = r"C:\Users\ry16123\Downloads\service-account.json"

# 📂 Carpeta con los Excels por año
CARPETA = r"C:\Users\ry16123\salidas_excel"

# 🧱 Tabla destino
PROJECT_ID = "eventos-479403"
DATASET = "eventos_pozos"
TABLE = "eventos"
TABLE_ID = f"{PROJECT_ID}.{DATASET}.{TABLE}"

# ==========================
# 1) Cliente y esquema BQ
# ==========================
client = bigquery.Client(project=PROJECT_ID)

# Traemos la tabla para ver el esquema real
table = client.get_table(TABLE_ID)

NUMERIC_TYPES = {"FLOAT", "FLOAT64", "INTEGER", "INT64", "NUMERIC", "BIGNUMERIC"}
STRING_TYPES = {"STRING"}

numeric_fields = [f.name for f in table.schema if f.field_type.upper() in NUMERIC_TYPES]
string_fields  = [f.name for f in table.schema if f.field_type.upper() in STRING_TYPES]

print("Columnas numéricas según BigQuery:")
for col in numeric_fields:
    print("  -", col)

print("\nColumnas string según BigQuery:")
for col in string_fields:
    print("  -", col)

# ==========================
# 2) Buscar todos los Excels
# ==========================
archivos = sorted(glob.glob(os.path.join(CARPETA, "*.xlsx")))

if not archivos:
    print("⚠ No se encontraron .xlsx en la carpeta:", CARPETA)
    raise SystemExit

print("\nArchivos encontrados para subir a BigQuery:")
for a in archivos:
    print("  -", a)

# ==========================
# 3) Subir uno por uno
# ==========================
primero = True

for archivo in archivos:
    print(f"\nSubiendo archivo: {archivo}")
    df = pd.read_excel(archivo)

    # 🔧 3.1. Forzar NUMÉRICAS donde BQ espera numérico
    for col in numeric_fields:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # 🔧 3.2. Forzar STRING donde BQ espera string
    for col in string_fields:
        if col in df.columns:
            df[col] = df[col].astype("string")  # dtype string/obj, todo texto

    # Primer archivo pisa la tabla, el resto APPEND
    if primero:
        write_mode = bigquery.WriteDisposition.WRITE_TRUNCATE
        primero = False
        print("  → Modo: WRITE_TRUNCATE (crear/pisar tabla)")
    else:
        write_mode = bigquery.WriteDisposition.WRITE_APPEND
        print("  → Modo: WRITE_APPEND (agregar filas)")

    job_config = bigquery.LoadJobConfig(
        write_disposition=write_mode,
        autodetect=True,
    )

    job = client.load_table_from_dataframe(df, TABLE_ID, job_config=job_config)
    job.result()  # esperar a que termine

    print(f"✔ Cargado OK: {os.path.basename(archivo)}")

print("\n🔥 LISTO: todos los Excels se cargaron respetando el esquema de BigQuery.")





# In[10]:


#esta consulta es la que funciona, carga todo a una tabla denominada eventos_fix
import os
import glob
import pandas as pd
from google.cloud import bigquery

# =====================================================
# 1) CONFIG
# =====================================================

# 🔐 Credenciales de servicio
os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = r"C:\Users\ry16123\Downloads\service-account.json"

# 📂 Carpeta donde tenés TODOS los Excels (1996, 1999, 2003, etc.)
CARPETA_EXCEL = r"C:\Users\ry16123\salidas_excel"

# 🔭 Proyecto / dataset / tabla destino
PROJECT_ID = "eventos-479403"
DATASET    = "eventos_pozos"
TABLE      = "eventos_fix"
TABLE_ID   = f"{PROJECT_ID}.{DATASET}.{TABLE}"

# Columnas NUMÉRICAS reales
NUMERIC_COLS = [
    "activity_duration",
    "cost_code",
    "pickup_weight",
    "step_no",
]

# Columnas FECHA-HORA (TIMESTAMP en BQ)
DATETIME_COLS = [
    "time_from",
    "time_to",
    "date_time_off_location",
    "date_rig_pickup",
]

# Columnas FECHA (DATE en BQ)
DATE_COLS = [
    "date_ops_end",
    "date_ops_start",
    "date_report",
]

# Columnas que siempre deben ir como TEXTO / CÓDIGO
STRING_COLS = [
    "contractor_name",
    "rig_name",
    "loc_fed_lease_no",
    "loc_region",
    "site_name",
    "field_name",
    "well_legal_name",
    "activity_class",
    "activity_class_desc",
    "activity_code",
    "activity_code_desc",
    "activity_group",
    "expr1",
    "activity_phase",
    "billing_code",
    "activity_subcode",
    "activity_subcode2",
    "event_code",
    "event_id",
    "event_objective_1",
    "event_objective_2",
    "event_type",
    "status_end",
    "well_id",
    "entity_type",
]

# Todas las columnas definidas en la tabla eventos_fix
ALL_COLS = [
    "contractor_name",
    "rig_name",
    "loc_fed_lease_no",
    "loc_region",
    "site_name",
    "field_name",
    "well_legal_name",
    "activity_class",
    "activity_class_desc",
    "activity_code",
    "activity_code_desc",
    "activity_duration",
    "activity_group",
    "expr1",
    "activity_phase",
    "billing_code",
    "activity_subcode",
    "activity_subcode2",
    "cost_code",
    "pickup_weight",
    "step_no",
    "time_from",
    "time_to",
    "date_ops_end",
    "date_ops_start",
    "event_code",
    "event_id",
    "event_objective_1",
    "event_objective_2",
    "event_type",
    "status_end",
    "well_id",
    "date_time_off_location",
    "date_report",
    "entity_type",
    "date_rig_pickup",
]

# =====================================================
# 2) CLIENTE BIGQUERY
# =====================================================
client = bigquery.Client(project=PROJECT_ID)

# =====================================================
# 3) BUSCAR EXCELS
# =====================================================
archivos = sorted(glob.glob(os.path.join(CARPETA_EXCEL, "*.xlsx")))

if not archivos:
    raise SystemExit(f"⚠ No se encontraron .xlsx en la carpeta: {CARPETA_EXCEL}")

print("\n📂 Archivos a cargar en BigQuery:")
for a in archivos:
    print("  -", a)

# =====================================================
# 4) CARGAR UNO POR UNO
# =====================================================
primero = True

for archivo in archivos:
    print("\n=================================================")
    print(f"📁 Procesando archivo: {archivo}")

    # Leemos el Excel
    df = pd.read_excel(archivo)

    # Nos quedamos SOLO con las columnas que existen en la tabla
    cols_presentes = [c for c in ALL_COLS if c in df.columns]
    df = df[cols_presentes].copy()

    # ---- Forzar STRING donde corresponde ----
    for col in STRING_COLS:
        if col in df.columns:
            df[col] = df[col].astype("string")

    # ---- Convertir NUMÉRICAS ----
    for col in NUMERIC_COLS:
        if col in df.columns:
            print(f"  · Numérica → {col}")
            serie = df[col].astype(str).str.replace(",", ".", regex=False)
            df[col] = pd.to_numeric(serie, errors="coerce")

    # ---- Convertir DATETIME (timestamp) ----
    for col in DATETIME_COLS:
        if col in df.columns:
            print(f"  · Datetime → {col}")
            df[col] = pd.to_datetime(df[col], errors="coerce")

    # ---- Convertir DATE ----
    for col in DATE_COLS:
        if col in df.columns:
            print(f"  · Date → {col}")
            df[col] = pd.to_datetime(df[col], errors="coerce").dt.date

    # ---- Modo de escritura ----
    if primero:
        write_mode = bigquery.WriteDisposition.WRITE_TRUNCATE
        primero = False
        print("  🚀 Modo de carga: WRITE_TRUNCATE (pisar/crear eventos_fix)")
    else:
        write_mode = bigquery.WriteDisposition.WRITE_APPEND
        print("  🚀 Modo de carga: WRITE_APPEND (agregar filas)")

    job_config = bigquery.LoadJobConfig(
        write_disposition=write_mode,
        autodetect=False,  # usamos el esquema de la tabla ya creada
    )

    job = client.load_table_from_dataframe(df, TABLE_ID, job_config=job_config)
    job.result()  # Espera a que termine

    print(f"✅ Cargado OK: {os.path.basename(archivo)}")

print("\n🔥 LISTO: todos los Excels fueron cargados en eventos_fix sin perder información.")


# In[ ]:





# In[ ]:




