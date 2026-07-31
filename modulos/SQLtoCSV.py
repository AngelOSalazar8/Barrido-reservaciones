"""
SQLtoCSV.py
Flujo unificado del Barrido Quoted Rate:
  1. Extrae reservas desde SQL Server (sin guardar CSV intermedio)
  2. Ejecuta el barrido en BlueZone/Wizard via modulos.barrido_wizard
  3. Guarda el CSV final con Fecha_Consulta en /Barridos/ y lo abre

Requiere Python 32-bit para la conexion con BlueZone (win32com).
"""

import json
import os
import sys

import pandas as pd
import pyodbc
from datetime import datetime

from . import barrido_wizard


# ============================================================
# RUTAS BASE
# ============================================================

# Raiz del proyecto — compatible con ejecucion como script y como exe compilado.
# Cuando PyInstaller genera --onefile, __file__ apunta al directorio temporal _MEI*.
# sys.executable apunta siempre al exe real junto al config.json.
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    # Script: subir un nivel desde modulos/ hasta la raiz del proyecto
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

SQL_PATH     = os.path.join(BASE_DIR, "Consultas", "WkSQLQuery.sql")
CONFIG_PATH  = os.path.join(BASE_DIR, "config.json")
BARRIDOS_DIR = os.path.join(BASE_DIR, "Barridos")


# ============================================================
# FLUJO PRINCIPAL
# ============================================================

def ejecutar_proceso_completo(days_back: int = 1, log=None, progress=None) -> pd.DataFrame:
    """
    Ejecuta el proceso completo de Barrido Quoted Rate en tres pasos:
      [1/3] Extraccion de reservas desde SQL Server.
      [2/3] Barrido en BlueZone/Wizard (P502).
      [3/3] Guardado del CSV final en /Barridos/.

    Parametros:
        days_back : Dias hacia atras para la consulta SQL.
                    1 = dia anterior, 3 = fin de semana, N = caso especial.
        log       : Funcion callback para reportar progreso a la GUI.
                    Recibe un string. Si es None, usa print().
        progress  : Funcion callback(actual, total) para la barra de progreso.

    Retorna:
        DataFrame con el resultado final del barrido.
    """
    def _log(msg):
        if log:
            log(msg)
        else:
            print(msg)

    fecha_str = datetime.now().strftime("%d-%m-%Y")
    csv_final = os.path.join(BARRIDOS_DIR, f"Barrido_reservas-{fecha_str}.csv")

    # ----------------------------------------------------------
    # PASO 1: Extraccion SQL
    # ----------------------------------------------------------
    _log(f"[1/3] Conectando con SQL Server... (offset: -{days_back} dias)")

    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            config = json.load(f)
        db_cfg = config["database"]
        conn_str = (
            f"DRIVER={{{db_cfg['driver']}}};"
            f"SERVER={db_cfg['server']};"
            f"DATABASE={db_cfg['dbname']};"
            f"UID={db_cfg['user']};"
            f"PWD={db_cfg['password']}"
        )
        conn = pyodbc.connect(conn_str)
    except Exception as e:
        _log(f"Error al conectar con SQL Server: {e}")
        raise

    try:
        with open(SQL_PATH, "r", encoding="utf-8") as f:
            query = f.read().replace("{DAYS_BACK}", str(days_back))

        _log("[1/3] Ejecutando consulta SQL...")
        df_reservas = pd.read_sql_query(query, conn)
    finally:
        conn.close()

    if "ClienteNacional" in df_reservas.columns:
        df_reservas["ClienteNacional"] = df_reservas["ClienteNacional"].astype(int)

    _log(f"[1/3] Reservas encontradas: {len(df_reservas)}")

    # ----------------------------------------------------------
    # PASO 2: Barrido BlueZone
    # ----------------------------------------------------------
    _log("[2/3] Iniciando barrido en BlueZone/Wizard...")
    df_resultado = barrido_wizard.ejecutar_barrido(df_reservas, log=_log, progress=progress)
    _log(f"[2/3] Barrido completado. Reservas procesadas: {len(df_resultado)}")

    # ----------------------------------------------------------
    # PASO 3: Guardar CSV final
    # ----------------------------------------------------------
    _log("[3/3] Guardando archivo final...")
    os.makedirs(BARRIDOS_DIR, exist_ok=True)
    df_resultado.to_csv(csv_final, index=False, encoding="utf-8-sig")
    _log(f"[3/3] Archivo guardado: {csv_final}")

    os.startfile(csv_final)
    _log("Proceso completado exitosamente.")

    return df_resultado


# ============================================================
# EJECUCION DIRECTA (sin GUI)
# ============================================================

if __name__ == "__main__":
    ejecutar_proceso_completo(days_back=1)
