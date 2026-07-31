"""
barrido_wizard.py
Responsabilidad: ejecutar el barrido de reservas en BlueZone/Wizard via win32com.
Recibe un DataFrame con columna 'Reservation', retorna un DataFrame con columnas:
Reserva, QUOTED_RATE, Rate, Currency, Fecha_Consulta.

Parametros de reconexion hardcodeados (version GUI — usuario presente):
  REINTENTOS_TERMINAL = 3
  ESPERA_TERMINAL_SEG = 30
"""

import re
import time
from datetime import datetime

import pandas as pd
import win32com.client


# ============================================================
# CONSTANTES
# ============================================================

REINTENTOS_TERMINAL = 3
ESPERA_TERMINAL_SEG = 30


# ============================================================
# UTILIDADES DE CONEXION A LA TERMINAL
# ============================================================

def _verificar_y_reconectar(bzhao, log=None) -> None:
    """
    Verifica si la terminal BlueZone esta conectada.
    Si no lo esta, intenta reconectar hasta REINTENTOS_TERMINAL veces
    con ESPERA_TERMINAL_SEG segundos entre cada intento.
    Lanza ConnectionError si agota los intentos.
    """
    if bzhao.Connected:
        return

    def _log(msg):
        if log:
            log(msg)

    _log("Terminal BlueZone desconectada. Intentando reconectar...")

    for intento in range(1, REINTENTOS_TERMINAL + 1):
        try:
            bzhao.connect("")
        except Exception as e:
            _log(f"Intento {intento}/{REINTENTOS_TERMINAL} - Error al reconectar: {e}")

        if bzhao.Connected:
            _log(f"Terminal reconectada exitosamente en el intento {intento}.")
            return

        _log(
            f"Intento {intento}/{REINTENTOS_TERMINAL}: sin conexion. "
            f"Esperando {ESPERA_TERMINAL_SEG}s..."
        )
        time.sleep(ESPERA_TERMINAL_SEG)

    msg = (
        f"La terminal BlueZone no pudo reconectarse tras "
        f"{REINTENTOS_TERMINAL} intentos."
    )
    if log:
        log(f"ERROR: {msg}")
    raise ConnectionError(msg)


# ============================================================
# UTILIDADES DE LECTURA DE PANTALLA
# ============================================================

def _encuentraDatos(bzhao) -> tuple:
    """
    Lee las 25 lineas de la pantalla actual de Wizard y extrae:
      - amt      : Quoted Rate (PER IS / PER=)
      - rate     : Rate Selected
      - currency : Currency
    Retorna: (amt, rate, currency) como strings (vacios si no se encuentran).
    """
    patron_rate     = r'RATE SELECTED = (\S+)'
    patron_currency = r'CURRENCY = (\S+)'
    patron_quoted1  = r'PER IS\s+(\S+)'
    patron_quoted2  = r'PER=\s+(\S+)'

    amt = rate = currency = texto = ''

    for j in range(25):
        linea = bzhao.ReadScreen("", 44, 2 + j, 37)[1].strip()
        texto = texto + ' ' + linea

    m = re.search(patron_quoted1, texto)
    if m:
        amt = m.group(1)

    m = re.search(patron_quoted2, texto)
    if m:
        amt = m.group(1)

    m = re.search(patron_rate, texto)
    if m:
        rate = m.group(1)

    m = re.search(patron_currency, texto)
    if m:
        currency = m.group(1)

    return amt, rate, currency


def _validaDatos(amt, rate, currency, amt2, rate2, currency2) -> tuple:
    """
    Consolida dos lecturas de pantalla.
    Los valores del segundo conjunto tienen prioridad si no estan vacios.
    """
    if amt2 != '':
        amt = amt2
    if rate2 != '':
        rate = rate2
    if currency2 != '':
        currency = currency2
    return amt, rate, currency


# ============================================================
# FUNCION PRINCIPAL DEL BARRIDO
# ============================================================

def ejecutar_barrido(df_reservas: pd.DataFrame, log=None, progress=None) -> pd.DataFrame:
    """
    Ejecuta el barrido P502 en BlueZone para cada reserva del DataFrame.

    Flujo por reserva:
      1. Verifica conexion de terminal (reconecta si es necesario)
      2. Navega a /for p502
      3. Consulta DR (Display Reservation)
      4. Consulta QR (Quoted Rate)
      5. Si hay pantalla adicional (mod=1), navega a ella
      6. Si hay MORE, presiona PA1 para paginar

    Parametros:
        df_reservas : DataFrame con columna 'Reservation'
        log         : Funcion callback para mensajes importantes (inicio, errores, fin)
        progress    : Funcion callback(actual, total) para actualizar la barra de progreso

    Retorna:
        DataFrame con columnas: Reserva, QUOTED_RATE, Rate, Currency, Fecha_Consulta
    """
    def _log(msg):
        if log:
            log(msg)

    lsreserva = list(df_reservas["Reservation"])
    fecha_hoy = datetime.today().strftime('%Y-%m-%d')
    total     = len(lsreserva)

    _log("Inicializando conexion con BlueZone/Wizard...")

    bzhao = win32com.client.Dispatch("BZWhll.WhllObj")

    try:
        bzhao.connect("")
    except Exception as e:
        _log(f"Advertencia al conectar inicialmente: {e}")

    _verificar_y_reconectar(bzhao, log=_log)
    _log(f"Terminal conectada. Iniciando barrido de {total} reservas...")

    L_reserva  = []
    L_amt      = []
    L_rate     = []
    L_currency = []

    for idx, reserva in enumerate(lsreserva, start=1):
        # Actualizar barra de progreso silenciosamente (sin llenar el log)
        if progress:
            progress(idx, total)

        _verificar_y_reconectar(bzhao, log=_log)

        amt = rate = currency = ''

        try:
            # Navegar a pantalla P502
            bzhao.sendkey("<Clear>")
            bzhao.WaitReady(0, 0.00001)
            bzhao.sendkey("/for p502")
            bzhao.WaitReady(0, 0.00001)
            bzhao.SendKey("<Enter>")
            bzhao.WaitReady(0, 0.00001)

            # DR - Display Reservation
            bzhao.WriteScreen("DR", 2, 2)
            bzhao.WriteScreen("R/" + reserva, 11, 6)
            bzhao.SendKey("<Enter>")
            bzhao.WaitReady(0, 0.00001)

            amt2, rate2, currency2 = _encuentraDatos(bzhao)
            amt, rate, currency    = _validaDatos(amt, rate, currency, amt2, rate2, currency2)

            # QR - Quoted Rate
            bzhao.WriteScreen("qr", 2, 2)
            bzhao.SendKey("<Enter>")
            bzhao.WaitReady(0, 0.00001)

            amt2, rate2, currency2 = _encuentraDatos(bzhao)
            amt, rate, currency    = _validaDatos(amt, rate, currency, amt2, rate2, currency2)

            # Verificar si hay modulo 1 adicional
            mod1 = bzhao.ReadScreen("", 1, 2, 37)[1].strip()
            if re.search(r'1\b', mod1):
                bzhao.WriteScreen("                               ", 11, 6)
                bzhao.WriteScreen("1", 11, 6)
                bzhao.SendKey("<Enter>")
                bzhao.WaitReady(0, 0.00001)

                amt2, rate2, currency2 = _encuentraDatos(bzhao)
                amt, rate, currency    = _validaDatos(amt, rate, currency, amt2, rate2, currency2)

            # Verificar paginacion (MORE)
            more = bzhao.ReadScreen("", 9, 23, 72)[1].strip()
            if re.search(r'MORE\b', more):
                bzhao.SendKey("<PA1>")
                bzhao.WaitReady(0, 0.00001)

                amt2, rate2, currency2 = _encuentraDatos(bzhao)
                amt, rate, currency    = _validaDatos(amt, rate, currency, amt2, rate2, currency2)

        except ConnectionError:
            raise

        except Exception as e:
            _log(f"Error al procesar reserva {reserva}: {e}")

        L_reserva.append(reserva)
        L_amt.append(amt)
        L_rate.append(rate)
        L_currency.append(currency)

    _log(f"Barrido completado. Total procesadas: {len(L_reserva)}")

    return pd.DataFrame({
        "Reserva":        L_reserva,
        "QUOTED_RATE":    L_amt,
        "Rate":           L_rate,
        "Currency":       L_currency,
        "Fecha_Consulta": fecha_hoy,
    })
