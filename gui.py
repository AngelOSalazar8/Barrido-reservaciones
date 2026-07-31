import tkinter as tk
from tkinter import ttk, messagebox
import threading
import sv_ttk
from modulos import SQLtoCSV as main
from datetime import datetime


# ============================================================
# CONSTANTES DE MODO
# ============================================================

MODO_DIA_ANTERIOR = "dia_anterior"
MODO_FIN_SEMANA   = "fin_semana"
MODO_N_DIAS       = "n_dias"


# ============================================================
# FUNCIONES DE LOGICA
# ============================================================

def log_message(msg):
    """
    Inserta un mensaje con timestamp en el log de forma segura desde
    cualquier hilo usando root.after para no bloquear la GUI.
    """
    root.after(0, _log_insert, msg)


def _log_insert(msg):
    """Funcion interna que escribe en el widget de texto (solo desde el hilo principal)."""
    timestamp = datetime.now().strftime("%H:%M:%S")
    log_box.configure(state="normal")
    log_box.insert(tk.END, f"{timestamp} - {msg}\n")
    log_box.see(tk.END)
    log_box.configure(state="disabled")


def update_progreso(actual, total):
    """
    Actualiza la barra de progreso y el contador de reservas.
    Llamado desde el hilo del barrido mediante root.after.
    """
    root.after(0, _progreso_insert, actual, total)


def _progreso_insert(actual, total):
    """Funcion interna que modifica los widgets de progreso (solo hilo principal)."""
    progreso_bar["maximum"] = total
    progreso_bar["value"]   = actual
    lbl_progreso.configure(text=f"{actual} / {total} reservas")


def obtener_days_back():
    """
    Lee la seleccion del usuario y retorna el numero de dias hacia atras.
    Lanza ValueError si el Spinbox tiene un valor invalido en modo N dias.
    """
    modo = modo_var.get()
    if modo == MODO_DIA_ANTERIOR:
        return 1
    elif modo == MODO_FIN_SEMANA:
        return 3
    else:
        try:
            val = int(spinbox_n.get())
            if val < 1:
                raise ValueError("El valor debe ser >= 1")
            return val
        except ValueError:
            raise ValueError("Ingresa un numero entero valido en el campo 'N dias'.")


def describir_modo(days_back):
    """Retorna una descripcion legible del modo seleccionado."""
    modo = modo_var.get()
    if modo == MODO_DIA_ANTERIOR:
        return f"Modo: Dia anterior | Consultando los ultimos {days_back} dia(s)"
    elif modo == MODO_FIN_SEMANA:
        return f"Modo: Fin de semana | Consultando los ultimos {days_back} dias (vie+sab+dom)"
    else:
        return f"Modo: N dias atras | Consultando los ultimos {days_back} dia(s)"


def iniciar_barrido():
    """
    Valida la configuracion de dias, deshabilita el boton y lanza
    el proceso completo en un hilo separado para no congelar la GUI.
    """
    try:
        days_back = obtener_days_back()
    except ValueError as e:
        messagebox.showerror("Error de configuracion", str(e))
        return

    btn_iniciar.configure(state="disabled")
    log_message(describir_modo(days_back))
    log_message("Iniciando proceso...")

    def _tarea():
        try:
            main.ejecutar_proceso_completo(
                days_back=days_back,
                log=log_message,
                progress=update_progreso
            )
            root.after(0, _on_exito)
        except ConnectionError as e:
            root.after(0, _on_error, f"Terminal BlueZone no disponible:\n{e}")
        except Exception as e:
            root.after(0, _on_error, f"Error inesperado:\n{e}")

    threading.Thread(target=_tarea, daemon=True).start()


def _on_exito():
    """Se ejecuta en el hilo principal al terminar exitosamente."""
    btn_iniciar.configure(state="normal")
    lbl_progreso.configure(text="Completado")
    messagebox.showinfo(
        "Proceso completado",
        "El barrido finalizo correctamente.\n\n"
        "El archivo CSV se ha abierto automaticamente desde la carpeta Barridos."
    )


def _on_error(mensaje):
    """Se ejecuta en el hilo principal si ocurre un error."""
    btn_iniciar.configure(state="normal")
    lbl_progreso.configure(text="Error")
    progreso_bar["value"] = 0
    messagebox.showerror("Error en el proceso", mensaje)


# ============================================================
# CALLBACKS DE UI
# ============================================================

def on_modo_change():
    """Habilita o deshabilita el Spinbox segun el modo elegido."""
    if modo_var.get() == MODO_N_DIAS:
        spinbox_n.configure(state="normal")
        lbl_spinbox.configure(foreground="white")
    else:
        spinbox_n.configure(state="disabled")
        lbl_spinbox.configure(foreground="#555555")


# ============================================================
# CONSTRUCCION DE LA INTERFAZ
# ============================================================

root = tk.Tk()
root.title("Barrido Quoted Rate")
root.geometry("640x540")
root.resizable(False, False)

sv_ttk.set_theme("dark")

# --- Titulo ---
ttk.Label(
    root,
    text="Barrido Quoted Rate",
    font=("Segoe UI", 15, "bold")
).pack(pady=(18, 14))

ttk.Separator(root, orient="horizontal").pack(fill="x", padx=20, pady=(0, 12))

# -------------------------------------------------------
# BLOQUE: Selector de rango de dias
# -------------------------------------------------------
frame_modo = ttk.LabelFrame(root, text="  Rango de consulta  ", padding=(16, 10))
frame_modo.pack(fill="x", padx=24, pady=(0, 14))

modo_var = tk.StringVar(value=MODO_DIA_ANTERIOR)

rb_dia = ttk.Radiobutton(
    frame_modo,
    text="Dia anterior  (-1 dia)   — uso normal de lunes a sabado",
    variable=modo_var,
    value=MODO_DIA_ANTERIOR,
    command=on_modo_change
)
rb_dia.grid(row=0, column=0, columnspan=3, sticky="w", pady=3)

rb_fin = ttk.Radiobutton(
    frame_modo,
    text="Fin de semana (-3 dias) — lunes: cubre viernes, sabado y domingo",
    variable=modo_var,
    value=MODO_FIN_SEMANA,
    command=on_modo_change
)
rb_fin.grid(row=1, column=0, columnspan=3, sticky="w", pady=3)

rb_n = ttk.Radiobutton(
    frame_modo,
    text="N dias atras            — puentes u otros casos especiales:",
    variable=modo_var,
    value=MODO_N_DIAS,
    command=on_modo_change
)
rb_n.grid(row=2, column=0, sticky="w", pady=3)

spinbox_n = ttk.Spinbox(
    frame_modo,
    from_=1,
    to=30,
    width=5,
    state="disabled"
)
spinbox_n.set(4)
spinbox_n.grid(row=2, column=1, padx=(8, 4), pady=3, sticky="w")

lbl_spinbox = ttk.Label(frame_modo, text="dias", foreground="#555555")
lbl_spinbox.grid(row=2, column=2, sticky="w")

# -------------------------------------------------------
# BLOQUE: Boton principal + barra de progreso
# -------------------------------------------------------
ttk.Separator(root, orient="horizontal").pack(fill="x", padx=20, pady=(0, 14))

btn_iniciar = ttk.Button(
    root,
    text="Iniciar Barrido Completo",
    command=iniciar_barrido,
    style="Accent.TButton"
)
btn_iniciar.pack(pady=4, ipadx=30, ipady=8)

ttk.Label(
    root,
    text="SQL   ->   BlueZone P502   ->   CSV final en /Barridos/",
    font=("Segoe UI", 8),
    foreground="#888888"
).pack(pady=(4, 6))

# Barra de progreso del barrido
frame_progreso = ttk.Frame(root)
frame_progreso.pack(fill="x", padx=24, pady=(0, 2))

progreso_bar = ttk.Progressbar(
    frame_progreso,
    mode="determinate",
    length=470
)
progreso_bar.pack(side="left", padx=(0, 10))

lbl_progreso = ttk.Label(
    frame_progreso,
    text="En espera...",
    font=("Segoe UI", 8),
    foreground="#888888",
    width=18
)
lbl_progreso.pack(side="left")

# -------------------------------------------------------
# BLOQUE: Registro de acciones (log)
# -------------------------------------------------------
ttk.Separator(root, orient="horizontal").pack(fill="x", padx=20, pady=(14, 8))

ttk.Label(
    root,
    text="Registro de acciones:",
    font=("Segoe UI", 9, "bold")
).pack(anchor="w", padx=24)

log_box = tk.Text(
    root,
    height=10,
    width=74,
    wrap="word",
    state="disabled",
    font=("Consolas", 9),
    relief="flat",
    borderwidth=0
)
log_box.pack(padx=24, pady=(4, 0))




# ============================================================
# INICIO
# ============================================================

root.mainloop()
