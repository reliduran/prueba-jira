import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd

# =============================================================================
# 1. CONFIGURACIÓN DE MAPEO
# =============================================================================
COLUMNAS_INTERNAS = [
    "Empresa",
    "Centro_Costo",
    "Nombre",
    "Vacio_1",
    "NIT_DUI",
    "Codigo_Ingreso",
    "Sueldo_Devengado",
    "Bonificaciones",
    "Retencion_Renta",
    "Aguinaldo_Exento",
    "Aguinaldo_Gravado",
    "AFP",
    "ISSS",
    "Col_13",
    "Col_14",
    "Col_15",
    "Col_16",
    "Col_17",
    "Control_1",
    "Control_2",
    "Control_3",
    "Control_4",
    "Fecha_Archivo",
]

MAPEO_FINAL = {
    "Empresa": "domiciliado",
    "Centro_Costo": "CODIGO DE PAIS",
    "Nombre": "APELLIDO NOMBRE",
    "Vacio_1": "Nit",
    "NIT_DUI": "Dui",
    "Codigo_Ingreso": "Codigo ingreso",
    "Sueldo_Devengado": "Monto Devengado",
    "Bonificaciones": "Monto devengado por bono etc",
    "Retencion_Renta": "impuesto retenido",
    "Aguinaldo_Exento": "Aguinaldo Exento",
    "Aguinaldo_Gravado": "Aguinaldo Gravado",
    "AFP": "AFP",
    "ISSS": "ISSS",
    "Col_13": "INPEP",
    "Col_14": "IPSFA",
    "Col_15": "CEFAFA",
    "Col_16": "BIENESTAR MAGISTERIAL",
    "Col_17": "ISSS IVM",
    "Control_1": "TIPO OPERAC",
    "Control_2": "CLASIFICACION",
    "Control_3": "SECTOR",
    "Control_4": "TIPO COSTO/GTO",
    "Fecha_Archivo": "PERIODO",
}


# =============================================================================
# 2. FUNCIONES
# =============================================================================
def mostrar_guia_adaptacion():
    vent_ayuda = tk.Toplevel()
    vent_ayuda.title("Guía de Adaptación de Columnas")
    vent_ayuda.geometry("500x600")

    # Truco para que la ventana de ayuda también salga al frente
    vent_ayuda.lift()
    vent_ayuda.focus_force()

    tk.Label(
        vent_ayuda,
        text="Adaptación Automática de Encabezados",
        font=("Arial", 12, "bold"),
    ).pack(pady=10)
    tk.Label(
        vent_ayuda,
        text="Transformación de datos crudos a formato Excel:",
        wraplength=480,
    ).pack(pady=5)

    frame_tabla = tk.Frame(vent_ayuda)
    frame_tabla.pack(fill="both", expand=True, padx=10, pady=10)

    tree = ttk.Treeview(
        frame_tabla, columns=("Original", "Adaptado"), show="headings", height=20
    )
    tree.heading("Original", text="Nombre Original")
    tree.heading("Adaptado", text="Nombre Excel Final")
    tree.column("Original", width=200)
    tree.column("Adaptado", width=250)

    for k, v in MAPEO_FINAL.items():
        tree.insert("", "end", values=(k, v))

    tree.pack(side="left", fill="both", expand=True)
    ttk.Scrollbar(frame_tabla, orient="vertical", command=tree.yview).pack(
        side="right", fill="y"
    )

    tk.Button(vent_ayuda, text="Cerrar", command=vent_ayuda.destroy).pack(pady=10)


def procesar_archivos():
    rutas_archivos = filedialog.askopenfilenames(
        title="Selecciona Reportes Mensuales (TXT/CSV)",
        filetypes=[("Archivos de Datos", "*.txt *.csv"), ("Todos", "*.*")],
    )
    if not rutas_archivos:
        return

    lista_dfs = []
    for ruta in rutas_archivos:
        try:
            df_temp = pd.read_csv(
                ruta,
                sep=";",
                header=None,
                names=COLUMNAS_INTERNAS,
                dtype=str,
                encoding="latin-1",
                on_bad_lines="skip",
            )
            lista_dfs.append(df_temp)
        except Exception as e:
            messagebox.showerror(
                "Error", f"Fallo al leer {os.path.basename(ruta)}:\n{str(e)}"
            )

    if not lista_dfs:
        return

    df_final = pd.concat(lista_dfs, ignore_index=True)

    # Limpieza
    cols_dinero = [
        "Sueldo_Devengado",
        "Retencion_Renta",
        "Aguinaldo_Exento",
        "Aguinaldo_Gravado",
        "AFP",
        "ISSS",
        "Bonificaciones",
    ]
    for col in cols_dinero:
        if col in df_final.columns:
            df_final[col] = df_final[col].astype(str).str.replace(",", "", regex=False)

    df_final.rename(columns=MAPEO_FINAL, inplace=True)

    ruta_guardado = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel", "*.xlsx")],
        title="Guardar Base Consolidada",
        initialfile="Base_Datos_RRHH_Adaptada.xlsx",
    )

    if ruta_guardado:
        try:
            df_final.to_excel(ruta_guardado, index=False)
            messagebox.showinfo("Éxito", f"Archivo generado en:\n{ruta_guardado}")
        except Exception as e:
            messagebox.showerror("Error al guardar", str(e))


# =============================================================================
# 3. INTERFAZ GRÁFICA (CON AUTO-ENFOQUE)
# =============================================================================
root = tk.Tk()
root.title("Conversor de Nómina F910 - Módulo 1")
root.geometry("400x300")

# --- BLOQUE PARA FORZAR QUE LA VENTANA SE ABRA ENCIMA ---
root.lift()  # Levantar ventana
root.attributes("-topmost", True)  # Mantener siempre encima (temporalmente)
root.after_idle(
    root.attributes, "-topmost", False
)  # Soltar el 'siempre encima' para permitir usar otras ventanas
root.focus_force()  # Forzar el foco del teclado
# --------------------------------------------------------

tk.Label(root, text="Módulo de Extracción y Estandarización", font=("Arial", 14)).pack(
    pady=20
)

btn_procesar = tk.Button(
    root,
    text="Seleccionar Archivos y Convertir",
    command=procesar_archivos,
    height=2,
    bg="#e1f5fe",
)
btn_procesar.pack(pady=10, fill="x", padx=50)

btn_info = tk.Button(
    root, text="ℹ Ver Guía de Adaptación", command=mostrar_guia_adaptacion, bg="#fff9c4"
)
btn_info.pack(pady=5, fill="x", padx=50)

btn_salir = tk.Button(root, text="Cerrar Módulo", command=root.destroy, width=20)
btn_salir.pack(pady=20)

root.mainloop()
