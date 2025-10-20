import tkinter as tk
from tkinter import ttk, messagebox
import ttkbootstrap as tb
import pandas as pd
from pathlib import Path

# Configuracion de rutas y hojas: Todas las hojas estan en el mismo archivo, revisar si es necesario cambiar porfa
RUTA_DATOS = Path("Ventas.xlsx")    
HOJA_INVENTARIO = "Inventario"

def inicializar_excel():
    pd.read_excel(RUTA_DATOS, sheet_name=HOJA_INVENTARIO)

def leer_inventario() -> pd.DataFrame:
    inicializar_excel()
    return pd.read_excel(RUTA_DATOS, sheet_name=HOJA_INVENTARIO, dtype={"codigo": str})

def escribir_inventario(df: pd.DataFrame):
    inicializar_excel()
    with pd.ExcelWriter(RUTA_DATOS, engine="openpyxl", mode="a", if_sheet_exists="replace") as libro:
        df.to_excel(libro, sheet_name=HOJA_INVENTARIO, index=False)

# Operaciones del inventario
def listar_productos(tabla: ttk.Treeview):
    df = leer_inventario()
    tabla.delete(*tabla.get_children())
    for _, fila in df.iterrows():
        tabla.insert("", "end", values=(fila.codigo, fila.nombre, fila.existencia, fila.proveedor, fila.precio))

def crear_producto(codigo, nombre, existencia, proveedor, precio, tabla):
    try:
        codigo = str(codigo).strip()
        nombre = str(nombre).strip()
        proveedor = str(proveedor).strip()
        existencia = int(existencia)
        precio = float(precio)

        if not codigo or not nombre:
            raise ValueError("Código y nombre son obligatorios.")

        df = leer_inventario()
        if (df["codigo"] == codigo).any():
            raise ValueError(f"Ya existe un producto con código {codigo}.")

        nuevo_producto = pd.DataFrame([{
            "codigo": codigo,
            "nombre": nombre,
            "existencia": existencia,
            "proveedor": proveedor,
            "precio": precio
        }])
        df = pd.concat([df, nuevo_producto], ignore_index=True)
        escribir_inventario(df)
        listar_productos(tabla)
        messagebox.showinfo("Éxito", "Producto creado correctamente.")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def actualizar_producto(tabla, nombre, existencia, proveedor, precio):
    try:
        seleccion = tabla.focus()
        if not seleccion:
            raise ValueError("Selecciona un producto en la tabla.")
        valores = tabla.item(seleccion, "values")
        codigo_sel = valores[0]

        df = leer_inventario()
        indice = df.index[df["codigo"].astype(str) == str(codigo_sel)]
        if indice.empty:
            raise ValueError("No se encontró el producto en el archivo.")

        df.loc[indice, ["nombre", "existencia", "proveedor", "precio"]] = [
            str(nombre).strip(),
            int(existencia),
            str(proveedor).strip(),
            float(precio)
        ]
        escribir_inventario(df)
        listar_productos(tabla)
        messagebox.showinfo("Éxito", "Producto actualizado correctamente.")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def eliminar_producto(tabla):
    try:
        seleccion = tabla.focus()
        if not seleccion:
            raise ValueError("Selecciona un producto en la tabla.")
        codigo_sel = tabla.item(seleccion, "values")[0]

        if not messagebox.askyesno("Confirmar", f"¿Eliminar producto {codigo_sel}?"):
            return

        df = leer_inventario()
        df["codigo"] = df["codigo"].astype(str)  
        df = df[df["codigo"] != str(codigo_sel)]
        escribir_inventario(df)
        listar_productos(tabla)
        messagebox.showinfo("Éxito", "Producto eliminado correctamente.")
    except Exception as e:
        messagebox.showerror("Error", str(e))

# Interfaz, creacion de tablas y botones de acciones
def crear_tabla(parent, columnas, encabezados):
    tabla = ttk.Treeview(parent, columns=columnas, show="headings", height=12)
    for c, h in zip(columnas, encabezados):
        tabla.heading(c, text=h)
        tabla.column(c, width=120, anchor="center")
    barra = ttk.Scrollbar(parent, orient="vertical", command=tabla.yview)
    tabla.configure(yscrollcommand=barra.set)
    tabla.grid(row=0, column=0, sticky="nsew")
    barra.grid(row=0, column=1, sticky="ns")
    parent.grid_rowconfigure(0, weight=1)
    parent.grid_columnconfigure(0, weight=1)
    return tabla

def pestaña_inventario(notebook):
    pestaña = ttk.Frame(notebook)
    notebook.add(pestaña, text="Inventario")

    # Tabla de productos
    contenedor_tabla = ttk.Frame(pestaña)
    contenedor_tabla.pack(fill="both", expand=True, padx=10, pady=8)
    columnas = ("codigo", "nombre", "existencia", "proveedor", "precio")
    tabla = crear_tabla(contenedor_tabla, columnas, ["Código", "Nombre", "Existencia", "Proveedor", "Precio"])
    listar_productos(tabla)

    # Formulario de creación
    formulario = ttk.Frame(pestaña)
    formulario.pack(fill="x", padx=10, pady=8)

    entrada_codigo = tb.Entry(formulario, width=12)
    entrada_nombre = tb.Entry(formulario, width=24)
    entrada_existencia = tb.Entry(formulario, width=8)
    entrada_proveedor = tb.Entry(formulario, width=20)
    entrada_precio = tb.Entry(formulario, width=10)

    tb.Label(formulario, text="Código:").grid(row=0, column=0, padx=4, pady=4, sticky="e"); entrada_codigo.grid(row=0, column=1, padx=4, pady=4)
    tb.Label(formulario, text="Nombre:").grid(row=0, column=2, padx=4, pady=4, sticky="e"); entrada_nombre.grid(row=0, column=3, padx=4, pady=4)
    tb.Label(formulario, text="Existencia:").grid(row=0, column=4, padx=4, pady=4, sticky="e"); entrada_existencia.grid(row=0, column=5, padx=4, pady=4)
    tb.Label(formulario, text="Proveedor:").grid(row=1, column=0, padx=4, pady=4, sticky="e"); entrada_proveedor.grid(row=1, column=1, padx=4, pady=4)
    tb.Label(formulario, text="Precio:").grid(row=1, column=2, padx=4, pady=4, sticky="e"); entrada_precio.grid(row=1, column=3, padx=4, pady=4)

    # edicion de productos
    def limpiar_formulario():
        for e in (entrada_codigo, entrada_nombre, entrada_existencia, entrada_proveedor, entrada_precio):
            e.configure(state="normal")
            e.delete(0, tk.END)

    def modo_creacion():
        limpiar_formulario()
        try:
            tabla.selection_remove(*tabla.selection())
        except Exception:
            pass
        entrada_codigo.focus_set()

    def modo_edicion(valores):
        limpiar_formulario()
        entrada_codigo.insert(0, valores[0]); entrada_codigo.configure(state="disabled")
        entrada_nombre.insert(0, valores[1])
        entrada_existencia.insert(0, valores[2])
        entrada_proveedor.insert(0, valores[3])
        entrada_precio.insert(0, valores[4])

    # botoncitos
    tb.Button(
        formulario, text="Crear", bootstyle="success",
        command=lambda: (
            crear_producto(
                entrada_codigo.get(), entrada_nombre.get(), entrada_existencia.get(),
                entrada_proveedor.get(), entrada_precio.get(), tabla
            ),
            # limpiar formulario y volver a modo creacion
            modo_creacion()
        )
    ).grid(row=2, column=1, pady=6)

    tb.Button(
        formulario, text="Actualizar selección", bootstyle="info",
        command=lambda: (
            actualizar_producto(
                tabla, entrada_nombre.get(), entrada_existencia.get(),
                entrada_proveedor.get(), entrada_precio.get()
            ),
            modo_creacion()
        )
    ).grid(row=2, column=3, pady=6)

    tb.Button(
        formulario, text="Eliminar selección", bootstyle="danger",
        command=lambda: (eliminar_producto(tabla), modo_creacion())
    ).grid(row=2, column=5, pady=6)

    tb.Button(
        formulario, text="Refrescar", bootstyle="secondary",
        command=lambda: (listar_productos(tabla), modo_creacion())
    ).grid(row=2, column=6, pady=6)

    # Cuando seleccionas en la tabla pasmaos al modo edicion
    def cargar_en_formulario(_evento):
        seleccion = tabla.focus()
        if not seleccion:
            return
        valores = tabla.item(seleccion, "values")
        modo_edicion(valores)

    tabla.bind("<<TreeviewSelect>>", cargar_en_formulario)

    # Al abrir la pestaña, empezamos en modo creación
    modo_creacion()

    return pestaña


# para el menu principal
def abrir_algoritmos(padre=None):
    """Abre la ventana del módulo de Inventario."""
    is_root = padre is None
    ventana = tb.Window(themename="superhero") if is_root else tb.Toplevel(padre)
    ventana.title("Módulo de Inventario")
    ventana.geometry("900x600")

    tb.Label(ventana, text="Control de Inventario", font=("Arial", 20, "bold")).pack(pady=10)

    cuaderno = tb.Notebook(ventana)
    cuaderno.pack(expand=True, fill="both", padx=10, pady=10)

    pestaña_inventario(cuaderno)

    if is_root:
        ventana.mainloop()