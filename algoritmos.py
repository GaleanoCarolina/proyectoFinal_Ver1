import tkinter as tk 
from tkinter import ttk, messagebox
import ttkbootstrap as tb
import pandas as pd
from pathlib import Path
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

RUTA_DATOS = Path("Ventas.xlsx")
HOJA_INVENTARIO = "Inventario"
HOJA_CLIENTES   = "Clientes"
HOJA_VENTAS     = "Ventas"

def crear_archivo_excel_si_no_existe():
    """Crea Ventas.xlsx con las hojas necesarias si no existe."""
    if not RUTA_DATOS.exists():
        with pd.ExcelWriter(RUTA_DATOS, engine="openpyxl") as writer:
            pd.DataFrame(columns=["codigo", "nombre", "existencia", "proveedor", "precio"])\
                .to_excel(writer, sheet_name=HOJA_INVENTARIO, index=False)
            pd.DataFrame(columns=["codigo", "nombre", "direccion"])\
                .to_excel(writer, sheet_name=HOJA_CLIENTES, index=False)
            # Hoja Ventas con columnas estándar + alias para reportes
            pd.DataFrame(columns=[
                "id","fecha","producto","cliente","cantidad","precio_unit","total","anulada",
                "Producto","Cliente"
            ]).to_excel(writer, sheet_name=HOJA_VENTAS, index=False)
        print(f"Creado {RUTA_DATOS}")

def inicializar_excel():
    """Valida que exista el archivo y al menos la hoja Inventario; si no, lo crea."""
    try:
        if not RUTA_DATOS.exists():
            crear_archivo_excel_si_no_existe()
        else:
            pd.read_excel(RUTA_DATOS, sheet_name=HOJA_INVENTARIO, dtype={"codigo": str})
    except Exception:
        crear_archivo_excel_si_no_existe()


# INVENTARIO

def leer_inventario() -> pd.DataFrame:
    inicializar_excel()
    return pd.read_excel(RUTA_DATOS, sheet_name=HOJA_INVENTARIO, dtype={"codigo": str})

def escribir_inventario(df: pd.DataFrame):
    inicializar_excel()
    with pd.ExcelWriter(RUTA_DATOS, engine="openpyxl", mode="a", if_sheet_exists="replace") as libro:
        df.to_excel(libro, sheet_name=HOJA_INVENTARIO, index=False)

def listar_productos(tabla: ttk.Treeview):
    df = leer_inventario()
    tabla.delete(*tabla.get_children())
    for _, fila in df.iterrows():
        tabla.insert("", "end", values=(fila["codigo"], fila["nombre"], fila["existencia"], fila["proveedor"], fila["precio"]))

def crear_producto(codigo, nombre, existencia, proveedor, precio, tabla):
    try:
        codigo = str(codigo).strip()
        nombre = str(nombre).strip()
        proveedor = str(proveedor).strip()
        existencia = int(float(existencia)) if str(existencia).strip() != "" else 0
        precio = float(precio) if str(precio).strip() != "" else 0.0

        if not codigo or not nombre:
            raise ValueError("Código y nombre son obligatorios.")

        df = leer_inventario()
        if (df["codigo"].astype(str) == codigo).any():
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
            int(float(existencia)),
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

    contenedor_tabla = ttk.Frame(pestaña)
    contenedor_tabla.pack(fill="both", expand=True, padx=10, pady=8)
    columnas = ("codigo", "nombre", "existencia", "proveedor", "precio")
    tabla = crear_tabla(contenedor_tabla, columnas, ["Código", "Nombre", "Existencia", "Proveedor", "Precio"])
    listar_productos(tabla)
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

    tb.Button(
        formulario, text="Crear", bootstyle="success",
        command=lambda: (
            crear_producto(
                entrada_codigo.get(), entrada_nombre.get(), entrada_existencia.get(),
                entrada_proveedor.get(), entrada_precio.get(), tabla
            )
        )
    ).grid(row=2, column=1, pady=6)

    tb.Button(
        formulario, text="Actualizar selección", bootstyle="info",
        command=lambda: (
            actualizar_producto(
                tabla, entrada_nombre.get(), entrada_existencia.get(),
                entrada_proveedor.get(), entrada_precio.get()
            )
        )
    ).grid(row=2, column=3, pady=6)

    tb.Button(
        formulario, text="Eliminar selección", bootstyle="danger",
        command=lambda: (eliminar_producto(tabla))
    ).grid(row=2, column=5, pady=6)

    tb.Button(
        formulario, text="Refrescar", bootstyle="secondary",
        command=lambda: (listar_productos(tabla))
    ).grid(row=2, column=6, pady=6)

    pestaña.bind("<Visibility>", lambda _e: listar_productos(tabla))

    return pestaña


# CLIENTES

def leer_clientes() -> pd.DataFrame:
    """Lee la hoja Clientes y normaliza columnas."""
    inicializar_excel()
    try:
        df = pd.read_excel(RUTA_DATOS, sheet_name=HOJA_CLIENTES, dtype={"codigo": str})
        df.columns = [c.lower().strip() for c in df.columns]
        for col in ["codigo", "nombre", "direccion"]:
            if col not in df.columns:
                df[col] = ""
        return df[["codigo", "nombre", "direccion"]]
    except Exception as e:
        print("Error al leer clientes:", e)
        return pd.DataFrame(columns=["codigo", "nombre", "direccion"])

def escribir_clientes(df: pd.DataFrame):
    """Escribe la hoja Clientes en el archivo sin tocar otras hojas."""
    inicializar_excel()
    try:
        with pd.ExcelWriter(RUTA_DATOS, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df.to_excel(writer, sheet_name=HOJA_CLIENTES, index=False)
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo escribir Clientes: {e}")

def listar_clientes(tabla: ttk.Treeview):
    df = leer_clientes()
    tabla.delete(*tabla.get_children())
    for _, fila in df.iterrows():
        tabla.insert("", "end", values=(fila["codigo"], fila["nombre"], fila["direccion"]))

def crear_cliente(codigo, nombre, direccion, tabla):
    try:
        codigo = str(codigo).strip()
        nombre = str(nombre).strip()
        direccion = str(direccion).strip()

        if not codigo or not nombre or not direccion:
            raise ValueError("Código, Nombre y Dirección son obligatorios.")

        df = leer_clientes()
        if (df["codigo"].astype(str) == codigo).any():
            raise ValueError(f"Ya existe un cliente con código {codigo}.")

        nuevo = pd.DataFrame([{"codigo": codigo, "nombre": nombre, "direccion": direccion}])
        df = pd.concat([df, nuevo], ignore_index=True)
        escribir_clientes(df)
        listar_clientes(tabla)
        messagebox.showinfo("Éxito", "Cliente agregado correctamente.")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def actualizar_cliente(tabla, nombre, direccion):
    try:
        seleccion = tabla.focus()
        if not seleccion:
            raise ValueError("Selecciona un cliente.")
        codigo_sel = tabla.item(seleccion, "values")[0]

        df = leer_clientes()
        idx = df.index[df["codigo"].astype(str) == str(codigo_sel)]
        if idx.empty:
            raise ValueError("No se encontró el cliente en el archivo.")

        if not nombre.strip() or not direccion.strip():
            raise ValueError("Nombre y Dirección son obligatorios.")

        df.loc[idx, ["nombre", "direccion"]] = [str(nombre).strip(), str(direccion).strip()]
        escribir_clientes(df)
        listar_clientes(tabla)
        messagebox.showinfo("Éxito", "Cliente actualizado correctamente.")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def eliminar_cliente(tabla):
    try:
        seleccion = tabla.focus()
        if not seleccion:
            raise ValueError("Selecciona un cliente.")
        codigo_sel = tabla.item(seleccion, "values")[0]

        if not messagebox.askyesno("Confirmar", f"¿Eliminar cliente {codigo_sel}?"):
            return

        df = leer_clientes()
        df = df[df["codigo"].astype(str) != str(codigo_sel)]
        escribir_clientes(df)
        listar_clientes(tabla)
        messagebox.showinfo("Éxito", "Cliente eliminado correctamente.")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def pestaña_clientes(notebook):
    pestaña = ttk.Frame(notebook)
    notebook.add(pestaña, text="Clientes")

    contenedor_tabla = ttk.Frame(pestaña)
    contenedor_tabla.pack(fill="both", expand=True, padx=10, pady=8)
    columnas = ("codigo", "nombre", "direccion")
    tabla = crear_tabla(contenedor_tabla, columnas, ["Código", "Nombre", "Dirección"])

    formulario = ttk.Frame(pestaña)
    formulario.pack(fill="x", padx=10, pady=8)

    entrada_codigo = tb.Entry(formulario, width=15)
    entrada_nombre = tb.Entry(formulario, width=30)
    entrada_direccion = tb.Entry(formulario, width=40)

    tb.Label(formulario, text="Código:").grid(row=0, column=0, padx=4, pady=4, sticky="e"); entrada_codigo.grid(row=0, column=1, padx=4, pady=4)
    tb.Label(formulario, text="Nombre:").grid(row=0, column=2, padx=4, pady=4, sticky="e"); entrada_nombre.grid(row=0, column=3, padx=4, pady=4)
    tb.Label(formulario, text="Dirección:").grid(row=0, column=4, padx=4, pady=4, sticky="e"); entrada_direccion.grid(row=0, column=5, padx=4, pady=4)

    tb.Button(formulario, text="Crear", bootstyle="success",
              command=lambda: (crear_cliente(entrada_codigo.get(), entrada_nombre.get(), entrada_direccion.get(), tabla))).grid(row=1, column=1, pady=6)
    tb.Button(formulario, text="Actualizar selección", bootstyle="info",
              command=lambda: (actualizar_cliente(tabla, entrada_nombre.get(), entrada_direccion.get()))).grid(row=1, column=3, pady=6)
    tb.Button(formulario, text="Eliminar selección", bootstyle="danger",
              command=lambda: (eliminar_cliente(tabla))).grid(row=1, column=5, pady=6)
    tb.Button(formulario, text="Refrescar", bootstyle="secondary",
              command=lambda: (listar_clientes(tabla))).grid(row=1, column=6, pady=6)

    listar_clientes(tabla)
    pestaña.bind("<Visibility>", lambda _e: listar_clientes(tabla))

    return pestaña

# VENTAS 

VENTAS_COLUMNS = ["id","fecha","producto","cliente","cantidad","precio_unit","total","anulada"]

def _asegurar_hoja_ventas():
    """Crea/normaliza la hoja Ventas sin tocar Inventario/Clientes."""
    crear_archivo_excel_si_no_existe()
    try:
        libro = pd.ExcelFile(RUTA_DATOS)
        if HOJA_VENTAS not in libro.sheet_names:
            with pd.ExcelWriter(RUTA_DATOS, engine="openpyxl", mode="a") as w:
                pd.DataFrame(columns=VENTAS_COLUMNS + ["Producto", "Cliente"]).to_excel(
                    w, sheet_name=HOJA_VENTAS, index=False
                )
            return

        df = pd.read_excel(RUTA_DATOS, sheet_name=HOJA_VENTAS)
        for col in VENTAS_COLUMNS:
            if col not in df.columns:
                if col in ("cantidad","precio_unit","total"):
                    df[col] = 0
                elif col == "anulada":
                    df[col] = False
                else:
                    df[col] = ""
        if "Producto" not in df.columns:
            df["Producto"] = df["producto"] if "producto" in df.columns else ""
        if "Cliente" not in df.columns:
            df["Cliente"] = df["cliente"] if "cliente" in df.columns else ""
        base = [c for c in (VENTAS_COLUMNS + ["Producto","Cliente"]) if c in df.columns]
        df = df[base + [c for c in df.columns if c not in base]]
        with pd.ExcelWriter(RUTA_DATOS, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
            df.to_excel(w, sheet_name=HOJA_VENTAS, index=False)
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo preparar la hoja Ventas: {e}")

def _ventas_leer_base() -> pd.DataFrame:
    """Lee y normaliza la hoja Ventas."""
    _asegurar_hoja_ventas()
    df = pd.read_excel(RUTA_DATOS, sheet_name=HOJA_VENTAS)
    for col in VENTAS_COLUMNS:
        if col not in df.columns:
            if col in ("cantidad","precio_unit","total"):
                df[col] = 0
            elif col == "anulada":
                df[col] = False
            else:
                df[col] = ""
    df["anulada"] = df["anulada"].fillna(False).astype(bool)
    df["Producto"] = df.get("producto", "").astype(str) if "producto" in df.columns else ""
    df["Cliente"]  = df.get("cliente", "").astype(str)  if "cliente"  in df.columns else ""
    return df

def escribir_ventas(df: pd.DataFrame) -> None:
    """Escribe el DataFrame completo en la hoja Ventas (compatibilidad con reportes)."""
    _asegurar_hoja_ventas()
    for col in VENTAS_COLUMNS:
        if col not in df.columns:
            if col in ("cantidad","precio_unit","total"):
                df[col] = 0
            elif col == "anulada":
                df[col] = False
            else:
                df[col] = ""
    df["Producto"] = df["producto"].astype(str)
    df["Cliente"]  = df["cliente"].astype(str)
    base = [c for c in (VENTAS_COLUMNS + ["Producto","Cliente"]) if c in df.columns]
    out = df[base + [c for c in df.columns if c not in base]]
    with pd.ExcelWriter(RUTA_DATOS, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
        out.to_excel(w, sheet_name=HOJA_VENTAS, index=False)

def calcular_total(cantidad, precio_unit) -> float:
    return float(cantidad) * float(precio_unit)

def validar_existencia(codigo, cantidad) -> bool:
    inv = leer_inventario().copy()
    inv["codigo"] = inv["codigo"].astype(str)
    try:
        cantidad = float(cantidad)
    except Exception:
        return False
    fila = inv.loc[inv["codigo"] == str(codigo)]
    if fila.empty:
        return False
    existencia = float(fila.iloc[0]["existencia"])
    return existencia >= cantidad

def actualizar_stock(codigo: str, cantidad_vendida: float) -> None:
    inv = leer_inventario()
    inv["codigo"] = inv["codigo"].astype(str)
    idx = inv.index[inv["codigo"] == str(codigo)]
    if idx.empty:
        raise ValueError(f"Producto {codigo} no existe en inventario.")
    i = idx[0]
    nuevo = float(inv.at[i, "existencia"]) - float(cantidad_vendida)
    if nuevo < 0:
        raise ValueError("Stock insuficiente.")
    inv.at[i, "existencia"] = nuevo
    escribir_inventario(inv)

def restaurar_stock(codigo: str, cantidad: float) -> None:
    inv = leer_inventario()
    inv["codigo"] = inv["codigo"].astype(str)
    idx = inv.index[inv["codigo"] == str(codigo)]
    if idx.empty:
        raise ValueError(f"Producto {codigo} no existe en inventario.")
    i = idx[0]
    nuevo = float(inv.at[i, "existencia"]) + float(cantidad)
    inv.at[i, "existencia"] = nuevo
    escribir_inventario(inv)

def _siguiente_id(df: pd.DataFrame) -> int:
    if df.empty or ("id" not in df.columns) or df["id"].isna().all():
        return 1
    return int(pd.to_numeric(df["id"], errors="coerce").max()) + 1

def _nombre_producto_por_codigo(codigo: str) -> str:
    inv = leer_inventario()
    inv["codigo"] = inv["codigo"].astype(str)
    fila = inv.loc[inv["codigo"] == str(codigo)]
    if fila.empty:
        raise ValueError(f"Código de producto {codigo} no existe.")
    return str(fila.iloc[0]["nombre"])

def _nombre_cliente_por_codigo(codigo: str) -> str:
    cli = leer_clientes()
    cli["codigo"] = cli["codigo"].astype(str)
    fila = cli.loc[cli["codigo"] == str(codigo)]
    if fila.empty:
        raise ValueError(f"Código de cliente {codigo} no existe.")
    return str(fila.iloc[0]["nombre"])

def _codigo_producto_por_nombre(nombre: str) -> str:
    inv = leer_inventario()
    fila = inv.loc[inv["nombre"] == str(nombre)]
    if fila.empty:
        raise ValueError("Producto no encontrado en inventario para ajuste.")
    return str(fila.iloc[0]["codigo"])

def listar_ventas(tabla: ttk.Treeview):
    df = _ventas_leer_base()
    tabla.delete(*tabla.get_children())
    if df.empty:
        return
    for _, f in df.iterrows():
        tabla.insert("", "end", values=(
            f.get("id", ""),
            f.get("fecha", ""),
            f.get("producto", ""),
            f.get("cliente", ""),
            f.get("cantidad", ""),
            f.get("precio_unit", ""),
            f.get("total", ""),
            "Sí" if bool(f.get("anulada", False)) else "No"
        ))

def crear_venta(producto_codigo, cliente_codigo, cantidad, precio_unit, tabla):
    """Crea venta: valida existencia, descuenta stock, registra total."""
    try:
        cantidad = float(cantidad)
        precio_unit = float(precio_unit)
        if cantidad <= 0 or precio_unit < 0:
            raise ValueError("Cantidad > 0 y precio_unit ≥ 0.")
        if not validar_existencia(producto_codigo, cantidad):
            raise ValueError("No hay existencia suficiente para esta venta.")

        ventas = _ventas_leer_base()
        vid = _siguiente_id(ventas)
        prod_nom = _nombre_producto_por_codigo(producto_codigo)
        cli_nom  = _nombre_cliente_por_codigo(cliente_codigo)
        total = calcular_total(cantidad, precio_unit)
        fecha_txt = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")

        actualizar_stock(producto_codigo, cantidad)  # regla: restar stock

        nueva = {
            "id": vid,
            "fecha": fecha_txt,
            "producto": prod_nom,
            "cliente": cli_nom,
            "cantidad": cantidad,
            "precio_unit": precio_unit,
            "total": total,
            "anulada": False
        }
        ventas = pd.concat([ventas, pd.DataFrame([nueva])], ignore_index=True)
        escribir_ventas(ventas)
        if tabla:
            listar_ventas(tabla)
        messagebox.showinfo("Éxito", "Venta creada correctamente.")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def actualizar_venta(tabla, id_venta: int, cantidad=None, precio_unit=None):
    """Actualiza cantidad/precio; ajusta stock por diferencia en cantidad."""
    try:
        ventas = _ventas_leer_base()
        idx = ventas.index[ventas["id"] == id_venta]
        if idx.empty:
            raise ValueError("ID de venta no existe.")
        i = idx[0]
        if bool(ventas.at[i, "anulada"]):
            raise ValueError("No se puede actualizar una venta anulada.")

        old_cant = float(ventas.at[i, "cantidad"])
        old_prec = float(ventas.at[i, "precio_unit"])
        new_cant = float(cantidad) if cantidad not in (None, "") else old_cant
        new_prec = float(precio_unit) if precio_unit not in (None, "") else old_prec
        if new_cant <= 0 or new_prec < 0:
            raise ValueError("Cantidad > 0 y precio_unit ≥ 0.")

        diff = new_cant - old_cant
        if diff != 0:
            prod_nom = str(ventas.at[i, "producto"])
            codigo = _codigo_producto_por_nombre(prod_nom)
            if diff > 0:
                if not validar_existencia(codigo, diff):
                    raise ValueError("Stock insuficiente para aumentar cantidad.")
                actualizar_stock(codigo, diff)
            else:
                restaurar_stock(codigo, -diff)

        ventas.at[i, "cantidad"] = new_cant
        ventas.at[i, "precio_unit"] = new_prec
        ventas.at[i, "total"] = calcular_total(new_cant, new_prec)
        escribir_ventas(ventas)
        if tabla:
            listar_ventas(tabla)
        messagebox.showinfo("Éxito", "Venta actualizada.")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def anular_venta(tabla, id_venta: int):
    """Marca anulada=True y repone stock."""
    try:
        ventas = _ventas_leer_base()
        idx = ventas.index[ventas["id"] == id_venta]
        if idx.empty:
            raise ValueError("ID de venta no existe.")
        i = idx[0]
        if bool(ventas.at[i, "anulada"]):
            raise ValueError("La venta ya estaba anulada.")

        prod_nom = str(ventas.at[i, "producto"])
        cant = float(ventas.at[i, "cantidad"])
        codigo = _codigo_producto_por_nombre(prod_nom)

        restaurar_stock(codigo, cant)
        ventas.at[i, "anulada"] = True
        escribir_ventas(ventas)
        if tabla:
            listar_ventas(tabla)
        messagebox.showinfo("Éxito", "Venta anulada y stock repuesto.")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def eliminar_venta(tabla, id_venta: int):
    """Elimina la venta. Si no estaba anulada, repone stock primero."""
    try:
        ventas = _ventas_leer_base()
        idx = ventas.index[ventas["id"] == id_venta]
        if idx.empty:
            raise ValueError("ID de venta no existe.")
        i = idx[0]
        if not bool(ventas.at[i, "anulada"]):
            prod_nom = str(ventas.at[i, "producto"])
            cant = float(ventas.at[i, "cantidad"])
            codigo = _codigo_producto_por_nombre(prod_nom)
            restaurar_stock(codigo, cant)

        ventas = ventas.drop(index=i).reset_index(drop=True)
        escribir_ventas(ventas)
        if tabla:
            listar_ventas(tabla)
        messagebox.showinfo("Éxito", "Venta eliminada.")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def pestaña_ventas(notebook):
    
    pestaña = ttk.Frame(notebook)
    notebook.add(pestaña, text="Ventas")

    cont = ttk.Frame(pestaña)
    cont.pack(fill="both", expand=True, padx=10, pady=8)
    cols = ("id","fecha","producto","cliente","cantidad","precio_unit","total","anulada")
    tabla = crear_tabla(cont, cols, ["ID","Fecha","Producto","Cliente","Cantidad","Precio Unit","Total","Anulada"])

    form = ttk.Frame(pestaña)
    form.pack(fill="x", padx=10, pady=8)

    inv = leer_inventario().copy(); inv["codigo"] = inv["codigo"].astype(str)
    cli = leer_clientes().copy();   cli["codigo"] = cli["codigo"].astype(str)
    productos_vals = [f'{r["codigo"]} | {r["nombre"]}' for _, r in inv.iterrows()]
    clientes_vals  = [f'{r["codigo"]} | {r["nombre"]}' for _, r in cli.iterrows()]

    cb_prod = ttk.Combobox(form, values=productos_vals, width=35)
    cb_cli  = ttk.Combobox(form, values=clientes_vals,  width=35)
    ent_cant = tb.Entry(form, width=10)
    ent_prec = tb.Entry(form, width=12)
    ent_id   = tb.Entry(form, width=10)

    tb.Label(form, text="Producto:").grid(row=0, column=0, padx=4, pady=4, sticky="e"); cb_prod.grid(row=0, column=1, padx=4, pady=4)
    tb.Label(form, text="Cliente:").grid(row=0, column=2, padx=4, pady=4, sticky="e"); cb_cli.grid(row=0, column=3, padx=4, pady=4)
    tb.Label(form, text="Cantidad:").grid(row=0, column=4, padx=4, pady=4, sticky="e"); ent_cant.grid(row=0, column=5, padx=4, pady=4)
    tb.Label(form, text="Precio unit:").grid(row=0, column=6, padx=4, pady=4, sticky="e"); ent_prec.grid(row=0, column=7, padx=4, pady=4)
    tb.Label(form, text="ID venta:").grid(row=1, column=0, padx=4, pady=4, sticky="e"); ent_id.grid(row=1, column=1, padx=4, pady=4)

    # --- AGREGADO: recargar combos dinámicamente y refrescar al mostrar la pestaña ---
    def _recargar_combobox():
        inv2 = leer_inventario().copy(); inv2["codigo"] = inv2["codigo"].astype(str)
        cli2 = leer_clientes().copy();   cli2["codigo"] = cli2["codigo"].astype(str)
        cb_prod["values"] = [f'{r["codigo"]} | {r["nombre"]}' for _, r in inv2.iterrows()]
        cb_cli["values"]  = [f'{r["codigo"]} | {r["nombre"]}' for _, r in cli2.iterrows()]

    cb_prod.configure(postcommand=_recargar_combobox, state="readonly")
    cb_cli.configure(postcommand=_recargar_combobox,  state="readonly")

    pestaña.bind("<Visibility>", lambda _e: (listar_ventas(tabla), _recargar_combobox()))

    def _cod(texto):
        return texto.split("|")[0].strip() if texto and "|" in texto else texto.strip()

    def _id_valido():
        txt = ent_id.get().strip()
        if not txt.isdigit():
            messagebox.showwarning("Atención", "Ingresa un ID numérico de venta.")
            return None
        return int(txt)

    tb.Button(form, text="Crear venta", bootstyle="success",
              command=lambda: crear_venta(_cod(cb_prod.get()), _cod(cb_cli.get()), ent_cant.get(), ent_prec.get(), tabla)
              ).grid(row=2, column=1, pady=6)

    tb.Button(form, text="Actualizar venta (opc.)", bootstyle="info",
              command=lambda: (lambda _id=_id_valido(): _id is not None and actualizar_venta(tabla, _id,
                                               cantidad=ent_cant.get().strip() or None,
                                               precio_unit=ent_prec.get().strip() or None))()
              ).grid(row=2, column=3, pady=6)

    tb.Button(form, text="Anular venta", bootstyle="warning",
              command=lambda: (lambda _id=_id_valido(): _id is not None and anular_venta(tabla, _id))()
              ).grid(row=2, column=5, pady=6)

    tb.Button(form, text="Eliminar venta (opc.)", bootstyle="danger",
              command=lambda: (lambda _id=_id_valido(): _id is not None and eliminar_venta(tabla, _id))()
              ).grid(row=2, column=7, pady=6)

    tb.Button(form, text="Refrescar", bootstyle="secondary",
              command=lambda: listar_ventas(tabla)
              ).grid(row=2, column=8, pady=6)

    listar_ventas(tabla)
    return pestaña


# REPORTES
def generar_reporte(columna, valor):
    """Genera reporte filtrando por la columna y valor indicados"""
    df = leer_ventas()
    if df is None:
        return None

    # Aseguramos que la columna exista
    if columna not in df.columns:
        messagebox.showerror("Error", f"No hay columna '{columna}' en la hoja Ventas.")
        return None

    # Filtro tipo contiene (case-insensitive)
    filtro = df[columna].astype(str).str.contains(str(valor), case=False, na=False)
    resultado = df[filtro]

    if resultado.empty:
        messagebox.showinfo("Sin resultados", f"No se encontraron ventas de '{valor}' en {columna}.")
        return None

    return resultado

def generar_txt_reporte(df: pd.DataFrame, ruta_archivo: Path):
    try:
        lineas = []
        # encabezados
        encabezado = " | ".join(df.columns.astype(str))
        lineas.append(encabezado)
        lineas.append("-" * len(encabezado))

        # filas
        for _, fila in df.iterrows():
            fila_txt = " | ".join(str(fila[col]) for col in df.columns)
            lineas.append(fila_txt)

        contenido = "\n".join(lineas)

        with open(ruta_archivo, "w", encoding="utf-8") as f:
            f.write(contenido)

        return ruta_archivo
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo generar el archivo de reporte: {e}")
        return None

def leer_ventas():
    """Lee los datos del archivo Ventas.xlsx (para reportes)."""
    if not RUTA_DATOS.exists():
        messagebox.showerror("Error", "No se encontró el archivo Ventas.xlsx")
        return None
    try:
        return pd.read_excel(RUTA_DATOS, sheet_name=HOJA_VENTAS)
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo leer el archivo: {e}")
        return None

def enviar_mail_con_adjunto(destinatario: str, asunto: str, cuerpo: str, ruta_adjunto: Path):
    try:
        remitente = os.getenv("GOOGLE_APP_EMAIL")
        password_app = os.getenv("GOOGLE_APP_PASS")

        msg = MIMEMultipart()
        msg["From"] = remitente
        msg["To"] = destinatario
        msg["Subject"] = asunto
        msg.attach(MIMEText(cuerpo, "plain"))

        with open(ruta_adjunto, "rb") as f:
            parte = MIMEBase("application", "octet-stream")
            parte.set_payload(f.read())
        encoders.encode_base64(parte)
        parte.add_header("Content-Disposition", f'attachment; filename="{ruta_adjunto.name}"')
        msg.attach(parte)

        servidor = smtplib.SMTP("smtp.gmail.com", 587)
        servidor.starttls()
        servidor.login(remitente, password_app)
        servidor.send_message(msg)
        servidor.quit()

        messagebox.showinfo("Éxito", f"Reporte enviado a {destinatario}")
        return True
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo enviar el correo: {e}")
        return False

def pestaña_reportes(notebook):
    frame = ttk.Frame(notebook)
    notebook.add(frame, text="Reportes")

    ttk.Label(
        frame,
        text="Enviar reportes de ventas por cliente",
        font=("Segoe UI", 14, "bold")
    ).grid(row=0, column=0, columnspan=3, pady=10)

    # Cargar lista de clientes
    try:
        df_cli = leer_clientes()
        clientes_nombres = sorted(df_cli["nombre"].dropna().astype(str).unique().tolist())
    except Exception as e:
        messagebox.showerror("Error", f"No se pudieron leer clientes: {e}")
        clientes_nombres = []

    ttk.Label(frame, text="Cliente:").grid(row=1, column=0, sticky="e", padx=4, pady=4)
    cb_cliente = ttk.Combobox(frame, values=clientes_nombres, width=35, state="readonly")
    cb_cliente.grid(row=1, column=1, padx=4, pady=4)

    ttk.Label(frame, text="Correo destino:").grid(row=2, column=0, sticky="e", padx=4, pady=4)
    entry_destinatario = ttk.Entry(frame, width=35)
    entry_destinatario.grid(row=2, column=1, padx=4, pady=4)

    # Recargar lista cuando se muestra la pestañas
    def recargar_clientes(_evt=None):
        try:
            df_cli2 = leer_clientes()
            cb_cliente["values"] = sorted(df_cli2["nombre"].dropna().astype(str).unique().tolist())
        except Exception as e:
            print("Error al recargar lista de clientes:", e)

    frame.bind("<Visibility>", recargar_clientes)

    # Enviar reporte
    def enviar_reporte_cliente():
        try:
            cliente = cb_cliente.get().strip()
            destinatario = entry_destinatario.get().strip()
            if not cliente:
                raise ValueError("Selecciona un cliente.")
            if not destinatario:
                raise ValueError("Escribe el correo destinatario.")

            df_rep = generar_reporte("Cliente", cliente)
            if df_rep is None:
                return

            nombre_archivo = f"reporte_cliente_{cliente}.txt".replace(" ", "_")
            ruta_archivo = Path(nombre_archivo)
            ruta_final = generar_txt_reporte(df_rep, ruta_archivo)
            if ruta_final is None:
                return

            asunto = f"Reporte de ventas del cliente: {cliente}"
            cuerpo = f"Adjunto encontrarás el reporte de ventas del cliente {cliente}."
            enviar_mail_con_adjunto(destinatario, asunto, cuerpo, ruta_final)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    tb.Button(
        frame,
        text="Enviar reporte de cliente por correo",
        bootstyle="info",
        command=enviar_reporte_cliente
    ).grid(row=3, column=0, columnspan=2, pady=12)

    return frame


# VENTANA PRINCIPAL

def abrir_algoritmos(padre=None):
    """Abre la ventana del módulo de Inventario, Clientes y Ventas."""
    crear_archivo_excel_si_no_existe()
    is_root = padre is None
    ventana = tb.Window(themename="superhero") if is_root else tb.Toplevel(padre)
    ventana.title("Sistema de Ventas - Grupo #5")
    ventana.geometry("1000x650")

    tb.Label(
        ventana,
        text="Control de Inventario",
        font=("Arial", 20, "bold")
    ).pack(pady=10)

    cuaderno = tb.Notebook(ventana)
    cuaderno.pack(expand=True, fill="both", padx=10, pady=10)

    pestaña_inventario(cuaderno)
    pestaña_clientes(cuaderno)
    _asegurar_hoja_ventas()
    pestaña_ventas(cuaderno)
    try:
        pestaña_reportes(cuaderno)
    except Exception:
        pass

    if is_root:
        ventana.mainloop()

if __name__ == "__main__":
    abrir_algoritmos()
