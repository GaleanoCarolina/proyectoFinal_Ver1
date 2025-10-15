
import tkinter as tk
from tkinter import ttk, messagebox
import ttkbootstrap as tb
import numpy as np

def leer_matriz(texto):
    filas = [list(map(float, f.split())) for f in texto.strip().split("\n") if f.strip()]
    if not filas:
        raise ValueError("Entrada vacía")
    ncols = len(filas[0])
    if any(len(f) != ncols for f in filas):
        raise ValueError("Todas las filas deben tener el mismo número de columnas")
    return np.array(filas, dtype=float)

def leer_vector(texto):
    valores = [float(x) for x in texto.strip().replace("\n", " ").split()]
    return np.array(valores, dtype=float).reshape(-1, 1)

def calcular_inversa():
    try:
        A = leer_matriz(txt_inv.get("1.0", tk.END))
        if A.shape[0] != A.shape[1]:
            salida.set("La matriz debe ser cuadrada")
            return
        det = np.linalg.det(A)
        if np.isclose(det, 0):
            salida.set("La matriz no tiene inversa (determinante = 0)")
            return
        inv = np.linalg.inv(A)
        salida.set("Matriz inversa:\n" + np.array2string(inv, precision=4, separator='  '))
    except Exception as e:
        salida.set("Error: " + str(e))

def multiplicar_matrices():
    try:
        A = leer_matriz(txt_mulA.get("1.0", tk.END))
        B = leer_matriz(txt_mulB.get("1.0", tk.END))
        if A.shape[1] != B.shape[0]:
            salida.set("No se pueden multiplicar: columnas de A ≠ filas de B")
            return
        R = A.dot(B)
        salida.set("Resultado de A × B:\n" + np.array2string(R, precision=4, separator='  '))
    except Exception as e:
        salida.set("Error: " + str(e))

def resolver_sistema():
    try:
        A = leer_matriz(txt_sysA.get("1.0", tk.END))
        b = leer_vector(txt_sysb.get("1.0", tk.END))
        metodo = combo_metodo.get()

        if A.shape[0] != A.shape[1]:
            salida.set("La matriz A debe ser cuadrada")
            return
        if b.shape[0] != A.shape[0]:
            salida.set(f"El vector b debe tener {A.shape[0]} filas")
            return

        detA = np.linalg.det(A)

        if not np.isclose(detA, 0):
            if metodo == "Gauss-Jordan":
                x = np.linalg.solve(A, b)
                salida.set("Solución única (Gauss-Jordan):\n" + np.array2string(x.flatten(), precision=4, separator='  '))
            else:
                soluciones = []
                for i in range(A.shape[0]):
                    Ai = A.copy()
                    Ai[:, i] = b.flatten()
                    soluciones.append(np.linalg.det(Ai) / detA)
                salida.set("Solución única (Regla de Cramer):\n" + np.array2string(np.array(soluciones), precision=4, separator='  '))
        else:
            rangoA = np.linalg.matrix_rank(A)
            rangoAb = np.linalg.matrix_rank(np.hstack([A, b]))
            if rangoA < rangoAb:
                salida.set("Sin solución (sistema incompatible)")
            elif rangoA == rangoAb and rangoA < A.shape[0]:
                salida.set("Infinitas soluciones (sistema indeterminado)")
            else:
                salida.set("Caso no determinado")
    except Exception as e:
        salida.set("Error: " + str(e))

def limpiar_campos():
    for t in [txt_inv, txt_mulA, txt_mulB, txt_sysA, txt_sysb]:
        t.delete("1.0", tk.END)
    salida.set("Campos limpiados correctamente")

def abrir_algebra(parent=None):
    is_root = parent is None
    if is_root:
        global app
        app = tb.Window(themename="superhero")
    else:
        app = tk.Toplevel(parent)

    app.title("Calculadora de Álgebra")
    app.geometry("820x700")

    global txt_inv, txt_mulA, txt_mulB, txt_sysA, txt_sysb, combo_metodo, salida
    salida = tk.StringVar()

    tb.Label(app, text="Calculadora de Álgebra", font=("Arial", 20, "bold")).pack(pady=10)

    notebook = tb.Notebook(app)
    notebook.pack(expand=True, fill="both", padx=15, pady=10)

    tab1 = ttk.Frame(notebook)
    notebook.add(tab1, text="Inversa")
    tb.Label(tab1, text="Ingrese la matriz (valores separados por espacios):", font=("Arial", 11)).pack(pady=5)
    txt_inv = tk.Text(tab1, height=7, width=70)
    txt_inv.pack(pady=5)
    tb.Button(tab1, text="Calcular Inversa", bootstyle="success", command=calcular_inversa).pack(pady=10)

    tab2 = ttk.Frame(notebook)
    notebook.add(tab2, text="Multiplicación")
    tb.Label(tab2, text="Matriz A:", font=("Arial", 11)).pack(pady=5)
    txt_mulA = tk.Text(tab2, height=6, width=70)
    txt_mulA.pack()
    tb.Label(tab2, text="Matriz B:", font=("Arial", 11)).pack(pady=5)
    txt_mulB = tk.Text(tab2, height=6, width=70)
    txt_mulB.pack()
    tb.Button(tab2, text="Multiplicar", bootstyle="info", command=multiplicar_matrices).pack(pady=10)

    tab3 = ttk.Frame(notebook)
    notebook.add(tab3, text="Sistemas de ecuaciones")
    tb.Label(tab3, text="Matriz de coeficientes (A):", font=("Arial", 11)).pack(pady=5)
    txt_sysA = tk.Text(tab3, height=6, width=70)
    txt_sysA.pack()
    tb.Label(tab3, text="Vector de resultados (b):", font=("Arial", 11)).pack(pady=5)
    txt_sysb = tk.Text(tab3, height=3, width=70)
    txt_sysb.pack()
    tb.Label(tab3, text="Método de resolución:", font=("Arial", 11)).pack(pady=5)
    combo_metodo = tb.Combobox(tab3, values=["Gauss-Jordan", "Regla de Cramer"], state="readonly")
    combo_metodo.current(0)
    combo_metodo.pack(pady=5)
    tb.Button(tab3, text="Resolver Sistema", bootstyle="warning", command=resolver_sistema).pack(pady=10)

    tb.Label(app, text="Resultado:", font=("Arial", 13, "bold")).pack()
    tb.Label(app, textvariable=salida, font=("Consolas", 11), justify="left", wraplength=760, bootstyle="light").pack(fill="both", padx=10, pady=10)

    tb.Button(app, text="Limpiar Todo", bootstyle="danger", command=limpiar_campos).pack(pady=8)

    if is_root:
        app.mainloop()