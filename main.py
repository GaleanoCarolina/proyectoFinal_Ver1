# menu.py
import tkinter as tk
import ttkbootstrap as tb
from matematica import abrir_matematicas
from algebra import abrir_algebra
from algoritmos import abrir_algoritmos

def abrir_modulo(modulo):
    if modulo == "Matemáticas":
        abrir_matematicas(root)
    elif modulo == "Álgebra":
        abrir_algebra(root)
    elif modulo == "Algoritmos":                
        abrir_algoritmos(root)

root = tb.Window(themename="superhero")
root.title("Menú Principal")
root.geometry("400x300")

tb.Label(root, text="Menú Principal", font=("Arial", 20, "bold")).pack(pady=20)

tb.Button(root, text="Matemáticas", bootstyle="primary", width=20,
          command=lambda: abrir_modulo("Matemáticas")).pack(pady=10)

tb.Button(root, text="Álgebra", bootstyle="secondary", width=20,
          command=lambda: abrir_modulo("Álgebra")).pack(pady=10)

tb.Button(root, text="Algoritmos", bootstyle="success", width=20,
          command=lambda: abrir_modulo("Algoritmos")).pack(pady=10)

tb.Button(root, text="Salir", bootstyle="danger", width=20,
          command=root.destroy).pack(pady=20)

root.mainloop()
