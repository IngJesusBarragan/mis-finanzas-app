import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog

archivo = "Mis finanzas en BS.xlsx"

try:
    df = pd.read_excel(archivo)
except FileNotFoundError:
    df = pd.DataFrame(columns=["Fecha", "Tipo", "Concepto", "Monto"])

def obtener_saldo():
    return df["Monto"].sum()

def guardar_df():
    df.to_excel(archivo, index=False)

def actualizar_tabla():
    for fila in tabla.get_children():
        tabla.delete(fila)
    ultimos = df.tail(10)
    for _, fila in ultimos.iterrows():
        monto = fila["Monto"]
        if monto >= 0:
            monto_str = f"+Bs.{monto:.2f}"
            color = "green"
        else:
            monto_str = f"-Bs.{abs(monto):.2f}"
            color = "red"
        tabla.insert("", "end", values=(
            fila["Fecha"],
            fila["Tipo"],
            fila["Concepto"],
            monto_str
        ), tags=(color,))
    tabla.tag_configure("green", foreground="green")
    tabla.tag_configure("red", foreground="red")

def actualizar_saldo_label():
    saldo = obtener_saldo()
    label_saldo.config(text=f"Bs.{saldo:.2f}")
    color = "green" if saldo >= 0 else "red"
    label_saldo.config(fg=color)

def ventana_primer_uso():
    global df
    ventana = tk.Toplevel(root)
    ventana.title("¡Bienvenido!")
    ventana.geometry("350x180")
    ventana.resizable(False, False)
    ventana.grab_set()
    tk.Label(ventana, text="¡Primera vez! ¿Con cuánto dinero empiezas?",
             wraplength=300, pady=10).pack()
    entry = tk.Entry(ventana, font=("Arial", 14), justify="center")
    entry.pack(pady=5)
    entry.focus()
    def confirmar():
        global df
        try:
            monto = float(entry.get())
            if monto <= 0:
                messagebox.showerror("Error", "Debe ser un número positivo.", parent=ventana)
                return
            fecha = datetime.now().strftime("%d/%m/%Y")
            fila = pd.DataFrame([{
                "Fecha": fecha,
                "Tipo": "Ingreso",
                "Concepto": "Presupuesto Inicial",
                "Monto": monto
            }])
            df = pd.concat([df, fila], ignore_index=True)
            guardar_df()
            ventana.destroy()
            actualizar_tabla()
            actualizar_saldo_label()
        except ValueError:
            messagebox.showerror("Error", "Ingresa un número válido.", parent=ventana)
    tk.Button(ventana, text="Comenzar", command=confirmar,
              bg="#4CAF50", fg="white", font=("Arial", 11)).pack(pady=10)

def abrir_anadir_gasto():
    global df
    saldo_actual = obtener_saldo()
    ventana = tk.Toplevel(root)
    ventana.title("Añadir Gasto")
    ventana.geometry("350x220")
    ventana.resizable(False, False)
    ventana.grab_set()
    tk.Label(ventana, text="¿Cuánto gastaste?", font=("Arial", 11)).pack(pady=(15, 2))
    entry_monto = tk.Entry(ventana, font=("Arial", 13), justify="center")
    entry_monto.pack()
    entry_monto.focus()
    tk.Label(ventana, text="¿En qué gastaste?", font=("Arial", 11)).pack(pady=(10, 2))
    entry_concepto = tk.Entry(ventana, font=("Arial", 13), justify="center")
    entry_concepto.pack()
    def confirmar():
        global df
        try:
            monto = float(entry_monto.get())
            concepto = entry_concepto.get().strip()
            if monto <= 0:
                messagebox.showerror("Error", "El monto debe ser positivo.", parent=ventana)
                return
            if monto > saldo_actual:
                messagebox.showerror("Error", "¡Fondos insuficientes!", parent=ventana)
                return
            if concepto == "":
                messagebox.showerror("Error", "El concepto no puede estar vacío.", parent=ventana)
                return
            fecha = datetime.now().strftime("%d/%m/%Y")
            fila = pd.DataFrame([{
                "Fecha": fecha,
                "Tipo": "Gasto",
                "Concepto": concepto,
                "Monto": -monto
            }])
            df = pd.concat([df, fila], ignore_index=True)
            guardar_df()
            ventana.destroy()
            actualizar_tabla()
            actualizar_saldo_label()
            messagebox.showinfo("Listo", f"Gasto de Bs.{monto:.2f} registrado.")
        except ValueError:
            messagebox.showerror("Error", "Ingresa un número válido.", parent=ventana)
    tk.Button(ventana, text="Registrar Gasto", command=confirmar,
              bg="#e53935", fg="white", font=("Arial", 11)).pack(pady=15)

def abrir_anadir_ingreso():
    global df
    ventana = tk.Toplevel(root)
    ventana.title("Añadir Ingreso")
    ventana.geometry("350x220")
    ventana.resizable(False, False)
    ventana.grab_set()
    tk.Label(ventana, text="¿Cuánto dinero recibes?", font=("Arial", 11)).pack(pady=(15, 2))
    entry_monto = tk.Entry(ventana, font=("Arial", 13), justify="center")
    entry_monto.pack()
    entry_monto.focus()
    tk.Label(ventana, text="¿Qué dinero es?", font=("Arial", 11)).pack(pady=(10, 2))
    entry_concepto = tk.Entry(ventana, font=("Arial", 13), justify="center")
    entry_concepto.pack()
    def confirmar():
        global df
        try:
            monto = float(entry_monto.get())
            concepto = entry_concepto.get().strip()
            if monto <= 0:
                messagebox.showerror("Error", "El monto debe ser positivo.", parent=ventana)
                return
            if concepto == "":
                messagebox.showerror("Error", "El concepto no puede estar vacío.", parent=ventana)
                return
            fecha = datetime.now().strftime("%d/%m/%Y")
            fila = pd.DataFrame([{
                "Fecha": fecha,
                "Tipo": "Ingreso",
                "Concepto": concepto,
                "Monto": monto
            }])
            df = pd.concat([df, fila], ignore_index=True)
            guardar_df()
            ventana.destroy()
            actualizar_tabla()
            actualizar_saldo_label()
            messagebox.showinfo("Listo", f"¡Se sumaron Bs.{monto:.2f} a tu saldo!")
        except ValueError:
            messagebox.showerror("Error", "Ingresa un número válido.", parent=ventana)
    tk.Button(ventana, text="Registrar Ingreso", command=confirmar,
              bg="#4CAF50", fg="white", font=("Arial", 11)).pack(pady=15)

# ============================================================
# VENTANA PRINCIPAL
# ============================================================
root = tk.Tk()
root.title("Mis Finanzas 💰")
root.geometry("600x500")
root.resizable(False, False)

frame_saldo = tk.Frame(root, bg="#1e1e2e", pady=15)
frame_saldo.pack(fill="x")
tk.Label(frame_saldo, text="Saldo Actual",
         bg="#1e1e2e", fg="white", font=("Arial", 12)).pack()
label_saldo = tk.Label(frame_saldo, text="Bs.0.00",
                       bg="#1e1e2e", fg="green", font=("Arial", 32, "bold"))
label_saldo.pack()

frame_tabla = tk.Frame(root)
frame_tabla.pack(fill="both", expand=True, padx=10, pady=10)
tabla = ttk.Treeview(frame_tabla,
                     columns=("Fecha", "Tipo", "Concepto", "Monto"),
                     show="headings", height=12)
tabla.heading("Fecha",    text="Fecha")
tabla.heading("Tipo",     text="Tipo")
tabla.heading("Concepto", text="Concepto")
tabla.heading("Monto",    text="Monto")
tabla.column("Fecha",    width=100, anchor="center")
tabla.column("Tipo",     width=80,  anchor="center")
tabla.column("Concepto", width=220, anchor="w")
tabla.column("Monto",    width=100, anchor="e")
scrollbar = ttk.Scrollbar(frame_tabla, orient="vertical", command=tabla.yview)
tabla.configure(yscrollcommand=scrollbar.set)
tabla.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

frame_botones = tk.Frame(root, pady=10)
frame_botones.pack()
tk.Button(frame_botones, text="➕ Añadir Gasto",
          command=abrir_anadir_gasto,
          bg="#e53935", fg="white",
          font=("Arial", 11), width=16).grid(row=0, column=0, padx=8)
tk.Button(frame_botones, text="💵 Añadir Ingreso",
          command=abrir_anadir_ingreso,
          bg="#4CAF50", fg="white",
          font=("Arial", 11), width=16).grid(row=0, column=1, padx=8)
tk.Button(frame_botones, text="🚪 Salir",
          command=root.quit,
          bg="#555", fg="white",
          font=("Arial", 11), width=10).grid(row=0, column=2, padx=8)

if df.empty:
    root.after(100, ventana_primer_uso)
actualizar_tabla()
actualizar_saldo_label()
root.mainloop()
print("¡Has salido con éxito!")