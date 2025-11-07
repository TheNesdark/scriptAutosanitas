import customtkinter as ctk
from tkinter import filedialog, messagebox
from src.procesador import procesar_archivo

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

def seleccionar_archivo_entrada():
    archivo = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
    if archivo:
        entrada_var.set(archivo)

def seleccionar_archivo_salida():
    archivo = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
    if archivo:
        salida_var.set(archivo)

def procesar():
    entrada = entrada_var.get()
    salida = salida_var.get()
    
    if not entrada or not salida:
        messagebox.showerror("Error", "Selecciona ambos archivos")
        return
    
    try:
        total = procesar_archivo(entrada, salida)
        messagebox.showinfo("Éxito", f"Procesados {total} registros\nArchivo guardado: {salida}")
    except Exception as e:
        messagebox.showerror("Error", f"Error al procesar: {str(e)}")

def iniciar_gui():
    global entrada_var, salida_var
    
    ventana = ctk.CTk()
    ventana.title("Procesador Autosanitas")
    ventana.geometry("650x400")
    ventana.resizable(False, False)

    entrada_var = ctk.StringVar()
    salida_var = ctk.StringVar(value="NV.xlsx")

    titulo = ctk.CTkLabel(ventana, text="Procesador Autosanitas", font=("Segoe UI", 24, "bold"))
    titulo.pack(pady=30)

    frame = ctk.CTkFrame(ventana)
    frame.pack(padx=40, pady=10, fill="both", expand=True)

    ctk.CTkLabel(frame, text="Archivo de entrada:", font=("Segoe UI", 12)).pack(anchor="w", padx=20, pady=(20, 5))
    frame_entrada = ctk.CTkFrame(frame, fg_color="transparent")
    frame_entrada.pack(padx=20, pady=(0, 15), fill="x")
    ctk.CTkEntry(frame_entrada, textvariable=entrada_var, height=35, font=("Segoe UI", 11)).pack(side="left", fill="x", expand=True, padx=(0, 10))
    ctk.CTkButton(frame_entrada, text="📁 Buscar", command=seleccionar_archivo_entrada, width=100, height=35, font=("Segoe UI", 11, "bold")).pack(side="left")

    ctk.CTkLabel(frame, text="Archivo de salida:", font=("Segoe UI", 12)).pack(anchor="w", padx=20, pady=(10, 5))
    frame_salida = ctk.CTkFrame(frame, fg_color="transparent")
    frame_salida.pack(padx=20, pady=(0, 20), fill="x")
    ctk.CTkEntry(frame_salida, textvariable=salida_var, height=35, font=("Segoe UI", 11)).pack(side="left", fill="x", expand=True, padx=(0, 10))
    ctk.CTkButton(frame_salida, text="💾 Guardar", command=seleccionar_archivo_salida, width=100, height=35, font=("Segoe UI", 11, "bold")).pack(side="left")

    ctk.CTkButton(ventana, text="▶  Procesar Archivo", command=procesar, height=45, font=("Segoe UI", 14, "bold"), fg_color="#27ae60", hover_color="#229954").pack(pady=20, padx=40, fill="x")

    ventana.mainloop()
