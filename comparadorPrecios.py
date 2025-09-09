import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os

class ComparadorPreciosApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Comparador de Precios")
        self.root.geometry("900x600")
        self.root.configure(bg="black")

        self.archivo_actual = None
        self.archivo_nuevo = None
        self.df_resultado = None

        # Frame superior centrado
        frame_top = tk.Frame(root, bg="black")
        frame_top.pack(pady=10)

        frame_top.grid_columnconfigure(0, weight=1)
        frame_top.grid_columnconfigure(1, weight=1)

        tk.Button(frame_top, text="Cargar Excel Actual", command=self.cargar_actual, bg="gray20", fg="white").grid(row=0, column=0, padx=5, sticky="e")
        self.label_actual = tk.Label(frame_top, text="No cargado", bg="black", fg="white")
        self.label_actual.grid(row=0, column=1, padx=5, sticky="w")

        tk.Button(frame_top, text="Cargar Excel Nuevo", command=self.cargar_nuevo, bg="gray20", fg="white").grid(row=1, column=0, padx=5, sticky="e")
        self.label_nuevo = tk.Label(frame_top, text="No cargado", bg="black", fg="white")
        self.label_nuevo.grid(row=1, column=1, padx=5, sticky="w")

        tk.Button(frame_top, text="Comparar", command=self.comparar, bg="gray30", fg="white").grid(row=2, column=0, pady=10, sticky="e")
        tk.Button(frame_top, text="Exportar Resultado", command=self.exportar, bg="gray30", fg="white").grid(row=2, column=1, pady=10, sticky="w")

        # Tabla de resultados
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", background="black", foreground="white", fieldbackground="black", rowheight=25)
        style.configure("Treeview.Heading", background="gray15", foreground="white")
        style.map("Treeview", background=[("selected", "gray40")])

        self.tabla = ttk.Treeview(root, columns=("codigo", "descripcion", "precio_viejo", "precio_nuevo", "diferencia"), show="headings")
        self.tabla.heading("codigo", text="Código")
        self.tabla.heading("descripcion", text="Descripción")
        self.tabla.heading("precio_viejo", text="Precio Viejo")
        self.tabla.heading("precio_nuevo", text="Precio Nuevo")
        self.tabla.heading("diferencia", text="Diferencia")

        self.tabla.column("codigo", width=100, anchor="center")
        self.tabla.column("descripcion", width=350, anchor="center")
        self.tabla.column("precio_viejo", width=120, anchor="center")
        self.tabla.column("precio_nuevo", width=120, anchor="center")
        self.tabla.column("diferencia", width=150, anchor="center")

        self.tabla.pack(fill=tk.BOTH, expand=True)

        # Estilo para colores
        self.tabla.tag_configure("subio", background="firebrick")
        self.tabla.tag_configure("bajo", background="darkgreen")
        self.tabla.tag_configure("igual", background="gray25")

    def cargar_actual(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            try:
                df = pd.read_excel(path, sheet_name="Hoja1")
                if df.shape[1] < 4:
                    raise ValueError("Formato incorrecto")
                self.archivo_actual = df
                self.label_actual.config(text=os.path.basename(path))
            except Exception:
                messagebox.showwarning("Advertencia", f"El archivo {os.path.basename(path)} no tiene el formato esperado")

    def cargar_nuevo(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            try:
                df = pd.read_excel(path, sheet_name="Hoja1")
                if df.shape[1] < 4:
                    raise ValueError("Formato incorrecto")
                self.archivo_nuevo = df
                self.label_nuevo.config(text=os.path.basename(path))
            except Exception:
                messagebox.showwarning("Advertencia", f"El archivo {os.path.basename(path)} no tiene el formato esperado")

    def comparar(self):
        if self.archivo_actual is None or self.archivo_nuevo is None:
            messagebox.showerror("Error", "Debes cargar ambos archivos válidos")
            return

        # Renombrar columnas para estandarizar
        columnas = ["codigo", "descripcion", "cantidad", "precio"]
        self.archivo_actual.columns = columnas
        self.archivo_nuevo.columns = columnas

        # Merge por código
        df = pd.merge(self.archivo_actual, self.archivo_nuevo, on="codigo", suffixes=("_viejo", "_nuevo"), how="outer")

        # Detectar productos faltantes
        faltantes_actual = df[df["precio_viejo"].isna()]["codigo"].tolist()
        faltantes_nuevo = df[df["precio_nuevo"].isna()]["codigo"].tolist()

        if faltantes_actual or faltantes_nuevo:
            msg = """⚠️ Productos faltantes:\n"""
            if faltantes_actual:
                msg += f"\nNo estaban en archivo actual: {faltantes_actual}"
            if faltantes_nuevo:
                msg += f"\nNo estaban en archivo nuevo: {faltantes_nuevo}"
            messagebox.showwarning("Productos faltantes", msg)

        # Reemplazar NaN por "No encontrado"
        df.fillna("No encontrado", inplace=True)

        # Calcular diferencia (evitar error si no son números)
        def calc_diff(row):
            try:
                return float(row["precio_nuevo"]) - float(row["precio_viejo"])
            except:
                return "No calculado"
        df["diferencia"] = df.apply(calc_diff, axis=1)

        # Limpiar tabla
        for row in self.tabla.get_children():
            self.tabla.delete(row)

        # Insertar filas con colores
        subio = bajo = igual = 0
        for _, row in df.iterrows():
            tag = ""
            if isinstance(row["diferencia"], str):
                tag = "igual"
            elif row["precio_nuevo"] > row["precio_viejo"]:
                tag = "subio"; subio += 1
            elif row["precio_nuevo"] < row["precio_viejo"]:
                tag = "bajo"; bajo += 1
            else:
                tag = "igual"; igual += 1

            self.tabla.insert("", tk.END, values=(row["codigo"], row["descripcion_viejo"] if row["descripcion_viejo"] != "No encontrado" else row["descripcion_nuevo"], row["precio_viejo"], row["precio_nuevo"], row["diferencia"]), tags=(tag,))

        self.df_resultado = df

        # Alert resumen
        messagebox.showinfo("Resultado", f"📈 Subieron: {subio}\n📉 Bajaron: {bajo}\n✅ Iguales: {igual}")

    def exportar(self):
        if self.df_resultado is None:
            messagebox.showerror("Error", "No hay datos para exportar")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.df_resultado.to_excel(path, index=False)
            messagebox.showinfo("Exportado", f"Archivo guardado en: {path}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ComparadorPreciosApp(root)
    root.mainloop()