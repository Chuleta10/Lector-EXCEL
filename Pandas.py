import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

class ExcelReaderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Reader")
        self.filtered_df = None
        self.create_widgets()
        
    def create_widgets(self):
        self.frame = tk.Frame(self.root, padx=20, pady=20)
        self.frame.pack()
        
        # Sección de selección de archivo
        self.file_label = tk.Label(self.frame, text="Seleccionar archivo Excel:")
        self.file_label.grid(row=0, column=0, sticky='w')
        
        self.file_entry = tk.Entry(self.frame, width=40, state='readonly')
        self.file_entry.grid(row=0, column=1)
        
        self.browse_button = tk.Button(self.frame, text="Examinar", command=self.browse_file)
        self.browse_button.grid(row=0, column=2)
        
        # Sección de selección de hoja
        self.sheet_label = tk.Label(self.frame, text="Seleccionar hoja:")
        self.sheet_label.grid(row=1, column=0, sticky='w')
        
        self.sheet_combobox = ttk.Combobox(self.frame, width=37, state='readonly')
        self.sheet_combobox.grid(row=1, column=1, columnspan=2)
        
        # Sección de filtro
        self.filter_label = tk.Label(self.frame, text="Filtro:")
        self.filter_label.grid(row=2, column=0, sticky='w')
        
        self.filter_entry = tk.Entry(self.frame, width=40)
        self.filter_entry.grid(row=2, column=1)
        
        self.search_options = tk.StringVar()
        self.search_options.set("Buscar en todas las columnas")
        self.search_options_menu = ttk.OptionMenu(self.frame, self.search_options, "Buscar en todas las columnas",
                                                   "Buscar en todas las columnas", "Buscar en columna específica")
        self.search_options_menu.grid(row=2, column=2)
        
        self.apply_button = tk.Button(self.frame, text="Aplicar Filtro", command=self.apply_filter)
        self.apply_button.grid(row=2, column=3)
        
        # Botón de exportación CSV
        self.export_button = tk.Button(self.frame, text="Exportar CSV", command=self.export_csv)
        self.export_button.grid(row=3, column=1, pady=10)
        
        # Etiqueta de estado
        self.status_label = tk.Label(self.frame, text="", fg="blue")
        self.status_label.grid(row=4, column=0, columnspan=4)
        
        # Tabla de resultados
        self.output_tree = ttk.Treeview(self.frame)
        self.output_tree.grid(row=5, column=0, columnspan=4)
        
    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_path:
            self.file_entry.configure(state='normal')
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, file_path)
            self.file_entry.configure(state='readonly')
            self.load_excel_sheets(file_path)
        
    def load_excel_sheets(self, file_path):
        try:
            sheets = pd.ExcelFile(file_path).sheet_names
            self.sheet_combobox['values'] = sheets
            if sheets:
                self.sheet_combobox.current(0)
        except Exception as e:
            self.show_status(f"No se pudieron cargar las hojas: {str(e)}", "error")
        
    def apply_filter(self):
        file_path = self.file_entry.get()
        sheet_name = self.sheet_combobox.get()
        filter_value = self.filter_entry.get()
        search_option = self.search_options.get()
        
        if not all([file_path, sheet_name, filter_value]):
            self.show_status("Por favor complete todas las opciones.", "warning")
            return
        
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            if search_option == "Buscar en todas las columnas":
                self.filtered_df = df[df.apply(lambda row: self.check_filter(row, filter_value), axis=1)]
            else:
                column_name = simpledialog.askstring("Buscar en columna específica", "Ingrese el nombre de la columna:")
                if column_name is None or column_name.strip() == "":
                    self.show_status("Por favor ingrese el nombre de la columna.", "warning")
                    return
                if column_name not in df.columns:
                    self.show_status("La columna especificada no existe.", "warning")
                    return
                self.filtered_df = df[df[column_name].apply(lambda cell: filter_value.lower() in str(cell).lower())]
            
            if self.filtered_df.empty:
                self.show_status("No se encontraron coincidencias.", "info")
            else:
                self.show_status("Filtro aplicado exitosamente.", "success")
                
            self.display_results(self.filtered_df)
        except Exception as e:
            self.show_status(f"Error: {str(e)}", "error")
            self.display_results(pd.DataFrame())
    
    def check_filter(self, row, filter_value):
        for value in row:
            if filter_value.lower() in str(value).lower():
                return True
        return False
    
    def display_results(self, df):
        self.output_tree.delete(*self.output_tree.get_children())
        
        if not df.empty:
            # Agregar etiquetas de columna
            self.output_tree["columns"] = list(df.columns)
            self.output_tree.column("#0", width=0, stretch=tk.NO)  
            for column in df.columns:
                self.output_tree.heading(column, text=column)
            # Agregar filas
            for index, row_data in df.iterrows():
                self.output_tree.insert("", "end", values=list(row_data))
    
    def show_status(self, message, status_type):
        color = {"error": "red", "success": "green", "info": "black", "warning": "orange"}
        fg_color = color.get(status_type, "black")
        self.status_label.config(text=message, fg=fg_color)
        messagebox.showinfo(status_type.capitalize(), message)
        
    def export_csv(self):
        if self.filtered_df is None:
            self.show_status("No hay datos filtrados para exportar.", "warning")
            return
        
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if file_path:
            try:
                self.filtered_df.to_csv(file_path, index=False)
                self.show_status("CSV exportado exitosamente.", "success")
            except Exception as e:
                self.show_status(f"Error al exportar CSV: {str(e)}", "error")
                
def main():
    root = tk.Tk()
    app = ExcelReaderApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
