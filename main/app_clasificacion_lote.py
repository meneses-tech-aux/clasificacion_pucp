import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
import threading
import csv
from datetime import datetime
import sys
import os

# Si la aplicación se ejecuta desde un ejecutable creado por PyInstaller,
# los archivos se extraen en un directorio temporal (`sys._MEIPASS`).
# En ese caso, configuramos `PLAYWRIGHT_BROWSERS_PATH` antes de importar
# Playwright para que Playwright pueda localizar los navegadores empaquetados
# dentro del ejecutable.
if getattr(sys, "frozen", False):
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        os.environ["PLAYWRIGHT_BROWSERS_PATH"] = os.path.join(meipass, "ms-playwright")

from playwright.sync_api import Playwright, sync_playwright, expect

# Importar la función principal del script
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from Clasificacion_en_Lote import funcion_registro_clasificacion, fecha_y_hora_actual

class ClasificacionAppLote:
    def __init__(self, root):
        self.root = root
        self.root.title("Clasificación en Lote - PUCP")
        self.root.geometry("800x700")
        self.root.resizable(True, True)
        
        # Configurar estilos
        style = ttk.Style()
        style.theme_use('clam')
        
        # Variables
        self.username = tk.StringVar()
        self.password = tk.StringVar()
        self.url_pandora = tk.StringVar(value="https://pandora.pucp.edu.pe/pucp/login?TARGET=https%3A%2F%2Feros.pucp.edu.pe%2Fpucp%2Fjsp%2FIntranet.jsp")
        self.curso_codigo_data = []
        self.is_running = False
        
        # Crear interfaz
        self.crear_widgets()
        
    def crear_widgets(self):
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configurar peso de filas y columnas
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(6, weight=1)
        
        # ===== SECCIÓN DE CREDENCIALES =====
        cred_frame = ttk.LabelFrame(main_frame, text="Credenciales", padding="10")
        cred_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        cred_frame.columnconfigure(1, weight=1)
        
        # Usuario
        ttk.Label(cred_frame, text="Usuario:").grid(row=0, column=0, sticky=tk.W, padx=5)
        username_entry = ttk.Entry(cred_frame, textvariable=self.username, width=30)
        username_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)
        
        # Contraseña
        ttk.Label(cred_frame, text="Contraseña:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        password_entry = ttk.Entry(cred_frame, textvariable=self.password, width=30, show="*")
        password_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        
        # URL Pandora
        ttk.Label(cred_frame, text="URL Pandora:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        url_entry = ttk.Entry(cred_frame, textvariable=self.url_pandora, width=30)
        url_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        
        # ===== SECCIÓN DE ARCHIVO =====
        file_frame = ttk.LabelFrame(main_frame, text="Archivo de Datos", padding="10")
        file_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        file_frame.columnconfigure(1, weight=1)
        
        self.file_label = ttk.Label(file_frame, text="Ningún archivo seleccionado", foreground="gray")
        self.file_label.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), padx=5, pady=5)
        
        ttk.Button(file_frame, text="Cargar archivo CSV", command=self.cargar_archivo).grid(row=1, column=0, padx=5, pady=5)
        ttk.Button(file_frame, text="Ver datos cargados", command=self.ver_datos).grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        
        # Información de datos
        self.info_label = ttk.Label(file_frame, text="Registros cargados: 0")
        self.info_label.grid(row=2, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)
        
        # ===== SECCIÓN DE CONTROL =====
        control_frame = ttk.Frame(main_frame)
        control_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        self.start_btn = ttk.Button(control_frame, text="Iniciar Proceso", command=self.iniciar_proceso)
        self.start_btn.pack(side=tk.LEFT, padx=5)
        
        self.stop_btn = ttk.Button(control_frame, text="Detener", command=self.detener_proceso, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(control_frame, text="Limpiar Consola", command=self.limpiar_consola).pack(side=tk.LEFT, padx=5)
        
        # ===== SECCIÓN DE ESTADO =====
        status_frame = ttk.LabelFrame(main_frame, text="Estado", padding="10")
        status_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        status_frame.columnconfigure(0, weight=1)
        
        self.progress = ttk.Progressbar(status_frame, mode='determinate')
        self.progress.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=5, pady=5)
        
        self.status_label = ttk.Label(status_frame, text="Estado: Listo", foreground="blue")
        self.status_label.grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        
        # ===== SECCIÓN DE CONSOLE/LOG =====
        log_frame = ttk.LabelFrame(main_frame, text="Consola", padding="10")
        log_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        self.console = scrolledtext.ScrolledText(log_frame, height=15, wrap=tk.WORD, bg="black", fg="white", font=("Courier", 9))
        self.console.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        
        # Redireccionar print a la consola
        sys.stdout = ConsoleRedirector(self.console)
        sys.stderr = ConsoleRedirector(self.console)
        
    def cargar_archivo(self):
        file_path = filedialog.askopenfilename(
            title="Seleccionar archivo CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                self.curso_codigo_data = []
                with open(file_path, 'r', encoding='utf-8') as f:
                    reader = csv.reader(f)
                    # Saltar encabezado si existe
                    next(reader, None)
                    for row in reader:
                        if row:  # Ignorar filas vacías
                            self.curso_codigo_data.append(row)
                
                if not self.curso_codigo_data:
                    raise ValueError("El archivo CSV está vacío")
                
                self.file_label.config(text=f"✓ Archivo: {os.path.basename(file_path)}", foreground="green")
                self.info_label.config(text=f"Registros cargados: {len(self.curso_codigo_data)}")
                print(f"Archivo cargado exitosamente: {file_path}")
                print(f"Total de registros: {len(self.curso_codigo_data)}\n")
                
            except Exception as e:
                messagebox.showerror("Error", f"Error al cargar el archivo: {str(e)}")
                print(f"Error: {str(e)}")
    
    def ver_datos(self):
        if not self.curso_codigo_data:
            messagebox.showwarning("Advertencia", "No hay datos cargados")
            return
        
        # Crear ventana emergente
        data_window = tk.Toplevel(self.root)
        data_window.title("Datos Cargados")
        data_window.geometry("600x400")
        
        # Tabla de datos
        columns = ("Código Curso", "Código Alumno", "Observaciones")
        tree = ttk.Treeview(data_window, columns=columns, height=15)
        tree.column("#0", width=0, stretch=tk.NO)
        tree.column("Código Curso", anchor=tk.W, width=150)
        tree.column("Código Alumno", anchor=tk.W, width=150)
        tree.column("Observaciones", anchor=tk.W, width=250)
        
        tree.heading("#0", text="", anchor=tk.W)
        tree.heading("Código Curso", text="Código Curso", anchor=tk.W)
        tree.heading("Código Alumno", text="Código Alumno", anchor=tk.W)
        tree.heading("Observaciones", text="Observaciones", anchor=tk.W)
        
        for i, registro in enumerate(self.curso_codigo_data):
            if isinstance(registro, (list, tuple)) and len(registro) >= 3:
                tree.insert(parent='', index='end', iid=i, text='',
                           values=(registro[0], registro[1], registro[2][:50]))
        
        tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(data_window, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
    
    def iniciar_proceso(self):
        # Validar datos
        if not self.username.get():
            messagebox.showerror("Error", "Ingrese el usuario")
            return
        if not self.password.get():
            messagebox.showerror("Error", "Ingrese la contraseña")
            return
        if not self.curso_codigo_data:
            messagebox.showerror("Error", "Cargue un archivo de datos")
            return
        
        # Deshabilitar botones
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.is_running = True
        
        # Ejecutar en thread separado
        thread = threading.Thread(target=self.ejecutar_clasificacion)
        thread.daemon = True
        thread.start()
    
    def ejecutar_clasificacion(self):
        try:
            self.status_label.config(text="Estado: Ejecutando...", foreground="blue")
            print(f"\n{'='*60}")
            print(f"Iniciando proceso de clasificación en lote")
            print(f"Fecha y hora: {fecha_y_hora_actual()}")
            print(f"{'='*60}\n")
            
            with sync_playwright() as playwright:
                respuesta = funcion_registro_clasificacion(
                    playwright,
                    self.url_pandora.get(),
                    self.username.get(),
                    self.password.get(),
                    self.curso_codigo_data
                )
                
                # Mostrar resultados
                print(f"\n{'='*60}")
                print("RESULTADOS DEL PROCESO:")
                print(f"{'='*60}")
                for resultado in respuesta:
                    print(f"Alumno: {resultado[0]} | Curso: {resultado[1]} | Estado: {resultado[2]}")
                print(f"{'='*60}")
                print(f"Proceso finalizado: {fecha_y_hora_actual()}\n")
                
                self.status_label.config(text="Estado: Completado ✓", foreground="green")
                messagebox.showinfo("Éxito", "Proceso de clasificación completado exitosamente")
                
        except Exception as e:
            print(f"\n[ERROR] {str(e)}")
            self.status_label.config(text=f"Estado: Error", foreground="red")
            messagebox.showerror("Error", f"Error en el proceso: {str(e)}")
        
        finally:
            self.start_btn.config(state=tk.NORMAL)
            self.stop_btn.config(state=tk.DISABLED)
            self.is_running = False
    
    def detener_proceso(self):
        self.is_running = False
        self.status_label.config(text="Estado: Detenido", foreground="orange")
        print("\n[ADVERTENCIA] Proceso detenido por el usuario")
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
    
    def limpiar_consola(self):
        self.console.delete('1.0', tk.END)


class ConsoleRedirector:
    """Redirige salida de consola a la ventana de texto"""
    def __init__(self, text_widget):
        self.text_widget = text_widget
    
    def write(self, message):
        self.text_widget.insert(tk.END, message)
        self.text_widget.see(tk.END)
        self.text_widget.update()
    
    def flush(self):
        pass


if __name__ == "__main__":
    root = tk.Tk()
    app = ClasificacionAppLote(root)
    root.mainloop()
