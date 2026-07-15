import customtkinter as ctk

# Configuración inicial de apariencia
ctk.set_appearance_mode("Light")  # Modos: "System" (estándar), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Temas: "blue" (estándar), "green", "dark-blue"

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Configuración de la ventana principal
        self.title("AutoSRT 900/15")
        self.geometry("1000x700")
        self.minsize(900, 600)

        # Configuración del grid layout principal (1 fila, 2 columnas)
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        # Tipografía base
        self.font_title = ctk.CTkFont(family="Inter", size=24, weight="bold")
        self.font_subtitle = ctk.CTkFont(family="Inter", size=18, weight="bold")
        self.font_normal = ctk.CTkFont(family="Inter", size=14)
        
        # Colores personalizados
        self.sidebar_color = "#1E293B"  # Slate 800 (Azul noche / gris oscuro)
        self.sidebar_text_color = "#F8FAFC"
        self.content_bg_color = "#F1F5F9"  # Slate 100 (Gris claro)
        
        # Configurar color de fondo del área de contenido
        self.configure(fg_color=self.content_bg_color)

        # ==================== MENU LATERAL ====================
        self.sidebar_frame = ctk.CTkFrame(self, width=220, corner_radius=0, fg_color=self.sidebar_color)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(5, weight=1) # Espacio vacío abajo

        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="AutoSRT 900/15", font=self.font_title, text_color=self.sidebar_text_color)
        self.logo_label.grid(row=0, column=0, padx=20, pady=(30, 30))

        # Botones del menú lateral
        self.btn_datos = ctk.CTkButton(self.sidebar_frame, text="Datos Generales", font=self.font_normal,
                                       fg_color="transparent", text_color=self.sidebar_text_color, hover_color="#334155",
                                       anchor="w", command=lambda: self.select_frame("datos"))
        self.btn_datos.grid(row=1, column=0, padx=10, pady=5, sticky="ew")

        self.btn_pat = ctk.CTkButton(self.sidebar_frame, text="Mediciones PAT", font=self.font_normal,
                                     fg_color="transparent", text_color=self.sidebar_text_color, hover_color="#334155",
                                     anchor="w", command=lambda: self.select_frame("pat"))
        self.btn_pat.grid(row=2, column=0, padx=10, pady=5, sticky="ew")

        self.btn_dif = ctk.CTkButton(self.sidebar_frame, text="Diferenciales", font=self.font_normal,
                                     fg_color="transparent", text_color=self.sidebar_text_color, hover_color="#334155",
                                     anchor="w", command=lambda: self.select_frame("dif"))
        self.btn_dif.grid(row=3, column=0, padx=10, pady=5, sticky="ew")

        self.btn_reporte = ctk.CTkButton(self.sidebar_frame, text="Generar Reporte", font=self.font_normal,
                                         fg_color="transparent", text_color=self.sidebar_text_color, hover_color="#334155",
                                         anchor="w", command=lambda: self.select_frame("reporte"))
        self.btn_reporte.grid(row=4, column=0, padx=10, pady=5, sticky="ew")


        # ==================== VISTAS (FRAMES) ====================
        
        # 1. Vista Datos Generales
        self.frame_datos = ctk.CTkFrame(self, corner_radius=10, fg_color="white")
        self.setup_frame_datos()

        # 2. Vista Mediciones PAT
        self.frame_pat = ctk.CTkFrame(self, corner_radius=10, fg_color="white")
        self.setup_frame_pat()

        # 3. Vista Diferenciales
        self.frame_dif = ctk.CTkFrame(self, corner_radius=10, fg_color="white")
        self.setup_frame_dif()

        # 4. Vista Generar Reporte
        self.frame_reporte = ctk.CTkFrame(self, corner_radius=10, fg_color="white")
        self.setup_frame_reporte()

        # Seleccionar la primera vista por defecto
        self.select_frame("datos")

    # --- Métodos de configuración de cada vista ---

    def setup_frame_datos(self):
        self.frame_datos.grid_columnconfigure((0, 1), weight=1)
        
        lbl_titulo = ctk.CTkLabel(self.frame_datos, text="Datos Generales del Cliente", font=self.font_subtitle)
        lbl_titulo.grid(row=0, column=0, columnspan=2, padx=30, pady=(30, 20), sticky="w")

        # Razón Social
        ctk.CTkLabel(self.frame_datos, text="Razón Social:", font=self.font_normal).grid(row=1, column=0, padx=30, pady=(10, 0), sticky="w")
        self.entry_razon_social = ctk.CTkEntry(self.frame_datos, placeholder_text="Ej. Empresa S.A.", font=self.font_normal)
        self.entry_razon_social.grid(row=2, column=0, padx=30, pady=(0, 20), sticky="ew")

        # CUIT
        ctk.CTkLabel(self.frame_datos, text="CUIT:", font=self.font_normal).grid(row=1, column=1, padx=30, pady=(10, 0), sticky="w")
        self.entry_cuit = ctk.CTkEntry(self.frame_datos, placeholder_text="Ej. 30-12345678-9", font=self.font_normal)
        self.entry_cuit.grid(row=2, column=1, padx=30, pady=(0, 20), sticky="ew")

        # Dirección
        ctk.CTkLabel(self.frame_datos, text="Dirección:", font=self.font_normal).grid(row=3, column=0, padx=30, pady=(10, 0), sticky="w")
        self.entry_direccion = ctk.CTkEntry(self.frame_datos, placeholder_text="Ej. Av. Ejemplo 742", font=self.font_normal)
        self.entry_direccion.grid(row=4, column=0, padx=30, pady=(0, 20), sticky="ew")

        # Localidad
        ctk.CTkLabel(self.frame_datos, text="Localidad:", font=self.font_normal).grid(row=3, column=1, padx=30, pady=(10, 0), sticky="w")
        self.entry_localidad = ctk.CTkEntry(self.frame_datos, placeholder_text="Ej. Capital", font=self.font_normal)
        self.entry_localidad.grid(row=4, column=1, padx=30, pady=(0, 20), sticky="ew")

        # Provincia (ComboBox)
        ctk.CTkLabel(self.frame_datos, text="Provincia:", font=self.font_normal).grid(row=5, column=0, padx=30, pady=(10, 0), sticky="w")
        provincias = ["San Juan", "Buenos Aires", "CABA", "Córdoba", "Santa Fe", "Mendoza", "Tucumán"] # Completar...
        self.combo_provincia = ctk.CTkComboBox(self.frame_datos, values=provincias, font=self.font_normal)
        self.combo_provincia.grid(row=6, column=0, padx=30, pady=(0, 20), sticky="ew")

        # N° de Contrato
        ctk.CTkLabel(self.frame_datos, text="N° de Contrato:", font=self.font_normal).grid(row=5, column=1, padx=30, pady=(10, 0), sticky="w")
        self.entry_contrato = ctk.CTkEntry(self.frame_datos, placeholder_text="Ej. 12345/26", font=self.font_normal)
        self.entry_contrato.grid(row=6, column=1, padx=30, pady=(0, 20), sticky="ew")

        # Botón Siguiente
        btn_siguiente = ctk.CTkButton(self.frame_datos, text="Siguiente ➔", font=self.font_normal, width=120, 
                                      command=lambda: self.select_frame("pat"))
        btn_siguiente.grid(row=7, column=1, padx=30, pady=(40, 30), sticky="e")


    def setup_frame_pat(self):
        self.frame_pat.grid_columnconfigure((0, 1, 2, 3, 4), weight=1)
        self.frame_pat.grid_rowconfigure(3, weight=1)

        lbl_titulo = ctk.CTkLabel(self.frame_pat, text="Carga de Mediciones de Puesta a Tierra", font=self.font_subtitle)
        lbl_titulo.grid(row=0, column=0, columnspan=5, padx=30, pady=(30, 20), sticky="w")

        # --- Entradas para nueva medición ---
        # Sector
        self.pat_sector = ctk.CTkEntry(self.frame_pat, placeholder_text="Sector", font=self.font_normal)
        self.pat_sector.grid(row=1, column=0, padx=(30, 5), pady=10, sticky="ew")
        
        # Condición
        self.pat_condicion = ctk.CTkComboBox(self.frame_pat, values=["Normal", "Húmedo", "Seco"], font=self.font_normal)
        self.pat_condicion.grid(row=1, column=1, padx=5, pady=10, sticky="ew")

        # Esquema
        self.pat_esquema = ctk.CTkComboBox(self.frame_pat, values=["TT", "TN", "IT"], font=self.font_normal)
        self.pat_esquema.grid(row=1, column=2, padx=5, pady=10, sticky="ew")

        # Valor (Ohms)
        self.pat_valor = ctk.CTkEntry(self.frame_pat, placeholder_text="Valor (Ω)", font=self.font_normal)
        self.pat_valor.grid(row=1, column=3, padx=5, pady=10, sticky="ew")

        # Cumple
        self.pat_cumple = ctk.CTkComboBox(self.frame_pat, values=["SI", "NO", "N/A"], font=self.font_normal)
        self.pat_cumple.grid(row=1, column=4, padx=(5, 30), pady=10, sticky="ew")

        # Botón Agregar
        btn_agregar_pat = ctk.CTkButton(self.frame_pat, text="Agregar Medición", font=self.font_normal,
                                        command=self.agregar_medicion_pat)
        btn_agregar_pat.grid(row=2, column=3, columnspan=2, padx=30, pady=10, sticky="e")

        # --- Tabla (ScrollableFrame imitando tabla) ---
        self.scroll_pat = ctk.CTkScrollableFrame(self.frame_pat, fg_color="#F8FAFC", corner_radius=5)
        self.scroll_pat.grid(row=3, column=0, columnspan=5, padx=30, pady=(10, 20), sticky="nsew")
        self.scroll_pat.grid_columnconfigure((0,1,2,3,4), weight=1)
        
        # Encabezados de tabla
        encabezados = ["Sector", "Condición", "Esquema", "Valor (Ω)", "Cumple"]
        for i, texto in enumerate(encabezados):
            lbl = ctk.CTkLabel(self.scroll_pat, text=texto, font=ctk.CTkFont(family="Inter", size=13, weight="bold"))
            lbl.grid(row=0, column=i, padx=5, pady=5, sticky="w")

        # TODO: Cargar lista de mediciones desde backend aquí
        # Para simular, agregamos un elemento vacío por ahora
        self.fila_actual_pat = 1

        # Botón Eliminar Seleccionado
        btn_eliminar_pat = ctk.CTkButton(self.frame_pat, text="Eliminar Última", font=self.font_normal,
                                         fg_color="#EF4444", hover_color="#DC2626") # Rojo
        # TODO: Conectar command con backend para eliminar
        btn_eliminar_pat.grid(row=4, column=0, padx=30, pady=(0, 30), sticky="w")


    def setup_frame_dif(self):
        self.frame_dif.grid_columnconfigure((0, 1, 2, 3, 4), weight=1)
        self.frame_dif.grid_rowconfigure(3, weight=1)

        lbl_titulo = ctk.CTkLabel(self.frame_dif, text="Ensayos de Disyuntores Diferenciales", font=self.font_subtitle)
        lbl_titulo.grid(row=0, column=0, columnspan=5, padx=30, pady=(30, 20), sticky="w")

        # --- Entradas para nuevo ensayo ---
        # Ubicación
        self.dif_ubicacion = ctk.CTkEntry(self.frame_dif, placeholder_text="Ubicación/Sector", font=self.font_normal)
        self.dif_ubicacion.grid(row=1, column=0, padx=(30, 5), pady=10, sticky="ew")
        
        # Marca
        self.dif_marca = ctk.CTkEntry(self.frame_dif, placeholder_text="Marca", font=self.font_normal)
        self.dif_marca.grid(row=1, column=1, padx=5, pady=10, sticky="ew")

        # Sensibilidad (mA)
        self.dif_sensibilidad = ctk.CTkComboBox(self.frame_dif, values=["30", "10", "300"], font=self.font_normal)
        self.dif_sensibilidad.grid(row=1, column=2, padx=5, pady=10, sticky="ew")

        # Tiempo (ms)
        self.dif_tiempo = ctk.CTkEntry(self.frame_dif, placeholder_text="Tiempo (ms)", font=self.font_normal)
        self.dif_tiempo.grid(row=1, column=3, padx=5, pady=10, sticky="ew")

        # Cumple
        self.dif_cumple = ctk.CTkComboBox(self.frame_dif, values=["SI", "NO", "N/A"], font=self.font_normal)
        self.dif_cumple.grid(row=1, column=4, padx=(5, 30), pady=10, sticky="ew")

        # Botón Agregar
        btn_agregar_dif = ctk.CTkButton(self.frame_dif, text="Agregar Ensayo", font=self.font_normal,
                                        command=self.agregar_medicion_dif)
        btn_agregar_dif.grid(row=2, column=3, columnspan=2, padx=30, pady=10, sticky="e")

        # --- Tabla (ScrollableFrame imitando tabla) ---
        self.scroll_dif = ctk.CTkScrollableFrame(self.frame_dif, fg_color="#F8FAFC", corner_radius=5)
        self.scroll_dif.grid(row=3, column=0, columnspan=5, padx=30, pady=(10, 20), sticky="nsew")
        self.scroll_dif.grid_columnconfigure((0,1,2,3,4), weight=1)
        
        # Encabezados de tabla
        encabezados = ["Ubicación", "Marca", "Sens. (mA)", "Tiempo (ms)", "Cumple"]
        for i, texto in enumerate(encabezados):
            lbl = ctk.CTkLabel(self.scroll_dif, text=texto, font=ctk.CTkFont(family="Inter", size=13, weight="bold"))
            lbl.grid(row=0, column=i, padx=5, pady=5, sticky="w")

        # TODO: Cargar lista de ensayos desde backend aquí
        self.fila_actual_dif = 1

        # Botón Eliminar Seleccionado
        btn_eliminar_dif = ctk.CTkButton(self.frame_dif, text="Eliminar Último", font=self.font_normal,
                                         fg_color="#EF4444", hover_color="#DC2626") # Rojo
        # TODO: Conectar command con backend para eliminar
        btn_eliminar_dif.grid(row=4, column=0, padx=30, pady=(0, 30), sticky="w")


    def setup_frame_reporte(self):
        self.frame_reporte.grid_columnconfigure(0, weight=1)
        self.frame_reporte.grid_rowconfigure(2, weight=1)
        
        lbl_titulo = ctk.CTkLabel(self.frame_reporte, text="Generar Excel", font=self.font_subtitle)
        lbl_titulo.grid(row=0, column=0, padx=30, pady=(30, 20))

        lbl_desc = ctk.CTkLabel(self.frame_reporte, 
                                text="Todos los datos han sido validados. Al presionar el botón,\n"
                                     "se inyectarán en la plantilla 'Mediciones 2026.xlsx' y se generará el reporte final.",
                                font=self.font_normal, justify="center")
        lbl_desc.grid(row=1, column=0, padx=30, pady=20)

        # Botón Generar
        self.btn_generar = ctk.CTkButton(self.frame_reporte, text="Generar y Descargar Protocolo", 
                                         font=ctk.CTkFont(family="Inter", size=16, weight="bold"),
                                         height=50, width=250,
                                         command=self.generar_reporte)
        self.btn_generar.grid(row=2, column=0, pady=20)

        # Consola de estado
        ctk.CTkLabel(self.frame_reporte, text="Estado del proceso:", font=ctk.CTkFont(family="Inter", size=12, weight="bold")).grid(row=3, column=0, padx=30, pady=(20,0), sticky="w")
        self.textbox_log = ctk.CTkTextbox(self.frame_reporte, height=150, font=ctk.CTkFont(family="Consolas", size=12), state="disabled", fg_color="#F8FAFC")
        self.textbox_log.grid(row=4, column=0, padx=30, pady=(5, 30), sticky="ew")

    # --- Lógica de la GUI ---

    def select_frame(self, name):
        # Actualizar colores de botones del menú
        color_activo = "#334155" # Color cuando está seleccionado
        color_inactivo = "transparent"

        self.btn_datos.configure(fg_color=color_activo if name == "datos" else color_inactivo)
        self.btn_pat.configure(fg_color=color_activo if name == "pat" else color_inactivo)
        self.btn_dif.configure(fg_color=color_activo if name == "dif" else color_inactivo)
        self.btn_reporte.configure(fg_color=color_activo if name == "reporte" else color_inactivo)

        # Mostrar frame seleccionado y ocultar los demás
        if name == "datos":
            self.frame_datos.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")
        else:
            self.frame_datos.grid_forget()

        if name == "pat":
            self.frame_pat.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")
        else:
            self.frame_pat.grid_forget()

        if name == "dif":
            self.frame_dif.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")
        else:
            self.frame_dif.grid_forget()

        if name == "reporte":
            self.frame_reporte.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")
        else:
            self.frame_reporte.grid_forget()

    # --- Acciones (Simulaciones y TODOs para backend) ---

    def agregar_medicion_pat(self):
        # Obtener valores
        sector = self.pat_sector.get()
        cond = self.pat_condicion.get()
        esq = self.pat_esquema.get()
        val = self.pat_valor.get()
        cumple = self.pat_cumple.get()

        if not sector or not val:
            # Podrías agregar un messagebox aquí
            return

        # TODO: Guardar en la lista del backend (ej. self.datos['mediciones_pat'].append(...))
        
        # Mostrar en la tabla (UI)
        ctk.CTkLabel(self.scroll_pat, text=sector, font=self.font_normal).grid(row=self.fila_actual_pat, column=0, padx=5, pady=2, sticky="w")
        ctk.CTkLabel(self.scroll_pat, text=cond, font=self.font_normal).grid(row=self.fila_actual_pat, column=1, padx=5, pady=2, sticky="w")
        ctk.CTkLabel(self.scroll_pat, text=esq, font=self.font_normal).grid(row=self.fila_actual_pat, column=2, padx=5, pady=2, sticky="w")
        ctk.CTkLabel(self.scroll_pat, text=val, font=self.font_normal).grid(row=self.fila_actual_pat, column=3, padx=5, pady=2, sticky="w")
        ctk.CTkLabel(self.scroll_pat, text=cumple, font=self.font_normal).grid(row=self.fila_actual_pat, column=4, padx=5, pady=2, sticky="w")
        
        self.fila_actual_pat += 1
        
        # Limpiar campos
        self.pat_sector.delete(0, 'end')
        self.pat_valor.delete(0, 'end')

    def agregar_medicion_dif(self):
        ubicacion = self.dif_ubicacion.get()
        marca = self.dif_marca.get()
        sens = self.dif_sensibilidad.get()
        tiempo = self.dif_tiempo.get()
        cumple = self.dif_cumple.get()

        if not ubicacion or not tiempo:
            return

        # TODO: Guardar en la lista del backend
        
        # Mostrar en la tabla (UI)
        ctk.CTkLabel(self.scroll_dif, text=ubicacion, font=self.font_normal).grid(row=self.fila_actual_dif, column=0, padx=5, pady=2, sticky="w")
        ctk.CTkLabel(self.scroll_dif, text=marca, font=self.font_normal).grid(row=self.fila_actual_dif, column=1, padx=5, pady=2, sticky="w")
        ctk.CTkLabel(self.scroll_dif, text=sens, font=self.font_normal).grid(row=self.fila_actual_dif, column=2, padx=5, pady=2, sticky="w")
        ctk.CTkLabel(self.scroll_dif, text=tiempo, font=self.font_normal).grid(row=self.fila_actual_dif, column=3, padx=5, pady=2, sticky="w")
        ctk.CTkLabel(self.scroll_dif, text=cumple, font=self.font_normal).grid(row=self.fila_actual_dif, column=4, padx=5, pady=2, sticky="w")
        
        self.fila_actual_dif += 1
        
        # Limpiar campos
        self.dif_ubicacion.delete(0, 'end')
        self.dif_marca.delete(0, 'end')
        self.dif_tiempo.delete(0, 'end')

    def log_message(self, message):
        self.textbox_log.configure(state="normal")
        self.textbox_log.insert("end", message + "\n")
        self.textbox_log.see("end")
        self.textbox_log.configure(state="disabled")

    def generar_reporte(self):
        self.log_message("Iniciando generación de reporte...")
        self.log_message("Recolectando datos de la interfaz...")
        
        # TODO: Aquí llamas a tu función `generar_protocolo()` del archivo `excel_builder.py`
        # pasándole los datos recolectados de la GUI en lugar del `get_mock_data()`.
        
        self.log_message("Abriendo plantilla 'Mediciones 2026.xlsx'...")
        self.log_message("Escribiendo mediciones PAT y Diferenciales...")
        self.log_message("Guardando archivo final...")
        self.log_message("¡Completado! Protocolo_SRT_Generado.xlsx generado con éxito.")


if __name__ == "__main__":
    app = App()
    app.mainloop()
