import flet as ft
import traceback

# Import defensivo: en Android, openpyxl puede fallar al cargar.
# Si falla, la app muestra el error en pantalla en vez de pantalla negra.
try:
    from excel_builder import generar_protocolo
    _excel_disponible = True
    _excel_error = None
except Exception as _e:
    _excel_disponible = False
    _excel_error = traceback.format_exc()

def main(page: ft.Page):
    page.title = "AutoSRT 900/15"
    page.theme_mode = ft.ThemeMode.LIGHT
    page.scroll = ft.ScrollMode.AUTO

    # Si openpyxl no pudo cargarse, mostrar el error en pantalla
    if not _excel_disponible:
        page.add(
            ft.Container(
                content=ft.Column([
                    ft.Icon(ft.icons.ERROR, color=ft.colors.RED, size=48),
                    ft.Text("Error al iniciar la aplicación", size=20, weight=ft.FontWeight.BOLD, color=ft.colors.RED),
                    ft.Text("No se pudo cargar el módulo de generación de Excel:", size=14),
                    ft.Text(_excel_error, size=11, selectable=True, color=ft.colors.ORANGE_700),
                ], scroll=ft.ScrollMode.AUTO),
                padding=20,
                expand=True,
            )
        )
        return

    # ================= ESTADO DE LA APLICACIÓN =================
    state = {
        "datos_cliente": {
            "razon_social": "", "cuit": "", "direccion": "",
            "localidad": "", "provincia": "Buenos Aires", "contrato": ""
        },
        "mediciones_pat": [],
        "mediciones_dif": [],
        "fila_actual_pat": 1
    }

    # ================= COMPONENTES COMPARTIDOS =================
    
    def on_file_saved(e: ft.FilePickerResultEvent):
        if e.path:
            # e.path contiene la ruta donde el usuario eligió guardar
            try:
                # Armar diccionario final
                datos_final = {
                    "datos_cliente": state["datos_cliente"],
                    "mediciones_pat": state["mediciones_pat"],
                    "mediciones_diferenciales": state["mediciones_dif"],
                    "conclusiones": {
                        "observaciones": "Generado desde AutoSRT Móvil (Flet)",
                        "fecha_medicion": "",
                        "instrumento_utilizado": ""
                    }
                }
                # Llamar a excel_builder
                generar_protocolo("Mediciones 2026.xlsx", e.path, datos_final)
                
                snack = ft.SnackBar(ft.Text(f"¡Protocolo guardado con éxito!"), bgcolor=ft.colors.GREEN_700)
                page.overlay.append(snack)
                snack.open = True
                page.update()
            except Exception as ex:
                snack = ft.SnackBar(ft.Text(f"Error: {str(ex)}"), bgcolor=ft.colors.RED_700)
                page.overlay.append(snack)
                snack.open = True
                page.update()

    file_picker = ft.FilePicker(on_result=on_file_saved)
    page.overlay.append(file_picker)

    # ================= VISTAS =================

    # --- Vista: Datos Generales ---
    def update_datos(e):
        state["datos_cliente"]["razon_social"] = input_razon.value
        state["datos_cliente"]["cuit"] = input_cuit.value
        state["datos_cliente"]["direccion"] = input_dir.value
        state["datos_cliente"]["localidad"] = input_loc.value
        state["datos_cliente"]["provincia"] = input_prov.value
        state["datos_cliente"]["contrato"] = input_cont.value

    input_razon = ft.TextField(label="Razón Social", on_change=update_datos)
    input_cuit = ft.TextField(label="CUIT", on_change=update_datos)
    input_dir = ft.TextField(label="Dirección", on_change=update_datos)
    input_loc = ft.TextField(label="Localidad", on_change=update_datos)
    input_prov = ft.Dropdown(
        label="Provincia",
        options=[ft.dropdown.Option(x) for x in ["Buenos Aires", "CABA", "Córdoba", "Santa Fe", "Mendoza", "Tucumán"]],
        value="Buenos Aires",
        on_change=update_datos
    )
    input_cont = ft.TextField(label="N° de Contrato", on_change=update_datos)

    view_datos = ft.Column([
        ft.Text("Datos Generales del Cliente", size=20, weight=ft.FontWeight.BOLD),
        input_razon, input_cuit, input_dir, input_loc, input_prov, input_cont
    ], visible=True)

    # --- Vista: PAT ---
    pat_sector = ft.TextField(label="Sector", expand=True)
    pat_cond = ft.Dropdown(label="Cond.", options=[ft.dropdown.Option(x) for x in ["Normal", "Húmedo", "Seco"]], value="Normal", expand=True)
    pat_esq = ft.Dropdown(label="Esq.", options=[ft.dropdown.Option(x) for x in ["TT", "TN", "IT"]], value="TT", expand=True)
    pat_val = ft.TextField(label="Valor (Ω)", expand=True)
    pat_cumple = ft.Dropdown(label="Cumple", options=[ft.dropdown.Option(x) for x in ["SI", "NO", "N/A"]], value="SI", expand=True)
    
    pat_list_view = ft.ListView(expand=True, spacing=10, height=300)

    def refrescar_lista_pat():
        pat_list_view.controls.clear()
        for i, m in enumerate(state["mediciones_pat"]):
            pat_list_view.controls.append(
                ft.ListTile(
                    title=ft.Text(f"{m['sector']} - {m['valor_ohms']}Ω"),
                    subtitle=ft.Text(f"Cond: {m['condicion_terreno']} | Esq: {m['esquema']} | Cumple: {m['cumple']}"),
                    leading=ft.Icon(ft.icons.ELECTRIC_BOLT),
                )
            )
        page.update()

    def add_pat(e):
        if not pat_sector.value or not pat_val.value:
            snack = ft.SnackBar(ft.Text("Sector y Valor son obligatorios"), bgcolor=ft.colors.RED_700)
            page.overlay.append(snack)
            snack.open = True
            page.update()
            return
            
        state["mediciones_pat"].append({
            "nro_toma": state["fila_actual_pat"],
            "sector": pat_sector.value,
            "condicion_terreno": pat_cond.value,
            "esquema": pat_esq.value,
            "valor_ohms": pat_val.value,
            "cumple": pat_cumple.value
        })
        state["fila_actual_pat"] += 1
        
        pat_sector.value = ""
        pat_val.value = ""
        refrescar_lista_pat()
        
    def remove_pat(e):
        if state["mediciones_pat"]:
            state["mediciones_pat"].pop()
            state["fila_actual_pat"] -= 1
            refrescar_lista_pat()

    view_pat = ft.Column([
        ft.Text("Puesta a Tierra", size=20, weight=ft.FontWeight.BOLD),
        pat_sector,
        ft.Row([pat_cond, pat_esq]),
        ft.Row([pat_val, pat_cumple]),
        ft.Row([
            ft.ElevatedButton("Agregar", on_click=add_pat, icon=ft.icons.ADD),
            ft.ElevatedButton("Borrar Último", on_click=remove_pat, icon=ft.icons.DELETE, color=ft.colors.RED)
        ]),
        ft.Divider(),
        ft.Text("Registros:", weight=ft.FontWeight.BOLD),
        pat_list_view
    ], visible=False)

    # --- Vista: Diferenciales ---
    dif_sec = ft.TextField(label="Ubicación/Sector", expand=True)
    dif_marca = ft.TextField(label="Marca", expand=True)
    dif_sens = ft.Dropdown(label="Sens.(mA)", options=[ft.dropdown.Option(x) for x in ["30", "10", "300"]], value="30", expand=True)
    dif_tiempo = ft.TextField(label="Tiempo(ms)", expand=True)
    dif_cumple = ft.Dropdown(label="Cumple", options=[ft.dropdown.Option(x) for x in ["SI", "NO", "N/A"]], value="SI", expand=True)
    
    dif_list_view = ft.ListView(expand=True, spacing=10, height=300)

    def refrescar_lista_dif():
        dif_list_view.controls.clear()
        for m in state["mediciones_dif"]:
            dif_list_view.controls.append(
                ft.ListTile(
                    title=ft.Text(f"{m['sector']} - {m['tiempo_disparo_ms']}ms"),
                    subtitle=ft.Text(f"Marca: {m['marca']} | Sens: {m['corriente_fuga_ma']}mA | Cumple: {m['cumple']}"),
                    leading=ft.Icon(ft.icons.SHIELD),
                )
            )
        page.update()

    def add_dif(e):
        if not dif_sec.value or not dif_tiempo.value:
            snack = ft.SnackBar(ft.Text("Ubicación y Tiempo son obligatorios"), bgcolor=ft.colors.RED_700)
            page.overlay.append(snack)
            snack.open = True
            page.update()
            return
            
        state["mediciones_dif"].append({
            "sector": dif_sec.value,
            "marca": dif_marca.value,
            "corriente_nominal_a": "",
            "corriente_fuga_ma": dif_sens.value,
            "tiempo_disparo_ms": dif_tiempo.value,
            "cumple": dif_cumple.value
        })
        
        dif_sec.value = ""
        dif_marca.value = ""
        dif_tiempo.value = ""
        refrescar_lista_dif()
        
    def remove_dif(e):
        if state["mediciones_dif"]:
            state["mediciones_dif"].pop()
            refrescar_lista_dif()

    view_dif = ft.Column([
        ft.Text("Diferenciales", size=20, weight=ft.FontWeight.BOLD),
        dif_sec, dif_marca,
        ft.Row([dif_sens, dif_tiempo]),
        dif_cumple,
        ft.Row([
            ft.ElevatedButton("Agregar", on_click=add_dif, icon=ft.icons.ADD),
            ft.ElevatedButton("Borrar Último", on_click=remove_dif, icon=ft.icons.DELETE, color=ft.colors.RED)
        ]),
        ft.Divider(),
        ft.Text("Registros:", weight=ft.FontWeight.BOLD),
        dif_list_view
    ], visible=False)

    # --- Vista: Reporte ---
    def guardar_excel(e):
        # Al presionar, abre el diálogo para elegir dónde guardar
        file_picker.save_file(
            dialog_title="Guardar Protocolo SRT",
            file_name="Protocolo_SRT_Generado.xlsx",
            file_type=ft.FilePickerFileType.CUSTOM,
            allowed_extensions=["xlsx"]
        )

    view_reporte = ft.Column([
        ft.Text("Generar Excel", size=20, weight=ft.FontWeight.BOLD),
        ft.Text("Se generará el archivo inyectando todas las mediciones en la plantilla base."),
        ft.Container(height=30),
        ft.ElevatedButton(
            "Generar y Guardar",
            icon=ft.icons.SAVE_ALT,
            on_click=guardar_excel,
            style=ft.ButtonStyle(
                color=ft.colors.WHITE,
                bgcolor=ft.colors.BLUE_700,
                padding=20
            ),
            width=300
        )
    ], visible=False, horizontal_alignment=ft.CrossAxisAlignment.CENTER)

    # ================= NAVEGACIÓN =================
    views = [view_datos, view_pat, view_dif, view_reporte]

    def on_nav_change(e):
        # Ocultar todas
        for v in views:
            v.visible = False
        # Mostrar la seleccionada
        views[e.control.selected_index].visible = True
        page.update()

    page.navigation_bar = ft.NavigationBar(
        destinations=[
            ft.NavigationDestination(icon=ft.icons.PERSON, label="Datos"),
            ft.NavigationDestination(icon=ft.icons.ELECTRIC_BOLT, label="PAT"),
            ft.NavigationDestination(icon=ft.icons.SHIELD, label="Dif"),
            ft.NavigationDestination(icon=ft.icons.PICTURE_AS_PDF, label="Reporte"),
        ],
        on_change=on_nav_change,
        selected_index=0
    )

    # Contenedor principal
    page.add(
        ft.Container(
            content=ft.Column([view_datos, view_pat, view_dif, view_reporte], expand=True),
            padding=20,
            expand=True
        )
    )

# ft.app debe estar a nivel de módulo (NO solo dentro de __main__)
# En Android, Flet importa el módulo directamente sin ejecutar __main__,
# por lo que sin esta línea Flutter nunca recibe la función main → pantalla negra.
ft.app(target=main)
