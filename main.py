import datetime
import inspect
import os
import subprocess
import sys
import traceback
from pathlib import Path
from typing import Any, Dict, List, Optional

import flet as ft

from srt_logic import (
    CONDICIONES_TERRENO,
    ESQUEMAS_CONEXION,
    PROTECCIONES_PAT,
    PROVINCIAS_ARGENTINA,
    SECTORES_SUGERIDOS,
    USOS_PAT,
    cargar_borrador,
    cargar_perfil_electricista,
    construir_nombre_archivo,
    directorio_datos_app,
    evaluar_diferencial,
    evaluar_parametro_rcd,
    evaluar_valor_pat,
    generar_textos_sugeridos,
    guardar_borrador,
    guardar_perfil_electricista,
    limpiar_borrador,
    listar_protocolos_xlsx,
    parsear_float,
    resumir_relevamiento,
    verificar_calibracion,
)

# Import defensivo: en Android, openpyxl o pdf_builder pueden fallar si falta alguna dependencia.
try:
    from excel_builder import (
        cargar_protocolo_excel,
        es_medicion_pat_valida,
        generar_protocolo,
        ruta_plantilla_default,
    )
    from pdf_builder import generar_protocolo_pdf

    _modulos_disponibles = True
    _modulos_error = None
except Exception:
    _modulos_disponibles = False
    _modulos_error = traceback.format_exc()


# ==============================================================================
# HELPERS DE COMPATIBILIDAD FLET (0.28.3 Android CI <-> 0.86+ Desktop)
# ==============================================================================

_TF_PARAMS = inspect.signature(ft.TextField.__init__).parameters
_DD_PARAMS = inspect.signature(ft.Dropdown.__init__).parameters
_ET_PARAMS = inspect.signature(ft.ExpansionTile.__init__).parameters
_Btn = getattr(ft, "Button", ft.ElevatedButton)


def _color(nombre: str, fallback: Any = None) -> Any:
    return getattr(ft.Colors, nombre, fallback)


def _borde(width: float = 1, color: Any = None) -> Any:
    c = color or _color("OUTLINE_VARIANT", ft.Colors.GREY_400)
    if hasattr(ft, "Border") and hasattr(ft.Border, "all"):
        return ft.Border.all(width, c)
    return ft.border.all(width, c)


def _padding_sim(horizontal: float = 0, vertical: float = 0) -> Any:
    if hasattr(ft, "Padding") and hasattr(ft.Padding, "symmetric"):
        return ft.Padding.symmetric(horizontal=horizontal, vertical=vertical)
    return ft.padding.symmetric(horizontal=horizontal, vertical=vertical)


def _crear_textfield(
    label: str,
    value: str = "",
    hint_text: Optional[str] = None,
    helper_text: Optional[str] = None,
    error_text: Optional[str] = None,
    suffix_text: Optional[str] = None,
    on_change=None,
    on_submit=None,
    keyboard_type=None,
    multiline: bool = False,
    min_lines: Optional[int] = None,
    max_lines: Optional[int] = None,
    expand: bool = False,
    col: Any = None,
    dense: bool = False,
    text_size: Optional[float] = None,
) -> ft.TextField:
    kwargs: Dict[str, Any] = {
        "label": label,
        "value": value,
        "expand": expand,
        "multiline": multiline,
        "dense": dense,
    }
    if hint_text is not None:
        kwargs["hint_text"] = hint_text
    if helper_text is not None:
        kwargs["helper" if "helper" in _TF_PARAMS else "helper_text"] = helper_text
    if error_text is not None:
        kwargs["error" if "error" in _TF_PARAMS else "error_text"] = error_text
    if suffix_text is not None:
        kwargs["suffix" if "suffix" in _TF_PARAMS else "suffix_text"] = suffix_text
    if on_change is not None:
        kwargs["on_change"] = on_change
    if on_submit is not None:
        kwargs["on_submit"] = on_submit
    if keyboard_type is not None:
        kwargs["keyboard_type"] = keyboard_type
    if min_lines is not None:
        kwargs["min_lines"] = min_lines
    if max_lines is not None:
        kwargs["max_lines"] = max_lines
    if col is not None and "col" in _TF_PARAMS:
        kwargs["col"] = col
    if text_size is not None:
        kwargs["text_size"] = text_size
    return ft.TextField(**kwargs)


def _set_tf_error(tf: ft.TextField, msg: Optional[str]) -> None:
    if "error" in _TF_PARAMS:
        tf.error = msg
    else:
        tf.error_text = msg


def _set_tf_helper(tf: ft.TextField, msg: Optional[str]) -> None:
    if "helper" in _TF_PARAMS:
        tf.helper = msg
    else:
        tf.helper_text = msg


def _crear_dropdown(
    label: str,
    opciones: List[str],
    value: str,
    on_change=None,
    expand: bool = True,
    col: Any = None,
    dense: bool = False,
) -> ft.Dropdown:
    kwargs: Dict[str, Any] = {
        "label": label,
        "options": [ft.dropdown.Option(x, text=x if x else "(Sin especificar)") for x in opciones],
        "value": value if value in opciones else (opciones[0] if opciones else None),
        "expand": expand,
    }
    if dense and "dense" in _DD_PARAMS:
        kwargs["dense"] = dense
    if col is not None and "col" in _DD_PARAMS:
        kwargs["col"] = col
    if on_change is not None:
        if "on_select" in _DD_PARAMS:
            kwargs["on_select"] = on_change
        elif "on_change" in _DD_PARAMS:
            kwargs["on_change"] = on_change
    return ft.Dropdown(**kwargs)


def _crear_expansion_tile(
    title: ft.Control,
    controls: List[ft.Control],
    subtitle: Optional[ft.Control] = None,
    expanded: bool = False,
) -> ft.ExpansionTile:
    kwargs: Dict[str, Any] = {"title": title, "controls": controls}
    if subtitle is not None:
        kwargs["subtitle"] = subtitle
    if "expanded" in _ET_PARAMS:
        kwargs["expanded"] = expanded
    elif "initially_expanded" in _ET_PARAMS:
        kwargs["initially_expanded"] = expanded
    return ft.ExpansionTile(**kwargs)


def _set_btn_text(btn: Any, texto: str) -> None:
    if hasattr(btn, "text") and not isinstance(getattr(btn, "content", None), ft.Control):
        btn.text = texto
    if hasattr(btn, "content"):
        if isinstance(btn.content, ft.Text):
            btn.content.value = texto
        elif isinstance(btn.content, str) or not hasattr(btn, "text"):
            btn.content = texto


def _set_icon(ctrl: ft.Icon, icon_val: Any) -> None:
    if hasattr(ctrl, "icon"):
        ctrl.icon = icon_val
    else:
        ctrl.name = icon_val


def _align_center() -> Any:
    if hasattr(ft, "Alignment") and hasattr(ft.Alignment, "CENTER"):
        return ft.Alignment.CENTER
    if hasattr(ft, "alignment") and hasattr(ft.alignment, "center"):
        return ft.alignment.center
    return None


def _enfocar(page: ft.Page, ctrl: ft.Control) -> None:
    try:
        if inspect.iscoroutinefunction(ctrl.focus):
            if hasattr(page, "run_task"):
                page.run_task(ctrl.focus)
        else:
            ctrl.focus()
    except Exception:
        pass


def _abrir_dialogo(page: ft.Page, dlg: ft.AlertDialog) -> None:
    if hasattr(page, "open"):
        page.open(dlg)
    elif hasattr(page, "show_dialog"):
        page.show_dialog(dlg)
    else:
        page.overlay.append(dlg)
        dlg.open = True
        page.update()


def _cerrar_dialogo(page: ft.Page, dlg: ft.AlertDialog) -> None:
    if hasattr(page, "close"):
        page.close(dlg)
    else:
        dlg.open = False
        try:
            dlg.update()
        except Exception:
            page.update()


def _card(
    titulo: str,
    contenido: ft.Control,
    icono: Any = None,
    col: Any = None,
    trailing: Optional[ft.Control] = None,
) -> ft.Container:
    cabecera_items: List[ft.Control] = []
    if icono is not None:
        cabecera_items.append(ft.Icon(icono, size=20, color=_color("PRIMARY", ft.Colors.BLUE_700)))
    cabecera_items.append(ft.Text(titulo, size=16, weight=ft.FontWeight.BOLD, expand=True))
    if trailing is not None:
        cabecera_items.append(trailing)

    return ft.Container(
        content=ft.Column(
            [
                ft.Row(cabecera_items, vertical_alignment=ft.CrossAxisAlignment.CENTER),
                ft.Divider(height=1),
                contenido,
            ],
            spacing=10,
        ),
        padding=14,
        border_radius=12,
        bgcolor=_color("SURFACE_CONTAINER_LOW", None),
        border=_borde(1, _color("OUTLINE_VARIANT", ft.Colors.GREY_400)),
        col=col,
    )


def _badge(texto: str, bgcolor: Any, text_color: Any = ft.Colors.WHITE) -> ft.Container:
    return ft.Container(
        content=ft.Text(texto, size=11, weight=ft.FontWeight.BOLD, color=text_color),
        padding=_padding_sim(horizontal=8, vertical=3),
        border_radius=8,
        bgcolor=bgcolor,
    )


def _directorio_salida_default() -> Path:
    """Obtiene una carpeta de salida accesible en PC o Android."""
    base = Path(__file__).resolve().parent
    descargas = Path.home() / "Descargas"
    downloads = Path.home() / "Downloads"
    for carpeta in [
        base / "salidas",
        descargas,
        downloads,
        directorio_datos_app() / "salidas",
        base,
    ]:
        try:
            carpeta.mkdir(parents=True, exist_ok=True)
            if os.access(carpeta, os.W_OK):
                return carpeta
        except Exception:
            continue
    return base


def _abrir_ruta_externa(page: ft.Page, ruta: Path) -> bool:
    """Intenta abrir un archivo o carpeta de forma segura tanto en escritorio como en móvil."""
    try:
        plat = str(getattr(page, "platform", "") or "").upper()
        es_movil = "ANDROID" in plat or "IOS" in plat
        if not es_movil:
            if sys.platform.startswith("linux"):
                subprocess.Popen(["xdg-open", str(ruta)])
                return True
            if sys.platform == "darwin":
                subprocess.Popen(["open", str(ruta)])
                return True
            if os.name == "nt":
                os.startfile(str(ruta))  # type: ignore[attr-defined]
                return True
        if hasattr(page, "launch_url"):
            page.launch_url(ruta.resolve().as_uri())
            return True
    except Exception:
        return False
    return False


# ==============================================================================
# APLICACIÓN PRINCIPAL
# ==============================================================================


def main(page: ft.Page):
    page.title = "AutoSRT 900/15"
    page.theme = ft.Theme(color_scheme_seed=ft.Colors.BLUE_700)
    page.dark_theme = ft.Theme(color_scheme_seed=ft.Colors.BLUE_400)
    page.padding = 0
    page.scroll = ft.ScrollMode.AUTO

    if not _modulos_disponibles:
        page.theme_mode = ft.ThemeMode.LIGHT
        page.add(
            ft.Container(
                content=ft.Column(
                    [
                        ft.Icon(ft.Icons.ERROR, color=ft.Colors.RED, size=48),
                        ft.Text(
                            "Error al iniciar la aplicación",
                            size=20,
                            weight=ft.FontWeight.BOLD,
                            color=ft.Colors.RED,
                        ),
                        ft.Text("No se pudieron cargar los módulos de generación:", size=14),
                        ft.Text(
                            _modulos_error or "Error desconocido",
                            size=11,
                            selectable=True,
                            color=ft.Colors.ORANGE_700,
                        ),
                    ],
                    scroll=ft.ScrollMode.AUTO,
                ),
                padding=20,
                expand=True,
            )
        )
        return

    perfil = cargar_perfil_electricista()
    page.theme_mode = ft.ThemeMode.DARK if perfil.get("tema_oscuro") else ft.ThemeMode.LIGHT

    hoy_str = datetime.date.today().strftime("%d/%m/%Y")
    borrador_inicial = cargar_borrador()

    def _estado_limpio() -> Dict[str, Any]:
        return {
            "datos_cliente": {
                "razon_social": "",
                "cuit": "",
                "direccion": "",
                "localidad": perfil.get("localidad", "Capital") or "Capital",
                "cp": perfil.get("cp", "5400") or "5400",
                "provincia": perfil.get("provincia", "San Juan") or "San Juan",
                "contrato": "",
                "plano": "",
                "condicion_terreno": "ARENOSO SECO",
            },
            "conclusiones": {
                "fecha_medicion": datetime.date.today().strftime("%d/%m/%Y"),
                "hora_inicio": "",
                "hora_fin": "",
                "instrumento_utilizado": perfil.get("instrumento_utilizado", ""),
                "fecha_calibracion": perfil.get("fecha_calibracion", ""),
                "metodologia": perfil.get(
                    "metodologia",
                    "Medición de impedancia por lazo de falla y comprobación de RCD",
                ),
                "observaciones_datos": "",
                "conclusiones_texto": "",
                "recomendaciones_texto": "",
                "adjunta_calibracion": bool(perfil.get("fecha_calibracion", "")),
                "adjunta_plano": False,
            },
            "mediciones_pat": [],
            "mediciones_dif": [],
        }

    state: Dict[str, Any] = _estado_limpio()
    runtime = {
        "textos_editados_manualmente": False,
        "pat_filtro": "todas",
        "dif_fase_activa": "R",
        "dif_edit_idx": None,
        "ultimos_archivos": [],
        "suspend_autosave": False,
    }

    if borrador_inicial:
        state["datos_cliente"].update(borrador_inicial.get("datos_cliente", {}) or {})
        state["conclusiones"].update(borrador_inicial.get("conclusiones", {}) or {})
        state["mediciones_pat"] = list(borrador_inicial.get("mediciones_pat", []) or [])
        state["mediciones_dif"] = list(borrador_inicial.get("mediciones_dif", []) or [])
        runtime["textos_editados_manualmente"] = bool(
            state["conclusiones"].get("textos_editados_manualmente", False)
        )

    def notificar(mensaje: str, color: Any = ft.Colors.GREEN_700):
        snack = ft.SnackBar(
            ft.Text(mensaje, color=ft.Colors.WHITE),
            bgcolor=color,
            duration=2600,
        )
        if hasattr(page, "open"):
            page.open(snack)
        elif hasattr(page, "show_dialog"):
            page.show_dialog(snack)
        else:
            page.overlay.append(snack)
            snack.open = True
            page.update()

    # ================= APP BAR & NAVIGATION BAR =================
    appbar_titulo = ft.Text("Nuevo relevamiento", size=17, weight=ft.FontWeight.BOLD)
    appbar_subtitulo = ft.Text(
        "SRT 900/15",
        size=11,
        color=_color("ON_SURFACE_VARIANT", ft.Colors.GREY_600),
    )

    dest_datos = ft.NavigationBarDestination(icon=ft.Icons.PERSON, label="Datos")
    dest_pat = ft.NavigationBarDestination(icon=ft.Icons.ELECTRIC_BOLT, label="PAT (0)")
    dest_dif = ft.NavigationBarDestination(icon=ft.Icons.SHIELD, label="RCD (0)")
    dest_exportar = ft.NavigationBarDestination(icon=ft.Icons.PICTURE_AS_PDF, label="Exportar")

    def actualizar_chrome():
        razon = (state["datos_cliente"].get("razon_social") or "").strip()
        appbar_titulo.value = razon if razon else "Nuevo relevamiento"
        dest_pat.label = f"PAT ({len(state['mediciones_pat'])})"
        dest_dif.label = f"RCD ({len(state['mediciones_dif'])})"

    def autoguardar():
        if runtime["suspend_autosave"]:
            return
        hora = guardar_borrador(state)
        if hora:
            appbar_subtitulo.value = f"SRT 900/15 · Guardado {hora}"
        actualizar_chrome()

    # ================= VISTA 1: DATOS GENERALES =================
    cal_icono = ft.Icon(ft.Icons.INFO_OUTLINE, size=18)
    cal_texto = ft.Text("Sin fecha de calibración", size=12)
    cal_banner = ft.Container(
        content=ft.Row([cal_icono, cal_texto], spacing=8),
        padding=_padding_sim(horizontal=10, vertical=8),
        border_radius=8,
        border=_borde(1, _color("OUTLINE_VARIANT", ft.Colors.GREY_400)),
        col={"xs": 12},
    )

    def refrescar_banner_calibracion():
        info = verificar_calibracion(input_cal.value)
        est = info["estado"]
        cal_texto.value = info["mensaje"]
        if est == "vigente":
            _set_icon(cal_icono, ft.Icons.CHECK_CIRCLE)
            cal_icono.color = ft.Colors.GREEN_700
            cal_banner.bgcolor = _color("GREEN_50", None) if page.theme_mode == ft.ThemeMode.LIGHT else None
            cal_banner.border = _borde(1, ft.Colors.GREEN_700)
        elif est in ("por_vencer", "invalida"):
            _set_icon(cal_icono, ft.Icons.WARNING_AMBER)
            cal_icono.color = ft.Colors.AMBER_700
            cal_banner.bgcolor = _color("AMBER_50", None) if page.theme_mode == ft.ThemeMode.LIGHT else None
            cal_banner.border = _borde(1, ft.Colors.AMBER_700)
        elif est == "vencida":
            _set_icon(cal_icono, ft.Icons.ERROR)
            cal_icono.color = ft.Colors.RED_700
            cal_banner.bgcolor = _color("RED_50", None) if page.theme_mode == ft.ThemeMode.LIGHT else None
            cal_banner.border = _borde(1, ft.Colors.RED_700)
        else:
            _set_icon(cal_icono, ft.Icons.INFO_OUTLINE)
            cal_icono.color = _color("ON_SURFACE_VARIANT", ft.Colors.GREY_600)
            cal_banner.bgcolor = None
            cal_banner.border = _borde(1, _color("OUTLINE_VARIANT", ft.Colors.GREY_400))

    def update_datos(e=None):
        state["datos_cliente"]["razon_social"] = input_razon.value or ""
        state["datos_cliente"]["cuit"] = input_cuit.value or ""
        state["datos_cliente"]["direccion"] = input_dir.value or ""
        state["datos_cliente"]["localidad"] = input_loc.value or ""
        state["datos_cliente"]["cp"] = input_cp.value or ""
        state["datos_cliente"]["provincia"] = input_prov.value or "San Juan"
        state["datos_cliente"]["contrato"] = input_cont.value or ""
        state["datos_cliente"]["plano"] = input_plano.value or ""
        state["datos_cliente"]["condicion_terreno"] = input_terreno.value or "ARENOSO SECO"
        state["conclusiones"]["fecha_medicion"] = input_fecha.value or ""
        state["conclusiones"]["hora_inicio"] = input_hora_ini.value or ""
        state["conclusiones"]["hora_fin"] = input_hora_fin.value or ""
        state["conclusiones"]["instrumento_utilizado"] = input_inst.value or ""
        state["conclusiones"]["fecha_calibracion"] = input_cal.value or ""
        refrescar_banner_calibracion()
        autoguardar()
        page.update()

    input_razon = _crear_textfield(
        label="Razón Social / Establecimiento *",
        value=state["datos_cliente"]["razon_social"],
        hint_text="Ej. Ortopedia Cuyo S.R.L.",
        on_change=update_datos,
        col={"xs": 12},
    )
    input_cuit = _crear_textfield(
        label="C.U.I.T.",
        value=state["datos_cliente"]["cuit"],
        hint_text="30-12345678-9",
        keyboard_type=ft.KeyboardType.NUMBER,
        on_change=update_datos,
        col={"xs": 12, "sm": 6},
    )
    input_cont = _crear_textfield(
        label="N° de Contrato",
        value=state["datos_cliente"]["contrato"],
        on_change=update_datos,
        col={"xs": 12, "sm": 6},
    )
    input_dir = _crear_textfield(
        label="Dirección",
        value=state["datos_cliente"]["direccion"],
        on_change=update_datos,
        col={"xs": 12},
    )
    input_loc = _crear_textfield(
        label="Localidad",
        value=state["datos_cliente"]["localidad"],
        on_change=update_datos,
        col={"xs": 12, "sm": 6},
    )
    input_cp = _crear_textfield(
        label="CP",
        value=state["datos_cliente"]["cp"],
        keyboard_type=ft.KeyboardType.NUMBER,
        on_change=update_datos,
        col={"xs": 6, "sm": 3},
    )
    input_plano = _crear_textfield(
        label="Plano N°",
        value=state["datos_cliente"]["plano"],
        on_change=update_datos,
        col={"xs": 6, "sm": 3},
    )
    input_prov = _crear_dropdown(
        label="Provincia",
        opciones=PROVINCIAS_ARGENTINA,
        value=state["datos_cliente"]["provincia"],
        on_change=update_datos,
        col={"xs": 12, "sm": 6},
    )
    input_terreno = _crear_dropdown(
        label="Condición del Terreno (Grilla)",
        opciones=CONDICIONES_TERRENO,
        value=state["datos_cliente"]["condicion_terreno"],
        on_change=update_datos,
        col={"xs": 12, "sm": 6},
    )

    input_fecha = _crear_textfield(
        label="Fecha Medición (DD/MM/AAAA)",
        value=state["conclusiones"]["fecha_medicion"] or hoy_str,
        on_change=update_datos,
        col={"xs": 8},
    )

    def poner_fecha_hoy(e):
        input_fecha.value = datetime.date.today().strftime("%d/%m/%Y")
        update_datos()

    def poner_hora_ahora(campo_tf: ft.TextField):
        campo_tf.value = datetime.datetime.now().strftime("%H:%M")
        update_datos()

    btn_fecha_hoy = ft.OutlinedButton(
        "Hoy",
        icon=ft.Icons.CALENDAR_TODAY,
        on_click=poner_fecha_hoy,
        col={"xs": 4},
        height=48,
    )
    input_hora_ini = _crear_textfield(
        label="Hora Inicio",
        value=state["conclusiones"]["hora_inicio"],
        hint_text="09:00",
        on_change=update_datos,
        col={"xs": 6},
    )
    input_hora_fin = _crear_textfield(
        label="Hora Fin",
        value=state["conclusiones"]["hora_fin"],
        hint_text="11:30",
        on_change=update_datos,
        col={"xs": 6},
    )
    btn_ini_ahora = ft.OutlinedButton(
        "Inicio: Ahora",
        icon=ft.Icons.ACCESS_TIME,
        on_click=lambda _: poner_hora_ahora(input_hora_ini),
        col={"xs": 6},
    )
    btn_fin_ahora = ft.OutlinedButton(
        "Fin: Ahora",
        icon=ft.Icons.ACCESS_TIME,
        on_click=lambda _: poner_hora_ahora(input_hora_fin),
        col={"xs": 6},
    )
    input_inst = _crear_textfield(
        label="(7) Marca, modelo y N° serie del instrumento",
        value=state["conclusiones"]["instrumento_utilizado"],
        hint_text="Ej. Metrel MI 3102H BT S/N ...",
        on_change=update_datos,
        col={"xs": 12},
    )
    input_cal = _crear_textfield(
        label="(8) Fecha Calibración (DD/MM/AAAA)",
        value=state["conclusiones"]["fecha_calibracion"],
        hint_text="DD/MM/AAAA",
        on_change=update_datos,
        col={"xs": 12},
    )

    def guardar_mi_instrumental(e):
        update_datos()
        perfil["instrumento_utilizado"] = state["conclusiones"]["instrumento_utilizado"]
        perfil["fecha_calibracion"] = state["conclusiones"]["fecha_calibracion"]
        perfil["provincia"] = state["datos_cliente"]["provincia"]
        perfil["localidad"] = state["datos_cliente"]["localidad"]
        perfil["cp"] = state["datos_cliente"]["cp"]
        if guardar_perfil_electricista(perfil):
            notificar("Instrumental y zona guardados por defecto para futuros relevamientos")
        else:
            notificar("No se pudo guardar la configuración en disco", ft.Colors.RED_700)

    btn_guardar_perfil = ft.OutlinedButton(
        "Guardar como Mi Instrumental",
        icon=ft.Icons.SAVE,
        on_click=guardar_mi_instrumental,
        col={"xs": 12},
        height=46,
    )

    card_establecimiento = _card(
        "Establecimiento",
        ft.ResponsiveRow(
            [
                input_razon,
                input_cuit,
                input_cont,
                input_dir,
                input_loc,
                input_cp,
                input_plano,
                input_prov,
                input_terreno,
            ],
            spacing=10,
            run_spacing=10,
        ),
        icono=ft.Icons.PERSON,
        col={"xs": 12, "md": 6},
    )

    card_instrumental = _card(
        "Instrumental y Horarios (Ítems 7 a 11)",
        ft.ResponsiveRow(
            [
                input_fecha,
                btn_fecha_hoy,
                input_hora_ini,
                input_hora_fin,
                btn_ini_ahora,
                btn_fin_ahora,
                input_inst,
                input_cal,
                cal_banner,
                btn_guardar_perfil,
            ],
            spacing=10,
            run_spacing=10,
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
        ),
        icono=ft.Icons.BUILD,
        col={"xs": 12, "md": 6},
    )

    view_datos = ft.Column(
        [
            ft.ResponsiveRow([card_establecimiento, card_instrumental], spacing=12, run_spacing=12),
            ft.Row(
                [
                    _Btn(
                        "Continuar a Puesta a Tierra",
                        icon=ft.Icons.ARROW_FORWARD,
                        on_click=lambda _: ir_a_vista(1),
                        bgcolor=ft.Colors.BLUE_700,
                        color=ft.Colors.WHITE,
                        height=48,
                        expand=True,
                    )
                ]
            ),
        ],
        visible=True,
        spacing=12,
    )

    # ================= VISTA 2: PUESTA A TIERRA (PAT) =================
    pat_sector = _crear_textfield(
        label="Sector / Ubicación *",
        hint_text="Ej. Oficina, Cocina, Tablero...",
        col={"xs": 12},
    )
    sw_mantener_sector = ft.Switch(label="Mantener sector al agregar", value=True)
    chips_sectores_row = ft.Row(wrap=True, spacing=6, run_spacing=6)

    def seleccionar_sector_chip(nombre_sector: str):
        pat_sector.value = nombre_sector
        _set_tf_error(pat_sector, None)
        _enfocar(page, pat_val)
        page.update()

    def refrescar_chips_sectores():
        chips_sectores_row.controls.clear()
        combinados: List[str] = []
        for s in list(perfil.get("sectores_recientes", [])) + SECTORES_SUGERIDOS:
            s_clean = str(s or "").strip()
            if s_clean and s_clean.lower() not in [x.lower() for x in combinados]:
                combinados.append(s_clean)
        for sec in combinados[:10]:
            chips_sectores_row.controls.append(
                ft.Chip(
                    label=ft.Text(sec, size=12),
                    on_click=lambda e, s_val=sec: seleccionar_sector_chip(s_val),
                )
            )

    def registrar_sector_reciente(sec: str):
        s_clean = sec.strip()
        if not s_clean:
            return
        rec = [x for x in perfil.get("sectores_recientes", []) if x.lower() != s_clean.lower()]
        rec.insert(0, s_clean)
        perfil["sectores_recientes"] = rec[:6]
        guardar_perfil_electricista(perfil)
        refrescar_chips_sectores()

    pat_eval_txt = ft.Text(
        "Ingresá un valor en Ohms (≤ 40 Ω) o usá los botones rápidos",
        size=12,
        color=_color("ON_SURFACE_VARIANT", ft.Colors.GREY_600),
    )

    def on_pat_val_change(e=None):
        ev = evaluar_valor_pat(pat_val.value)
        est = ev["estado"]
        if est == "vacio":
            pat_eval_txt.value = "Ingresá un valor en Ohms (≤ 40 Ω) o usá los botones rápidos"
            pat_eval_txt.color = _color("ON_SURFACE_VARIANT", ft.Colors.GREY_600)
            _set_tf_error(pat_val, None)
        elif est == "apto":
            pat_eval_txt.value = ev["mensaje"]
            pat_eval_txt.color = ft.Colors.GREEN_700
            _set_tf_error(pat_val, None)
        elif est in ("no_cumple", "invalido"):
            pat_eval_txt.value = ev["mensaje"]
            pat_eval_txt.color = ft.Colors.RED_700
            _set_tf_error(pat_val, None)
        else:
            pat_eval_txt.value = ev["mensaje"]
            pat_eval_txt.color = ft.Colors.ORANGE_800
            _set_tf_error(pat_val, None)
        page.update()

    pat_val = _crear_textfield(
        label="Valor medido (Ω)",
        hint_text="0.88",
        suffix_text="Ω",
        keyboard_type=ft.KeyboardType.NUMBER,
        on_change=on_pat_val_change,
        on_submit=lambda e: add_pat(e),
        expand=True,
        text_size=18,
    )

    pat_esq = _crear_dropdown(
        label="Esquema", opciones=ESQUEMAS_CONEXION, value="TT", col={"xs": 6}
    )
    pat_prot = _crear_dropdown(
        label="Protección", opciones=PROTECCIONES_PAT, value="DD", col={"xs": 6}
    )
    pat_uso = _crear_dropdown(
        label="Uso de PAT (Opcional)", opciones=USOS_PAT, value="", col={"xs": 12}
    )
    pat_cont = _crear_dropdown(
        label="Circuito continuo", opciones=["SI", "NO"], value="SI", col={"xs": 6}
    )
    pat_cap = _crear_dropdown(
        label="Capacidad carga", opciones=["SI", "NO"], value="SI", col={"xs": 6}
    )

    exp_pat_avanzado = _crear_expansion_tile(
        title=ft.Text("Opciones técnicas (Esquema / Protección / Uso)", size=13, weight=ft.FontWeight.W_600),
        subtitle=ft.Text("Por defecto: TT · DD · Continuidad SI", size=11),
        expanded=False,
        controls=[
            ft.Container(
                padding=8,
                content=ft.ResponsiveRow(
                    [pat_esq, pat_prot, pat_uso, pat_cont, pat_cap],
                    spacing=8,
                    run_spacing=8,
                ),
            )
        ],
    )

    pat_resumen_row = ft.Row(wrap=True, spacing=8, run_spacing=6)
    pat_list_col = ft.Column(spacing=6)

    def reenumerar_pat():
        for idx, item in enumerate(state["mediciones_pat"], start=1):
            item["nro_toma"] = idx

    def _clasificar_pat_item(m: Dict[str, Any]) -> str:
        es_num, val_f = es_medicion_pat_valida(m.get("valor_ohms"))
        if not es_num or val_f is None:
            return "anomalia"
        cont = str(m.get("continuidad", "SI") or "SI").upper()
        cap = str(m.get("capacidad_carga", "SI") or "SI").upper()
        cumple = str(m.get("cumple", "SI") or "SI").upper()
        if val_f <= 0 or val_f > 40.0 or cumple == "NO" or cont == "NO" or cap == "NO":
            return "no_cumple"
        return "apto"

    def abrir_editor_pat(indice: int):
        if not (0 <= indice < len(state["mediciones_pat"])):
            return
        item = state["mediciones_pat"][indice]
        ed_sector = _crear_textfield(label="Sector / Ubicación", value=str(item.get("sector", "")))
        ed_valor = _crear_textfield(
            label="Valor (Ω) o Anomalía",
            value=str(item.get("valor_ohms", "")),
            helper_text="Número en Ω o texto ej. INVERTIDO / SIN PE / NO TIENE",
        )
        ed_cond = _crear_dropdown(
            label="Condición del Terreno",
            opciones=CONDICIONES_TERRENO,
            value=str(item.get("condicion_terreno", "") or state["datos_cliente"]["condicion_terreno"]),
        )
        ed_esq = _crear_dropdown(
            label="Esquema",
            opciones=ESQUEMAS_CONEXION,
            value=str(item.get("esquema", "TT") or "TT"),
        )
        ed_prot = _crear_dropdown(
            label="Protección",
            opciones=PROTECCIONES_PAT,
            value=str(item.get("proteccion", "DD") or "DD"),
        )
        ed_uso = _crear_dropdown(
            label="Uso de PAT",
            opciones=USOS_PAT,
            value=str(item.get("uso_pat", "") or ""),
        )
        ed_cont = _crear_dropdown(
            label="Continuidad",
            opciones=["SI", "NO", ""],
            value=str(item.get("continuidad", "SI") or ""),
        )
        ed_cap = _crear_dropdown(
            label="Capacidad de carga",
            opciones=["SI", "NO", ""],
            value=str(item.get("capacidad_carga", "SI") or ""),
        )

        dlg_ref: Dict[str, Any] = {}

        def guardar_edicion(ev):
            sec_n = (ed_sector.value or "").strip()
            val_n = (ed_valor.value or "").strip()
            if not sec_n or not val_n:
                _set_tf_error(ed_sector, "Obligatorio" if not sec_n else None)
                _set_tf_error(ed_valor, "Obligatorio" if not val_n else None)
                page.update()
                return
            es_num, val_f = es_medicion_pat_valida(val_n)
            item["sector"] = sec_n
            item["condicion_terreno"] = ed_cond.value or "ARENOSO SECO"
            item["esquema"] = ed_esq.value or "TT"
            item["uso_pat"] = ed_uso.value or ""
            if es_num and val_f is not None:
                cont_v = ed_cont.value or "SI"
                cap_v = ed_cap.value or "SI"
                item["valor_ohms"] = val_f
                item["continuidad"] = cont_v
                item["capacidad_carga"] = cap_v
                item["proteccion"] = ed_prot.value or "DD"
                item["cumple"] = "SI" if (0 < val_f <= 40.0 and cont_v != "NO" and cap_v != "NO") else "NO"
            else:
                item["valor_ohms"] = val_n
                item["continuidad"] = ""
                item["capacidad_carga"] = ""
                item["proteccion"] = ""
                item["cumple"] = "NO"
            _cerrar_dialogo(page, dlg_ref["dlg"])
            refrescar_lista_pat()
            autoguardar()
            page.update()
            notificar(f"Toma #{item['nro_toma']} actualizada")

        dlg = ft.AlertDialog(
            modal=True,
            title=ft.Text(f"Editar Toma #{item.get('nro_toma', indice + 1)}"),
            content=ft.Container(
                width=420,
                content=ft.Column(
                    [
                        ed_sector,
                        ed_valor,
                        ft.ResponsiveRow([ed_esq, ed_prot], spacing=8, run_spacing=8),
                        ed_cond,
                        ed_uso,
                        ft.ResponsiveRow([ed_cont, ed_cap], spacing=8, run_spacing=8),
                    ],
                    tight=True,
                    spacing=8,
                    scroll=ft.ScrollMode.AUTO,
                ),
            ),
            actions=[
                ft.TextButton("Cancelar", on_click=lambda _: _cerrar_dialogo(page, dlg)),
                _Btn("Guardar", on_click=guardar_edicion, bgcolor=ft.Colors.BLUE_700, color=ft.Colors.WHITE),
            ],
        )
        dlg_ref["dlg"] = dlg
        _abrir_dialogo(page, dlg)

    def refrescar_lista_pat():
        reenumerar_pat()
        pat_list_col.controls.clear()
        total = len(state["mediciones_pat"])
        aptas = 0
        obs = 0

        for m in state["mediciones_pat"]:
            if _clasificar_pat_item(m) == "apto":
                aptas += 1
            else:
                obs += 1

        hojas_nec = max(1, (total + 15) // 16)
        pat_resumen_row.controls = [
            _badge(f"Total: {total}", ft.Colors.BLUE_700),
            _badge(f"Aptas: {aptas}", ft.Colors.GREEN_700),
            _badge(f"Obs./Fallas: {obs}", ft.Colors.ORANGE_800 if obs else ft.Colors.GREY_600),
            _badge(f"{hojas_nec} hoja/s GRILLA", _color("SECONDARY", ft.Colors.BLUE_GREY_700)),
        ]

        filtro = runtime["pat_filtro"]
        indices_visibles = []
        for idx in range(total - 1, -1, -1):
            cat = _clasificar_pat_item(state["mediciones_pat"][idx])
            if filtro == "aptas" and cat != "apto":
                continue
            if filtro == "obs" and cat == "apto":
                continue
            indices_visibles.append(idx)

        if not indices_visibles:
            pat_list_col.controls.append(
                ft.Container(
                    padding=20,
                    alignment=_align_center(),
                    content=ft.Text(
                        "No hay mediciones en esta vista todavía."
                        if total == 0
                        else "No hay mediciones para el filtro seleccionado.",
                        color=_color("ON_SURFACE_VARIANT", ft.Colors.GREY_600),
                    ),
                )
            )
        else:
            for idx in indices_visibles:
                m = state["mediciones_pat"][idx]
                cat = _clasificar_pat_item(m)
                if cat == "apto":
                    icono = ft.Icon(ft.Icons.CHECK_CIRCLE, color=ft.Colors.GREEN_700)
                    val_txt = f"{m['valor_ohms']} Ω · CUMPLE"
                    borde_c = _color("OUTLINE_VARIANT", ft.Colors.GREY_400)
                elif cat == "no_cumple":
                    icono = ft.Icon(ft.Icons.ERROR, color=ft.Colors.RED_700)
                    val_txt = f"{m['valor_ohms']} Ω · NO CUMPLE"
                    borde_c = ft.Colors.RED_400
                else:
                    icono = ft.Icon(ft.Icons.WARNING_AMBER, color=ft.Colors.ORANGE_800)
                    val_txt = f"{m['valor_ohms']} · OBSERVADO"
                    borde_c = ft.Colors.ORANGE_400

                def _crear_borrador(i_del):
                    def _borrar(e):
                        if 0 <= i_del < len(state["mediciones_pat"]):
                            state["mediciones_pat"].pop(i_del)
                            refrescar_lista_pat()
                            autoguardar()
                            page.update()

                    return _borrar

                sub_partes = [
                    f"Esq: {m.get('esquema', 'TT')}",
                    f"Prot: {m.get('proteccion') or '-'}",
                    f"Terreno: {m.get('condicion_terreno', '')}",
                ]
                if m.get("continuidad") == "NO":
                    sub_partes.append("Cont: NO")
                if m.get("capacidad_carga") == "NO":
                    sub_partes.append("Cap: NO")
                if m.get("uso_pat"):
                    sub_partes.append(f"Uso: {m['uso_pat']}")

                pat_list_col.controls.append(
                    ft.Container(
                        border=_borde(1, borde_c),
                        border_radius=10,
                        padding=_padding_sim(horizontal=10, vertical=6),
                        content=ft.Row(
                            [
                                icono,
                                ft.Column(
                                    [
                                        ft.Text(
                                            f"#{m['nro_toma']} · {m['sector']}  —  {val_txt}",
                                            weight=ft.FontWeight.W_600,
                                            size=13.5,
                                        ),
                                        ft.Text(
                                            " | ".join(sub_partes),
                                            size=11.5,
                                            color=_color("ON_SURFACE_VARIANT", ft.Colors.GREY_600),
                                        ),
                                    ],
                                    spacing=2,
                                    expand=True,
                                ),
                                ft.IconButton(
                                    ft.Icons.EDIT,
                                    tooltip="Editar toma",
                                    on_click=lambda e, i_ed=idx: abrir_editor_pat(i_ed),
                                ),
                                ft.IconButton(
                                    ft.Icons.DELETE_OUTLINE,
                                    icon_color=ft.Colors.RED_600,
                                    tooltip="Eliminar toma",
                                    on_click=_crear_borrador(idx),
                                ),
                            ],
                            vertical_alignment=ft.CrossAxisAlignment.CENTER,
                            spacing=8,
                        ),
                    )
                )
        actualizar_chrome()

    def set_filtro_pat(nuevo: str):
        runtime["pat_filtro"] = nuevo
        refrescar_chips_filtro_pat()
        refrescar_lista_pat()
        page.update()

    chip_f_todas = ft.Chip(label=ft.Text("Todas"), selected=True, on_select=lambda _: set_filtro_pat("todas"))
    chip_f_aptas = ft.Chip(label=ft.Text("Aptas"), selected=False, on_select=lambda _: set_filtro_pat("aptas"))
    chip_f_obs = ft.Chip(label=ft.Text("Observadas"), selected=False, on_select=lambda _: set_filtro_pat("obs"))

    def refrescar_chips_filtro_pat():
        f = runtime["pat_filtro"]
        chip_f_todas.selected = f == "todas"
        chip_f_aptas.selected = f == "aptas"
        chip_f_obs.selected = f == "obs"

    def registrar_pat(valor_ingresado: str, es_anomalia_rapida: bool = False):
        sector_txt = (pat_sector.value or "").strip()
        val_txt = (valor_ingresado or "").strip()
        if not sector_txt:
            _set_tf_error(pat_sector, "Ingresá el sector")
            _enfocar(page, pat_sector)
            page.update()
            notificar("Ingresá el Sector antes de registrar la toma", ft.Colors.RED_700)
            return
        _set_tf_error(pat_sector, None)

        if not val_txt:
            _set_tf_error(pat_val, "Ingresá Ω o tocá una anomalía")
            _enfocar(page, pat_val)
            page.update()
            notificar("Ingresá el Valor (Ω) o tocá un botón rápido", ft.Colors.RED_700)
            return
        _set_tf_error(pat_val, None)

        es_num, val_f = es_medicion_pat_valida(val_txt)
        if es_num and val_f is not None and val_f <= 0:
            _set_tf_error(pat_val, "El valor debe ser mayor a 0 Ω")
            _enfocar(page, pat_val)
            page.update()
            return

        if es_num and val_f is not None:
            cont_v = pat_cont.value or "SI"
            cap_v = pat_cap.value or "SI"
            cumple_auto = "SI" if (val_f <= 40.0 and cont_v != "NO" and cap_v != "NO") else "NO"
            prot_auto = pat_prot.value or "DD"
        else:
            cumple_auto = "NO"
            prot_auto = ""
            cont_v = ""
            cap_v = ""

        state["mediciones_pat"].append(
            {
                "nro_toma": len(state["mediciones_pat"]) + 1,
                "sector": sector_txt,
                "condicion_terreno": state["datos_cliente"]["condicion_terreno"],
                "uso_pat": pat_uso.value or "",
                "esquema": pat_esq.value or "TT",
                "valor_ohms": val_f if es_num else val_txt,
                "cumple": cumple_auto,
                "continuidad": cont_v,
                "capacidad_carga": cap_v,
                "proteccion": prot_auto,
            }
        )
        registrar_sector_reciente(sector_txt)

        if not sw_mantener_sector.value:
            pat_sector.value = ""
        pat_val.value = ""
        on_pat_val_change()
        refrescar_lista_pat()
        autoguardar()
        if not es_anomalia_rapida:
            _enfocar(page, pat_val if sw_mantener_sector.value else pat_sector)
        page.update()

    def add_pat(e):
        registrar_pat(pat_val.value or "", es_anomalia_rapida=False)

    def remove_pat(e):
        if state["mediciones_pat"]:
            state["mediciones_pat"].pop()
            refrescar_lista_pat()
            autoguardar()
            page.update()
            notificar("Última toma eliminada", ft.Colors.ORANGE_800)

    btn_add_pat = _Btn(
        "Agregar",
        on_click=add_pat,
        icon=ft.Icons.ADD_CIRCLE,
        bgcolor=ft.Colors.BLUE_700,
        color=ft.Colors.WHITE,
        height=52,
    )

    botones_anomalias = ft.ResponsiveRow(
        [
            _Btn(
                "INVERTIDO",
                icon=ft.Icons.SWAP_HORIZ,
                on_click=lambda _: registrar_pat("INVERTIDO", True),
                bgcolor=ft.Colors.AMBER_700,
                color=ft.Colors.WHITE,
                height=48,
                col={"xs": 6},
            ),
            _Btn(
                "SIN PE",
                icon=ft.Icons.LINK_OFF,
                on_click=lambda _: registrar_pat("SIN PE", True),
                bgcolor=ft.Colors.ORANGE_800,
                color=ft.Colors.WHITE,
                height=48,
                col={"xs": 6},
            ),
            _Btn(
                "NO TIENE",
                icon=ft.Icons.BLOCK,
                on_click=lambda _: registrar_pat("NO TIENE", True),
                bgcolor=ft.Colors.RED_700,
                color=ft.Colors.WHITE,
                height=48,
                col={"xs": 6},
            ),
            _Btn(
                "INV / SIN PE",
                icon=ft.Icons.REPORT_PROBLEM,
                on_click=lambda _: registrar_pat("INVERTIDO/SIN PE", True),
                bgcolor=ft.Colors.DEEP_ORANGE_900,
                color=ft.Colors.WHITE,
                height=48,
                col={"xs": 6},
            ),
        ],
        spacing=8,
        run_spacing=8,
    )

    card_pat_carga = _card(
        "Carga Rápida PAT",
        ft.Column(
            [
                pat_sector,
                sw_mantener_sector,
                ft.Text("Sectores rápidos:", size=12, color=_color("ON_SURFACE_VARIANT", ft.Colors.GREY_600)),
                chips_sectores_row,
                ft.Row([pat_val, btn_add_pat], spacing=8, vertical_alignment=ft.CrossAxisAlignment.START),
                pat_eval_txt,
                ft.Divider(height=1),
                ft.Text("Anomalías en 1 toque:", size=12, weight=ft.FontWeight.W_600),
                botones_anomalias,
                exp_pat_avanzado,
            ],
            spacing=8,
        ),
        icono=ft.Icons.ELECTRIC_BOLT,
        col={"xs": 12, "md": 5},
    )

    card_pat_lista = _card(
        "Tomas Registradas",
        ft.Column(
            [
                pat_resumen_row,
                ft.Row(
                    [
                        chip_f_todas,
                        chip_f_aptas,
                        chip_f_obs,
                        ft.OutlinedButton("Deshacer última", icon=ft.Icons.UNDO, on_click=remove_pat),
                    ],
                    spacing=6,
                    run_spacing=6,
                    wrap=True,
                    vertical_alignment=ft.CrossAxisAlignment.CENTER,
                ),
                pat_list_col,
            ],
            spacing=10,
        ),
        icono=ft.Icons.LIST_ALT if hasattr(ft.Icons, "LIST_ALT") else ft.Icons.TABLE_CHART,
        col={"xs": 12, "md": 7},
    )

    view_pat = ft.Column(
        [
            ft.ResponsiveRow([card_pat_carga, card_pat_lista], spacing=12, run_spacing=12),
        ],
        visible=False,
        spacing=12,
    )

    # ================= VISTA 3: DIFERENCIALES (RCD) =================
    def _fases_vacias() -> Dict[str, Dict[str, Dict[str, str]]]:
        return {
            f: {
                "0": {"dt_05": "ND", "dt_1": "", "dt_5": "", "id_ma": ""},
                "180": {"dt_05": "ND", "dt_1": "", "dt_5": "", "id_ma": ""},
            }
            for f in ("R", "S", "T")
        }

    dif_draft_fases: Dict[str, Dict[str, Dict[str, str]]] = _fases_vacias()

    def _construir_med_borrador() -> Dict[str, Any]:
        tipo_sel = dif_tipo.value or "TRIFASICO"
        fases_out: Dict[str, Any] = {}
        if tipo_sel == "TRIFASICO":
            for f_l in ("R", "S", "T"):
                fases_out[f_l] = {
                    "0": {**dif_draft_fases[f_l]["0"], "t_ms": None},
                    "180": {**dif_draft_fases[f_l]["180"], "t_ms": None},
                }
        else:
            bip = {
                "0": {**dif_draft_fases["R"]["0"], "t_ms": None},
                "180": {**dif_draft_fases["R"]["180"], "t_ms": None},
            }
            fases_out["R"] = bip
            fases_out["BIPOLAR"] = bip
        return {
            "tipo": tipo_sel,
            "sector": (dif_sec.value or "").strip() or "TP",
            "marca": (dif_marca.value or "").strip(),
            "corriente_nominal_a": (dif_in.value or "").strip(),
            "corriente_fuga_ma": dif_sens.value or "30",
            "tiempo_disparo_ms": dif_draft_fases["R"]["0"]["dt_1"],
            "fases": fases_out,
        }

    dif_eval_icono = ft.Icon(ft.Icons.INFO_OUTLINE, size=18)
    dif_eval_txt = ft.Text("Cargá los tiempos de disparo para evaluar el diferencial", size=12, expand=True)
    dif_eval_banner = ft.Container(
        content=ft.Row([dif_eval_icono, dif_eval_txt], spacing=8),
        padding=_padding_sim(horizontal=10, vertical=8),
        border_radius=8,
        border=_borde(1, _color("OUTLINE_VARIANT", ft.Colors.GREY_400)),
    )

    def _sync_campos_a_borrador(e=None):
        f_act = runtime["dif_fase_activa"]
        d_f = dif_draft_fases[f_act]
        if not btn_0_dt05_estado_disparo["0"]:
            d_f["0"]["dt_05"] = "ND"
        else:
            d_f["0"]["dt_05"] = (fld_0_dt05_ms.value or "").strip() or "DISPARA"

        if not btn_0_dt05_estado_disparo["180"]:
            d_f["180"]["dt_05"] = "ND"
        else:
            d_f["180"]["dt_05"] = (fld_180_dt05_ms.value or "").strip() or "DISPARA"

        d_f["0"]["dt_1"] = (fld_0_dt1.value or "").strip()
        d_f["180"]["dt_1"] = (fld_180_dt1.value or "").strip()
        d_f["0"]["dt_5"] = (fld_0_dt5.value or "").strip()
        d_f["180"]["dt_5"] = (fld_180_dt5.value or "").strip()
        d_f["0"]["id_ma"] = (fld_0_id.value or "").strip()
        d_f["180"]["id_ma"] = (fld_180_id.value or "").strip()
        refrescar_validacion_rcd()
        page.update()

    def _cargar_borrador_en_campos():
        f_act = runtime["dif_fase_activa"]
        d_f = dif_draft_fases[f_act]
        for pol_k, fld_ms in (("0", fld_0_dt05_ms), ("180", fld_180_dt05_ms)):
            val05 = str(d_f[pol_k].get("dt_05", "ND") or "ND").strip()
            es_nd = val05.upper() in ("", "ND", "NO DISPARA")
            btn_0_dt05_estado_disparo[pol_k] = not es_nd
            fld_ms.value = "" if (es_nd or val05.upper() == "DISPARA") else val05
            fld_ms.visible = not es_nd

        fld_0_dt1.value = str(d_f["0"].get("dt_1", "") or "")
        fld_180_dt1.value = str(d_f["180"].get("dt_1", "") or "")
        fld_0_dt5.value = str(d_f["0"].get("dt_5", "") or "")
        fld_180_dt5.value = str(d_f["180"].get("dt_5", "") or "")
        fld_0_id.value = str(d_f["0"].get("id_ma", "") or "")
        fld_180_id.value = str(d_f["180"].get("id_ma", "") or "")
        _actualizar_botones_dt05()
        refrescar_validacion_rcd()

    def _actualizar_botones_dt05():
        for pol_k, btn in (("0", btn_0_dt05), ("180", btn_180_dt05)):
            disparo = btn_0_dt05_estado_disparo[pol_k]
            if disparo:
                _set_btn_text(btn, f"{pol_k}°: Disparó (NO APTO)")
                btn.bgcolor = ft.Colors.RED_700
                btn.color = ft.Colors.WHITE
            else:
                _set_btn_text(btn, f"{pol_k}°: ND (Apto)")
                btn.bgcolor = ft.Colors.GREEN_700
                btn.color = ft.Colors.WHITE

    def toggle_dt05(pol_k: str):
        btn_0_dt05_estado_disparo[pol_k] = not btn_0_dt05_estado_disparo[pol_k]
        fld_ms = fld_0_dt05_ms if pol_k == "0" else fld_180_dt05_ms
        fld_ms.visible = btn_0_dt05_estado_disparo[pol_k]
        if not fld_ms.visible:
            fld_ms.value = ""
        _actualizar_botones_dt05()
        _sync_campos_a_borrador()

    btn_0_dt05_estado_disparo = {"0": False, "180": False}
    btn_0_dt05 = _Btn(
        "0°: ND (Apto)",
        on_click=lambda _: toggle_dt05("0"),
        bgcolor=ft.Colors.GREEN_700,
        color=ft.Colors.WHITE,
        height=44,
        col={"xs": 6},
    )
    btn_180_dt05 = _Btn(
        "180°: ND (Apto)",
        on_click=lambda _: toggle_dt05("180"),
        bgcolor=ft.Colors.GREEN_700,
        color=ft.Colors.WHITE,
        height=44,
        col={"xs": 6},
    )
    fld_0_dt05_ms = _crear_textfield(
        label="0° ΔTx0.5 disparó (ms)",
        keyboard_type=ft.KeyboardType.NUMBER,
        on_change=_sync_campos_a_borrador,
        col={"xs": 6},
    )
    fld_0_dt05_ms.visible = False
    fld_180_dt05_ms = _crear_textfield(
        label="180° ΔTx0.5 disparó (ms)",
        keyboard_type=ft.KeyboardType.NUMBER,
        on_change=_sync_campos_a_borrador,
        col={"xs": 6},
    )
    fld_180_dt05_ms.visible = False

    fld_0_dt1 = _crear_textfield(
        label="0° (+) ΔTx1 (ms)",
        hint_text="< 300",
        keyboard_type=ft.KeyboardType.NUMBER,
        on_change=_sync_campos_a_borrador,
        col={"xs": 6},
    )
    fld_180_dt1 = _crear_textfield(
        label="180° (-) ΔTx1 (ms)",
        hint_text="< 300",
        keyboard_type=ft.KeyboardType.NUMBER,
        on_change=_sync_campos_a_borrador,
        col={"xs": 6},
    )
    fld_0_dt5 = _crear_textfield(
        label="0° (+) ΔTx5 (ms)",
        hint_text="< 40",
        keyboard_type=ft.KeyboardType.NUMBER,
        on_change=_sync_campos_a_borrador,
        col={"xs": 6},
    )
    fld_180_dt5 = _crear_textfield(
        label="180° (-) ΔTx5 (ms)",
        hint_text="< 40",
        keyboard_type=ft.KeyboardType.NUMBER,
        on_change=_sync_campos_a_borrador,
        col={"xs": 6},
    )
    fld_0_id = _crear_textfield(
        label="0° (+) Rampa ID (mA)",
        hint_text="15 - 30",
        keyboard_type=ft.KeyboardType.NUMBER,
        on_change=_sync_campos_a_borrador,
        col={"xs": 6},
    )
    fld_180_id = _crear_textfield(
        label="180° (-) Rampa ID (opc.)",
        hint_text="opc.",
        keyboard_type=ft.KeyboardType.NUMBER,
        on_change=_sync_campos_a_borrador,
        col={"xs": 6},
    )

    def cambiar_tipo_o_sens(e=None):
        es_tri = (dif_tipo.value or "TRIFASICO") == "TRIFASICO"
        fila_tabs_fases.visible = es_tri
        if not es_tri:
            runtime["dif_fase_activa"] = "R"
            _cargar_borrador_en_campos()
        else:
            refrescar_validacion_rcd()
        lbl_fase_titulo.value = (
            f"Ensayo Fase {runtime['dif_fase_activa']} (0° y 180°)"
            if es_tri
            else "Ensayo Bipolar Monofásico (0° y 180°)"
        )
        page.update()

    dif_tipo = _crear_dropdown(
        label="Tipo de Disyuntor",
        opciones=["TRIFASICO", "BIPOLAR"],
        value="TRIFASICO",
        on_change=cambiar_tipo_o_sens,
        col={"xs": 12, "sm": 4},
    )
    def _on_dif_sec_change(e=None):
        refrescar_validacion_rcd()
        page.update()

    dif_sec = _crear_textfield(
        label="Ubicación *",
        value="TP",
        hint_text="TP, TG, TS1",
        on_change=_on_dif_sec_change,
        col={"xs": 6, "sm": 4},
    )
    dif_sens = _crear_dropdown(
        label="IΔn (mA)",
        opciones=["30", "10", "100", "300"],
        value="30",
        on_change=cambiar_tipo_o_sens,
        col={"xs": 6, "sm": 4},
    )
    dif_marca = _crear_textfield(label="Marca", hint_text="Ej. Schneider / ABB", col={"xs": 7, "sm": 6})
    dif_in = _crear_textfield(
        label="In (A)",
        hint_text="40",
        keyboard_type=ft.KeyboardType.NUMBER,
        col={"xs": 5, "sm": 6},
    )

    def seleccionar_fase(fase_l: str):
        runtime["dif_fase_activa"] = fase_l
        lbl_fase_titulo.value = f"Ensayo Fase {fase_l} (0° y 180°)"
        _cargar_borrador_en_campos()
        page.update()

    btn_fase_r = _Btn("Fase R ○", on_click=lambda _: seleccionar_fase("R"), expand=True, height=46)
    btn_fase_s = _Btn("Fase S ○", on_click=lambda _: seleccionar_fase("S"), expand=True, height=46)
    btn_fase_t = _Btn("Fase T ○", on_click=lambda _: seleccionar_fase("T"), expand=True, height=46)
    fila_tabs_fases = ft.Row([btn_fase_r, btn_fase_s, btn_fase_t], spacing=8, visible=True)

    lbl_fase_titulo = ft.Text("Ensayo Fase R (0° y 180°)", size=14, weight=ft.FontWeight.BOLD)
    lbl_rango_id = ft.Text("Rampa ID (mA) — rango válido: 15 a 30 mA", size=12, weight=ft.FontWeight.W_600)

    def refrescar_validacion_rcd():
        ian = parsear_float(dif_sens.value or "30") or 30.0
        lbl_rango_id.value = f"Rampa ID (mA) — rango válido: {0.5 * ian:g} < ID ≤ {ian:g} mA"

        # Validar campos visibles de la fase activa
        f_act = runtime["dif_fase_activa"]
        d_f = dif_draft_fases[f_act]
        for fld, param_k, pol_k in (
            (fld_0_dt1, "dt_1", "0"),
            (fld_180_dt1, "dt_1", "180"),
            (fld_0_dt5, "dt_5", "0"),
            (fld_180_dt5, "dt_5", "180"),
            (fld_0_id, "id_ma", "0"),
            (fld_180_id, "id_ma", "180"),
        ):
            ok, msg = evaluar_parametro_rcd(param_k, d_f[pol_k].get(param_k, ""), ian)
            _set_tf_error(fld, msg if ok is False else None)

        # Estado de cada pestaña de fase
        for f_l, btn in (("R", btn_fase_r), ("S", btn_fase_s), ("T", btn_fase_t)):
            sub_med = {
                "tipo": "BIPOLAR",
                "corriente_fuga_ma": str(ian),
                "fases": {"BIPOLAR": dif_draft_fases[f_l]},
            }
            ev_f = evaluar_diferencial(sub_med)
            if not ev_f["tiene_mediciones"]:
                simbolo = "○"
            elif ev_f["cumple"] == "SI":
                simbolo = "✓"
            else:
                simbolo = "⚠"
            _set_btn_text(btn, f"Fase {f_l} {simbolo}")
            if f_l == f_act:
                btn.bgcolor = ft.Colors.BLUE_700
                btn.color = ft.Colors.WHITE
            else:
                btn.bgcolor = None
                btn.color = None

        # Evaluación global del borrador
        med_borrador = _construir_med_borrador()
        ev_g = evaluar_diferencial(med_borrador)
        if not ev_g["tiene_mediciones"]:
            _set_icon(dif_eval_icono, ft.Icons.INFO_OUTLINE)
            dif_eval_icono.color = _color("ON_SURFACE_VARIANT", ft.Colors.GREY_600)
            dif_eval_txt.value = "Cargá los tiempos de disparo para validar según SRT 900/15"
            dif_eval_banner.bgcolor = None
            dif_eval_banner.border = _borde(1, _color("OUTLINE_VARIANT", ft.Colors.GREY_400))
        elif ev_g["cumple"] == "SI":
            _set_icon(dif_eval_icono, ft.Icons.CHECK_CIRCLE)
            dif_eval_icono.color = ft.Colors.GREEN_700
            dif_eval_txt.value = "APTO según SRT 900/15 / RAEA"
            dif_eval_banner.bgcolor = _color("GREEN_50", None) if page.theme_mode == ft.ThemeMode.LIGHT else None
            dif_eval_banner.border = _borde(1, ft.Colors.GREEN_700)
        else:
            _set_icon(dif_eval_icono, ft.Icons.ERROR)
            dif_eval_icono.color = ft.Colors.RED_700
            dif_eval_txt.value = "NO CUMPLE: " + " | ".join(ev_g["motivos"][:2])
            dif_eval_banner.bgcolor = _color("RED_50", None) if page.theme_mode == ft.ThemeMode.LIGHT else None
            dif_eval_banner.border = _borde(1, ft.Colors.RED_700)

    dif_list_col = ft.Column(spacing=8)

    def limpiar_formulario_dif(e=None):
        runtime["dif_edit_idx"] = None
        runtime["dif_fase_activa"] = "R"
        for f in ("R", "S", "T"):
            dif_draft_fases[f] = _fases_vacias()[f]
        _set_btn_text(btn_guardar_dif, "Agregar Diferencial")
        _set_btn_text(btn_limpiar_dif, "Limpiar")
        lbl_fase_titulo.value = (
            "Ensayo Fase R (0° y 180°)"
            if (dif_tipo.value or "TRIFASICO") == "TRIFASICO"
            else "Ensayo Bipolar Monofásico (0° y 180°)"
        )
        _cargar_borrador_en_campos()
        page.update()

    def editar_dif(indice: int):
        if not (0 <= indice < len(state["mediciones_dif"])):
            return
        m = state["mediciones_dif"][indice]
        runtime["dif_edit_idx"] = indice
        runtime["dif_fase_activa"] = "R"
        tipo_m = str(m.get("tipo", "BIPOLAR") or "BIPOLAR").upper()
        dif_tipo.value = "TRIFASICO" if tipo_m.startswith("TRI") else "BIPOLAR"
        dif_sec.value = str(m.get("sector", "TP") or "TP")
        dif_marca.value = str(m.get("marca", "") or "")
        dif_in.value = str(m.get("corriente_nominal_a", "") or "")
        sens_val = str(m.get("corriente_fuga_ma", "30") or "30")
        dif_sens.value = sens_val if sens_val in ("10", "30", "100", "300") else "30"

        fases_m = m.get("fases") or {}
        for f in ("R", "S", "T"):
            dif_draft_fases[f] = _fases_vacias()[f]
        if tipo_m.startswith("TRI"):
            for f in ("R", "S", "T"):
                fd = fases_m.get(f, {}) or {}
                for pol_k in ("0", "180"):
                    pd = fd.get(pol_k, {}) or {}
                    for k in ("dt_05", "dt_1", "dt_5", "id_ma"):
                        v = pd.get(k)
                        if v is not None:
                            dif_draft_fases[f][pol_k][k] = str(v)
        else:
            fd = fases_m.get("BIPOLAR", fases_m.get("R", {})) or {}
            for pol_k in ("0", "180"):
                pd = fd.get(pol_k, {}) or {}
                for k in ("dt_05", "dt_1", "dt_5", "id_ma"):
                    v = pd.get(k)
                    if v is not None:
                        dif_draft_fases["R"][pol_k][k] = str(v)
            if not fases_m and m.get("tiempo_disparo_ms"):
                dif_draft_fases["R"]["0"]["dt_1"] = str(m["tiempo_disparo_ms"])

        _set_btn_text(btn_guardar_dif, "Guardar cambios")
        _set_btn_text(btn_limpiar_dif, "Cancelar edición")
        fila_tabs_fases.visible = dif_tipo.value == "TRIFASICO"
        lbl_fase_titulo.value = (
            "Ensayo Fase R (0° y 180°)"
            if dif_tipo.value == "TRIFASICO"
            else "Ensayo Bipolar Monofásico (0° y 180°)"
        )
        _cargar_borrador_en_campos()
        page.update()
        notificar(f"Editando diferencial #{indice + 1} ({dif_sec.value})", ft.Colors.BLUE_700)

    def refrescar_lista_dif():
        dif_list_col.controls.clear()
        if not state["mediciones_dif"]:
            dif_list_col.controls.append(
                ft.Container(
                    padding=20,
                    content=ft.Text(
                        "No hay ensayos de diferenciales cargados todavía.",
                        color=_color("ON_SURFACE_VARIANT", ft.Colors.GREY_600),
                    ),
                )
            )
            actualizar_chrome()
            return

        for idx, m in enumerate(state["mediciones_dif"]):
            ev = evaluar_diferencial(m)
            m["cumple"] = ev["cumple"]
            es_apto = ev["cumple"] == "SI"
            badge = _badge(
                "APTO" if es_apto else "NO CUMPLE",
                ft.Colors.GREEN_700 if es_apto else ft.Colors.RED_700,
            )
            sub_lineas = [
                f"Marca: {m.get('marca') or '-'} | In: {m.get('corriente_nominal_a') or '-'}A | ΔTx1 ref: {m.get('tiempo_disparo_ms') or '-'} ms"
            ]
            if not es_apto and ev["motivos"]:
                sub_lineas.append("Fallas: " + "; ".join(ev["motivos"][:2]))

            def _crear_borrador_dif(i_del):
                def _borrar(e):
                    if 0 <= i_del < len(state["mediciones_dif"]):
                        state["mediciones_dif"].pop(i_del)
                        if runtime["dif_edit_idx"] == i_del:
                            limpiar_formulario_dif()
                        elif runtime["dif_edit_idx"] is not None and runtime["dif_edit_idx"] > i_del:
                            runtime["dif_edit_idx"] -= 1
                        refrescar_lista_dif()
                        autoguardar()
                        page.update()

                return _borrar

            dif_list_col.controls.append(
                ft.Container(
                    border=_borde(1, ft.Colors.GREEN_400 if es_apto else ft.Colors.RED_400),
                    border_radius=10,
                    padding=_padding_sim(horizontal=10, vertical=8),
                    content=ft.Row(
                        [
                            ft.Icon(
                                ft.Icons.SHIELD,
                                color=ft.Colors.GREEN_700 if es_apto else ft.Colors.RED_700,
                            ),
                            ft.Column(
                                [
                                    ft.Row(
                                        [
                                            ft.Text(
                                                f"#{idx + 1} [{m.get('tipo', 'BIPOLAR')}] {m['sector']} · {m['corriente_fuga_ma']}mA",
                                                weight=ft.FontWeight.W_600,
                                                size=13.5,
                                                expand=True,
                                            ),
                                            badge,
                                        ],
                                        vertical_alignment=ft.CrossAxisAlignment.CENTER,
                                    ),
                                    ft.Text(
                                        "\n".join(sub_lineas),
                                        size=11.5,
                                        color=_color("ON_SURFACE_VARIANT", ft.Colors.GREY_600),
                                    ),
                                ],
                                spacing=3,
                                expand=True,
                            ),
                            ft.IconButton(
                                ft.Icons.EDIT,
                                tooltip="Editar diferencial",
                                on_click=lambda e, i_ed=idx: editar_dif(i_ed),
                            ),
                            ft.IconButton(
                                ft.Icons.DELETE_OUTLINE,
                                icon_color=ft.Colors.RED_600,
                                tooltip="Eliminar diferencial",
                                on_click=_crear_borrador_dif(idx),
                            ),
                        ],
                        vertical_alignment=ft.CrossAxisAlignment.CENTER,
                        spacing=8,
                    ),
                )
            )
        actualizar_chrome()

    def add_or_update_dif(e):
        ubic = (dif_sec.value or "").strip()
        if not ubic:
            _set_tf_error(dif_sec, "Ubicación obligatoria")
            _enfocar(page, dif_sec)
            page.update()
            notificar("Ingresá la Ubicación del disyuntor (ej. TP, TG)", ft.Colors.RED_700)
            return
        _set_tf_error(dif_sec, None)

        med = _construir_med_borrador()
        ev = evaluar_diferencial(med)
        if not ev["tiene_mediciones"]:
            notificar("Cargá al menos un tiempo ΔTx1 (ms) en Fase R / Bipolar", ft.Colors.RED_700)
            return

        med["cumple"] = ev["cumple"]
        edit_idx = runtime["dif_edit_idx"]
        if edit_idx is not None and 0 <= edit_idx < len(state["mediciones_dif"]):
            state["mediciones_dif"][edit_idx] = med
            msg_ok = "Diferencial actualizado"
        else:
            state["mediciones_dif"].append(med)
            msg_ok = "Ensayo de diferencial agregado"

        limpiar_formulario_dif()
        refrescar_lista_dif()
        autoguardar()
        page.update()
        notificar(
            msg_ok if med["cumple"] == "SI" else f"{msg_ok} (NO CUMPLE norma)",
            ft.Colors.GREEN_700 if med["cumple"] == "SI" else ft.Colors.ORANGE_800,
        )

    btn_guardar_dif = _Btn(
        "Agregar Diferencial",
        on_click=add_or_update_dif,
        icon=ft.Icons.ADD,
        bgcolor=ft.Colors.BLUE_700,
        color=ft.Colors.WHITE,
        height=48,
        expand=True,
    )
    btn_limpiar_dif = ft.OutlinedButton(
        "Limpiar",
        icon=ft.Icons.CLOSE,
        on_click=limpiar_formulario_dif,
        height=48,
    )

    card_dif_form = _card(
        "Ensayo de Dispositivo Diferencial (RCD)",
        ft.Column(
            [
                ft.ResponsiveRow([dif_tipo, dif_sec, dif_sens, dif_marca, dif_in], spacing=10, run_spacing=10),
                fila_tabs_fases,
                lbl_fase_titulo,
                ft.Text("ΔTx 0,5 × IΔn (No debe disparar):", size=12, weight=ft.FontWeight.W_600),
                ft.ResponsiveRow([btn_0_dt05, btn_180_dt05, fld_0_dt05_ms, fld_180_dt05_ms], spacing=8, run_spacing=8),
                ft.Text("ΔTx 1 × IΔn (ms) — Límite RAEA: < 300 ms", size=12, weight=ft.FontWeight.W_600),
                ft.ResponsiveRow([fld_0_dt1, fld_180_dt1], spacing=8, run_spacing=8),
                ft.Text("ΔTx 5 × IΔn (ms) — Límite RAEA: < 40 ms", size=12, weight=ft.FontWeight.W_600),
                ft.ResponsiveRow([fld_0_dt5, fld_180_dt5], spacing=8, run_spacing=8),
                lbl_rango_id,
                ft.ResponsiveRow([fld_0_id, fld_180_id], spacing=8, run_spacing=8),
                dif_eval_banner,
                ft.Row([btn_guardar_dif, btn_limpiar_dif], spacing=8),
            ],
            spacing=8,
        ),
        icono=ft.Icons.SHIELD,
        col={"xs": 12, "md": 7},
    )

    card_dif_lista = _card(
        "Diferenciales Registrados",
        dif_list_col,
        icono=ft.Icons.CHECK_CIRCLE,
        col={"xs": 12, "md": 5},
    )

    view_dif = ft.Column(
        [
            ft.ResponsiveRow([card_dif_form, card_dif_lista], spacing=12, run_spacing=12),
        ],
        visible=False,
        spacing=12,
    )

    # ================= VISTA 4: REVISIÓN, CONCLUSIONES Y EXPORTAR =================
    checklist_col = ft.Column(spacing=6)

    def _on_texto_manual_change(e=None):
        runtime["textos_editados_manualmente"] = True
        state["conclusiones"]["textos_editados_manualmente"] = True
        state["conclusiones"]["observaciones_datos"] = txt_obs_datos.value or ""
        state["conclusiones"]["conclusiones_texto"] = txt_conclusiones.value or ""
        state["conclusiones"]["recomendaciones_texto"] = txt_recomendaciones.value or ""
        autoguardar()
        page.update()

    def _on_chk_adj_change(e=None):
        state["conclusiones"]["adjunta_calibracion"] = bool(chk_adj_cal.value)
        state["conclusiones"]["adjunta_plano"] = bool(chk_adj_plano.value)
        autoguardar()

    txt_obs_datos = _crear_textfield(
        label="(13) Observaciones generales",
        multiline=True,
        min_lines=3,
        max_lines=6,
        on_change=_on_texto_manual_change,
    )
    txt_conclusiones = _crear_textfield(
        label="Conclusiones",
        multiline=True,
        min_lines=3,
        max_lines=6,
        on_change=_on_texto_manual_change,
    )
    txt_recomendaciones = _crear_textfield(
        label="Recomendaciones para la adecuación",
        multiline=True,
        min_lines=3,
        max_lines=6,
        on_change=_on_texto_manual_change,
    )
    chk_adj_cal = ft.Checkbox(
        label="Adjunta (14) Certificado de Calibración",
        value=bool(state["conclusiones"].get("adjunta_calibracion")),
        on_change=_on_chk_adj_change,
    )
    chk_adj_plano = ft.Checkbox(
        label="Adjunta (15) Plano o croquis",
        value=bool(state["conclusiones"].get("adjunta_plano")),
        on_change=_on_chk_adj_change,
    )

    def armar_datos_finales() -> Dict[str, Any]:
        state["datos_cliente"]["razon_social"] = input_razon.value or ""
        state["datos_cliente"]["cuit"] = input_cuit.value or ""
        state["datos_cliente"]["direccion"] = input_dir.value or ""
        state["datos_cliente"]["localidad"] = input_loc.value or ""
        state["datos_cliente"]["cp"] = input_cp.value or ""
        state["datos_cliente"]["provincia"] = input_prov.value or "San Juan"
        state["datos_cliente"]["contrato"] = input_cont.value or ""
        state["datos_cliente"]["plano"] = input_plano.value or ""
        state["datos_cliente"]["condicion_terreno"] = input_terreno.value or "ARENOSO SECO"
        state["conclusiones"]["fecha_medicion"] = input_fecha.value or ""
        state["conclusiones"]["hora_inicio"] = input_hora_ini.value or ""
        state["conclusiones"]["hora_fin"] = input_hora_fin.value or ""
        state["conclusiones"]["instrumento_utilizado"] = input_inst.value or ""
        state["conclusiones"]["fecha_calibracion"] = input_cal.value or ""
        state["conclusiones"]["observaciones_datos"] = txt_obs_datos.value or ""
        state["conclusiones"]["conclusiones_texto"] = txt_conclusiones.value or ""
        state["conclusiones"]["recomendaciones_texto"] = txt_recomendaciones.value or ""
        state["conclusiones"]["adjunta_calibracion"] = bool(chk_adj_cal.value)
        state["conclusiones"]["adjunta_plano"] = bool(chk_adj_plano.value)
        return {
            "datos_cliente": state["datos_cliente"],
            "mediciones_pat": state["mediciones_pat"],
            "mediciones_diferenciales": state["mediciones_dif"],
            "conclusiones": state["conclusiones"],
        }

    def aplicar_textos_sugeridos(forzar: bool = False):
        datos = armar_datos_finales()
        sug = generar_textos_sugeridos(datos)
        if forzar or not runtime["textos_editados_manualmente"]:
            txt_obs_datos.value = sug["observaciones_datos"]
            txt_conclusiones.value = sug["conclusiones_texto"]
            txt_recomendaciones.value = sug["recomendaciones_texto"]
            state["conclusiones"]["observaciones_datos"] = sug["observaciones_datos"]
            state["conclusiones"]["conclusiones_texto"] = sug["conclusiones_texto"]
            state["conclusiones"]["recomendaciones_texto"] = sug["recomendaciones_texto"]
            if forzar:
                runtime["textos_editados_manualmente"] = False
                state["conclusiones"]["textos_editados_manualmente"] = False
                autoguardar()
                page.update()
                notificar("Observaciones y conclusiones actualizadas según las mediciones")

    def refrescar_vista_reporte():
        datos = armar_datos_finales()
        if not runtime["textos_editados_manualmente"]:
            aplicar_textos_sugeridos(forzar=False)

        chk_adj_cal.value = bool(
            state["conclusiones"].get(
                "adjunta_calibracion", bool((input_cal.value or "").strip())
            )
        )
        chk_adj_plano.value = bool(
            state["conclusiones"].get(
                "adjunta_plano", bool((input_plano.value or "").strip())
            )
        )

        res = resumir_relevamiento(datos)
        checklist_col.controls.clear()
        for item in res["checklist"]:
            nv = item["nivel"]
            if nv == "ok":
                ic = ft.Icon(ft.Icons.CHECK_CIRCLE, size=18, color=ft.Colors.GREEN_700)
            elif nv == "warn":
                ic = ft.Icon(ft.Icons.WARNING_AMBER, size=18, color=ft.Colors.AMBER_700)
            else:
                ic = ft.Icon(ft.Icons.ERROR, size=18, color=ft.Colors.RED_700)
            checklist_col.controls.append(
                ft.Row([ic, ft.Text(item["texto"], size=13, expand=True)], spacing=8)
            )

    chk_exp_pdf = ft.Checkbox(label="Exportar PDF (.pdf) — listo para firmar/enviar", value=True)
    chk_exp_xlsx = ft.Checkbox(label="Exportar Excel (.xlsx) — plantilla SRT 900/15", value=True)

    resultado_col = ft.Column(
        [
            ft.Text(
                "Elegí los formatos y presioná 'Generar Protocolo'. Se guardará con fecha para no sobrescribir relevamientos previos.",
                size=12.5,
                color=_color("ON_SURFACE_VARIANT", ft.Colors.GREY_600),
            )
        ],
        spacing=8,
    )
    resultado_container = ft.Container(
        content=resultado_col,
        padding=12,
        border_radius=10,
        border=_borde(1, _color("OUTLINE_VARIANT", ft.Colors.GREY_400)),
    )

    def exportar_protocolo(e):
        quiere_pdf = bool(chk_exp_pdf.value)
        quiere_xlsx = bool(chk_exp_xlsx.value)
        if not quiere_pdf and not quiere_xlsx:
            notificar("Seleccioná al menos un formato (PDF o Excel)", ft.Colors.RED_700)
            return

        try:
            datos_final = armar_datos_finales()
            carpeta = _directorio_salida_default()
            base = construir_nombre_archivo(
                datos_final["datos_cliente"].get("razon_social", "Cliente"),
                datos_final["conclusiones"].get("fecha_medicion", ""),
            )
            archivos_creados: List[Path] = []
            ruta_pdf_final: Optional[Path] = None

            if quiere_xlsx:
                ruta_xlsx = carpeta / f"{base}.xlsx"
                generar_protocolo(
                    ruta_plantilla_default("Mediciones 2026.xlsx"),
                    str(ruta_xlsx),
                    datos_final,
                )
                archivos_creados.append(ruta_xlsx)

            if quiere_pdf:
                ruta_pdf = carpeta / f"{base}.pdf"
                generar_protocolo_pdf(str(ruta_pdf), datos_final)
                archivos_creados.append(ruta_pdf)
                ruta_pdf_final = ruta_pdf

            runtime["ultimos_archivos"] = [str(p) for p in archivos_creados]
            resultado_col.controls.clear()
            resultado_col.controls.append(
                ft.Row(
                    [
                        ft.Icon(ft.Icons.CHECK_CIRCLE, color=ft.Colors.GREEN_700),
                        ft.Text("¡Protocolo generado con éxito!", weight=ft.FontWeight.BOLD, color=ft.Colors.GREEN_700),
                    ],
                    spacing=8,
                )
            )
            for p in archivos_creados:
                resultado_col.controls.append(ft.Text(f"• {p}", size=12, selectable=True))

            acciones_salida: List[ft.Control] = []
            if ruta_pdf_final is not None:
                acciones_salida.append(
                    _Btn(
                        "Abrir PDF",
                        icon=ft.Icons.OPEN_IN_NEW,
                        on_click=lambda _, r=ruta_pdf_final: (
                            notificar("Abriendo PDF...")
                            if _abrir_ruta_externa(page, r)
                            else notificar(f"PDF guardado en: {r}", ft.Colors.BLUE_700)
                        ),
                        bgcolor=ft.Colors.BLUE_700,
                        color=ft.Colors.WHITE,
                    )
                )
            acciones_salida.append(
                ft.OutlinedButton(
                    "Abrir carpeta",
                    icon=ft.Icons.FOLDER_OPEN,
                    on_click=lambda _, c=carpeta: (
                        notificar("Abriendo carpeta de salidas...")
                        if _abrir_ruta_externa(page, c)
                        else notificar(f"Carpeta: {c}", ft.Colors.BLUE_700)
                    ),
                )
            )
            resultado_col.controls.append(ft.Row(acciones_salida, wrap=True, spacing=8))
            notificar("¡Protocolo generado con éxito!")
            page.update()
        except Exception as ex:
            resultado_col.controls.clear()
            resultado_col.controls.append(
                ft.Row(
                    [
                        ft.Icon(ft.Icons.ERROR, color=ft.Colors.RED_700),
                        ft.Text(f"No se pudo generar el protocolo: {ex}", color=ft.Colors.RED_700, expand=True),
                    ],
                    spacing=8,
                )
            )
            resultado_col.controls.append(
                _crear_expansion_tile(
                    title=ft.Text("Ver detalle técnico", size=12),
                    expanded=False,
                    controls=[
                        ft.Container(
                            padding=8,
                            content=ft.Text(traceback.format_exc(), size=11, selectable=True),
                        )
                    ],
                )
            )
            notificar(f"Error al generar: {ex}", ft.Colors.RED_700)
            page.update()

    # ================= CARGA DE PROTOCOLOS / EJEMPLO / NUEVO =================
    def _aplicar_datos_cargados_en_ui(datos_cargados: Dict[str, Any], textos_manuales: bool = False):
        runtime["suspend_autosave"] = True
        try:
            nuevo = _estado_limpio()
            nuevo["datos_cliente"].update(datos_cargados.get("datos_cliente", {}) or {})
            nuevo["conclusiones"].update(datos_cargados.get("conclusiones", {}) or {})
            nuevo["mediciones_pat"] = list(datos_cargados.get("mediciones_pat", []) or [])
            difs_in = (
                datos_cargados.get("mediciones_diferenciales")
                if datos_cargados.get("mediciones_diferenciales") is not None
                else datos_cargados.get("mediciones_dif", [])
            ) or []
            nuevo["mediciones_dif"] = list(difs_in)

            state.clear()
            state.update(nuevo)

            input_razon.value = state["datos_cliente"].get("razon_social", "")
            input_cuit.value = state["datos_cliente"].get("cuit", "")
            input_cont.value = state["datos_cliente"].get("contrato", "")
            input_dir.value = state["datos_cliente"].get("direccion", "")
            input_loc.value = state["datos_cliente"].get("localidad", "Capital")
            input_cp.value = state["datos_cliente"].get("cp", "5400")
            input_plano.value = state["datos_cliente"].get("plano", "")
            prov_v = state["datos_cliente"].get("provincia", "San Juan")
            input_prov.value = prov_v if prov_v in PROVINCIAS_ARGENTINA else "San Juan"
            terr_v = state["datos_cliente"].get("condicion_terreno", "ARENOSO SECO")
            input_terreno.value = terr_v if terr_v in CONDICIONES_TERRENO else "ARENOSO SECO"

            input_fecha.value = state["conclusiones"].get("fecha_medicion") or datetime.date.today().strftime("%d/%m/%Y")
            input_hora_ini.value = state["conclusiones"].get("hora_inicio", "")
            input_hora_fin.value = state["conclusiones"].get("hora_fin", "")
            if state["conclusiones"].get("instrumento_utilizado"):
                input_inst.value = state["conclusiones"]["instrumento_utilizado"]
            if state["conclusiones"].get("fecha_calibracion"):
                input_cal.value = state["conclusiones"]["fecha_calibracion"]

            txt_obs_datos.value = state["conclusiones"].get("observaciones_datos", "")
            txt_conclusiones.value = state["conclusiones"].get("conclusiones_texto", "")
            txt_recomendaciones.value = state["conclusiones"].get("recomendaciones_texto", "")
            runtime["textos_editados_manualmente"] = textos_manuales
            state["conclusiones"]["textos_editados_manualmente"] = textos_manuales

            refrescar_banner_calibracion()
            refrescar_lista_pat()
            refrescar_lista_dif()
            refrescar_vista_reporte()
        finally:
            runtime["suspend_autosave"] = False
        autoguardar()
        page.update()

    def _cargar_archivo_excel_en_app(ruta_xlsx: str, nombre_visible: str):
        try:
            datos_cargados = cargar_protocolo_excel(ruta_xlsx)
            if not datos_cargados["datos_cliente"].get("razon_social") and "ORTOPEDIA" in nombre_visible.upper():
                datos_cargados["datos_cliente"]["razon_social"] = "ORTOPEDIA CUYO"
            _aplicar_datos_cargados_en_ui(datos_cargados, textos_manuales=False)
            notificar(
                f"Cargado '{nombre_visible}': {len(state['mediciones_pat'])} PAT, {len(state['mediciones_dif'])} RCD"
            )
        except Exception as ex:
            notificar(f"No se pudo abrir '{nombre_visible}': {ex}", ft.Colors.RED_700)

    def _confirmar_si_hay_datos(titulo: str, mensaje: str, accion_confirmada):
        hay_datos = bool(
            (input_razon.value or "").strip()
            or state["mediciones_pat"]
            or state["mediciones_dif"]
        )
        if not hay_datos:
            accion_confirmada()
            return

        def proceder(e):
            _cerrar_dialogo(page, dlg)
            accion_confirmada()

        dlg = ft.AlertDialog(
            modal=True,
            title=ft.Text(titulo),
            content=ft.Text(mensaje),
            actions=[
                ft.TextButton("Cancelar", on_click=lambda _: _cerrar_dialogo(page, dlg)),
                _Btn("Continuar", on_click=proceder, bgcolor=ft.Colors.RED_700, color=ft.Colors.WHITE),
            ],
        )
        _abrir_dialogo(page, dlg)

    def abrir_dialogo_cargar(e=None):
        carpeta_sal = _directorio_salida_default()
        archivos = listar_protocolos_xlsx(carpeta_sal)
        items_lista: List[ft.Control] = []
        dlg_ref: Dict[str, Any] = {}

        def elegir_y_cargar(ruta_str: str, label_str: str):
            _cerrar_dialogo(page, dlg_ref["dlg"])
            _confirmar_si_hay_datos(
                "¿Reemplazar relevamiento actual?",
                f"Se cargará '{label_str}' y se reemplazarán los datos actuales en pantalla.",
                lambda: _cargar_archivo_excel_en_app(ruta_str, label_str),
            )

        ruta_ej = ruta_plantilla_default("MEDICIONES_ORTOPEDIA_CUYO.xlsx")
        items_lista.append(
            ft.ListTile(
                leading=ft.Icon(ft.Icons.UPLOAD_FILE, color=ft.Colors.BLUE_700),
                title=ft.Text("Ortopedia Cuyo (Ejemplo de referencia)"),
                subtitle=ft.Text("30 tomas PAT + 1 diferencial trifásico"),
                on_click=lambda _: elegir_y_cargar(ruta_ej, "MEDICIONES_ORTOPEDIA_CUYO.xlsx"),
            )
        )
        items_lista.append(ft.Divider(height=1))

        if archivos:
            for p in archivos[:12]:
                fecha_mod = datetime.datetime.fromtimestamp(p.stat().st_mtime).strftime("%d/%m/%Y %H:%M")
                items_lista.append(
                    ft.ListTile(
                        leading=ft.Icon(ft.Icons.TABLE_CHART, color=ft.Colors.GREEN_700),
                        title=ft.Text(p.name),
                        subtitle=ft.Text(f"Modificado: {fecha_mod}"),
                        on_click=lambda _, rp=str(p), np=p.name: elegir_y_cargar(rp, np),
                    )
                )
        else:
            items_lista.append(
                ft.Container(
                    padding=12,
                    content=ft.Text(
                        f"No hay archivos .xlsx guardados todavía en {carpeta_sal}",
                        size=12,
                        color=_color("ON_SURFACE_VARIANT", ft.Colors.GREY_600),
                    ),
                )
            )

        dlg = ft.AlertDialog(
            title=ft.Text("Abrir relevamiento (.xlsx)"),
            content=ft.Container(
                width=420,
                height=340,
                content=ft.Column(items_lista, scroll=ft.ScrollMode.AUTO, spacing=4),
            ),
            actions=[
                ft.TextButton("Cerrar", on_click=lambda _: _cerrar_dialogo(page, dlg)),
            ],
        )
        dlg_ref["dlg"] = dlg
        _abrir_dialogo(page, dlg)

    def confirmar_nuevo_relevamiento(e=None):
        def reiniciar():
            limpiar_borrador()
            limpio = _estado_limpio()
            _aplicar_datos_cargados_en_ui(limpio, textos_manuales=False)
            limpiar_formulario_dif()
            appbar_subtitulo.value = "SRT 900/15 · Nuevo relevamiento"
            page.update()
            notificar("Relevamiento reiniciado (se conservó tu instrumental)")

        _confirmar_si_hay_datos(
            "¿Iniciar nuevo relevamiento?",
            "Se limpiarán los datos del cliente y las mediciones cargadas. Tu instrumental guardado se conservará.",
            reiniciar,
        )

    def alternar_tema(e=None):
        es_oscuro_ahora = page.theme_mode == ft.ThemeMode.DARK
        page.theme_mode = ft.ThemeMode.LIGHT if es_oscuro_ahora else ft.ThemeMode.DARK
        btn_tema.icon = ft.Icons.DARK_MODE if es_oscuro_ahora else ft.Icons.LIGHT_MODE
        perfil["tema_oscuro"] = not es_oscuro_ahora
        guardar_perfil_electricista(perfil)
        refrescar_banner_calibracion()
        refrescar_validacion_rcd()
        page.update()

    btn_tema = ft.IconButton(
        ft.Icons.LIGHT_MODE if page.theme_mode == ft.ThemeMode.DARK else ft.Icons.DARK_MODE,
        tooltip="Cambiar modo claro / oscuro",
        on_click=alternar_tema,
    )

    page.appbar = ft.AppBar(
        title=ft.Column([appbar_titulo, appbar_subtitulo], spacing=1),
        actions=[
            ft.IconButton(
                ft.Icons.FOLDER_OPEN,
                tooltip="Abrir relevamiento o ejemplo",
                on_click=abrir_dialogo_cargar,
            ),
            ft.IconButton(
                ft.Icons.RESTART_ALT,
                tooltip="Nuevo relevamiento",
                on_click=confirmar_nuevo_relevamiento,
            ),
            btn_tema,
        ],
    )

    card_revision = _card(
        "Revisión Pre-Exportación y Conclusiones",
        ft.Column(
            [
                checklist_col,
                ft.Divider(height=1),
                ft.Row(
                    [
                        ft.Text("Textos del protocolo", weight=ft.FontWeight.W_600),
                        ft.OutlinedButton(
                            "Actualizar textos según mediciones",
                            icon=ft.Icons.REFRESH,
                            on_click=lambda _: aplicar_textos_sugeridos(forzar=True),
                        ),
                    ],
                    alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                    spacing=8,
                    run_spacing=8,
                    wrap=True,
                    vertical_alignment=ft.CrossAxisAlignment.CENTER,
                ),
                txt_obs_datos,
                txt_conclusiones,
                txt_recomendaciones,
                chk_adj_cal,
                chk_adj_plano,
            ],
            spacing=10,
        ),
        icono=ft.Icons.DESCRIPTION,
        col={"xs": 12, "md": 7},
    )

    card_exportar = _card(
        "Exportar Protocolo SRT 900/15",
        ft.Column(
            [
                chk_exp_pdf,
                chk_exp_xlsx,
                ft.Row(
                    [
                        _Btn(
                            "Generar Protocolo",
                            icon=ft.Icons.DONE_ALL,
                            on_click=exportar_protocolo,
                            bgcolor=ft.Colors.GREEN_700,
                            color=ft.Colors.WHITE,
                            height=54,
                            expand=True,
                        )
                    ]
                ),
                resultado_container,
                ft.Divider(height=1),
                ft.OutlinedButton(
                    "Abrir protocolo (.xlsx) / Ejemplo",
                    icon=ft.Icons.UPLOAD_FILE,
                    on_click=abrir_dialogo_cargar,
                    height=44,
                ),
            ],
            spacing=10,
        ),
        icono=ft.Icons.PICTURE_AS_PDF,
        col={"xs": 12, "md": 5},
    )

    view_reporte = ft.Column(
        [
            ft.ResponsiveRow([card_revision, card_exportar], spacing=12, run_spacing=12),
        ],
        visible=False,
        spacing=12,
    )

    # ================= NAVEGACIÓN =================
    views = [view_datos, view_pat, view_dif, view_reporte]

    def ir_a_vista(idx: int):
        for i, v in enumerate(views):
            v.visible = i == idx
        if page.navigation_bar is not None:
            page.navigation_bar.selected_index = idx
        if idx == 3:
            refrescar_vista_reporte()
        page.update()

    def on_nav_change(e):
        ir_a_vista(e.control.selected_index)

    page.navigation_bar = ft.NavigationBar(
        destinations=[dest_datos, dest_pat, dest_dif, dest_exportar],
        on_change=on_nav_change,
        selected_index=0,
    )

    # Inicialización de estado visual en pantalla
    refrescar_chips_sectores()
    refrescar_banner_calibracion()
    refrescar_lista_pat()
    _cargar_borrador_en_campos()
    refrescar_lista_dif()
    if state["conclusiones"].get("conclusiones_texto") or state["conclusiones"].get("observaciones_datos"):
        txt_obs_datos.value = state["conclusiones"].get("observaciones_datos", "")
        txt_conclusiones.value = state["conclusiones"].get("conclusiones_texto", "")
        txt_recomendaciones.value = state["conclusiones"].get("recomendaciones_texto", "")
    actualizar_chrome()
    if borrador_inicial and (borrador_inicial.get("ultimo_guardado")):
        appbar_subtitulo.value = f"SRT 900/15 · Borrador ({borrador_inicial['ultimo_guardado']})"

    page.add(
        ft.Container(
            content=ft.Column(views, spacing=12),
            padding=12,
        )
    )

    if borrador_inicial and (
        state["mediciones_pat"]
        or state["mediciones_dif"]
        or (state["datos_cliente"].get("razon_social") or "").strip()
    ):
        notificar(
            f"Borrador restaurado ({len(state['mediciones_pat'])} PAT, {len(state['mediciones_dif'])} RCD)",
            ft.Colors.BLUE_700,
        )


# Compatibilidad entre Flet 0.28.3 (ft.app) y Flet 0.80+ (ft.run)
if hasattr(ft, "run"):
    ft.run(main)
else:
    ft.app(target=main)
