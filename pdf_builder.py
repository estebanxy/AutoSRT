import datetime
import logging
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from excel_builder import es_medicion_pat_valida
from srt_logic import (
    evaluar_diferencial,
    evaluar_parametro_rcd,
    generar_textos_sugeridos,
    parsear_float,
)

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas as pdfcanvas
    from reportlab.platypus import (
        KeepTogether,
        Paragraph,
        SimpleDocTemplate,
        Spacer,
        Table,
        TableStyle,
    )

    _REPORTLAB_DISPONIBLE = True
except ImportError:
    _REPORTLAB_DISPONIBLE = False


# ==============================================================================
# HELPERS COMPARTIDOS PARA AMBOS MOTORES PDF
# ==============================================================================


def _preparar_contexto_pdf(datos: Dict[str, Any]) -> Dict[str, Any]:
    """Normaliza los datos del protocolo y calcula textos/estados normativos compartidos."""
    datos_cliente = datos.get("datos_cliente", {}) or {}
    mediciones_pat = datos.get("mediciones_pat", []) or []
    mediciones_dif = (
        datos.get("mediciones_diferenciales")
        if datos.get("mediciones_diferenciales") is not None
        else datos.get("mediciones_dif", [])
    ) or []
    conclusiones = datos.get("conclusiones", {}) or {}

    razon_social = str(datos_cliente.get("razon_social", "") or "-").strip()
    cuit = str(datos_cliente.get("cuit", "") or "-").strip()
    direccion = str(datos_cliente.get("direccion", "") or "-").strip()
    localidad = str(datos_cliente.get("localidad", "Capital") or "Capital").strip()
    provincia = str(datos_cliente.get("provincia", "San Juan") or "San Juan").strip()
    cp = str(datos_cliente.get("cp", "5400") or "5400").strip()
    contrato = str(
        datos_cliente.get("contrato", datos_cliente.get("nro_contrato", "")) or "-"
    ).strip()
    plano = str(datos_cliente.get("plano", "") or "-").strip()
    condicion_gral = str(
        datos_cliente.get("condicion_terreno", "ARENOSO SECO") or "ARENOSO SECO"
    ).strip()

    hoy_str = datetime.date.today().strftime("%d/%m/%Y")
    fecha_med = str(
        conclusiones.get("fecha_medicion", datos_cliente.get("fecha_medicion", hoy_str))
        or hoy_str
    ).strip()
    hora_ini = str(
        conclusiones.get("hora_inicio", datos_cliente.get("hora_inicio", "")) or "-"
    ).strip()
    hora_fin = str(
        conclusiones.get("hora_fin", datos_cliente.get("hora_fin", "")) or "-"
    ).strip()
    instrumento = str(
        conclusiones.get("instrumento_utilizado", datos_cliente.get("instrumento", ""))
        or "Telurímetro / Multifunción certificado"
    ).strip()
    fecha_cal = str(
        conclusiones.get("fecha_calibracion", datos_cliente.get("fecha_calibracion", ""))
        or "-"
    ).strip()
    metodologia = str(
        conclusiones.get("metodologia", "")
        or "Medición de impedancia por lazo de falla y comprobación de RCD"
    ).strip()

    sugeridos = generar_textos_sugeridos(datos)
    obs_texto = str(
        conclusiones.get("observaciones_datos", "") or sugeridos["observaciones_datos"]
    ).strip()
    txt_conc = str(
        conclusiones.get("conclusiones_texto", conclusiones.get("observaciones", ""))
        or sugeridos["conclusiones_texto"]
    ).strip()
    txt_rec = str(
        conclusiones.get("recomendaciones_texto", "") or sugeridos["recomendaciones_texto"]
    ).strip()

    adj_cal = bool(
        conclusiones.get("adjunta_calibracion", bool(fecha_cal and fecha_cal != "-"))
    )
    adj_plano = bool(conclusiones.get("adjunta_plano", bool(plano and plano != "-")))

    return {
        "datos_cliente": datos_cliente,
        "mediciones_pat": mediciones_pat,
        "mediciones_dif": mediciones_dif,
        "conclusiones": conclusiones,
        "razon_social": razon_social,
        "cuit": cuit,
        "direccion": direccion,
        "localidad": localidad,
        "provincia": provincia,
        "cp": cp,
        "contrato": contrato,
        "plano": plano,
        "condicion_gral": condicion_gral,
        "fecha_med": fecha_med,
        "hora_ini": hora_ini,
        "hora_fin": hora_fin,
        "instrumento": instrumento,
        "fecha_cal": fecha_cal,
        "metodologia": metodologia,
        "obs_texto": obs_texto,
        "txt_conc": txt_conc,
        "txt_rec": txt_rec,
        "adj_cal": adj_cal,
        "adj_plano": adj_plano,
    }


def _formatear_tipo_marca_rcd(d: Dict[str, Any]) -> str:
    tipo = str(d.get("tipo", "BIPOLAR") or "BIPOLAR").upper()
    marca = str(d.get("marca", "") or "").strip()
    in_a = str(d.get("corriente_nominal_a", "") or "").strip()
    partes = [tipo]
    if marca and marca != "-":
        partes.append(marca)
    desc = " ".join(partes)
    if in_a:
        desc += f" ({in_a}A)"
    return desc


def _filas_dispositivo_rcd(idx_d: int, d: Dict[str, Any]) -> Tuple[List[List[str]], List[Tuple[int, int]]]:
    """
    Construye las filas de tabla para un dispositivo RCD y devuelve además una lista
    de coordenadas (fila_local, col) que deben resaltarse por estar fuera de norma.
    """
    ubic = str(d.get("sector", d.get("ubicacion", f"RCD {idx_d}"))).strip()
    tipo = str(d.get("tipo", "BIPOLAR") or "BIPOLAR").upper()
    ian_str = str(d.get("corriente_fuga_ma", "30") or "30").strip()
    ian_val = parsear_float(ian_str) or 30.0
    desc_marca = _formatear_tipo_marca_rcd(d)

    eval_rcd = evaluar_diferencial(d)
    estado_txt = "APTO" if eval_rcd["cumple"] == "SI" else "NO CUMPLE"

    filas: List[List[str]] = []
    alertas: List[Tuple[int, int]] = []

    def _fmt(v: Any, default: str = "-") -> str:
        if v is None:
            return default
        s = str(v).strip()
        return s if s else default

    fases_dict = d.get("fases", {}) or {}
    if tipo.startswith("TRI") and fases_dict:
        for fase_l in ("R", "S", "T"):
            fd = fases_dict.get(fase_l, {}) or {}
            for pol_k, pol_label in (("0", "Fase 0° (+)"), ("180", "Fase 180° (-)")):
                p_data = fd.get(pol_k, {}) or {}
                row_idx = len(filas)
                dt05 = p_data.get("dt_05")
                dt1 = p_data.get("dt_1")
                dt5 = p_data.get("dt_5")
                id_ma = p_data.get("id_ma")
                t_ms = p_data.get("t_ms")

                es_primera = fase_l == "R" and pol_k == "0"
                estado_celda = str(t_ms) if t_ms is not None else (estado_txt if es_primera else "")
                filas.append(
                    [
                        ubic if es_primera else "",
                        desc_marca if es_primera else "",
                        ian_str if es_primera else "",
                        fase_l if pol_k == "0" else "",
                        pol_label,
                        _fmt(dt05, "-"),
                        _fmt(dt1, "-"),
                        _fmt(dt5, "-"),
                        _fmt(id_ma, "-"),
                        estado_celda,
                    ]
                )
                for col_i, (param_k, raw_v) in (
                    (5, ("dt_05", dt05)),
                    (6, ("dt_1", dt1)),
                    (7, ("dt_5", dt5)),
                    (8, ("id_ma", id_ma)),
                ):
                    ok, _ = evaluar_parametro_rcd(param_k, raw_v, ian_val)
                    if ok is False:
                        alertas.append((row_idx, col_i))
                if es_primera and eval_rcd["cumple"] == "NO":
                    alertas.append((row_idx, 9))
    elif fases_dict and ("BIPOLAR" in fases_dict or "R" in fases_dict):
        fd = fases_dict.get("BIPOLAR", fases_dict.get("R", {})) or {}
        for pol_k, pol_label in (("0", "Fase 0° (+)"), ("180", "Fase 180° (-)")):
            p_data = fd.get(pol_k, {}) or {}
            row_idx = len(filas)
            dt05 = p_data.get("dt_05")
            dt1 = p_data.get("dt_1")
            dt5 = p_data.get("dt_5")
            id_ma = p_data.get("id_ma")
            t_ms = p_data.get("t_ms")
            es_primera = pol_k == "0"
            estado_celda = str(t_ms) if t_ms is not None else (estado_txt if es_primera else "")
            filas.append(
                [
                    ubic if es_primera else "",
                    desc_marca if es_primera else "",
                    ian_str if es_primera else "",
                    "Monof." if es_primera else "",
                    pol_label,
                    _fmt(dt05, "-"),
                    _fmt(dt1, "-"),
                    _fmt(dt5, "-"),
                    _fmt(id_ma, "-"),
                    estado_celda,
                ]
            )
            for col_i, (param_k, raw_v) in (
                (5, ("dt_05", dt05)),
                (6, ("dt_1", dt1)),
                (7, ("dt_5", dt5)),
                (8, ("id_ma", id_ma)),
            ):
                ok, _ = evaluar_parametro_rcd(param_k, raw_v, ian_val)
                if ok is False:
                    alertas.append((row_idx, col_i))
            if es_primera and eval_rcd["cumple"] == "NO":
                alertas.append((row_idx, 9))
    else:
        t_ms = d.get("tiempo_disparo_ms")
        filas.append(
            [
                ubic,
                desc_marca,
                ian_str,
                "Monof.",
                "Fase 0° (+)",
                "ND",
                _fmt(t_ms, "-"),
                "-",
                "-",
                estado_txt,
            ]
        )
        ok, _ = evaluar_parametro_rcd("dt_1", t_ms, ian_val)
        if ok is False:
            alertas.append((0, 6))
        if eval_rcd["cumple"] == "NO":
            alertas.append((0, 9))

    return filas, alertas


# ==============================================================================
# MOTOR 1: REPORTLAB (ESCRITORIO)
# ==============================================================================

if _REPORTLAB_DISPONIBLE:

    class _NumberedCanvas(pdfcanvas.Canvas):
        """Canvas que numera 'Hoja X/Y' y dibuja la línea de firma en todas las páginas."""

        def __init__(self, *args, **kwargs):
            super().__init__(*args, **kwargs)
            self._saved_page_states: List[Dict[str, Any]] = []

        def showPage(self):
            self._saved_page_states.append(dict(self.__dict__))
            self._startPage()

        def save(self):
            total_pages = len(self._saved_page_states)
            for idx, state in enumerate(self._saved_page_states, start=1):
                self.__dict__.update(state)
                self._dibujar_pie(idx, total_pages)
                super().showPage()
            super().save()

        def _dibujar_pie(self, page_num: int, total_pages: int):
            self.saveState()
            ancho, _ = self._pagesize
            self.setFont("Helvetica", 8)
            self.setStrokeColorRGB(0.4, 0.4, 0.4)
            x_firma = ancho - 95 * mm
            y_firma = 16 * mm
            self.line(x_firma, y_firma, ancho - 20 * mm, y_firma)
            self.drawCentredString(
                (x_firma + ancho - 20 * mm) / 2,
                12 * mm,
                "Firma, Aclaración y Registro del Profesional Interviniente",
            )
            self.drawString(
                15 * mm, 12 * mm, f"Protocolo SRT 900/15  |  Hoja {page_num}/{total_pages}"
            )
            self.restoreState()


def _generar_pdf_reportlab(ruta_salida: str, datos: Dict[str, Any]) -> str:
    """Genera el protocolo SRT 900/15 completo en PDF usando ReportLab."""
    Path(ruta_salida).parent.mkdir(parents=True, exist_ok=True)
    ctx = _preparar_contexto_pdf(datos)

    doc = SimpleDocTemplate(
        ruta_salida,
        pagesize=landscape(A4),
        leftMargin=14 * mm,
        rightMargin=14 * mm,
        topMargin=12 * mm,
        bottomMargin=24 * mm,
        title=f"Protocolo SRT 900/15 - {ctx['razon_social']}",
    )

    styles = getSampleStyleSheet()
    st_titulo = ParagraphStyle(
        "TituloSRT",
        parent=styles["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=12,
        leading=15,
        alignment=1,
        textColor=colors.HexColor("#0F172A"),
        spaceAfter=6,
    )
    st_subtitulo = ParagraphStyle(
        "SubtituloSRT",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=10,
        leading=13,
        textColor=colors.HexColor("#1E293B"),
        spaceBefore=6,
        spaceAfter=4,
    )
    st_celda = ParagraphStyle(
        "CeldaSRT",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8.5,
        leading=10.5,
    )
    st_celda_bold = ParagraphStyle(
        "CeldaBoldSRT",
        parent=st_celda,
        fontName="Helvetica-Bold",
    )
    st_header_tabla = ParagraphStyle(
        "HeaderTablaSRT",
        parent=st_celda_bold,
        fontSize=8,
        leading=9.5,
        alignment=1,
        textColor=colors.white,
    )

    story: List[Any] = []

    def bloque_encabezado_cliente():
        t_head = Table(
            [
                [
                    Paragraph(
                        "<b>PROTOCOLO DE MEDICIÓN DE PUESTA A TIERRA Y CONTINUIDAD DE LAS MASAS (SRT 900/15)</b>",
                        st_titulo,
                    )
                ]
            ],
            colWidths=[269 * mm],
        )
        t_head.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#E2E8F0")),
                    ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#334155")),
                    ("TOPPADDING", (0, 0), (-1, -1), 5),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ]
            )
        )

        t_cli = Table(
            [
                [
                    Paragraph(f"<b>Razón Social:</b> {ctx['razon_social']}", st_celda),
                    Paragraph(f"<b>C.U.I.T.:</b> {ctx['cuit']}", st_celda),
                    Paragraph(f"<b>N° Contrato:</b> {ctx['contrato']}", st_celda),
                    Paragraph(f"<b>Fecha:</b> {ctx['fecha_med']}", st_celda),
                ],
                [
                    Paragraph(f"<b>Dirección:</b> {ctx['direccion']}", st_celda),
                    Paragraph(f"<b>Localidad:</b> {ctx['localidad']}", st_celda),
                    Paragraph(f"<b>CP:</b> {ctx['cp']}", st_celda),
                    Paragraph(f"<b>Provincia:</b> {ctx['provincia']}", st_celda),
                ],
            ],
            colWidths=[105 * mm, 62 * mm, 52 * mm, 50 * mm],
        )
        t_cli.setStyle(
            TableStyle(
                [
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#94A3B8")),
                    ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F8FAFC")),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("TOPPADDING", (0, 0), (-1, -1), 4),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ]
            )
        )
        return [t_head, Spacer(1, 3 * mm), t_cli, Spacer(1, 4 * mm)]

    # ==================== SECCIÓN 1: DATOS Y METODOLOGÍA ====================
    story.extend(bloque_encabezado_cliente())
    story.append(Paragraph("1. Datos para la Medición y Metodología", st_subtitulo))

    chk_cal = "[X]" if ctx["adj_cal"] else "[ ]"
    chk_pla = "[X]" if ctx["adj_plano"] else "[ ]"

    tabla_datos_med = Table(
        [
            [
                Paragraph(
                    f"<b>(7) Marca, modelo y N° de serie del instrumento utilizado:</b><br/>{ctx['instrumento']}",
                    st_celda,
                ),
                Paragraph(
                    f"<b>(8) Fecha de Calibración del Instrumental:</b><br/>{ctx['fecha_cal']}",
                    st_celda,
                ),
                Paragraph(f"<b>Plano N°:</b><br/>{ctx['plano']}", st_celda),
            ],
            [
                Paragraph(f"<b>(9) Fecha de la medición:</b> {ctx['fecha_med']}", st_celda),
                Paragraph(f"<b>(10) Hora de inicio:</b> {ctx['hora_ini']}", st_celda),
                Paragraph(f"<b>(11) Hora finalización:</b> {ctx['hora_fin']}", st_celda),
            ],
            [
                Paragraph(
                    f"<b>(12) Metodología utilizada:</b> {ctx['metodologia']}",
                    st_celda,
                ),
                "",
                "",
            ],
            [
                Paragraph(f"<b>(13) Observaciones:</b> {ctx['obs_texto']}", st_celda),
                "",
                "",
            ],
            [
                Paragraph(
                    "<b>Documentación que se adjuntará a la medición:</b>  "
                    f"{chk_cal} (14) Certificado de Calibración     {chk_pla} (15) Plano o croquis",
                    st_celda,
                ),
                "",
                "",
            ],
        ],
        colWidths=[125 * mm, 90 * mm, 54 * mm],
    )
    tabla_datos_med.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#64748B")),
                ("SPAN", (0, 2), (2, 2)),
                ("SPAN", (0, 3), (2, 3)),
                ("SPAN", (0, 4), (2, 4)),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    story.append(tabla_datos_med)
    story.append(Spacer(1, 5 * mm))

    # ==================== SECCIÓN 2: GRILLA DE PUESTA A TIERRA ====================
    story.append(
        Paragraph(
            "2. Planilla de Medición de Puesta a Tierra y Continuidad de las Masas",
            st_subtitulo,
        )
    )

    encabezado_grilla = [
        Paragraph("<b>N° Toma</b>", st_header_tabla),
        Paragraph("<b>Sector / Ubicación</b>", st_header_tabla),
        Paragraph("<b>Condición del Terreno</b>", st_header_tabla),
        Paragraph("<b>Esquema</b>", st_header_tabla),
        Paragraph("<b>Valor (Ω) / Estado</b>", st_header_tabla),
        Paragraph("<b>Cumple (≤40Ω)</b>", st_header_tabla),
        Paragraph("<b>Circuito Continuo</b>", st_header_tabla),
        Paragraph("<b>Capacidad de Carga</b>", st_header_tabla),
        Paragraph("<b>Protección</b>", st_header_tabla),
    ]
    filas_grilla: List[List[Any]] = [encabezado_grilla]

    estilos_grilla: List[Any] = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1E293B")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#94A3B8")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 1), (0, -1), "CENTER"),
        ("ALIGN", (3, 1), (-1, -1), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, -1), 3.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
    ]

    mediciones_pat = ctx["mediciones_pat"]
    if not mediciones_pat:
        filas_grilla.append(
            [
                "-",
                Paragraph("Sin mediciones PAT registradas", st_celda),
                "-",
                "TT",
                "-",
                "-",
                "-",
                "-",
                "-",
            ]
        )
    else:
        for i, m in enumerate(mediciones_pat, start=1):
            nro = str(m.get("nro_toma", i))
            sec = str(m.get("sector", "") or "").strip()
            cond = str(m.get("condicion_terreno", "") or ctx["condicion_gral"]).strip()
            esq = str(m.get("esquema", "TT") or "TT").strip()
            val_raw = m.get("valor_ohms", "")
            es_num, val_f = es_medicion_pat_valida(val_raw)

            if es_num and val_f is not None:
                val_str = f"{val_f:.2f} Ω" if val_f % 1 else f"{int(val_f)} Ω"
                cont = str(m.get("continuidad", "SI") or "SI").strip()
                cap = str(m.get("capacidad_carga", "SI") or "SI").strip()
                prot = str(m.get("proteccion", "DD") or "DD").strip()
                cumple_def = "SI" if (val_f <= 40.0 and cont != "NO" and cap != "NO") else "NO"
                cumple = str(m.get("cumple", cumple_def) or cumple_def).strip()
                if val_f > 40.0 or val_f <= 0 or cumple == "NO" or cont == "NO" or cap == "NO":
                    estilos_grilla.append(
                        ("BACKGROUND", (4, i), (5, i), colors.HexColor("#FEE2E2"))
                    )
            else:
                val_str = str(val_raw or "-").strip()
                cumple = "OBSERVADO"
                cont = "-"
                cap = "-"
                prot = "-"
                estilos_grilla.append(
                    ("BACKGROUND", (4, i), (5, i), colors.HexColor("#FEF3C7"))
                )

            filas_grilla.append(
                [
                    nro,
                    Paragraph(sec, st_celda),
                    Paragraph(cond, st_celda),
                    esq,
                    val_str,
                    cumple,
                    cont,
                    cap,
                    prot,
                ]
            )

    t_grilla = Table(
        filas_grilla,
        colWidths=[
            15 * mm,
            68 * mm,
            38 * mm,
            18 * mm,
            34 * mm,
            26 * mm,
            24 * mm,
            24 * mm,
            22 * mm,
        ],
        repeatRows=1,
    )
    t_grilla.setStyle(TableStyle(estilos_grilla))
    story.append(t_grilla)
    story.append(Spacer(1, 2 * mm))
    story.append(
        Paragraph(
            "<i>Nota: Si la impedancia de lazo es menor a 40 Ω, la puesta a tierra es APTA. "
            "Los puntos indicados como INVERTIDO / SIN PE / NO TIENE requieren adecuación eléctrica.</i>",
            st_celda,
        )
    )

    # ==================== SECCIÓN 3: DIFERENCIALES (RCD) ====================
    mediciones_dif = ctx["mediciones_dif"]
    if mediciones_dif:
        enc_dif = [
            Paragraph("<b>Ubicación</b>", st_header_tabla),
            Paragraph("<b>Tipo / Marca</b>", st_header_tabla),
            Paragraph("<b>IΔn (mA)</b>", st_header_tabla),
            Paragraph("<b>Fase</b>", st_header_tabla),
            Paragraph("<b>Polaridad</b>", st_header_tabla),
            Paragraph("<b>ΔTx 0,5</b>", st_header_tabla),
            Paragraph("<b>ΔTx 1 (ms)</b>", st_header_tabla),
            Paragraph("<b>ΔTx 5 (ms)</b>", st_header_tabla),
            Paragraph("<b>ID (mA)</b>", st_header_tabla),
            Paragraph("<b>T (ms) / Estado</b>", st_header_tabla),
        ]
        col_widths_dif = [
            35 * mm,
            42 * mm,
            18 * mm,
            16 * mm,
            28 * mm,
            22 * mm,
            26 * mm,
            26 * mm,
            24 * mm,
            32 * mm,
        ]

        for idx_d, d in enumerate(mediciones_dif, start=1):
            filas_disp, alertas_disp = _filas_dispositivo_rcd(idx_d, d)
            tabla_data = [enc_dif] + filas_disp
            estilo_disp: List[Any] = [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1E293B")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#94A3B8")),
                ("ALIGN", (2, 1), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("FONTSIZE", (0, 1), (-1, -1), 8.2),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
            for r_local, c_local in alertas_disp:
                estilo_disp.append(
                    ("BACKGROUND", (c_local, r_local + 1), (c_local, r_local + 1), colors.HexColor("#FEE2E2"))
                )
            # Resaltar en verde suave el estado APTO si no hubo fallas
            if not alertas_disp and len(filas_disp) > 0:
                estilo_disp.append(
                    ("BACKGROUND", (9, 1), (9, 1), colors.HexColor("#DCFCE7"))
                )

            t_dif = Table(tabla_data, colWidths=col_widths_dif, repeatRows=1)
            t_dif.setStyle(TableStyle(estilo_disp))

            if idx_d == 1:
                story.append(
                    KeepTogether(
                        [
                            Spacer(1, 5 * mm),
                            Paragraph("3. Ensayo de Dispositivos Diferenciales (RCD)", st_subtitulo),
                            t_dif,
                        ]
                    )
                )
            else:
                story.append(KeepTogether([Spacer(1, 2.5 * mm), t_dif]))

        story.append(Spacer(1, 2 * mm))
        story.append(
            Paragraph(
                "<i>Referencias RCD Apto: ΔTx 0,5 = NO DISPARA (ND) | ΔTx 1 &lt; 300 ms | "
                "ΔTx 5 &lt; 40 ms | ID (mA): 0,5 x IΔn &lt; ID ≤ 1 x IΔn.</i>",
                st_celda,
            )
        )

    # ==================== SECCIÓN 4: CONCLUSIONES ====================
    t_conc = Table(
        [
            [
                Paragraph("<b>Conclusiones</b>", st_celda_bold),
                Paragraph(
                    "<b>Recomendaciones para la adecuación a la legislación vigente</b>",
                    st_celda_bold,
                ),
            ],
            [
                Paragraph(ctx["txt_conc"], st_celda),
                Paragraph(ctx["txt_rec"], st_celda),
            ],
        ],
        colWidths=[134 * mm, 135 * mm],
    )
    t_conc.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E2E8F0")),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#64748B")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    story.append(
        KeepTogether(
            [
                Spacer(1, 5 * mm),
                Paragraph("4. Análisis de los Datos, Conclusiones y Recomendaciones", st_subtitulo),
                t_conc,
            ]
        )
    )

    doc.build(story, canvasmaker=_NumberedCanvas)
    return str(Path(ruta_salida).resolve())


# ==============================================================================
# MOTOR 2: PDF 1.4 VECTORIAL NATIVO (ANDROID / SIN DEPENDENCIAS EXTERNAS)
# ==============================================================================

_MM = 72.0 / 25.4
_PAGE_W = 842.0  # A4 landscape width (pt)
_PAGE_H = 595.0  # A4 landscape height (pt)

# Anchos aproximados en milésimas de em para Helvetica estándar (WinAnsi)
_ANCHOS_HELVETICA: Dict[str, int] = {
    " ": 278, "!": 278, '"': 355, "#": 556, "$": 556, "%": 889, "&": 667, "'": 191,
    "(": 333, ")": 333, "*": 389, "+": 584, ",": 278, "-": 333, ".": 278, "/": 278,
    "0": 556, "1": 556, "2": 556, "3": 556, "4": 556, "5": 556, "6": 556, "7": 556,
    "8": 556, "9": 556, ":": 278, ";": 278, "<": 584, "=": 584, ">": 584, "?": 556,
    "@": 1015, "A": 667, "B": 667, "C": 722, "D": 722, "E": 667, "F": 611, "G": 778,
    "H": 722, "I": 278, "J": 500, "K": 667, "L": 556, "M": 833, "N": 722, "O": 778,
    "P": 667, "Q": 778, "R": 722, "S": 667, "T": 611, "U": 722, "V": 667, "W": 944,
    "X": 667, "Y": 667, "Z": 611, "[": 278, "\\": 278, "]": 278, "^": 469, "_": 556,
    "`": 333, "a": 556, "b": 556, "c": 500, "d": 556, "e": 556, "f": 278, "g": 556,
    "h": 556, "i": 222, "j": 222, "k": 500, "l": 222, "m": 833, "n": 556, "o": 556,
    "p": 556, "q": 556, "r": 333, "s": 500, "t": 278, "u": 556, "v": 500, "w": 722,
    "x": 500, "y": 500, "z": 500, "{": 334, "|": 260, "}": 334, "~": 584,
    "°": 400, "º": 365, "×": 584, "á": 556, "é": 556, "í": 278, "ó": 556, "ú": 556,
    "Á": 667, "É": 667, "Í": 278, "Ó": 778, "Ú": 722, "ñ": 556, "Ñ": 722, "ü": 556,
}


def _normalizar_texto_winansi(texto: str) -> str:
    """Convierte símbolos Unicode fuera de cp1252 a equivalentes legibles en Helvetica."""
    reemplazos = (
        ("Ω", "Ohm"),
        ("∆", "d"),
        ("Δ", "d"),
        ("≤", "<="),
        ("≥", ">="),
        ("✔", "OK"),
        ("✘", "NO"),
        ("—", "-"),
        ("–", "-"),
        ("•", "-"),
        ("·", "-"),
    )
    out = str(texto or "")
    for orig, rep in reemplazos:
        out = out.replace(orig, rep)
    return out


def _escapar_pdf_winansi(texto: str) -> str:
    limpio = _normalizar_texto_winansi(texto)
    bytes_win = limpio.encode("cp1252", errors="replace").decode("cp1252")
    return bytes_win.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")


def _medir_texto_pt(texto: str, font_size: float, bold: bool = False) -> float:
    limpio = _normalizar_texto_winansi(texto)
    factor = 1.05 if bold else 1.0
    total = sum(_ANCHOS_HELVETICA.get(ch, 556) for ch in limpio)
    return (total / 1000.0) * font_size * factor


def _envolver_texto(texto: str, max_w: float, font_size: float, bold: bool = False) -> List[str]:
    """Divide texto en líneas respetando saltos explícitos y ancho máximo disponible."""
    limpio = _normalizar_texto_winansi(texto)
    if not limpio:
        return [""]
    resultado: List[str] = []
    for parrafo in limpio.splitlines():
        if not parrafo.strip():
            resultado.append("")
            continue
        palabras = parrafo.split()
        actual = ""
        for pal in palabras:
            candidato = f"{actual} {pal}".strip() if actual else pal
            if _medir_texto_pt(candidato, font_size, bold) <= max_w:
                actual = candidato
            else:
                if actual:
                    resultado.append(actual)
                    actual = ""
                if _medir_texto_pt(pal, font_size, bold) <= max_w:
                    actual = pal
                else:
                    # Cortar palabra larga que no entra en la columna
                    trozo = ""
                    for ch in pal:
                        if _medir_texto_pt(trozo + ch, font_size, bold) <= max_w or not trozo:
                            trozo += ch
                        else:
                            resultado.append(trozo)
                            trozo = ch
                    actual = trozo
        if actual:
            resultado.append(actual)
    return resultado or [""]


class _NativePDFCanvas:
    """Constructor de documento PDF 1.4 multipágina con tablas vectoriales y Helvetica."""

    def __init__(self):
        self.margin_l = 14.0 * _MM
        self.margin_r = 14.0 * _MM
        self.margin_t = 12.0 * _MM
        self.margin_b = 24.0 * _MM
        self.content_w = _PAGE_W - self.margin_l - self.margin_r
        self.y_top = _PAGE_H - self.margin_t
        self.y_min = self.margin_b
        self.pages: List[List[str]] = []
        self.y = self.y_top
        self.new_page()

    def new_page(self) -> None:
        self.pages.append([])
        self.y = self.y_top

    def _emit(self, cmd: str) -> None:
        self.pages[-1].append(cmd)

    def ensure_space(self, needed_pt: float, on_break=None) -> bool:
        if self.y - needed_pt < self.y_min:
            self.new_page()
            if on_break is not None:
                on_break()
            return True
        return False

    def draw_rect(
        self,
        x: float,
        y_bottom: float,
        w: float,
        h: float,
        fill_rgb: Optional[Tuple[float, float, float]] = None,
        stroke_rgb: Optional[Tuple[float, float, float]] = None,
        line_w: float = 0.5,
    ) -> None:
        cmds = ["q"]
        if fill_rgb is not None:
            cmds.append(f"{fill_rgb[0]:.3f} {fill_rgb[1]:.3f} {fill_rgb[2]:.3f} rg")
        if stroke_rgb is not None:
            cmds.append(
                f"{line_w:.2f} w {stroke_rgb[0]:.3f} {stroke_rgb[1]:.3f} {stroke_rgb[2]:.3f} RG"
            )
        cmds.append(f"{x:.2f} {y_bottom:.2f} {w:.2f} {h:.2f} re")
        if fill_rgb is not None and stroke_rgb is not None:
            cmds.append("B")
        elif fill_rgb is not None:
            cmds.append("f")
        else:
            cmds.append("S")
        cmds.append("Q")
        self._emit(" ".join(cmds))

    def draw_line(
        self,
        x1: float,
        y1: float,
        x2: float,
        y2: float,
        stroke_rgb: Tuple[float, float, float] = (0.4, 0.4, 0.4),
        line_w: float = 0.6,
    ) -> None:
        self._emit(
            f"q {line_w:.2f} w {stroke_rgb[0]:.3f} {stroke_rgb[1]:.3f} {stroke_rgb[2]:.3f} RG "
            f"{x1:.2f} {y1:.2f} m {x2:.2f} {y2:.2f} l S Q"
        )

    def draw_text(
        self,
        x: float,
        y: float,
        text: str,
        bold: bool = False,
        size: float = 8.5,
        rgb: Tuple[float, float, float] = (0.06, 0.09, 0.16),
        align: str = "left",
        box_w: Optional[float] = None,
    ) -> None:
        font_id = "F2" if bold else "F1"
        txt_w = _medir_texto_pt(text, size, bold)
        tx = x
        if box_w is not None:
            if align == "center":
                tx = x + max(0.0, (box_w - txt_w) / 2.0)
            elif align == "right":
                tx = x + max(0.0, box_w - txt_w)
        esc = _escapar_pdf_winansi(text)
        self._emit(
            f"BT /{font_id} {size:.1f} Tf {rgb[0]:.3f} {rgb[1]:.3f} {rgb[2]:.3f} rg "
            f"1 0 0 1 {tx:.2f} {y:.2f} Tm ({esc}) Tj ET"
        )

    def draw_heading(self, texto: str, size: float = 10.0) -> None:
        self.ensure_space(size + 14.0)
        self.y -= 4.0
        self.draw_text(self.margin_l, self.y - size, texto, bold=True, size=size, rgb=(0.118, 0.161, 0.231))
        self.y -= size + 5.0

    def draw_note(self, texto: str, size: float = 8.0) -> None:
        lineas = _envolver_texto(texto, self.content_w, size, bold=False)
        line_h = size * 1.25
        self.ensure_space(len(lineas) * line_h + 4.0)
        for ln in lineas:
            self.draw_text(self.margin_l, self.y - size, ln, bold=False, size=size, rgb=(0.2, 0.25, 0.33))
            self.y -= line_h
        self.y -= 2.0

    def measure_row(
        self,
        cells: List[str],
        col_widths: List[float],
        bold_flags: List[bool],
        font_size: float = 8.0,
        pad_x: float = 4.0,
        pad_y: float = 3.5,
    ) -> Tuple[float, List[List[str]]]:
        line_h = font_size * 1.25
        wrapped: List[List[str]] = []
        max_lines = 1
        for txt, w, b in zip(cells, col_widths, bold_flags):
            lines = _envolver_texto(txt, max(10.0, w - 2 * pad_x), font_size, bold=b)
            wrapped.append(lines)
            if len(lines) > max_lines:
                max_lines = len(lines)
        row_h = max_lines * line_h + 2 * pad_y
        return row_h, wrapped

    def draw_row(
        self,
        cells: List[str],
        col_widths: List[float],
        aligns: List[str],
        bold_flags: List[bool],
        bg_colors: List[Optional[Tuple[float, float, float]]],
        text_colors: List[Tuple[float, float, float]],
        font_size: float = 8.0,
        pad_x: float = 4.0,
        pad_y: float = 3.5,
        border_rgb: Tuple[float, float, float] = (0.58, 0.64, 0.72),
        valign: str = "middle",
        header_repeat_cb=None,
    ) -> None:
        line_h = font_size * 1.25
        row_h, wrapped = self.measure_row(cells, col_widths, bold_flags, font_size, pad_x, pad_y)
        self.ensure_space(row_h, on_break=header_repeat_cb)

        y_bottom = self.y - row_h
        x = self.margin_l
        for idx, (lines, w) in enumerate(zip(wrapped, col_widths)):
            self.draw_rect(
                x,
                y_bottom,
                w,
                row_h,
                fill_rgb=bg_colors[idx],
                stroke_rgb=border_rgb,
                line_w=0.5,
            )
            block_h = len(lines) * line_h
            if valign == "top":
                start_y = self.y - pad_y - font_size
            else:
                start_y = y_bottom + (row_h + block_h) / 2.0 - font_size + (line_h - font_size) * 0.25

            for l_idx, ln in enumerate(lines):
                if ln:
                    self.draw_text(
                        x + pad_x,
                        start_y - l_idx * line_h,
                        ln,
                        bold=bold_flags[idx],
                        size=font_size,
                        rgb=text_colors[idx],
                        align=aligns[idx],
                        box_w=max(1.0, w - 2 * pad_x),
                    )
            x += w
        self.y = y_bottom

    def save(self, ruta_salida: str) -> str:
        total = len(self.pages)
        # Dibujar pie de página uniforme en todas las páginas
        for idx in range(total):
            p_cmds = self.pages[idx]
            x_firma = _PAGE_W - 95.0 * _MM
            y_firma = 16.0 * _MM
            p_cmds.append(
                f"q 0.6 w 0.4 0.4 0.4 RG {x_firma:.2f} {y_firma:.2f} m "
                f"{_PAGE_W - 20.0 * _MM:.2f} {y_firma:.2f} l S Q"
            )
            firma_txt = _escapar_pdf_winansi(
                "Firma, Aclaración y Registro del Profesional Interviniente"
            )
            firma_w = _medir_texto_pt(
                "Firma, Aclaración y Registro del Profesional Interviniente", 8.0, False
            )
            fx = (x_firma + (_PAGE_W - 20.0 * _MM)) / 2.0 - firma_w / 2.0
            p_cmds.append(
                f"BT /F1 8.0 Tf 0.1 0.1 0.1 rg 1 0 0 1 {fx:.2f} {12.0 * _MM:.2f} Tm ({firma_txt}) Tj ET"
            )
            pie_txt = _escapar_pdf_winansi(f"Protocolo SRT 900/15  |  Hoja {idx + 1}/{total}")
            p_cmds.append(
                f"BT /F1 8.0 Tf 0.1 0.1 0.1 rg 1 0 0 1 {15.0 * _MM:.2f} {12.0 * _MM:.2f} Tm ({pie_txt}) Tj ET"
            )

        # Ensamblar objetos PDF 1.4
        # 1: Catalog, 2: Pages, 3: Font F1 (Helvetica), 4: Font F2 (Helvetica-Bold)
        objetos: List[bytes] = []
        objetos.append(b"<< /Type /Catalog /Pages 2 0 R >>")
        kids_refs = " ".join(f"{5 + i * 2} 0 R" for i in range(total))
        objetos.append(f"<< /Type /Pages /Kids [{kids_refs}] /Count {total} >>".encode("ascii"))
        objetos.append(
            b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>"
        )
        objetos.append(
            b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>"
        )

        for idx, p_cmds in enumerate(self.pages):
            page_obj_num = 5 + idx * 2
            content_obj_num = page_obj_num + 1
            stream_data = "\n".join(p_cmds).encode("cp1252", errors="replace")
            objetos.append(
                (
                    f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {_PAGE_W:.0f} {_PAGE_H:.0f}] "
                    f"/Resources << /Font << /F1 3 0 R /F2 4 0 R >> >> "
                    f"/Contents {content_obj_num} 0 R >>"
                ).encode("ascii")
            )
            objetos.append(
                f"<< /Length {len(stream_data)} >>\nstream\n".encode("ascii")
                + stream_data
                + b"\nendstream"
            )

        out = bytearray(b"%PDF-1.4\n")
        offsets = [0]
        for i, obj in enumerate(objetos, start=1):
            offsets.append(len(out))
            out.extend(f"{i} 0 obj\n".encode("ascii"))
            out.extend(obj)
            out.extend(b"\nendobj\n")
        xref_pos = len(out)
        out.extend(f"xref\n0 {len(objetos) + 1}\n0000000000 65535 f \n".encode("ascii"))
        for off in offsets[1:]:
            out.extend(f"{off:010d} 00000 n \n".encode("ascii"))
        out.extend(
            f"trailer\n<< /Size {len(objetos) + 1} /Root 1 0 R >>\nstartxref\n{xref_pos}\n%%EOF\n".encode(
                "ascii"
            )
        )
        Path(ruta_salida).write_bytes(bytes(out))
        return str(Path(ruta_salida).resolve())


def _generar_pdf_nativo_fallback(ruta_salida: str, datos: Dict[str, Any]) -> str:
    """
    Generador PDF 1.4 vectorial en Python puro (cero dependencias externas),
    utilizado en Android cuando ReportLab no está disponible.
    """
    Path(ruta_salida).parent.mkdir(parents=True, exist_ok=True)
    ctx = _preparar_contexto_pdf(datos)
    pdf = _NativePDFCanvas()

    c_dark = (0.118, 0.161, 0.231)  # #1E293B
    c_head_bg = (0.886, 0.910, 0.941)  # #E2E8F0
    c_light_bg = (0.973, 0.980, 0.988)  # #F8FAFC
    c_border_dark = (0.200, 0.255, 0.333)  # #334155
    c_border_mid = (0.580, 0.639, 0.722)  # #94A3B8
    c_red_bg = (0.996, 0.886, 0.886)  # #FEE2E2
    c_amber_bg = (0.996, 0.953, 0.780)  # #FEF3C7
    c_green_bg = (0.863, 0.988, 0.906)  # #DCFCE7
    c_text = (0.059, 0.090, 0.165)  # #0F172A
    c_white = (1.0, 1.0, 1.0)

    # ---- ENCABEZADO GENERAL ----
    pdf.draw_row(
        ["PROTOCOLO DE MEDICIÓN DE PUESTA A TIERRA Y CONTINUIDAD DE LAS MASAS (SRT 900/15)"],
        [pdf.content_w],
        ["center"],
        [True],
        [c_head_bg],
        [c_text],
        font_size=11.0,
        pad_y=5.5,
        border_rgb=c_border_dark,
    )
    pdf.y -= 3.0 * _MM

    cli_w = [105 * _MM, 62 * _MM, 52 * _MM, 50 * _MM]
    pdf.draw_row(
        [
            f"Razón Social: {ctx['razon_social']}",
            f"C.U.I.T.: {ctx['cuit']}",
            f"N° Contrato: {ctx['contrato']}",
            f"Fecha: {ctx['fecha_med']}",
        ],
        cli_w,
        ["left", "left", "left", "left"],
        [False, False, False, False],
        [c_light_bg] * 4,
        [c_text] * 4,
        font_size=8.3,
        border_rgb=c_border_mid,
    )
    pdf.draw_row(
        [
            f"Dirección: {ctx['direccion']}",
            f"Localidad: {ctx['localidad']}",
            f"CP: {ctx['cp']}",
            f"Provincia: {ctx['provincia']}",
        ],
        cli_w,
        ["left", "left", "left", "left"],
        [False, False, False, False],
        [c_light_bg] * 4,
        [c_text] * 4,
        font_size=8.3,
        border_rgb=c_border_mid,
    )
    pdf.y -= 3.0 * _MM

    # ---- SECCIÓN 1: DATOS Y METODOLOGÍA ----
    pdf.draw_heading("1. Datos para la Medición y Metodología")
    w_s1 = [125 * _MM, 90 * _MM, 54 * _MM]
    pdf.draw_row(
        [
            f"(7) Marca, modelo y N° de serie del instrumento utilizado:\n{ctx['instrumento']}",
            f"(8) Fecha de Calibración del Instrumental:\n{ctx['fecha_cal']}",
            f"Plano N°:\n{ctx['plano']}",
        ],
        w_s1,
        ["left", "left", "left"],
        [False, False, False],
        [None, None, None],
        [c_text, c_text, c_text],
        font_size=8.3,
        valign="top",
    )
    pdf.draw_row(
        [
            f"(9) Fecha de la medición: {ctx['fecha_med']}",
            f"(10) Hora de inicio: {ctx['hora_ini']}",
            f"(11) Hora finalización: {ctx['hora_fin']}",
        ],
        w_s1,
        ["left", "left", "left"],
        [False, False, False],
        [None, None, None],
        [c_text, c_text, c_text],
        font_size=8.3,
    )
    pdf.draw_row(
        [f"(12) Metodología utilizada: {ctx['metodologia']}"],
        [pdf.content_w],
        ["left"],
        [False],
        [None],
        [c_text],
        font_size=8.3,
    )
    pdf.draw_row(
        [f"(13) Observaciones: {ctx['obs_texto']}"],
        [pdf.content_w],
        ["left"],
        [False],
        [None],
        [c_text],
        font_size=8.3,
        valign="top",
    )
    chk_cal = "[X]" if ctx["adj_cal"] else "[ ]"
    chk_pla = "[X]" if ctx["adj_plano"] else "[ ]"
    pdf.draw_row(
        [
            f"Documentación que se adjuntará a la medición:  {chk_cal} (14) Certificado de Calibración     {chk_pla} (15) Plano o croquis"
        ],
        [pdf.content_w],
        ["left"],
        [False],
        [None],
        [c_text],
        font_size=8.3,
    )
    pdf.y -= 3.5 * _MM

    # ---- SECCIÓN 2: GRILLA DE PUESTA A TIERRA ----
    pdf.draw_heading("2. Planilla de Medición de Puesta a Tierra y Continuidad de las Masas")
    w_pat = [
        15 * _MM,
        68 * _MM,
        38 * _MM,
        18 * _MM,
        34 * _MM,
        26 * _MM,
        24 * _MM,
        24 * _MM,
        22 * _MM,
    ]
    h_pat = [
        "N° Toma",
        "Sector / Ubicación",
        "Condición del Terreno",
        "Esquema",
        "Valor (Ω) / Estado",
        "Cumple (≤40Ω)",
        "Circuito Continuo",
        "Capacidad de Carga",
        "Protección",
    ]
    align_pat = ["center", "left", "left", "center", "center", "center", "center", "center", "center"]

    def _draw_pat_header():
        pdf.draw_row(
            h_pat,
            w_pat,
            ["center"] * 9,
            [True] * 9,
            [c_dark] * 9,
            [c_white] * 9,
            font_size=7.8,
            border_rgb=c_border_mid,
        )

    _draw_pat_header()
    mediciones_pat = ctx["mediciones_pat"]
    if not mediciones_pat:
        pdf.draw_row(
            ["-", "Sin mediciones PAT registradas", "-", "TT", "-", "-", "-", "-", "-"],
            w_pat,
            align_pat,
            [False] * 9,
            [None] * 9,
            [c_text] * 9,
            font_size=8.0,
            header_repeat_cb=_draw_pat_header,
        )
    else:
        for i, m in enumerate(mediciones_pat, start=1):
            nro = str(m.get("nro_toma", i))
            sec = str(m.get("sector", "") or "").strip()
            cond = str(m.get("condicion_terreno", "") or ctx["condicion_gral"]).strip()
            esq = str(m.get("esquema", "TT") or "TT").strip()
            val_raw = m.get("valor_ohms", "")
            es_num, val_f = es_medicion_pat_valida(val_raw)
            bgs: List[Optional[Tuple[float, float, float]]] = [None] * 9
            if es_num and val_f is not None:
                val_str = f"{val_f:.2f} Ω" if val_f % 1 else f"{int(val_f)} Ω"
                cont = str(m.get("continuidad", "SI") or "SI").strip()
                cap = str(m.get("capacidad_carga", "SI") or "SI").strip()
                prot = str(m.get("proteccion", "DD") or "DD").strip()
                cumple_def = "SI" if (val_f <= 40.0 and cont != "NO" and cap != "NO") else "NO"
                cumple = str(m.get("cumple", cumple_def) or cumple_def).strip()
                if val_f > 40.0 or val_f <= 0 or cumple == "NO" or cont == "NO" or cap == "NO":
                    bgs[4] = c_red_bg
                    bgs[5] = c_red_bg
            else:
                val_str = str(val_raw or "-").strip()
                cumple = "OBSERVADO"
                cont = "-"
                cap = "-"
                prot = "-"
                bgs[4] = c_amber_bg
                bgs[5] = c_amber_bg

            pdf.draw_row(
                [nro, sec, cond, esq, val_str, cumple, cont, cap, prot],
                w_pat,
                align_pat,
                [False] * 9,
                bgs,
                [c_text] * 9,
                font_size=8.0,
                header_repeat_cb=_draw_pat_header,
            )

    pdf.y -= 2.0 * _MM
    pdf.draw_note(
        "Nota: Si la impedancia de lazo es menor a 40 Ω, la puesta a tierra es APTA. "
        "Los puntos indicados como INVERTIDO / SIN PE / NO TIENE requieren adecuación eléctrica."
    )

    # ---- SECCIÓN 3: DIFERENCIALES (RCD) ----
    mediciones_dif = ctx["mediciones_dif"]
    if mediciones_dif:
        w_dif = [
            35 * _MM,
            42 * _MM,
            18 * _MM,
            16 * _MM,
            28 * _MM,
            22 * _MM,
            26 * _MM,
            26 * _MM,
            24 * _MM,
            32 * _MM,
        ]
        h_dif = [
            "Ubicación",
            "Tipo / Marca",
            "IΔn (mA)",
            "Fase",
            "Polaridad",
            "ΔTx 0,5",
            "ΔTx 1 (ms)",
            "ΔTx 5 (ms)",
            "ID (mA)",
            "T (ms) / Estado",
        ]
        align_dif = ["left", "left", "center", "center", "center", "center", "center", "center", "center", "center"]

        h_hdr, _ = pdf.measure_row(h_dif, w_dif, [True] * 10, font_size=7.8)

        for idx_d, d in enumerate(mediciones_dif, start=1):
            filas_disp, alertas_disp = _filas_dispositivo_rcd(idx_d, d)
            alertas_set = set(alertas_disp)

            bloque_h = h_hdr + 2.5 * _MM
            for row_cells in filas_disp:
                rh, _ = pdf.measure_row(row_cells, w_dif, [False] * 10, font_size=8.0)
                bloque_h += rh
            if idx_d == 1:
                bloque_h += 24.0  # espacio exacto para título de sección 3 (draw_heading)

            pdf.ensure_space(bloque_h)
            pdf.y -= 2.5 * _MM
            if idx_d == 1:
                pdf.draw_heading("3. Ensayo de Dispositivos Diferenciales (RCD)")

            pdf.draw_row(
                h_dif,
                w_dif,
                ["center"] * 10,
                [True] * 10,
                [c_dark] * 10,
                [c_white] * 10,
                font_size=7.8,
            )
            for r_local, row_cells in enumerate(filas_disp):
                bgs: List[Optional[Tuple[float, float, float]]] = [None] * 10
                for c_idx in range(10):
                    if (r_local, c_idx) in alertas_set:
                        bgs[c_idx] = c_red_bg
                if r_local == 0 and not alertas_set:
                    bgs[9] = c_green_bg
                pdf.draw_row(
                    row_cells,
                    w_dif,
                    align_dif,
                    [False] * 10,
                    bgs,
                    [c_text] * 10,
                    font_size=8.0,
                )

        pdf.y -= 2.0 * _MM
        pdf.draw_note(
            "Referencias RCD Apto: ΔTx 0,5 = NO DISPARA (ND) | ΔTx 1 < 300 ms | "
            "ΔTx 5 < 40 ms | ID (mA): 0,5 x IΔn < ID ≤ 1 x IΔn."
        )

    # ---- SECCIÓN 4: CONCLUSIONES ----
    w_conc = [134 * _MM, 135 * _MM]
    h_c_hdr, _ = pdf.measure_row(
        ["Conclusiones", "Recomendaciones para la adecuación a la legislación vigente"],
        w_conc,
        [True, True],
        font_size=8.5,
    )
    h_c_body, _ = pdf.measure_row(
        [ctx["txt_conc"], ctx["txt_rec"]],
        w_conc,
        [False, False],
        font_size=8.3,
    )
    pdf.ensure_space(h_c_hdr + h_c_body + 24.0)
    pdf.y -= 2.5 * _MM
    pdf.draw_heading("4. Análisis de los Datos, Conclusiones y Recomendaciones")
    pdf.draw_row(
        ["Conclusiones", "Recomendaciones para la adecuación a la legislación vigente"],
        w_conc,
        ["left", "left"],
        [True, True],
        [c_head_bg, c_head_bg],
        [c_text, c_text],
        font_size=8.5,
    )
    pdf.draw_row(
        [ctx["txt_conc"], ctx["txt_rec"]],
        w_conc,
        ["left", "left"],
        [False, False],
        [None, None],
        [c_text, c_text],
        font_size=8.3,
        valign="top",
    )

    return pdf.save(ruta_salida)


def generar_protocolo_pdf(
    ruta_salida: str, datos: Dict[str, Any], forzar_nativo: bool = False
) -> str:
    """
    Genera el protocolo SRT 900/15 en formato PDF listo para imprimir o enviar.
    Utiliza ReportLab si está disponible (salvo forzar_nativo=True) y cae en el
    motor PDF vectorial nativo en caso contrario (ej. APK Android).
    """
    logging.info(f"Generando protocolo PDF en: {ruta_salida}")
    if _REPORTLAB_DISPONIBLE and not forzar_nativo:
        return _generar_pdf_reportlab(ruta_salida, datos)
    return _generar_pdf_nativo_fallback(ruta_salida, datos)
