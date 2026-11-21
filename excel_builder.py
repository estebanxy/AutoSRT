import copy
import datetime
import logging
import os
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from srt_logic import (
    VALORES_ANOMALIA_PAT,
    evaluar_diferencial,
    generar_textos_sugeridos,
)

# Configuración básica de logging para informar el progreso o errores
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")


def ruta_plantilla_default(nombre_archivo: str = "Mediciones 2026.xlsx") -> str:
    """ Resuelve la ruta absoluta de la plantilla buscando junto al módulo o en el directorio actual. """
    base_dir = Path(__file__).resolve().parent
    candidato = base_dir / nombre_archivo
    if candidato.exists():
        return str(candidato)
    return nombre_archivo


def escribir_celda(ws: Worksheet, coord: str, valor: Any) -> None:
    """
    Escribe un valor en una celda de manera segura. Si la celda es parte de un rango
    combinado (MergedCell), encuentra automáticamente la celda principal (superior izquierda)
    de ese rango y escribe el valor allí para evitar el error 'MergedCell is read-only'.
    """
    celda = ws[coord]
    if type(celda).__name__ == "MergedCell":
        for merged_range in ws.merged_cells.ranges:
            if coord in merged_range:
                top_left_coord = str(merged_range).split(":")[0]
                ws[top_left_coord] = valor
                return
    else:
        ws[coord] = valor


def parsear_valor_numerico(val: Any) -> Union[float, int, str, None]:
    """
    Convierte strings numéricos ('0.88', '0,88', '30') a float/int para que Excel
    los trate como números reales, preservando textos como 'INVERTIDO', 'SIN PE' o 'ND'.
    """
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return val
    texto = str(val).strip()
    if not texto:
        return None
    texto_num = texto.replace(",", ".")
    try:
        num = float(texto_num)
        return int(num) if num.is_integer() and "." not in texto_num else num
    except ValueError:
        return texto


def es_medicion_pat_valida(valor: Any) -> Tuple[bool, Optional[float]]:
    """
    Determina si el valor de puesta a tierra es un valor numérico en Ohms (ej. 0.88)
    o una anomalía de instalación ('INVERTIDO', 'NO TIENE', 'SIN PE', etc.).
    Retorna (es_numerico, valor_float).
    """
    parsed = parsear_valor_numerico(valor)
    if isinstance(parsed, (int, float)):
        return True, float(parsed)
    return False, None


def copiar_estilo_celda(origen, destino) -> None:
    """Copia fuente, bordes, alineación y relleno de una celda a otra."""
    if origen.has_style:
        destino.font = copy.copy(origen.font)
        destino.border = copy.copy(origen.border)
        destino.fill = copy.copy(origen.fill)
        destino.number_format = origen.number_format
        destino.protection = copy.copy(origen.protection)
        destino.alignment = copy.copy(origen.alignment)


def llenar_hoja_datos(
    ws: Worksheet,
    datos_cliente: Dict[str, Any],
    conclusiones: Optional[Dict[str, Any]] = None,
) -> None:
    """
    Inyecta los datos generales del establecimiento e instrumental en la hoja 'DATOS ',
    respetando las etiquetas de las celdas combinadas que son referenciadas por fórmula
    (='DATOS '!A4:G4 y ='DATOS '!A6:G6) desde el resto de las hojas de la plantilla.
    """
    logging.info("Llenando hoja 'DATOS '...")
    conclusiones = conclusiones or {}

    razon_social = str(datos_cliente.get("razon_social", "") or "").strip()
    direccion = str(datos_cliente.get("direccion", "") or "").strip()
    localidad = str(datos_cliente.get("localidad", "") or "").strip()
    provincia = str(datos_cliente.get("provincia", "San Juan") or "San Juan").strip()
    cp = str(datos_cliente.get("cp", "") or "").strip()
    cuit = str(datos_cliente.get("cuit", "") or "").strip()
    contrato = str(
        datos_cliente.get("contrato", datos_cliente.get("nro_contrato", "")) or ""
    ).strip()
    plano = str(datos_cliente.get("plano", "") or "").strip()

    instrumento = str(
        conclusiones.get(
            "instrumento_utilizado", datos_cliente.get("instrumento", "")
        )
        or ""
    ).strip()
    fecha_cal = str(
        conclusiones.get(
            "fecha_calibracion", datos_cliente.get("fecha_calibracion", "")
        )
        or ""
    ).strip()
    fecha_med = str(
        conclusiones.get("fecha_medicion", datos_cliente.get("fecha_medicion", ""))
        or ""
    ).strip()
    hora_ini = str(
        conclusiones.get("hora_inicio", datos_cliente.get("hora_inicio", "")) or ""
    ).strip()
    hora_fin = str(
        conclusiones.get("hora_fin", datos_cliente.get("hora_fin", "")) or ""
    ).strip()
    metodologia = str(conclusiones.get("metodologia", "") or "").strip()
    obs_datos = str(conclusiones.get("observaciones_datos", "") or "").strip()

    # Celdas combinadas principales de DATOS
    # A4:G4 y A6:G6 son referenciadas desde GRILLA, DIF. TRI, DIFER. COMB., DIFERENCIALES y CONCLUSIONES
    escribir_celda(ws, "A4", f" Razón Social: {razon_social}")
    escribir_celda(ws, "A6", f" Dirección:  {direccion}")
    escribir_celda(ws, "A8", f"Localidad: {localidad}")
    escribir_celda(ws, "E8", f"Plano N° : {plano}")
    escribir_celda(ws, "A10", f"Provincia: {provincia}")
    escribir_celda(ws, "A12", f"CP: {cp}")
    escribir_celda(ws, "C12", f" C.U.I.T.: {cuit:<26} N.C. : {contrato}")

    # Datos de instrumento y horarios (filas 15 a 20)
    escribir_celda(
        ws,
        "A15",
        f"(7) Marca, modelo y número de serie del instrumento utilizado:\n    {instrumento}",
    )
    escribir_celda(
        ws,
        "A16",
        f"(8) Fecha de Calibración del Instrumental utilizado:  {fecha_cal}",
    )
    escribir_celda(ws, "A17", f"(9) Fecha de la medición: \n    {fecha_med}")
    escribir_celda(ws, "D17", f"(10) Hora de inicio:  \n    {hora_ini}")
    escribir_celda(ws, "F17", f"(11) Hora finalización: \n    {hora_fin}")

    if metodologia:
        escribir_celda(ws, "A18", f"(12) Metodologia utilizada\n    {metodologia}")

    if obs_datos:
        escribir_celda(ws, "A20", f"(13) Observaciones: {obs_datos}")


def actualizar_encabezado_grilla(
    ws: Worksheet, datos_cliente: Dict[str, Any]
) -> None:
    """Actualiza los datos del encabezado en hojas tipo GRILLA."""
    localidad = str(datos_cliente.get("localidad", "") or "Capital").strip()
    cp = str(datos_cliente.get("cp", "") or "5400").strip()
    provincia = str(datos_cliente.get("provincia", "") or "San Juan").strip()
    cuit = str(datos_cliente.get("cuit", "") or "").strip()
    contrato = str(
        datos_cliente.get("contrato", datos_cliente.get("nro_contrato", "")) or ""
    ).strip()

    escribir_celda(ws, "A3", "='DATOS '!A4")
    escribir_celda(ws, "I3", f" C.U.I.T.: {cuit:<45} N.C. : {contrato}")
    escribir_celda(ws, "A4", "='DATOS '!A6")
    escribir_celda(ws, "G4", f"Localidad: {localidad}")
    escribir_celda(ws, "K4", f"CP: {cp}")
    escribir_celda(ws, "L4", f"Provincia: {provincia}")


def llenar_hoja_grilla(
    ws: Worksheet,
    mediciones_pat: List[Dict[str, Any]],
    datos_cliente: Optional[Dict[str, Any]] = None,
    condicion_terreno_general: str = "",
) -> None:
    """
    Inyecta hasta 16 mediciones de Puesta a Tierra en una hoja 'GRILLA'
    (filas 9 a 24, respetando las columnas combinadas reales de la plantilla:
    A=Nro, B:D=Sector, E:F=Condición del terreno, G:I=Uso PAT, J=Esquema,
    K=Valor Ohms / Estado, L=Cumple, M=Continuidad, N=Capacidad carga,
    O=Protección (DD/IA/Fus), P:R=Desconexión automática).
    """
    logging.info(f"Llenando hoja '{ws.title}' con {len(mediciones_pat)} mediciones PAT...")
    if datos_cliente:
        actualizar_encabezado_grilla(ws, datos_cliente)

    fila_inicio = 9
    max_filas_tabla = 16  # Filas 9 a 24 inclusive en la plantilla Mediciones 2026.xlsx

    # Limpiar primero el bloque de 16 filas (9..24) y fila 25 extra
    for idx in range(max_filas_tabla + 1):
        r = fila_inicio + idx
        if idx < max_filas_tabla:
            escribir_celda(ws, f"A{r}", idx + 1)
        else:
            escribir_celda(ws, f"A{r}", None)
        for col in ["B", "E", "G", "K", "L", "M", "N", "O", "P"]:
            escribir_celda(ws, f"{col}{r}", None)
        escribir_celda(ws, f"J{r}", "TT")

    for i, medicion in enumerate(mediciones_pat[:max_filas_tabla]):
        fila = fila_inicio + i
        nro_toma = medicion.get("nro_toma", i + 1)
        sector = str(medicion.get("sector", "") or "").strip()
        # En la planilla real (ej. MEDICIONES_ORTOPEDIA_CUYO.xlsx), la condición del terreno
        # se anota en la primera fila (E9) de cada hoja de grilla.
        cond_terreno = (
            str(medicion.get("condicion_terreno", "") or condicion_terreno_general).strip()
            if i == 0
            else str(medicion.get("condicion_terreno_fila", "") or "").strip()
        )
        uso_pat = str(medicion.get("uso_pat", "") or "").strip()
        esquema = str(medicion.get("esquema", "TT") or "TT").strip()
        valor_raw = medicion.get("valor_ohms", "")

        es_num, valor_float = es_medicion_pat_valida(valor_raw)
        valor_celda = valor_float if es_num else str(valor_raw or "").strip()

        escribir_celda(ws, f"A{fila}", parsear_valor_numerico(nro_toma))
        escribir_celda(ws, f"B{fila}", sector)
        if cond_terreno:
            escribir_celda(ws, f"E{fila}", cond_terreno)
        if uso_pat:
            escribir_celda(ws, f"G{fila}", uso_pat)
        escribir_celda(ws, f"J{fila}", esquema)
        escribir_celda(ws, f"K{fila}", valor_celda)

        if es_num and valor_float is not None:
            # Si es una medición numérica válida (ej. 0.88 Ω), se completan L, M, N y O
            # igual que en MEDICIONES_ORTOPEDIA_CUYO.xlsx
            cumple_default = "SI" if valor_float <= 40.0 else "NO"
            cumple = str(medicion.get("cumple", cumple_default) or cumple_default).strip()
            continuidad = str(medicion.get("continuidad", "SI") or "SI").strip()
            capacidad = str(medicion.get("capacidad_carga", "SI") or "SI").strip()
            proteccion = str(medicion.get("proteccion", "DD") or "DD").strip()
            desconexion = str(medicion.get("desconexion_auto", "") or "").strip()

            escribir_celda(ws, f"L{fila}", cumple)
            escribir_celda(ws, f"M{fila}", continuidad)
            escribir_celda(ws, f"N{fila}", capacidad)
            escribir_celda(ws, f"O{fila}", proteccion)
            if desconexion:
                escribir_celda(ws, f"P{fila}", desconexion)


def actualizar_encabezado_diferenciales(
    ws: Worksheet, datos_cliente: Dict[str, Any], fecha_medicion: str = ""
) -> None:
    """Actualiza el encabezado de las hojas 'DIF. TRI', 'DIFER. COMB.' y 'DIFERENCIALES'."""
    cuit = str(datos_cliente.get("cuit", "") or "").strip()
    contrato = str(
        datos_cliente.get("contrato", datos_cliente.get("nro_contrato", "")) or ""
    ).strip()

    escribir_celda(ws, "A4", "='DATOS '!A4")
    escribir_celda(ws, "A5", "='DATOS '!A6")
    escribir_celda(
        ws,
        "A6",
        f"Fecha: {fecha_medicion:<12} C.U.I.T.: {cuit:<35} N.C.: {contrato}",
    )


def _llenar_bloque_trifasico(
    ws: Worksheet, fila_base: int, med: Dict[str, Any], nro_orden: int = 1
) -> None:
    """
    Llena un bloque de 6 filas (R 0°, R 180°, S 0°, S 180°, T 0°, T 180°) para un
    disyuntor diferencial tetrapolar/trifásico en 'DIF. TRI' (fila_base=10)
    o 'DIFER. COMB.' (fila_base=9).
    """
    ubicacion = str(med.get("sector", med.get("ubicacion", "TP")) or "TP").strip()
    marca = str(med.get("marca", "") or "").strip()
    in_amp = str(med.get("corriente_nominal_a", "") or "").strip()
    ian_ma = str(med.get("corriente_fuga_ma", "30") or "30").strip()

    escribir_celda(ws, f"A{fila_base}", ubicacion)
    escribir_celda(ws, f"B{fila_base}", nro_orden)
    escribir_celda(
        ws,
        f"C{fila_base}",
        f"MARCA: {marca}\n\nI: {in_amp}\n\nI∆n: {ian_ma}",
    )

    fases_data = med.get("fases", {})
    # Soporte compatible si se pasó un diccionario plano (sin 'fases' detalladas)
    tiempo_simple = med.get("tiempo_disparo_ms", "")
    fases_orden = ["R", "S", "T"]

    for idx_fase, letra_fase in enumerate(fases_orden):
        f_0 = fila_base + idx_fase * 2
        f_180 = f_0 + 1
        datos_f = fases_data.get(letra_fase, {})

        # Valores para 0° (+)
        d0 = datos_f.get("0", {})
        dt05_0 = d0.get("dt_05", med.get(f"{letra_fase}_0_dt05", "ND" if tiempo_simple else None))
        dt1_0 = d0.get("dt_1", med.get(f"{letra_fase}_0_dt1", tiempo_simple if idx_fase == 0 else None))
        dt5_0 = d0.get("dt_5", med.get(f"{letra_fase}_0_dt5", None))
        id_0 = d0.get("id_ma", med.get(f"{letra_fase}_0_id", None))
        t_0 = d0.get("t_ms", med.get(f"{letra_fase}_0_t", None))

        # Valores para 180° (-)
        d180 = datos_f.get("180", {})
        dt05_180 = d180.get("dt_05", med.get(f"{letra_fase}_180_dt05", None))
        dt1_180 = d180.get("dt_1", med.get(f"{letra_fase}_180_dt1", None))
        dt5_180 = d180.get("dt_5", med.get(f"{letra_fase}_180_dt5", None))
        id_180 = d180.get("id_ma", med.get(f"{letra_fase}_180_id", None))
        t_180 = d180.get("t_ms", med.get(f"{letra_fase}_180_t", None))

        escribir_celda(ws, f"F{f_0}", parsear_valor_numerico(dt05_0))
        escribir_celda(ws, f"G{f_0}", parsear_valor_numerico(dt1_0))
        escribir_celda(ws, f"H{f_0}", parsear_valor_numerico(dt5_0))
        escribir_celda(ws, f"I{f_0}", parsear_valor_numerico(id_0))
        escribir_celda(ws, f"J{f_0}", parsear_valor_numerico(t_0))

        escribir_celda(ws, f"F{f_180}", parsear_valor_numerico(dt05_180))
        escribir_celda(ws, f"G{f_180}", parsear_valor_numerico(dt1_180))
        escribir_celda(ws, f"H{f_180}", parsear_valor_numerico(dt5_180))
        escribir_celda(ws, f"I{f_180}", parsear_valor_numerico(id_180))
        escribir_celda(ws, f"J{f_180}", parsear_valor_numerico(t_180))


def _llenar_bloque_bipolar(
    ws: Worksheet, fila_base: int, med: Dict[str, Any], nro_orden: int = 1
) -> None:
    """
    Llena un bloque de 2 filas (Fase 0° y Fase 180°) para un disyuntor diferencial
    bipolar en 'DIFERENCIALES' (fila_base=9 o 11) o 'DIFER. COMB.' (fila_base=17 o 19).
    """
    f_0 = fila_base
    f_180 = fila_base + 1

    # Si estamos usando filas adicionales (11..12 en DIFERENCIALES o 19..20 en DIFER. COMB.),
    # copiamos el estilo de las 2 filas anteriores para mantener bordes y fuentes idénticos.
    if fila_base in (11, 19):
        for col_idx in range(1, 11):
            col_l = get_column_letter(col_idx)
            copiar_estilo_celda(ws[f"{col_l}{fila_base - 2}"], ws[f"{col_l}{f_0}"])
            copiar_estilo_celda(ws[f"{col_l}{fila_base - 1}"], ws[f"{col_l}{f_180}"])

    ubicacion = str(med.get("sector", med.get("ubicacion", "TG")) or "TG").strip()
    marca = str(med.get("marca", "") or "").strip()
    in_amp = str(med.get("corriente_nominal_a", "") or "").strip()
    ian_ma = str(med.get("corriente_fuga_ma", "30") or "30").strip()

    bipolar_txt = f"{in_amp}A / {ian_ma}mA" if in_amp else f"{ian_ma} mA"

    escribir_celda(ws, f"A{f_0}", ubicacion)
    escribir_celda(ws, f"B{f_0}", nro_orden)
    escribir_celda(ws, f"C{f_0}", bipolar_txt)
    escribir_celda(ws, f"D{f_0}", marca)
    escribir_celda(ws, f"E{f_0}", "Fase 0º")
    escribir_celda(ws, f"E{f_180}", "Fase180º")

    fases_data = med.get("fases", {})
    datos_f = fases_data.get("BIPOLAR", fases_data.get("R", {}))
    tiempo_simple = med.get("tiempo_disparo_ms", "")

    d0 = datos_f.get("0", {})
    dt05_0 = d0.get("dt_05", med.get("0_dt05", "ND" if tiempo_simple else None))
    dt1_0 = d0.get("dt_1", med.get("0_dt1", tiempo_simple if tiempo_simple else None))
    dt5_0 = d0.get("dt_5", med.get("0_dt5", None))
    id_0 = d0.get("id_ma", med.get("0_id", None))
    t_0 = d0.get("t_ms", med.get("0_t", None))

    d180 = datos_f.get("180", {})
    dt05_180 = d180.get("dt_05", med.get("180_dt05", None))
    dt1_180 = d180.get("dt_1", med.get("180_dt1", None))
    dt5_180 = d180.get("dt_5", med.get("180_dt5", None))
    id_180 = d180.get("id_ma", med.get("180_id", None))
    t_180 = d180.get("t_ms", med.get("180_t", None))

    escribir_celda(ws, f"F{f_0}", parsear_valor_numerico(dt05_0))
    escribir_celda(ws, f"G{f_0}", parsear_valor_numerico(dt1_0))
    escribir_celda(ws, f"H{f_0}", parsear_valor_numerico(dt5_0))
    escribir_celda(ws, f"I{f_0}", parsear_valor_numerico(id_0))
    escribir_celda(ws, f"J{f_0}", parsear_valor_numerico(t_0))

    escribir_celda(ws, f"F{f_180}", parsear_valor_numerico(dt05_180))
    escribir_celda(ws, f"G{f_180}", parsear_valor_numerico(dt1_180))
    escribir_celda(ws, f"H{f_180}", parsear_valor_numerico(dt5_180))
    escribir_celda(ws, f"I{f_180}", parsear_valor_numerico(id_180))
    escribir_celda(ws, f"J{f_180}", parsear_valor_numerico(t_180))


def llenar_hojas_diferenciales(
    ws: Worksheet,
    mediciones_dif: List[Dict[str, Any]],
    datos_cliente: Optional[Dict[str, Any]] = None,
    fecha_medicion: str = "",
) -> None:
    """
    Inyecta los ensayos de disyuntores diferenciales en la hoja indicada ('DIF. TRI',
    'DIFER. COMB.' o 'DIFERENCIALES') respetando la estructura real de cada tabla.
    """
    logging.info(f"Llenando hoja '{ws.title}'...")
    if datos_cliente:
        actualizar_encabezado_diferenciales(ws, datos_cliente, fecha_medicion)

    # Limpiar basura residual conocida en K18 de DIF. TRI en la plantilla original
    if ws["K18"].value == "A18:K39C18B18:KA18:K39":
        ws["K18"].value = None

    trifasicos = [
        m for m in mediciones_dif if str(m.get("tipo", "")).upper().startswith("TRI")
    ]
    bipolares = [
        m for m in mediciones_dif if not str(m.get("tipo", "")).upper().startswith("TRI")
    ]

    titulo = ws.title.upper()
    if "DIF. TRI" in titulo:
        objetivo = trifasicos if trifasicos else mediciones_dif
        if objetivo:
            _llenar_bloque_trifasico(ws, 10, objetivo[0], nro_orden=1)
    elif "COMB" in titulo:
        if trifasicos:
            _llenar_bloque_trifasico(ws, 9, trifasicos[0], nro_orden=1)
        if bipolares:
            _llenar_bloque_bipolar(ws, 17, bipolares[0], nro_orden=1)
            if len(bipolares) > 1:
                _llenar_bloque_bipolar(ws, 19, bipolares[1], nro_orden=2)
    else:
        # Hoja DIFERENCIALES (bipolares)
        objetivo = bipolares if bipolares else mediciones_dif
        if objetivo:
            _llenar_bloque_bipolar(ws, 9, objetivo[0], nro_orden=1)
            if len(objetivo) > 1:
                _llenar_bloque_bipolar(ws, 11, objetivo[1], nro_orden=2)


def llenar_hoja_conclusiones(
    ws: Worksheet,
    conclusiones: Dict[str, Any],
    datos_cliente: Optional[Dict[str, Any]] = None,
) -> None:
    """
    Inyecta los encabezados, conclusiones y recomendaciones en la hoja 'CONCLUSIONES'
    respetando las celdas reales ('A8:G8' para Conclusiones, 'H8:L8' para Recomendaciones).
    """
    logging.info("Llenando hoja 'CONCLUSIONES'...")
    datos_cliente = datos_cliente or {}

    localidad = str(datos_cliente.get("localidad", "") or "Capital").strip()
    cp = str(datos_cliente.get("cp", "") or "5400").strip()
    provincia = str(datos_cliente.get("provincia", "") or "San Juan").strip()
    cuit = str(datos_cliente.get("cuit", "") or "").strip()
    contrato = str(
        datos_cliente.get("contrato", datos_cliente.get("nro_contrato", "")) or ""
    ).strip()

    escribir_celda(ws, "A3", "='DATOS '!A4")
    escribir_celda(ws, "I3", f"C.U.I.T.: {cuit:<26} N.C. : {contrato}")
    escribir_celda(ws, "A4", "='DATOS '!A6")
    escribir_celda(ws, "H4", f"Loc.: {localidad}")
    escribir_celda(ws, "J4", f"CP: {cp}")
    escribir_celda(ws, "K4", f"Provincia: {provincia}")

    texto_conclusiones = str(conclusiones.get("conclusiones_texto", "") or "").strip()
    if texto_conclusiones:
        escribir_celda(ws, "A8", texto_conclusiones)

    texto_recomendaciones = str(conclusiones.get("recomendaciones_texto", "") or "").strip()
    if texto_recomendaciones:
        escribir_celda(ws, "H8", texto_recomendaciones)


def actualizar_numeracion_hojas(wb: Workbook) -> None:
    """Actualiza los pies de página 'Hoja X/Y' en todas las hojas del libro."""
    total = len(wb.sheetnames)
    for idx, nombre in enumerate(wb.sheetnames, start=1):
        ws = wb[nombre]
        etiqueta = f"  Hoja {idx}/{total}"
        if nombre == "DATOS ":
            escribir_celda(ws, "G26", etiqueta)
        elif nombre.startswith("GRILLA") or "GRILLA" in nombre:
            # En la plantilla original está en O29 (o O32 si tiene filas extras)
            celda_hoja = "O32" if ws["O32"].value is not None else "O29"
            escribir_celda(ws, celda_hoja, etiqueta)
        elif nombre.startswith("DIF. TRI"):
            escribir_celda(ws, "I25", etiqueta)
        elif nombre.startswith("DIFER. COMB"):
            escribir_celda(ws, "I28", etiqueta)
        elif nombre.startswith("DIFERENCIALES"):
            escribir_celda(ws, "I20", etiqueta)
        elif nombre == "CONCLUSIONES":
            escribir_celda(ws, "L9", etiqueta)


def generar_protocolo(
    ruta_plantilla: str,
    ruta_salida: str,
    datos: Dict[str, Any],
    eliminar_hojas_no_usadas: bool = False,
) -> str:
    """
    Función principal que orquesta la lectura, llenado (con paginación automática
    de hojas GRILLA si hay más de 16 tomas PAT) y guardado del protocolo Excel.
    Retorna la ruta absoluta del archivo generado.
    """
    ruta_plantilla_real = ruta_plantilla
    if not os.path.exists(ruta_plantilla_real):
        ruta_plantilla_real = ruta_plantilla_default(os.path.basename(ruta_plantilla))
    if not os.path.exists(ruta_plantilla_real):
        raise FileNotFoundError(
            f"No se encontró el archivo de plantilla en la ruta: {ruta_plantilla}"
        )

    logging.info(f"Abriendo plantilla: {ruta_plantilla_real}")
    wb: Workbook = openpyxl.load_workbook(ruta_plantilla_real)

    datos_cliente = datos.get("datos_cliente", {})
    mediciones_pat = datos.get("mediciones_pat", [])
    mediciones_dif = datos.get("mediciones_diferenciales", [])
    conclusiones = dict(datos.get("conclusiones", {}) or {})

    sugeridos = generar_textos_sugeridos(datos)
    if not str(conclusiones.get("observaciones_datos", "") or "").strip():
        conclusiones["observaciones_datos"] = sugeridos["observaciones_datos"]
    if not str(
        conclusiones.get("conclusiones_texto", conclusiones.get("observaciones", ""))
        or ""
    ).strip():
        conclusiones["conclusiones_texto"] = sugeridos["conclusiones_texto"]
    if not str(conclusiones.get("recomendaciones_texto", "") or "").strip():
        conclusiones["recomendaciones_texto"] = sugeridos["recomendaciones_texto"]

    fecha_medicion = str(
        conclusiones.get(
            "fecha_medicion",
            datos_cliente.get(
                "fecha_medicion", datetime.date.today().strftime("%d/%m/%Y")
            ),
        )
        or ""
    ).strip()
    condicion_terreno_gral = str(
        datos_cliente.get("condicion_terreno", "ARENOSO SECO") or ""
    ).strip()

    # 1. Hoja DATOS
    if "DATOS " in wb.sheetnames:
        llenar_hoja_datos(wb["DATOS "], datos_cliente, conclusiones)
    else:
        logging.warning("No se encontró la hoja 'DATOS ' en la plantilla.")

    # 2. Hojas de GRILLA (Puesta a Tierra) - 16 mediciones por hoja
    if "GRILLA" in wb.sheetnames:
        ws_grilla_base = wb["GRILLA"]
        tam_pagina = 16
        bloques_pat = [
            mediciones_pat[i : i + tam_pagina]
            for i in range(0, max(1, len(mediciones_pat)), tam_pagina)
        ]

        # Si hay más de 1 bloque de 16 mediciones, clonamos la hoja GRILLA limpia
        # antes de modificar la primera, e insertamos las copias justo después de GRILLA.
        hojas_grilla_creadas = [ws_grilla_base]
        for idx_bloque in range(1, len(bloques_pat)):
            ws_copia = wb.copy_worksheet(ws_grilla_base)
            ws_copia.title = f"GRILLA ({idx_bloque + 1})"
            hojas_grilla_creadas.append(ws_copia)

        # Reordenar las hojas para que GRILLA (2), GRILLA (3), etc. queden junto a GRILLA
        if len(hojas_grilla_creadas) > 1:
            idx_grilla = wb._sheets.index(ws_grilla_base)
            for offset, ws_c in enumerate(hojas_grilla_creadas[1:], start=1):
                wb._sheets.remove(ws_c)
                wb._sheets.insert(idx_grilla + offset, ws_c)

        for idx_bloque, ws_g in enumerate(hojas_grilla_creadas):
            llenar_hoja_grilla(
                ws_g,
                bloques_pat[idx_bloque],
                datos_cliente=datos_cliente,
                condicion_terreno_general=condicion_terreno_gral,
            )

    # 3. Hojas de DIFERENCIALES
    trifasicos = [
        m for m in mediciones_dif if str(m.get("tipo", "")).upper().startswith("TRI")
    ]
    bipolares = [
        m for m in mediciones_dif if not str(m.get("tipo", "")).upper().startswith("TRI")
    ]

    # Actualizar encabezados en todas las hojas de diferenciales
    for nombre_hoja in ["DIF. TRI", "DIFER. COMB.", "DIFERENCIALES"]:
        if nombre_hoja in wb.sheetnames:
            actualizar_encabezado_diferenciales(
                wb[nombre_hoja], datos_cliente, fecha_medicion
            )
            if wb[nombre_hoja]["K18"].value == "A18:K39C18B18:KA18:K39":
                wb[nombre_hoja]["K18"].value = None

    if trifasicos and "DIF. TRI" in wb.sheetnames:
        ws_tri_base = wb["DIF. TRI"]
        hojas_tri = [ws_tri_base]
        for idx_t in range(1, len(trifasicos)):
            ws_copia = wb.copy_worksheet(ws_tri_base)
            ws_copia.title = f"DIF. TRI ({idx_t + 1})"
            hojas_tri.append(ws_copia)
        if len(hojas_tri) > 1:
            idx_pos = wb._sheets.index(ws_tri_base)
            for offset, ws_c in enumerate(hojas_tri[1:], start=1):
                wb._sheets.remove(ws_c)
                wb._sheets.insert(idx_pos + offset, ws_c)
        for idx_t, ws_t in enumerate(hojas_tri):
            llenar_hojas_diferenciales(
                ws_t, [trifasicos[idx_t]], datos_cliente, fecha_medicion
            )

    if bipolares and "DIFERENCIALES" in wb.sheetnames:
        ws_bip_base = wb["DIFERENCIALES"]
        bloques_bip = [bipolares[i : i + 2] for i in range(0, len(bipolares), 2)]
        hojas_bip = [ws_bip_base]
        for idx_b in range(1, len(bloques_bip)):
            ws_copia = wb.copy_worksheet(ws_bip_base)
            ws_copia.title = f"DIFERENCIALES ({idx_b + 1})"
            hojas_bip.append(ws_copia)
        if len(hojas_bip) > 1:
            idx_pos = wb._sheets.index(ws_bip_base)
            for offset, ws_c in enumerate(hojas_bip[1:], start=1):
                wb._sheets.remove(ws_c)
                wb._sheets.insert(idx_pos + offset, ws_c)
        for idx_b, ws_b in enumerate(hojas_bip):
            llenar_hojas_diferenciales(
                ws_b, bloques_bip[idx_b], datos_cliente, fecha_medicion
            )

    if trifasicos and bipolares and "DIFER. COMB." in wb.sheetnames:
        llenar_hojas_diferenciales(
            wb["DIFER. COMB."], mediciones_dif, datos_cliente, fecha_medicion
        )

    if eliminar_hojas_no_usadas:
        if not trifasicos and "DIF. TRI" in wb.sheetnames and len(wb.sheetnames) > 1:
            del wb["DIF. TRI"]
        if not bipolares and "DIFERENCIALES" in wb.sheetnames and len(wb.sheetnames) > 1:
            del wb["DIFERENCIALES"]
        if (
            not (trifasicos and bipolares)
            and "DIFER. COMB." in wb.sheetnames
            and len(wb.sheetnames) > 1
        ):
            del wb["DIFER. COMB."]

    # 4. Hoja CONCLUSIONES
    if "CONCLUSIONES" in wb.sheetnames:
        llenar_hoja_conclusiones(wb["CONCLUSIONES"], conclusiones, datos_cliente)

    # 5. Actualizar numeración de páginas en el pie de cada hoja
    actualizar_numeracion_hojas(wb)

    # Guardar archivo nuevo
    Path(ruta_salida).parent.mkdir(parents=True, exist_ok=True)
    logging.info(f"Guardando el protocolo generado en: {ruta_salida}")
    wb.save(ruta_salida)
    logging.info("¡Proceso completado con éxito!")
    return str(Path(ruta_salida).resolve())


def cargar_protocolo_excel(ruta_excel: str) -> Dict[str, Any]:
    """
    Lee un archivo Excel existente con formato de plantilla (como MEDICIONES_ORTOPEDIA_CUYO.xlsx)
    y extrae todos sus datos (cliente, mediciones PAT de todas las hojas GRILLA,
    diferenciales trifásicos/bipolares y conclusiones) para poder editarlos o exportarlos a PDF.
    """
    wb = openpyxl.load_workbook(ruta_excel, data_only=True)

    datos_cliente = {
        "razon_social": "",
        "cuit": "",
        "direccion": "",
        "localidad": "Capital",
        "provincia": "San Juan",
        "cp": "5400",
        "contrato": "",
        "plano": "",
        "condicion_terreno": "ARENOSO SECO",
    }
    conclusiones = {
        "observaciones": "",
        "observaciones_datos": "",
        "conclusiones_texto": "",
        "recomendaciones_texto": "",
        "metodologia": "",
        "fecha_medicion": datetime.date.today().strftime("%d/%m/%Y"),
        "instrumento_utilizado": "",
        "fecha_calibracion": "",
        "hora_inicio": "",
        "hora_fin": "",
    }

    if "DATOS " in wb.sheetnames:
        ws_d = wb["DATOS "]
        a4 = str(ws_d["A4"].value or "")
        if ":" in a4:
            datos_cliente["razon_social"] = a4.split(":", 1)[1].strip()
        a6 = str(ws_d["A6"].value or "")
        if ":" in a6:
            datos_cliente["direccion"] = a6.split(":", 1)[1].strip()
        a8 = str(ws_d["A8"].value or "")
        if ":" in a8:
            datos_cliente["localidad"] = a8.split(":", 1)[1].strip() or "Capital"
        e8 = str(ws_d["E8"].value or "")
        if ":" in e8:
            datos_cliente["plano"] = e8.split(":", 1)[1].strip()
        a10 = str(ws_d["A10"].value or "")
        if ":" in a10:
            datos_cliente["provincia"] = a10.split(":", 1)[1].strip() or "San Juan"
        a12 = str(ws_d["A12"].value or "")
        if ":" in a12:
            datos_cliente["cp"] = a12.split(":", 1)[1].strip() or "5400"
        c12 = str(ws_d["C12"].value or "")
        if "C.U.I.T.:" in c12:
            partes = c12.split("C.U.I.T.:", 1)[1]
            if "N.C." in partes:
                cuit_p, nc_p = partes.split("N.C.", 1)
                datos_cliente["cuit"] = cuit_p.strip()
                datos_cliente["contrato"] = nc_p.replace(":", "").strip()

        a15 = str(ws_d["A15"].value or "")
        if ":" in a15:
            conclusiones["instrumento_utilizado"] = a15.split(":", 1)[1].strip()
        a16 = str(ws_d["A16"].value or "")
        if ":" in a16:
            conclusiones["fecha_calibracion"] = a16.split(":", 1)[1].strip()
        a17 = str(ws_d["A17"].value or "")
        if ":" in a17:
            f_med = a17.split(":", 1)[1].strip()
            if f_med:
                conclusiones["fecha_medicion"] = f_med
        d17 = str(ws_d["D17"].value or "")
        if ":" in d17:
            conclusiones["hora_inicio"] = d17.split(":", 1)[1].strip()
        f17 = str(ws_d["F17"].value or "")
        if ":" in f17:
            conclusiones["hora_fin"] = f17.split(":", 1)[1].strip()
        a18 = str(ws_d["A18"].value or "")
        if "\n" in a18:
            conclusiones["metodologia"] = a18.split("\n", 1)[1].strip()
        a20 = str(ws_d["A20"].value or "")
        if "Observaciones:" in a20:
            obs_raw = a20.split("Observaciones:", 1)[1]
            conclusiones["observaciones_datos"] = re.sub(r" {2,}", " ", obs_raw).strip()

    # Leer todas las hojas de GRILLA
    mediciones_pat: List[Dict[str, Any]] = []
    nro_global = 1
    for sname in wb.sheetnames:
        if "GRILLA" in sname.upper():
            ws_g = wb[sname]
            cond_hoja = str(ws_g["E9"].value or "").strip()
            if cond_hoja:
                datos_cliente["condicion_terreno"] = cond_hoja
            for r in range(9, 26):
                sector = str(ws_g[f"B{r}"].value or "").strip()
                val_k = ws_g[f"K{r}"].value
                if not sector and val_k is None:
                    continue
                cond_r = str(ws_g[f"E{r}"].value or cond_hoja or "ARENOSO SECO").strip()
                esq_r = str(ws_g[f"J{r}"].value or "TT").strip()
                cumple_r = str(ws_g[f"L{r}"].value or "").strip()
                es_num, v_float = es_medicion_pat_valida(val_k)
                if not cumple_r:
                    cumple_r = "SI" if (es_num and v_float is not None and v_float <= 40) else "NO"
                mediciones_pat.append(
                    {
                        "nro_toma": nro_global,
                        "sector": sector or f"Toma {nro_global}",
                        "condicion_terreno": cond_r,
                        "uso_pat": str(ws_g[f"G{r}"].value or "").strip(),
                        "esquema": esq_r,
                        "valor_ohms": val_k if val_k is not None else "",
                        "cumple": cumple_r,
                        "continuidad": str(ws_g[f"M{r}"].value or ("SI" if es_num else "")).strip(),
                        "capacidad_carga": str(ws_g[f"N{r}"].value or ("SI" if es_num else "")).strip(),
                        "proteccion": str(ws_g[f"O{r}"].value or ("DD" if es_num else "")).strip(),
                    }
                )
                nro_global += 1

    # Leer Diferenciales Trifásicos si tienen datos cargados
    mediciones_dif: List[Dict[str, Any]] = []
    for sname in wb.sheetnames:
        if "DIF. TRI" in sname.upper():
            ws_t = wb[sname]
            # Verificar si hay al menos un valor medido en F10..J15
            tiene_datos = any(
                ws_t[f"{col}{r}"].value is not None
                for r in range(10, 16)
                for col in ["F", "G", "H", "I", "J"]
            )
            if tiene_datos:
                ubic = str(ws_t["A10"].value or "TP").strip()
                caract = str(ws_t["C10"].value or "")
                marca, in_a, ian_ma = "", "", "30"
                for linea in caract.splitlines():
                    if "MARCA:" in linea.upper():
                        marca = linea.split(":", 1)[1].strip()
                    elif "I∆N:" in linea.upper() or "IΔN:" in linea.upper():
                        ian_ma = linea.split(":", 1)[1].strip() or "30"
                    elif "I:" in linea.upper():
                        in_a = linea.split(":", 1)[1].strip()

                fases = {}
                for idx_f, fase_l in enumerate(["R", "S", "T"]):
                    r0 = 10 + idx_f * 2
                    r180 = r0 + 1
                    fases[fase_l] = {
                        "0": {
                            "dt_05": ws_t[f"F{r0}"].value,
                            "dt_1": ws_t[f"G{r0}"].value,
                            "dt_5": ws_t[f"H{r0}"].value,
                            "id_ma": ws_t[f"I{r0}"].value,
                            "t_ms": ws_t[f"J{r0}"].value,
                        },
                        "180": {
                            "dt_05": ws_t[f"F{r180}"].value,
                            "dt_1": ws_t[f"G{r180}"].value,
                            "dt_5": ws_t[f"H{r180}"].value,
                            "id_ma": ws_t[f"I{r180}"].value,
                            "t_ms": ws_t[f"J{r180}"].value,
                        },
                    }
                t_ref = ws_t["G10"].value or ""
                item_tri = {
                    "tipo": "TRIFASICO",
                    "sector": ubic,
                    "marca": marca,
                    "corriente_nominal_a": in_a,
                    "corriente_fuga_ma": ian_ma,
                    "tiempo_disparo_ms": t_ref,
                    "fases": fases,
                }
                item_tri["cumple"] = evaluar_diferencial(item_tri)["cumple"]
                mediciones_dif.append(item_tri)

        elif "DIFERENCIALES" in sname.upper():
            ws_b = wb[sname]
            for fila_base in (9, 11):
                r0 = fila_base
                r180 = fila_base + 1
                tiene_datos = any(
                    ws_b[f"{col}{r}"].value is not None
                    for r in (r0, r180)
                    for col in ["F", "G", "H", "I", "J"]
                )
                if not tiene_datos:
                    continue
                ubic = str(ws_b[f"A{r0}"].value or "TG").strip()
                c_raw = str(ws_b[f"C{r0}"].value or "").strip()
                in_a, ian_ma = "", "30"
                if "/" in c_raw:
                    izq, der = c_raw.split("/", 1)
                    in_a = re.sub(r"[^0-9.,]", "", izq).strip()
                    ian_ma = re.sub(r"[^0-9.,]", "", der).strip() or "30"
                elif c_raw:
                    ian_ma = re.sub(r"[^0-9.,]", "", c_raw).strip() or "30"
                marca = str(ws_b[f"D{r0}"].value or "").strip()
                fases = {
                    "BIPOLAR": {
                        "0": {
                            "dt_05": ws_b[f"F{r0}"].value,
                            "dt_1": ws_b[f"G{r0}"].value,
                            "dt_5": ws_b[f"H{r0}"].value,
                            "id_ma": ws_b[f"I{r0}"].value,
                            "t_ms": ws_b[f"J{r0}"].value,
                        },
                        "180": {
                            "dt_05": ws_b[f"F{r180}"].value,
                            "dt_1": ws_b[f"G{r180}"].value,
                            "dt_5": ws_b[f"H{r180}"].value,
                            "id_ma": ws_b[f"I{r180}"].value,
                            "t_ms": ws_b[f"J{r180}"].value,
                        },
                    }
                }
                t_ref = ws_b[f"G{r0}"].value or ""
                item_bip = {
                    "tipo": "BIPOLAR",
                    "sector": ubic,
                    "marca": marca,
                    "corriente_nominal_a": in_a,
                    "corriente_fuga_ma": ian_ma,
                    "tiempo_disparo_ms": t_ref,
                    "fases": fases,
                }
                item_bip["cumple"] = evaluar_diferencial(item_bip)["cumple"]
                mediciones_dif.append(item_bip)

    if "CONCLUSIONES" in wb.sheetnames:
        ws_c = wb["CONCLUSIONES"]
        conc_txt = str(ws_c["A8"].value or "").strip()
        rec_txt = str(ws_c["H8"].value or "").strip()
        conclusiones["conclusiones_texto"] = conc_txt
        conclusiones["observaciones"] = conc_txt
        conclusiones["recomendaciones_texto"] = rec_txt

    return {
        "datos_cliente": datos_cliente,
        "mediciones_pat": mediciones_pat,
        "mediciones_diferenciales": mediciones_dif,
        "conclusiones": conclusiones,
    }
