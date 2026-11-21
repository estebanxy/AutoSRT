import datetime
import json
import math
import os
import tempfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

PROVINCIAS_ARGENTINA: List[str] = [
    "San Juan",
    "Buenos Aires",
    "CABA",
    "Catamarca",
    "Chaco",
    "Chubut",
    "Córdoba",
    "Corrientes",
    "Entre Ríos",
    "Formosa",
    "Jujuy",
    "La Pampa",
    "La Rioja",
    "Mendoza",
    "Misiones",
    "Neuquén",
    "Río Negro",
    "Salta",
    "San Luis",
    "Santa Cruz",
    "Santa Fe",
    "Santiago del Estero",
    "Tierra del Fuego",
    "Tucumán",
]

CONDICIONES_TERRENO: List[str] = [
    "ARENOSO SECO",
    "ARENOSO HÚMEDO",
    "LECHO SECO",
    "ARCILLOSO",
    "PANTANOSO",
    "LLUVIAS RECIENTES",
    "NORMAL",
    "OTRO",
]

ESQUEMAS_CONEXION: List[str] = ["TT", "TN-S", "TN-C", "TN-C-S", "IT"]

PROTECCIONES_PAT: List[str] = ["DD", "IA", "Fus"]

USOS_PAT: List[str] = [
    "",
    "Toma de Tierra de Seguridad de las Masas",
    "Toma de Tierra del Neutro de Transformador",
    "De Protección de Equipos Electrónicos",
    "De Informática",
    "De Iluminación",
    "De Pararrayos",
    "Otros",
]

SECTORES_SUGERIDOS: List[str] = [
    "Oficina",
    "Cocina",
    "Baño",
    "Depósito",
    "Recepción",
    "Salón",
    "Tablero",
    "Lavadero",
    "Taller",
    "Pasillo",
]

VALORES_ANOMALIA_PAT = {
    "INVERTIDO",
    "NO TIENE",
    "NO TIEN",
    "SIN PE",
    "INVERTIDO/SIN PE",
    "-",
    "N/A",
}

_FORMATOS_FECHA = ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d")


def parsear_fecha(fecha_str: Any) -> Optional[datetime.date]:
    """Parsea fechas en formato DD/MM/AAAA, DD/MM/AA o AAAA-MM-DD."""
    texto = str(fecha_str or "").strip()
    if not texto:
        return None
    for fmt in _FORMATOS_FECHA:
        try:
            return datetime.datetime.strptime(texto, fmt).date()
        except ValueError:
            continue
    return None


def parsear_float(val: Any) -> Optional[float]:
    """Convierte strings numéricos ('0,88', '0.88', 30) a float; None si no es número válido."""
    if val is None or isinstance(val, bool):
        return None
    if isinstance(val, (int, float)):
        fval = float(val)
        return fval if math.isfinite(fval) else None
    texto = str(val).strip().replace(",", ".")
    if not texto:
        return None
    try:
        fval = float(texto)
        return fval if math.isfinite(fval) else None
    except ValueError:
        return None


def evaluar_valor_pat(valor_str: Any) -> Dict[str, Any]:
    """Evalúa una lectura de puesta a tierra (en Ohms o anomalía) según SRT 900/15 (≤ 40 Ω)."""
    texto = str(valor_str or "").strip()
    if not texto:
        return {"estado": "vacio", "mensaje": ""}

    val = parsear_float(texto)
    if val is not None:
        if val <= 0:
            return {
                "estado": "invalido",
                "mensaje": "Debe ser mayor a 0 Ω",
                "valor": val,
            }
        if val <= 40.0:
            return {
                "estado": "apto",
                "mensaje": f"{val:g} Ω · APTO (≤ 40 Ω)",
                "valor": val,
            }
        return {
            "estado": "no_cumple",
            "mensaje": f"{val:g} Ω · NO CUMPLE (> 40 Ω)",
            "valor": val,
        }

    return {
        "estado": "anomalia",
        "mensaje": f"Anomalía registrada: {texto}",
        "valor": texto,
    }


def evaluar_parametro_rcd(
    tipo_param: str, valor: Any, ian_ma: float = 30.0
) -> Tuple[Optional[bool], str]:
    """
    Evalúa un parámetro individual de ensayo RCD según RAEA / IEC 61008 / SRT 900/15.
    Retorna (cumple: Optional[bool], mensaje: str), donde None indica dato vacío sin evaluar.
    """
    ian = ian_ma if (ian_ma and ian_ma > 0) else 30.0
    texto = "" if valor is None else str(valor).strip()
    upper = texto.upper()

    if tipo_param == "dt_05":
        if upper in ("", "ND", "NO DISPARA", "-"):
            return True, "OK (No dispara)"
        if upper in ("DISPARA", "DISPARO", "DISPARÓ", "SI"):
            return False, "Disparó a 0,5×IΔn"
        num = parsear_float(texto)
        if num is not None:
            if num == 0:
                return False, "Valor inválido en ΔTx0.5"
            if num > 0:
                return False, f"Disparó a 0,5×IΔn ({num:g} ms)"
            return False, "Valor inválido en ΔTx0.5"
        return False, f"Valor inválido en ΔTx0.5 ({texto})"

    if upper in ("", "-"):
        return None, ""

    if tipo_param == "dt_1":
        if upper in ("ND", "NO DISPARA"):
            return False, "No disparó a 1×IΔn"
        num = parsear_float(texto)
        if num is None:
            return False, f"ΔTx1 inválido ({texto})"
        if 0 < num < 300.0:
            return True, "OK (<300 ms)"
        return False, f"ΔTx1={num:g} ms (≥300 ms)" if num >= 300.0 else f"ΔTx1={num:g} ms (inválido)"

    if tipo_param == "dt_5":
        if upper in ("ND", "NO DISPARA"):
            return False, "No disparó a 5×IΔn"
        num = parsear_float(texto)
        if num is None:
            return False, f"ΔTx5 inválido ({texto})"
        if 0 < num < 40.0:
            return True, "OK (<40 ms)"
        return False, f"ΔTx5={num:g} ms (≥40 ms)" if num >= 40.0 else f"ΔTx5={num:g} ms (inválido)"

    if tipo_param == "id_ma":
        if upper in ("ND", "NO DISPARA"):
            return False, "No disparó en rampa ID"
        num = parsear_float(texto)
        min_ma = 0.5 * ian
        if num is None:
            return False, f"ID inválido ({texto})"
        if min_ma < num <= ian:
            return True, f"OK ({min_ma:g}-{ian:g} mA)"
        return False, f"ID={num:g} mA fuera de ({min_ma:g}-{ian:g} mA)"

    return None, ""


def _es_no_vacio(val: Any) -> bool:
    return val is not None and str(val).strip() not in ("", "-")


def evaluar_diferencial(med: Dict[str, Any]) -> Dict[str, Any]:
    """
    Evalúa un dispositivo diferencial completo (trifásico R/S/T o bipolar, 0° y 180°).
    Retorna {'cumple': 'SI'|'NO', 'motivos': List[str], 'tiene_mediciones': bool}.
    """
    ian = parsear_float(med.get("corriente_fuga_ma", 30.0)) or 30.0
    tipo = str(med.get("tipo", "BIPOLAR") or "BIPOLAR").upper()
    fases = med.get("fases") or {}

    motivos: List[str] = []
    tiene_mediciones = False

    if fases:
        if tipo.startswith("TRI"):
            pares_fases = [(f"Fase {f}", fases.get(f, {})) for f in ("R", "S", "T")]
        else:
            datos_bip = fases.get("BIPOLAR", fases.get("R", {}))
            pares_fases = [("Bipolar", datos_bip)]

        for etiqueta_fase, datos_fase in pares_fases:
            if not isinstance(datos_fase, dict):
                continue
            for pol_k, pol_txt in (("0", "0°"), ("180", "180°")):
                pol_data = datos_fase.get(pol_k, {})
                if not isinstance(pol_data, dict):
                    continue

                if any(_es_no_vacio(pol_data.get(k)) for k in ("dt_1", "dt_5", "id_ma", "t_ms")):
                    tiene_mediciones = True

                for campo in ("dt_05", "dt_1", "dt_5", "id_ma"):
                    val_campo = pol_data.get(campo)
                    ok, msg = evaluar_parametro_rcd(campo, val_campo, ian)
                    if ok is False:
                        if campo == "dt_05":
                            tiene_mediciones = True
                        motivos.append(f"{etiqueta_fase} {pol_txt}: {msg}")
    else:
        t_simple = med.get("tiempo_disparo_ms")
        if _es_no_vacio(t_simple):
            tiene_mediciones = True
            ok, msg = evaluar_parametro_rcd("dt_1", t_simple, ian)
            if ok is False:
                motivos.append(f"0°: {msg}")

    return {
        "cumple": "NO" if motivos else "SI",
        "motivos": motivos,
        "tiene_mediciones": tiene_mediciones,
    }


def verificar_calibracion(
    fecha_str: Any, hoy: Optional[datetime.date] = None
) -> Dict[str, Any]:
    """Verifica si el certificado de calibración (vigencia anual de 365 días) está vigente."""
    texto = str(fecha_str or "").strip()
    if not texto:
        return {
            "estado": "vacio",
            "mensaje": "Sin fecha de calibración",
            "dias_restantes": None,
        }

    fecha_cal = parsear_fecha(texto)
    if fecha_cal is None:
        return {
            "estado": "invalida",
            "mensaje": "Formato esperado: DD/MM/AAAA",
            "dias_restantes": None,
        }

    ref_hoy = hoy or datetime.date.today()
    vencimiento = fecha_cal + datetime.timedelta(days=365)
    dias_restantes = (vencimiento - ref_hoy).days

    if dias_restantes < 0:
        return {
            "estado": "vencida",
            "mensaje": f"Calibración VENCIDA hace {abs(dias_restantes)} días",
            "dias_restantes": dias_restantes,
        }
    if dias_restantes <= 30:
        return {
            "estado": "por_vencer",
            "mensaje": f"Vence en {dias_restantes} días",
            "dias_restantes": dias_restantes,
        }
    return {
        "estado": "vigente",
        "mensaje": f"Vigente ({dias_restantes} días restantes)",
        "dias_restantes": dias_restantes,
    }


def resumir_relevamiento(datos: Dict[str, Any]) -> Dict[str, Any]:
    """Resume estadísticas del relevamiento PAT y RCD y genera el checklist de revisión."""
    datos_cliente = datos.get("datos_cliente", {}) or {}
    conclusiones = datos.get("conclusiones", {}) or {}
    pats = datos.get("mediciones_pat", []) or []
    difs = (
        datos.get("mediciones_diferenciales")
        if datos.get("mediciones_diferenciales") is not None
        else datos.get("mediciones_dif", [])
    ) or []

    total_pat = len(pats)
    pat_aptas = 0
    pat_no_cumplen = 0
    pat_anomalias = 0
    sectores_no_cumplen: List[str] = []
    sectores_anomalias: List[str] = []

    for idx, m in enumerate(pats, start=1):
        sector = str(m.get("sector", f"Toma {idx}") or f"Toma {idx}").strip()
        val_raw = m.get("valor_ohms", "")
        val_f = parsear_float(val_raw)
        if val_f is None:
            pat_anomalias += 1
            sectores_anomalias.append(f"#{m.get('nro_toma', idx)} {sector} ({val_raw})")
        else:
            cumple_str = str(m.get("cumple", "") or "").strip().upper()
            cont_str = str(m.get("continuidad", "SI") or "SI").strip().upper()
            cap_str = str(m.get("capacidad_carga", "SI") or "SI").strip().upper()
            if (
                val_f <= 0
                or val_f > 40.0
                or cumple_str == "NO"
                or cont_str == "NO"
                or cap_str == "NO"
            ):
                pat_no_cumplen += 1
                sectores_no_cumplen.append(f"#{m.get('nro_toma', idx)} {sector} ({val_f:g} Ω)")
            else:
                pat_aptas += 1

    hojas_grilla = max(1, (total_pat + 15) // 16)

    total_dif = len(difs)
    dif_aptos = 0
    dif_no_aptos = 0
    difs_fallidos: List[str] = []
    for idx, d in enumerate(difs, start=1):
        ev = evaluar_diferencial(d)
        if ev["cumple"] == "SI":
            dif_aptos += 1
        else:
            dif_no_aptos += 1
            ubic = str(d.get("sector", d.get("ubicacion", f"RCD #{idx}"))).strip()
            difs_fallidos.append(f"{ubic}: {'; '.join(ev['motivos'][:2])}")

    checklist: List[Dict[str, str]] = []

    razon = str(datos_cliente.get("razon_social", "") or "").strip()
    if razon:
        checklist.append({"nivel": "ok", "texto": f"Establecimiento: {razon}"})
    else:
        checklist.append({"nivel": "error", "texto": "Falta Razón Social / Establecimiento"})

    cuit = str(datos_cliente.get("cuit", "") or "").strip()
    if cuit:
        checklist.append({"nivel": "ok", "texto": f"C.U.I.T.: {cuit}"})
    else:
        checklist.append({"nivel": "warn", "texto": "Falta C.U.I.T. del establecimiento"})

    inst = str(
        conclusiones.get("instrumento_utilizado", datos_cliente.get("instrumento", "")) or ""
    ).strip()
    if inst:
        checklist.append({"nivel": "ok", "texto": f"Instrumento: {inst}"})
    else:
        checklist.append({"nivel": "warn", "texto": "Falta marca, modelo y serie del instrumento"})

    cal_info = verificar_calibracion(
        conclusiones.get("fecha_calibracion", datos_cliente.get("fecha_calibracion", ""))
    )
    if cal_info["estado"] == "vigente":
        checklist.append({"nivel": "ok", "texto": f"Calibración: {cal_info['mensaje']}"})
    elif cal_info["estado"] in ("por_vencer", "vacio", "invalida"):
        checklist.append({"nivel": "warn", "texto": f"Calibración: {cal_info['mensaje']}"})
    else:
        checklist.append({"nivel": "error", "texto": cal_info["mensaje"]})

    if total_pat == 0:
        checklist.append({"nivel": "warn", "texto": "Sin mediciones de Puesta a Tierra (PAT)"})
    elif pat_no_cumplen == 0 and pat_anomalias == 0:
        checklist.append(
            {
                "nivel": "ok",
                "texto": f"PAT: {total_pat} tomas cargadas ({hojas_grilla} hoja/s GRILLA) · todas aptas",
            }
        )
    else:
        checklist.append(
            {
                "nivel": "warn",
                "texto": (
                    f"PAT: {total_pat} tomas · {pat_aptas} aptas, "
                    f"{pat_no_cumplen} no cumplen, {pat_anomalias} anomalías"
                ),
            }
        )

    if total_dif == 0:
        checklist.append({"nivel": "warn", "texto": "Sin ensayos de diferenciales (RCD)"})
    elif dif_no_aptos == 0:
        checklist.append(
            {"nivel": "ok", "texto": f"RCD: {total_dif} disyuntor(es) · todos aptos"}
        )
    else:
        checklist.append(
            {
                "nivel": "error",
                "texto": f"RCD: {dif_no_aptos} de {total_dif} disyuntor(es) NO CUMPLE norma",
            }
        )

    return {
        "total_pat": total_pat,
        "pat_aptas": pat_aptas,
        "pat_no_cumplen": pat_no_cumplen,
        "pat_anomalias": pat_anomalias,
        "hojas_grilla": hojas_grilla,
        "sectores_no_cumplen": sectores_no_cumplen,
        "sectores_anomalias": sectores_anomalias,
        "total_dif": total_dif,
        "dif_aptos": dif_aptos,
        "dif_no_aptos": dif_no_aptos,
        "difs_fallidos": difs_fallidos,
        "checklist": checklist,
    }


def generar_textos_sugeridos(datos: Dict[str, Any]) -> Dict[str, str]:
    """Genera textos coherentes de Observaciones, Conclusiones y Recomendaciones según las mediciones."""
    res = resumir_relevamiento(datos)
    pats = datos.get("mediciones_pat", []) or []

    esquemas = []
    for m in pats:
        esq = str(m.get("esquema", "") or "").strip()
        if esq and esq not in esquemas:
            esquemas.append(esq)
    esq_txt = ", ".join(esquemas) if esquemas else "TT"

    hay_fallas_pat = res["pat_no_cumplen"] > 0 or res["pat_anomalias"] > 0
    hay_fallas_dif = res["dif_no_aptos"] > 0

    # --- Observaciones (Ítem 13) ---
    if not hay_fallas_pat and not hay_fallas_dif:
        obs = (
            f"El Esquema de conexión a tierra es {esq_txt}. "
            "RPaT: Es menor a 40 ohm en todos los puntos medidos. "
            "CONTINUIDAD DE LAS MASAS: Se comprobó que las masas eléctricas y las extrañas están conectadas "
            "a la puesta a tierra, como así también el borne de tierra de cada tomacorriente verificado. "
            "TIEMPO Y CORRIENTE DE DISPARO DE DIFERENCIALES: Cumple con los tiempos máximos de desconexión "
            "indicados por la RAEA (no dispara a 1/2×IΔn y dispara a 1×IΔn en tiempo menor a 300 ms)."
        )
    else:
        partes_obs = [f"El Esquema de conexión a tierra es {esq_txt}."]
        if res["total_pat"] > 0:
            if not hay_fallas_pat:
                partes_obs.append("RPaT: Es menor a 40 ohm en todos los puntos medidos.")
            else:
                detalle_pat = (
                    f"RPaT y CONTINUIDAD: Sobre {res['total_pat']} puntos relevados, {res['pat_aptas']} "
                    f"cumplen con resistencia ≤ 40 ohm y continuidad adecuada"
                )
                if res["pat_no_cumplen"] > 0:
                    detalle_pat += f", {res['pat_no_cumplen']} exceden 40 ohm o presentan falla de continuidad/carga"
                if res["pat_anomalias"] > 0:
                    detalle_pat += f" y {res['pat_anomalias']} presentan anomalías de polaridad o ausencia de conductor PE"
                partes_obs.append(detalle_pat + ".")
        if res["total_dif"] > 0:
            if not hay_fallas_dif:
                partes_obs.append(
                    "DIFERENCIALES (RCD): Los dispositivos ensayados cumplen con los tiempos y corrientes de disparo de la RAEA."
                )
            else:
                partes_obs.append(
                    f"DIFERENCIALES (RCD): {res['dif_no_aptos']} de {res['total_dif']} dispositivo(s) "
                    "ensayado(s) presentaron valores fuera de tolerancia normativa."
                )
        obs = " ".join(partes_obs)

    # --- Conclusiones ---
    if not hay_fallas_pat and not hay_fallas_dif:
        conc = (
            "Cumpliendo con las recomendaciones y siguiendo el protocolo de la Resolución SRT 900/2015, "
            "SE CONSIDERAN SATISFACTORIOS los valores obtenidos en las mediciones realizadas."
        )
    else:
        problemas = []
        if res["pat_no_cumplen"] > 0:
            muestra = ", ".join(res["sectores_no_cumplen"][:4])
            suf = "..." if len(res["sectores_no_cumplen"]) > 4 else ""
            problemas.append(f"{res['pat_no_cumplen']} punto(s) con RPaT/continuidad NO APTA ({muestra}{suf})")
        if res["pat_anomalias"] > 0:
            muestra = ", ".join(res["sectores_anomalias"][:4])
            suf = "..." if len(res["sectores_anomalias"]) > 4 else ""
            problemas.append(f"{res['pat_anomalias']} toma(s) con anomalías ({muestra}{suf})")
        if res["dif_no_aptos"] > 0:
            muestra = "; ".join(res["difs_fallidos"][:3])
            problemas.append(f"{res['dif_no_aptos']} diferencial(es) fuera de norma ({muestra})")

        conc = (
            "Siguiendo el protocolo de la Resolución SRT 900/2015, SE OBSERVAN NO CONFORMIDADES que requieren "
            "adecuación eléctrica: " + "; ".join(problemas) + "."
        )
        if res["pat_aptas"] > 0:
            conc += f" Los restantes {res['pat_aptas']} punto(s) PAT medidos arrojaron valores satisfactorios (≤ 40 Ω)."

    # --- Recomendaciones ---
    recs = []
    if res["pat_no_cumplen"] > 0:
        recs.append(
            "Adecuar el sistema de puesta a tierra y/o continuidad en los sectores que superaron los 40 Ω hasta alcanzar valores reglamentarios."
        )
    if res["pat_anomalias"] > 0:
        recs.append(
            "Corregir de forma prioritaria las polaridades invertidas y proveer conductor de protección (PE) efectivo en los tomacorrientes observados."
        )
    if res["dif_no_aptos"] > 0:
        recs.append(
            "Reemplazar o verificar el/los interruptor(es) diferencial(es) cuyo disparo estuvo fuera de los límites RAEA/IEC, y repetir el ensayo."
        )
    if hay_fallas_pat or hay_fallas_dif:
        recs.append("Realizar una nueva verificación de los puntos corregidos tras ejecutar las adecuaciones.")
    recs.append(
        "Repetir el protocolo de medición de puesta a tierra (SRT 900/2015) anualmente y pulsar el botón de prueba mecánica de los diferenciales bimestralmente."
    )

    return {
        "observaciones_datos": obs,
        "conclusiones_texto": conc,
        "recomendaciones_texto": " ".join(recs),
    }


def directorio_datos_app() -> Path:
    """Devuelve una carpeta escribible para guardar borrador y perfil del electricista."""
    candidatos: List[Path] = []
    flet_data_dir = (os.getenv("FLET_APP_STORAGE_DATA") or "").strip()
    if flet_data_dir:
        candidatos.append(Path(flet_data_dir) / "autosrt_data")
    candidatos.extend(
        [
            Path(__file__).resolve().parent / ".autosrt_data",
            Path.home() / ".autosrt",
            Path(tempfile.gettempdir()) / "autosrt",
        ]
    )
    for carpeta in candidatos:
        try:
            carpeta.mkdir(parents=True, exist_ok=True)
            if os.access(carpeta, os.W_OK):
                return carpeta
        except Exception:
            continue
    return Path(tempfile.gettempdir())


def guardar_borrador(
    state: Dict[str, Any], path: Optional[Path] = None
) -> Optional[str]:
    """Guarda el estado actual en JSON; retorna la hora HH:MM del guardado o None si falló."""
    destino = path or (directorio_datos_app() / "borrador_actual.json")
    ahora_hm = datetime.datetime.now().strftime("%H:%M")
    payload = {
        "ultimo_guardado": ahora_hm,
        "datos_cliente": state.get("datos_cliente", {}),
        "conclusiones": state.get("conclusiones", {}),
        "mediciones_pat": state.get("mediciones_pat", []),
        "mediciones_dif": state.get("mediciones_dif", []),
    }
    try:
        destino.parent.mkdir(parents=True, exist_ok=True)
        destino.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        return ahora_hm
    except Exception:
        return None


def cargar_borrador(path: Optional[Path] = None) -> Optional[Dict[str, Any]]:
    """Lee el borrador guardado si existe y es válido; caso contrario retorna None."""
    origen = path or (directorio_datos_app() / "borrador_actual.json")
    if not origen.exists():
        return None
    try:
        data = json.loads(origen.read_text(encoding="utf-8"))
        if isinstance(data, dict):
            return data
    except Exception:
        return None
    return None


def limpiar_borrador(path: Optional[Path] = None) -> None:
    """Elimina el archivo de borrador actual si existe."""
    destino = path or (directorio_datos_app() / "borrador_actual.json")
    try:
        if destino.exists():
            destino.unlink()
    except Exception:
        pass


def _perfil_default() -> Dict[str, Any]:
    return {
        "instrumento_utilizado": "",
        "fecha_calibracion": "",
        "metodologia": "Medición de impedancia por lazo de falla y comprobación de RCD",
        "provincia": "San Juan",
        "localidad": "Capital",
        "cp": "5400",
        "tema_oscuro": False,
        "sectores_recientes": [],
    }


def guardar_perfil_electricista(
    perfil: Dict[str, Any], path: Optional[Path] = None
) -> bool:
    """Guarda el perfil del electricista (instrumental, ubicación base, tema y sectores frecuentes)."""
    destino = path or (directorio_datos_app() / "perfil_electricista.json")
    merged = _perfil_default()
    merged.update(perfil or {})
    try:
        destino.parent.mkdir(parents=True, exist_ok=True)
        destino.write_text(json.dumps(merged, ensure_ascii=False, indent=2), encoding="utf-8")
        return True
    except Exception:
        return False


def cargar_perfil_electricista(path: Optional[Path] = None) -> Dict[str, Any]:
    """Carga el perfil del electricista o retorna valores por defecto."""
    origen = path or (directorio_datos_app() / "perfil_electricista.json")
    base = _perfil_default()
    if not origen.exists():
        return base
    try:
        data = json.loads(origen.read_text(encoding="utf-8"))
        if isinstance(data, dict):
            base.update(data)
    except Exception:
        pass
    return base


def construir_nombre_archivo(razon_social: str, fecha_medicion: str = "") -> str:
    """Construye el nombre base MEDICIONES_{CLIENTE}_{YYYYMMDD} para evitar sobrescrituras."""
    razon = (razon_social or "Cliente").strip()
    limpio = "".join(c if c.isalnum() else "_" for c in razon)
    while "__" in limpio:
        limpio = limpio.replace("__", "_")
    limpio = limpio.strip("_") or "Cliente"
    fecha = parsear_fecha(fecha_medicion) or datetime.date.today()
    return f"MEDICIONES_{limpio.upper()}_{fecha.strftime('%Y%m%d')}"


def listar_protocolos_xlsx(carpeta: Path) -> List[Path]:
    """Lista los archivos .xlsx de una carpeta ordenados por fecha de modificación descendente."""
    if not carpeta.exists() or not carpeta.is_dir():
        return []
    archivos = [
        p
        for p in carpeta.iterdir()
        if p.is_file() and p.suffix.lower() == ".xlsx" and not p.name.startswith("~$")
    ]
    archivos.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return archivos
