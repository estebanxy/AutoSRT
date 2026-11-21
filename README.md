# AutoSRT 900/15

Aplicación responsive (smartphones Android, tablets Android y escritorio) para relevamiento y generación de protocolos de medición de puesta a tierra, continuidad de masas y ensayo de dispositivos diferenciales (RCD) según **Resolución SRT 900/2015**.

## Estructura del proyecto

- `main.py`: Interfaz Flet responsive con persistencia offline de borrador y perfil de instrumental, validación normativa en tiempo real y exportación.
- `srt_logic.py`: Reglas normativas SRT 900/15 / RAEA, cálculo de vigencia de calibración, generación de conclusiones dinámicas y persistencia local.
- `excel_builder.py`: Lectura y llenado de la plantilla `Mediciones 2026.xlsx` con paginación automática de hojas `GRILLA`, `DIF. TRI` y `DIFERENCIALES`.
- `pdf_builder.py`: Generación de protocolo PDF A4 apaisado (vía ReportLab en escritorio o motor PDF 1.4 vectorial nativo sin dependencias externas en Android).

## Requisitos y ejecución

```bash
pip install -r requirements.txt
python3 main.py
```

Para generar el APK de Android se utiliza el flujo definido en `.github/workflows/build-apk.yml` con `flet==0.28.3`.
