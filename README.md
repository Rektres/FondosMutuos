# Descargador de Fondos Mutuos CMF

Aplicación de escritorio (Tkinter) que descarga la cartola diaria de fondos mutuos publicada por la [CMF Chile](https://www.cmfchile.cl/institucional/estadisticas/fondos_cartola_diaria.php), permite filtrarla y exportarla a Excel.

## Requisitos

- Python 3.12
- Google Chrome instalado (Selenium abre un Chrome real para que el usuario resuelva el captcha de la CMF; no hay API pública ni forma de evitarlo)

## Instalación

```bash
python -m venv venv
venv\Scripts\pip install -r requirements.txt
```

## Ejecución

```bash
venv\Scripts\python fondos_mutuos_gui.py
```

## Uso

1. **Descargar datos**: elige el rango de fechas y presiona "Descargar datos". Se abre Chrome; resuelve el captcha y presiona "GENERAR ARCHIVO" en la página de la CMF. La app espera la descarga y carga los datos automáticamente.
2. **Entidad / Fondos**: filtra qué administradoras y qué fondos (y series) se muestran en la tabla.
3. **Filtros de columna**: cada columna de la tabla tiene su propio buscador en la cabecera, y es ordenable.
4. **Preferencias**: guarda la selección de entidades/fondos/series para la próxima vez que se abra la app.
5. **Plantilla**: exporta o carga la selección de fondos desde un archivo Excel, para compartir la misma configuración entre equipos.
6. **Exportar Excel**: genera un `.xlsx` agrupado por entidad con las columnas Fecha, N° Fondo, Nombre Fondo, Serie y Valor Cuota.

## Generar un ejecutable (opcional)

```bash
venv\Scripts\pip install pyinstaller
venv\Scripts\python -m PyInstaller --onefile --windowed --name "Fondos Mutuos CMF" ^
  --icon app.ico --add-data "Logo.png;." --collect-all selenium fondos_mutuos_gui.py
```

El resultado queda en `dist/Fondos Mutuos CMF.exe`. `--collect-all selenium` es necesario: Selenium resuelve `webdriver.Chrome`/`ChromeOptions` con importaciones perezosas que PyInstaller no detecta por análisis estático.
