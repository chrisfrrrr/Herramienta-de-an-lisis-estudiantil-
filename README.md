# Herramienta de análisis estudiantil - AVE

Aplicación migrada a Streamlit para procesar reportes PDF del listado del curso y generar un Excel con indicadores de actividad estudiantil.

## Créditos

Desarrollado por Ing. Christian Pocol Asesor AVE  
Universidad del Valle de Guatemala UVG

## Funciones principales

- Carga de reporte PDF del curso.
- Selección de semana de análisis de 1 a 5.
- Detección automática de fecha y hora del sistema.
- Filtrado de registros con rol Estudiante.
- Cálculo de riesgo de abandono.
- Alerta de 72 horas o más sin actividad.
- Clasificación de avance por actividad total.
- Comparación contra horas mínimas esperadas por semana.
- Visualización de métricas y gráficas en pantalla.
- Descarga de Excel con hojas de reporte, alertas, estadísticas y criterios.

## Ejecutar localmente en PowerShell

Desde la carpeta del proyecto:

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
streamlit run app.py
```

Si PowerShell bloquea la activación del entorno virtual:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
.\venv\Scripts\Activate.ps1
```

## Ejecutar sin entorno virtual

```powershell
python -m pip install -r requirements.txt
streamlit run app.py
```

## Publicar en Streamlit Community Cloud

1. Subir estos archivos a un repositorio de GitHub:
   - `app.py`
   - `requirements.txt`
   - carpeta `assets`
2. Entrar a Streamlit Community Cloud.
3. Crear una nueva app conectando el repositorio.
4. Seleccionar `app.py` como archivo principal.
5. Publicar.

## Notas

La aplicación usa el formato del reporte de Canvas/UVG que contiene columnas como Nombre, Identificador de inicio de sesión, SIS, Sección, Rol, Última actividad y Actividad total.
