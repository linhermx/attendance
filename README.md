# attendance

![GitHub release](https://img.shields.io/github/v/release/linhermx/attendance)
![Platform](https://img.shields.io/badge/platform-Windows-blue)
![Python](https://img.shields.io/badge/python-3.10%2B-blue)

Herramienta interna para analizar el **control diario de asistencia** a partir de los archivos exportados por el sistema de checador, desarrollada en Python.

El sistema cruza la BBDD de personal con los eventos del día y responde de forma clara:

- quién asistió y quién faltó
- quién llegó tarde
- quién omitió checadas importantes
- quién salió antes del horario programado
- quién generó horas extra pagables

Incluye:

- aplicación **Windows (.exe)** para usuarios no técnicos
- interfaz gráfica (GUI)
- actualización automática mediante launcher
- uso por línea de comandos (CLI) para usuarios técnicos
- generación de reporte en Excel y log de ejecución

---

## Características

- lectura directa de archivos `.xls` del sistema de checador
- cálculo diario de asistencias, retardos, faltas e incidencias
- detección de omisiones de checada:
  - salida a comida
  - regreso de comida
  - salida final
- cálculo de salida anticipada
- cálculo de horas trabajadas
- cálculo de **horas extra pagables** solo por horas completas después de cumplir la jornada
- vista rápida en Excel para compartir con jefatura
- exclusión automática de registros inválidos en la BBDD
- depuración de checadas duplicadas conservando la más temprana
- validación de columnas requeridas y detección de incidencias globales
- exportación de resultados a **Excel**
- launcher con actualización automática por **GitHub Releases**

---

## Uso en Windows (Recomendado)

### Descarga

1. Ir a **Releases**:
   https://github.com/linhermx/attendance/releases
2. Descargar:
   **`attendance_launcher.exe`**

> No necesitas instalar Python ni dependencias.

---

### Primer uso

1. Ejecuta `attendance_launcher.exe`
2. El launcher:
   - revisa si hay una versión más reciente
   - pregunta si deseas actualizar
   - guarda sus archivos internos en `%LOCALAPPDATA%\LINHER\Attendance`
3. Acepta y el sistema se actualiza automáticamente

Después se abre la aplicación principal.
Siempre debes abrir `attendance_launcher.exe`; los archivos internos se administran automáticamente.

---

### Uso de la aplicación

1. Selecciona el archivo **Personal**
2. Selecciona el archivo **Eventos del día**
3. Selecciona la **carpeta de salida**
4. Haz clic en **Analizar asistencia**

Salidas generadas:

- `reporte_asistencia.xlsx`
- `run_log.txt`
- `reporte_horas_extra.xlsx` solo si existen horas extra pagables

---

## Formato de archivos de entrada

La aplicación trabaja directamente con los archivos exportados por el sistema de checador.

### Archivo de personal

Patrón esperado:

- `Personal_*.xls`

Hoja requerida:

- primera hoja disponible o `data`

Columnas requeridas:

| Columna | Descripción |
|---|---|
| `ID de usuario` | Identificador del trabajador |
| `Nombre` | Nombre del trabajador |
| `Apellido` | Apellido del trabajador |

Columnas opcionales que también se leen cuando existen:

- `Número de tarjeta`
- `No. de departamento`
- `Departamento`

### Archivo de eventos

Patrón esperado:

- `Eventos de hoy_*.xls`

Hoja requerida:

- primera hoja disponible o `data`

Columnas requeridas:

| Columna | Descripción |
|---|---|
| `Tiempo` | Fecha y hora de la checada |
| `ID de Usuario` | Identificador del trabajador |
| `Nombre` | Nombre registrado en el checador |
| `Apellido` | Apellido registrado en el checador |
| `Estado` | Tipo de evento (`Entrada`, `Salida`, etc.) |

Columnas opcionales que también se leen cuando existen:

- `Dispositivo`
- `Punto del evento`
- `Verificación`
- `Evento`
- `Notas`

---

## Reglas de negocio implementadas

### Horario base

- **Lunes a viernes**
  - entrada `08:00`
  - salida a comida `12:00`
  - regreso de comida `12:45`
  - salida `17:00`
  - jornada objetivo `9 horas`

- **Sábado**
  - entrada `08:00`
  - salida a comida `12:00`
  - regreso de comida `12:30`
  - salida `14:00`
  - jornada objetivo `6 horas`

### Reglas de asistencia

- retardo a partir de `08:01`
- falta cuando no existe ninguna checada del trabajador en el día
- se excluyen del análisis los registros de personal sin nombre usable
- si existen checadas duplicadas, se conserva la más temprana

### Reglas de comida

- lunes a viernes:
  - mínimo `30 min`
  - máximo `45 min`
- sábado:
  - mínimo `30 min`
  - máximo `30 min`
- comida menor al mínimo no se eleva como incidencia
- comida mayor al máximo sí se reporta como incidencia

### Reglas de salida y horas extra

- salida anticipada cuando la checada final queda antes del horario programado
- el reporte principal muestra también las `horas trabajadas`
- horas extra solo por **horas completas**
- las horas extra ya no se toman solo por salir después del horario
- primero se exige cumplir la jornada diaria
- existe una holgura de `5 min` para cumplir jornada, sin eliminar el retardo
- solo el tiempo posterior a esa jornada cuenta como **hora extra pagable**

Ejemplo:

- si alguien entra tarde y sale más tarde solo para compensar su jornada, eso no genera horas extra
- si alguien entra `08:05` y sale `17:00`, conserva el `retardo`, pero sí cumple su jornada
- si alguien cumple su jornada y además rebasa ese punto por una hora completa, sí se genera `1` hora extra

---

## Uso técnico / desarrolladores (CLI)

### Requisitos

- Python **3.10+**
- Windows

Dependencias principales:

- `pandas`
- `openpyxl`
- `xlsxwriter`
- `requests`
- `xlrd`

---

### Instalación

Se recomienda usar un entorno virtual (`venv`).

```powershell
git clone https://github.com/linhermx/attendance.git
cd attendance

python -m venv venv
.\venv\Scripts\Activate.ps1

pip install -r requirements.txt
pip install -r requirements-dev.txt
```

---

### Ejecutar GUI en desarrollo

```powershell
python .\attendance_gui.py
```

### Ejecutar launcher en desarrollo

```powershell
python .\attendance_launcher.py
```

Si no existe una release publicada, el launcher cae localmente a `attendance_gui.py`.

---

### Uso básico por CLI

```powershell
python .\attendance_cli.py `
  --personal "Personal_2026050411583.xls" `
  --events "Eventos de hoy_20260504133033.xls" `
  --outdir "salida" `
  --overwrite
```

---

## Parámetros CLI

| Parámetro | Descripción |
|---|---|
| `--personal` | Ruta al archivo de personal |
| `--events` | Ruta al archivo de eventos del día |
| `--outdir` | Carpeta de salida. Default: `salida` |
| `--overwrite` | Sobrescribe el reporte si ya existe |

---

## Reportes generados

### `reporte_asistencia.xlsx`

Incluye estas hojas:

- `Resumen`
- `Vista rápida`
- `Faltas`
- `Retardos`
- `Incidencias`
- `Detalle diario`

### `reporte_horas_extra.xlsx`

Solo se genera cuando existen horas extra pagables.

Incluye:

- `Resumen`
- `Horas extra`

### `run_log.txt`

Resume:

- fecha analizada
- horario aplicado
- total de empleados válidos
- asistencias
- retardos
- faltas
- personal con incidencias
- personal con horas extra
- horas extra totales
- observaciones globales detectadas

---

## Flujo recomendado

1. Exportar la BBDD de personal vigente
2. Exportar los eventos del día desde el checador
3. Ejecutar la aplicación
4. Revisar el corte del día
5. Compartir la `Vista rápida`
6. Revisar `Faltas`, `Retardos` e `Incidencias`
7. Consultar el reporte de horas extra si aplica

---

## Build para Windows

```powershell
.\scripts\build_windows.ps1
.\scripts\build_launcher.ps1
```

Los binarios se generan en `dist/`:

- `attendance_windows.exe`
- `attendance_launcher.exe`

Ambos scripts crean `.\venv` automáticamente si todavía no existe.
