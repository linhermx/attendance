# attendance

![GitHub release](https://img.shields.io/github/v/release/linhermx/attendance)
![Platform](https://img.shields.io/badge/platform-Windows-blue)
![Python](https://img.shields.io/badge/python-3.10%2B-blue)

Herramienta interna para analizar el **control de asistencia** a partir de los archivos exportados por el sistema de checador, desarrollada en Python.

El sistema cruza la BBDD de personal con los eventos del checador y responde de forma clara:

- quién asistió y quién faltó
- quién llegó tarde
- quién omitió checadas importantes
- quién salió antes del horario programado
- quién generó horas extra pagables
- cómo se comporta el personal en un día o en un rango de fechas

Incluye:

- aplicación **Windows (.exe)** para usuarios no técnicos
- interfaz gráfica (GUI)
- actualización automática mediante launcher
- uso por línea de comandos (CLI) para usuarios técnicos
- generación de reportes en Excel

---

## Características

- lectura directa de archivos `.xls` exportados por el checador
- análisis en dos modos:
  - `Diario`
  - `Rango`
- cálculo de asistencias, retardos, faltas e incidencias
- detección de omisiones de checada:
  - entrada
  - salida a comida
  - regreso de comida
  - salida final
- cálculo de horas trabajadas
- cálculo de **horas extra pagables** solo por horas completas después de cumplir la jornada
- vista rápida en Excel para compartir con jefatura
- vista histórica por rango con empleados en horizontal e historial en vertical
- exclusión automática de registros inválidos en la BBDD
- depuración de checadas duplicadas conservando la más temprana
- inferencia inteligente cuando falta solo una checada y existen `3` registros coherentes
- detección de archivos invertidos entre personal y eventos
- exportación de resultados a **Excel**
- launcher con actualización automática por **GitHub Releases**
- instalador con `Inno Setup` para primer uso en Windows

---

## Uso en Windows (Recomendado)

### Descarga

1. Ir a **Releases**:
   https://github.com/linhermx/attendance/releases
2. Descargar:
   **`attendance_setup.exe`**

> No necesitas instalar Python ni dependencias.

---

### Primer uso

1. Ejecuta `attendance_setup.exe`
2. Completa la instalación
3. Abre `Attendance` desde el acceso directo creado por el instalador
4. El launcher:
   - revisa si hay una versión más reciente
   - pregunta si deseas actualizar
   - guarda sus archivos internos en `%LOCALAPPDATA%\LINHER\Attendance`
5. Acepta y el sistema se actualiza automáticamente

Después se abre la aplicación principal.
Siempre debes abrir `Attendance` desde su acceso directo o desde `attendance_launcher.exe`; los archivos internos se administran automáticamente.

---

### Uso de la aplicacion

La aplicación tiene dos modos principales:

- `Diario`
- `Rango`

#### Modo Diario

1. Selecciona el archivo **Personal**
2. Selecciona el archivo **Eventos del día**
3. Selecciona la **carpeta de salida**
4. Haz clic en **Analizar asistencia diaria**

Salidas generadas:

- `reporte_asistencia.xlsx`
- `reporte_horas_extra.xlsx` solo si existen horas extra pagables

#### Modo Rango

1. Selecciona el archivo **Personal**
2. Selecciona el archivo **Rango**
3. Selecciona la **carpeta de salida**
4. Haz clic en **Analizar rango**

Salidas generadas:

- `reporte_asistencia_rango.xlsx`
- `reporte_horas_extra_rango.xlsx` solo si existen horas extra pagables

---

## Formato de archivos de entrada

La aplicación trabaja directamente con los archivos exportados por el sistema de checador.

### Archivo de personal

Patron esperado:

- `Personal_*.xls`

Hoja requerida:

- primera hoja disponible o `data`

Columnas requeridas:

| Columna | Descripcion |
|---|---|
| `ID de usuario` | Identificador del trabajador |
| `Nombre` | Nombre del trabajador |
| `Apellido` | Apellido del trabajador |

Columnas opcionales que tambien se leen cuando existen:

- `Numero de tarjeta`
- `No. de departamento`
- `Departamento`

### Archivo de eventos diarios

Patron esperado:

- `Eventos de hoy_*.xls`

### Archivo de eventos por rango

Patron esperado:

- `Rango*.xls`

### Columnas requeridas para eventos

Hoja requerida:

- primera hoja disponible o `data`

| Columna | Descripcion |
|---|---|
| `Tiempo` | Fecha y hora de la checada |
| `ID de Usuario` | Identificador del trabajador |
| `Nombre` | Nombre registrado en el checador |
| `Apellido` | Apellido registrado en el checador |
| `Estado` | Tipo de evento (`Entrada`, `Salida`, etc.) |

Columnas opcionales que tambien se leen cuando existen:

- `Dispositivo`
- `Punto del evento`
- `Verificacion`
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

- **Sabado**
  - entrada `08:00`
  - salida a comida `12:00`
  - regreso de comida `12:30`
  - salida `14:00`
  - jornada objetivo `6 horas`

### Reglas de asistencia

- retardo a partir de `08:01`
- falta cuando no existe ninguna checada del trabajador en el dia
- se excluyen del analisis los registros de personal sin nombre usable
- si existen checadas duplicadas, se conserva la mas temprana
- si existe exactamente `1` checada faltante y el patron de `3` checadas encaja con los horarios esperados, el sistema la infiere
- las entradas inferidas se muestran con `~` en el reporte

### Reglas de comida

- lunes a viernes:
  - minimo `30 min`
  - maximo `45 min`
- sabado:
  - minimo `30 min`
  - maximo `30 min`
- comida menor al minimo no se eleva como incidencia
- comida mayor al maximo si se reporta como incidencia
- el exceso de comida se muestra solo como minutos por encima del maximo

### Reglas de salida y horas extra

- salida anticipada cuando la checada final queda antes del horario programado
- el reporte principal muestra tambien las `horas trabajadas`
- horas extra solo por **horas completas**
- primero se exige cumplir la jornada diaria
- existe una holgura de `5 min` para cumplir jornada, sin eliminar el retardo
- solo el tiempo posterior a esa jornada cuenta como **hora extra pagable**

Ejemplos:

- si alguien entra tarde y sale mas tarde solo para compensar su jornada, eso no genera horas extra
- si alguien entra `08:05` y sale `17:00`, conserva el `retardo`, pero si cumple su jornada
- si alguien cumple su jornada y ademas rebasa ese punto por una hora completa, si se genera `1` hora extra

### Reglas del modo Rango

- el periodo se toma desde la fecha minima hasta la fecha maxima del archivo
- se incluyen todos los dias laborales de `lunes` a `sabado`
- los `domingos` se excluyen
- si el ultimo dia viene incompleto, se marca como `corte parcial`
- si un dia laboral del rango no tiene registros globales, se marca como `Sin operacion`
- los dias `Sin operacion` no se contabilizan como falta masiva
- la `Vista historica` se ordena por `ID` ascendente

---

## Uso tecnico / desarrolladores (CLI)

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

### Instalacion

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

### Uso basico por CLI

#### Diario

```powershell
python .\attendance_cli.py `
  --mode daily `
  --personal "Personal_2026050411583.xls" `
  --events "Eventos de hoy_20260504133033.xls" `
  --outdir "salida" `
  --overwrite
```

#### Rango

```powershell
python .\attendance_cli.py `
  --mode range `
  --personal "Personal_2026050411583.xls" `
  --range-events "Rango.xls" `
  --outdir "salida_rango" `
  --overwrite
```

---

## Parametros CLI

| Parametro | Descripcion |
|---|---|
| `--mode` | Modo de analisis: `daily` o `range` |
| `--personal` | Ruta al archivo de personal |
| `--events` | Ruta al archivo de eventos del dia |
| `--range-events` | Ruta al archivo de eventos por rango |
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

### `reporte_asistencia_rango.xlsx`

Incluye estas hojas:

- `Resumen`
- `Vista histórica`
- `Faltas`
- `Retardos`
- `Incidencias`
- `Detalle consolidado`

### `reporte_horas_extra_rango.xlsx`

Solo se genera cuando existen horas extra pagables en el rango.

Incluye:

- `Resumen`
- `Detalle`

---

## Flujo recomendado

1. Exportar la BBDD de personal vigente
2. Elegir si el análisis será `Diario` o `Rango`
3. Exportar el archivo correspondiente desde el checador
4. Ejecutar la aplicación
5. Revisar el resumen principal
6. Compartir la `Vista rápida` o la `Vista histórica`
7. Revisar faltas, retardos e incidencias
8. Consultar el reporte de horas extra si aplica

---

## Build para Windows

```powershell
.\scripts\build_windows.ps1
.\scripts\build_launcher.ps1
.\scripts\build_installer.ps1
```

Los binarios y paquetes se generan en `dist/`:

- `attendance_windows\`
- `attendance_windows.zip`
- `attendance_launcher\`
- `attendance_launcher_portable.zip`
- `installer\attendance_setup.exe`

Los scripts crean `.\venv` automáticamente si todavía no existe.
