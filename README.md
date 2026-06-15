# attendance

![GitHub release](https://img.shields.io/github/v/release/linhermx/attendance)
![Platform](https://img.shields.io/badge/platform-Windows-blue)
![Python](https://img.shields.io/badge/python-3.10%2B-blue)

Herramienta para analizar control de asistencia a partir de archivos exportados por el checador.

## Funcionalidad

- análisis diario y por rango;
- clasificación contextual de entrada, inicio de comida, regreso de comida y salida;
- uso de estado registrado por el checador y dispositivo de origen como señales auditables de clasificación;
- normalización de checadas duplicadas cercanas antes de clasificar eventos;
- detección de retardos, faltas, omisiones, registros ambiguos y salidas anticipadas;
- cálculo de horas trabajadas únicamente con cuatro checadas reales, completas y ordenadas;
- reportes Excel con detalle operativo y una hoja separada de auditoría técnica;
- configuración de clasificación general, por turno y por empleado;
- domingos tratados como días no laborables, con checadas conservadas únicamente para revisión;
- GUI de Windows y CLI.

## Reglas de clasificación

- Entrada y salida son eventos rígidos y se evalúan contra el horario del turno.
- Inicio y regreso de comida son eventos flexibles.
- La hora programada de comida se usa solamente como referencia débil para clasificación.
- Una pareja cronológica coherente puede clasificarse como comida aunque ocurra después de la referencia.
- El estado registrado por el checador se usa como señal fuerte cuando identifica explícitamente entrada, salida a descanso, regreso de descanso o salida.
- Los estados genéricos `Entrada` y `Salida` se validan contra horario, secuencia y duración para evitar clasificaciones incorrectas.
- Si existen cuatro checadas utilizables, la secuencia entrada-comida-regreso-salida refuerza la hipótesis, pero no sustituye las reglas de contexto.
- La primera checada del día no se considera entrada automáticamente.
- La primera checada anterior a la referencia de comida se prioriza como entrada tardía cuando existe una salida final posterior plausible.
- Una checada aislada dentro del bloque flexible de comida se asigna al evento de comida más cercano cuando la diferencia es decisiva.
- Una selección incorrecta del tipo de checada puede corregirse por contexto y queda explicada en auditoría técnica.
- Las checadas ambiguas no se fuerzan.
- Nunca se crean ni sustituyen horas faltantes.

## Reglas operativas

### Entrada y salida

- El retardo se calcula únicamente con una entrada real clasificada.
- Un retardo de 60 minutos o más se muestra como `Retardo grave`.
- La salida anticipada se compara contra la salida programada.

### Comida

- La duración máxima se selecciona automáticamente por jornada: 45 minutos de lunes a viernes y 30 minutos el sábado.
- El regreso permitido se calcula desde el inicio real de comida.
- De lunes a viernes, `45:00` es válido y `45:01` genera `Exceso de comida (+1 min)`.
- El sábado, `30:00` es válido y `30:01` genera `Exceso de comida (+1 min)`.
- La comparación usa segundos reales.
- En cortes parciales no se reporta comida faltante solamente porque pasó la hora teórica.

### Horas trabajadas

- Requieren entrada, inicio de comida, regreso de comida y salida reales.
- Las cuatro checadas deben tener orden cronológico válido.
- Si falta cualquier evento, el campo queda vacío.

### Domingo

- El domingo es un día no laborable.
- No se aplican horarios de entrada, comida o salida.
- No se generan asistencias, faltas, retardos, incidencias ni horas trabajadas.
- Si existen checadas, el empleado aparece con estatus `Revisión` y detalle `Checadas en día no laborable`.
- Las horas registradas se conservan exclusivamente en la auditoría técnica.
- Las checadas dominicales no se deduplican, para conservar toda la evidencia disponible.
- Los empleados sin checadas aparecen con estatus neutral `Día no laborable`.

## Reportes

Modo diario:

- `reporte_asistencia.xlsx`

Modo rango:

- `reporte_asistencia_rango.xlsx`

Los reportes principales incluyen:

- `Resumen`
- `Vista rápida` o `Vista histórica`
- `Faltas`
- `Retardos`
- `Incidencias`
- `Detalle diario` o `Detalle consolidado`
- `Auditoría clasificación`

`Detalle`, `Detalle diario`, `Detalle consolidado` y las vistas de la GUI contienen únicamente información operativa. Scores, referencias horarias, alternativas, estado registrado por el checador, dispositivo de origen, checadas dominicales y razones de asignación se conservan exclusivamente en la hoja `Auditoría clasificación`.

La Vista rápida se construye únicamente con IDs y nombres presentes en la BBDD de personal cargada. Los archivos y carpetas de salida ubicados en testing, fixtures, mocks, demo, examples, evidence o casos no pueden utilizarse para un reporte operativo.

En reportes por rango, los domingos no incrementan días laborales ni métricas de asistencia. Solamente se incluyen cuando contienen checadas que requieren revisión.

## Horarios predeterminados

Lunes a viernes:

- entrada: `08:00`
- referencia de inicio de comida: `12:00`
- referencia de regreso: `12:45`
- salida: `17:00`

Sábado:

- entrada: `08:00`
- referencia de inicio de comida: `12:00`
- referencia de regreso: `12:30`
- salida: `14:00`
- duración máxima de comida: `30 minutos`

Domingo:

- día no laborable;
- sin horario normal;
- checadas únicamente para revisión.

## Uso en Windows

1. Descarga `attendance_setup.exe` desde Releases.
2. Instala y abre `Attendance`.
3. Selecciona modo diario o rango.
4. Carga el archivo de personal, el archivo de eventos y la carpeta de salida.
5. Ejecuta el análisis.

## Uso por CLI

Instalación:

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Reporte diario:

```powershell
python .\attendance_cli.py `
  --mode daily `
  --personal "Personal.xls" `
  --events "Eventos.xls" `
  --classification-config "examples\classification_config.example.json" `
  --outdir "salida" `
  --overwrite
```

Reporte por rango:

```powershell
python .\attendance_cli.py `
  --mode range `
  --personal "Personal.xls" `
  --range-events "Rango.xls" `
  --outdir "salida_rango" `
  --overwrite
```

## Configuración

La opción `--classification-config` acepta un JSON con políticas:

- `predeterminada`
- `turnos`
- `empleados`

La precedencia es `empleado > turno > predeterminada`. Consulta
[`examples/classification_config.example.json`](examples/classification_config.example.json) y
[`docs/clasificacion_contextual.md`](docs/clasificacion_contextual.md).

## Desarrollo

```powershell
$env:PYTHONPATH = "src"
python -m unittest discover -s tests -v
python .\attendance_gui.py
```

Build para Windows:

```powershell
.\scripts\build_windows.ps1
.\scripts\build_launcher.ps1
.\scripts\build_installer.ps1
```
