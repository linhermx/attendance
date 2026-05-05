# Attendance

Herramienta interna para analizar asistencias diarias a partir de los archivos exportados por el sistema de checador.

## Entradas

- `Personal_*.xls`
- `Eventos de hoy_*.xls`

## Reglas base

- Lunes a viernes:
  - entrada `08:00`
  - salida a comida `12:00`
  - regreso de comida `12:45`
  - salida `17:00`
- Sabado:
  - entrada `08:00`
  - salida a comida `12:00`
  - regreso de comida `12:30`
  - salida `14:00`
- Retardo desde `08:01`
- Falta cuando no existe ninguna checada del trabajador en el día
- Horas extra solo por horas completas después de cumplir la jornada diaria

## Desarrollo local

```powershell
python .\attendance_gui.py
```

## Salidas

- `reporte_asistencia.xlsx`
  - `Resumen`
  - `Vista rápida`
  - `Faltas`
  - `Retardos`
  - `Incidencias`
  - `Detalle diario`
- `reporte_horas_extra.xlsx`
  - solo se genera cuando existen horas extra

## Builds

```powershell
.\scripts\build_windows.ps1
.\scripts\build_launcher.ps1
```
