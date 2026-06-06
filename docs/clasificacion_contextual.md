# Clasificación contextual de checadas

## Arquitectura

La lógica se divide en dos capas.

### Clasificación técnica

`ClassificationResult` contiene asignaciones, scores, alternativas, confianza, referencias utilizadas, checadas ambiguas o no utilizadas, flags técnicos y razones auditables. Esta capa no produce incidencias operativas.

### Evaluación de negocio

`BusinessEvaluation` contiene incidencias operativas, status, detalle visible, retardo, duración real de comida, salida anticipada y horas trabajadas.

El campo `Detalle` responde qué problema tiene el registro. La hoja `Auditoría clasificación` explica cómo razonó el motor.

## Eventos rígidos

Entrada y salida se clasifican por cercanía a ventanas configurables. Las ventanas pueden definirse a nivel general, turno o empleado.

Una checada fuera del límite máximo de un evento rígido no se fuerza a ocuparlo. Si existe riesgo de confundir una entrada muy tardía con comida, el resultado operativo conserva `Registro ambiguo`.

La primera checada anterior a la referencia teórica de comida tiene prioridad como entrada tardía contextual cuando existe una salida final posterior plausible y no hay una marca explícita de comida. Una única checada intermedia sin pareja queda sin asignar en auditoría y produce `Registro incompleto`, evitando duraciones de comida falsas.

## Eventos flexibles

Inicio y regreso de comida se clasifican como parejas cronológicas:

1. El inicio debe ser anterior al regreso.
2. La pareja debe ser coherente con entrada y salida, cuando existan.
3. La duración tiene mayor peso que la hora teórica.
4. La referencia programada de comida actúa solo como desempate débil.
5. Una pareja coherente puede clasificarse a cualquier hora dentro de la jornada.

Ejemplos válidos de lunes a viernes:

```text
12:00 -> 12:45
12:20 -> 13:05
12:40 -> 13:25
13:10 -> 13:55
14:00 -> 14:45
15:00 -> 15:45
```

Ejemplos válidos de sábado:

```text
12:00 -> 12:30
12:40 -> 13:10
13:10 -> 13:40
```

## Duración de comida

El regreso permitido se calcula dinámicamente:

```text
regreso_permitido = inicio_real_comida + máximo_de_la_jornada
```

La jornada se selecciona automáticamente por fecha: de lunes a viernes el máximo es 45 minutos y el sábado es 30 minutos. La comparación usa segundos:

```text
Lunes a viernes: 45:00 -> válido
Lunes a viernes: 45:01 -> Exceso de comida (+1 min)
Sábado: 30:00 -> válido
Sábado: 30:01 -> Exceso de comida (+1 min)
```

Los minutos de exceso visibles se redondean hacia arriba para no perder segundos excedidos.

## Omisiones y cortes parciales

`Sin inicio de comida` y `Sin regreso de comida` se determinan cuando termina la jornada, existe salida final o existe evidencia posterior incompatible.

Un corte parcial no genera esas omisiones únicamente porque pasó la referencia teórica de comida.

## Domingo no laborable

El domingo no utiliza un horario de respaldo ni pasa por el clasificador de eventos.

- No se evalúan entrada, comida o salida.
- No se generan asistencias, faltas, retardos, incidencias ni horas trabajadas.
- Un empleado sin checadas conserva el estatus neutral `Día no laborable`.
- Un empleado con checadas conserva el estatus neutral de revisión `Revisión`.
- El detalle visible indica `Checadas en día no laborable`.
- Las horas reales se almacenan únicamente en `Auditoría clasificación`.
- Las checadas dominicales no se deduplican, porque deben conservarse completas para revisión.

En modo rango, el domingo no cuenta como día laboral. Solo se incorpora al detalle cuando existen checadas dominicales para revisión.

## Horas trabajadas

Solo se calculan con entrada, inicio de comida, regreso de comida y salida reales en secuencia cronológica válida. Nunca se sustituyen ni inventan checadas.

## Catálogo operativo

- `Sin entrada`
- `Sin salida final`
- `Sin inicio de comida`
- `Sin regreso de comida`
- `Retardo (N min)`
- `Retardo grave (N min)`
- `Exceso de comida (+N min)`
- `Salida anticipada (N min)`
- `Registro ambiguo`
- `Registro incompleto`
- `Checada no reconocida`
- `Secuencia inválida`

Los estados neutrales `Día no laborable` y `Revisión` no forman parte del catálogo de incidencias.

## Auditoría

Por cada checada se conserva evento asignado, score, tipo de referencia, distancia, alternativas y razón. También se guarda el regreso permitido calculado desde el inicio real y la duración real de comida en segundos.

La información técnica no se copia al campo operativo `Detalle`.

## Configuración

```json
{
  "predeterminada": {
    "score_minimo": 30,
    "margen_ambiguedad": 8,
    "comida_aislada_minutos_decisivos": 15,
    "maximo_par_comida_minutos": 240
  },
  "turnos": {
    "Lunes a viernes": {
      "entrada": {
        "antes": 120,
        "despues": 120,
        "max_antes": 180,
        "max_despues": 210
      }
    }
  }
}
```

Las referencias de comida configuradas en un turno modifican únicamente el desempate técnico, no la evaluación operativa.

## Checada aislada en horario de comida

Para una única checada cercana a la referencia de comida, por ejemplo `12:01:00`:

```text
Entrada: --
Inicio comida: 12:01:00
Fin comida: --
Salida: --
Status: Incidencia
Detalle: Sin entrada | Sin regreso de comida | Sin salida final
```

La auditoría conserva por qué la checada fue tomada como inicio de comida y no como entrada.
