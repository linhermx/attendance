# Clasificación contextual de checadas

## Arquitectura

La lógica se divide en dos capas.

### Clasificación técnica

`ClassificationResult` contiene asignaciones, scores, alternativas, confianza, referencias utilizadas, checadas ambiguas o no utilizadas, flags técnicos y razones auditables. Esta capa no produce incidencias operativas.

El clasificador recibe la hora de la checada, el estado registrado por el checador y el dispositivo de origen. Cuando el estado registrado coincide con el catálogo operativo válido, se toma como fuente primaria de clasificación. Cuando falta o no es válido, el sistema usa horario, secuencia, duración real, duplicados cercanos y zona protegida de comida como fallback contextual.

### Evaluación de negocio

`BusinessEvaluation` contiene incidencias operativas, status, detalle visible, retardo, duración real de comida, salida anticipada y horas trabajadas.

El campo `Detalle` responde qué problema tiene el registro. La hoja `Auditoría clasificación` explica cómo razonó el motor.

## Eventos rígidos

Entrada y salida se clasifican por cercanía a ventanas configurables. Las ventanas pueden definirse a nivel general, turno o empleado.

Una checada fuera del límite máximo de un evento rígido no se fuerza a ocuparlo. Si existe riesgo de confundir una entrada muy tardía con comida, el resultado operativo conserva `Registro ambiguo`.

La primera checada anterior a la referencia teórica de comida tiene prioridad como entrada tardía contextual cuando existe una salida final posterior plausible y no hay una marca explícita de comida. Una única checada intermedia sin pareja queda sin asignar en auditoría y produce `Registro incompleto`, evitando duraciones de comida falsas.

Una checada aislada dentro del bloque flexible de comida no se deja en blanco si una de las referencias (`inicio de comida` o `fin de comida`) gana de forma clara por cercanía. En ese caso se asigna al evento flexible más cercano y la evaluación operativa reporta únicamente las omisiones realmente observables.

## Estado y dispositivo del checador

Los estados válidos del checador se usan como fuente primaria:

| Estado registrado | Evento sugerido |
| --- | --- |
| `Entrada` | entrada |
| `Salida a descanso` | inicio de comida |
| `Regreso descanso` | regreso de comida |
| `Salida` | salida final |

Las marcas explícitas de comida no vuelven rígida la comida; solamente identifican el tipo de evento que el trabajador seleccionó. La duración máxima de comida sigue siendo la regla operativa.

Aunque el estado declarado sea válido, se valida contra la secuencia completa del día. Si una marca de comida deja sin clasificar una salida final clara, el clasificador puede corregirla por contexto y conservar la diferencia solamente en auditoría técnica.

Si el estado registrado no pertenece al catálogo válido, el clasificador vuelve al análisis contextual y conserva la evidencia en la auditoría técnica. El campo operativo `Detalle` no muestra esta información técnica.

## Eventos flexibles

Inicio y regreso de comida se clasifican como parejas cronológicas:

1. El inicio debe ser anterior al regreso.
2. La pareja debe ser coherente con entrada y salida, cuando existan.
3. La duración tiene mayor peso que la hora teórica.
4. La referencia programada de comida actúa solo como desempate débil.
5. Una pareja coherente puede clasificarse a cualquier hora dentro de la jornada.

Cuando existen cuatro checadas utilizables, la secuencia cronológica entrada-comida-regreso-salida refuerza una hipótesis válida, pero no se usa como regla automática ni reemplaza los demás criterios.

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

Se calculan cuando existe una entrada real y una salida real en orden cronológico válido.

Si también existe el par completo de comida, se descuenta la duración real entre inicio y regreso de comida. Si no existe comida registrada, el tiempo trabajado se calcula como salida menos entrada. Si la comida está incompleta o la secuencia es inválida, el campo queda vacío.

Para el cálculo exclusivamente, una entrada registrada hasta `08:00:59` se considera como `08:00:00`. Desde `08:01:00` se usa la hora real registrada. Nunca se sustituyen ni inventan checadas faltantes.

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

Por cada checada se conserva evento asignado, score, tipo de referencia, distancia, alternativas, estado registrado, dispositivo de origen y razón. También se guarda el regreso permitido calculado desde el inicio real y la duración real de comida en segundos.

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

Para una única checada cercana al regreso esperado, por ejemplo `12:44:56` entre semana o `12:29:58` en sábado:

```text
Entrada: --
Inicio comida: --
Fin comida: 12:44:56
Salida: --
Status: Incidencia
Detalle: Sin entrada | Sin inicio de comida | Sin salida final
```
