# AGENTS.md

Reglas obligatorias para cualquier agente que trabaje en este repositorio.

## Procedimiento

- Leer este archivo antes de analizar, modificar, compilar, commitear, etiquetar o publicar.
- No hacer push, crear tags remotos ni publicar releases sin aprobacion explicita del usuario.
- No modificar tags o releases ya publicados. Todo bug posterior a produccion debe tratarse como hotfix nuevo.
- Antes de cerrar un cambio, ejecutar pruebas unitarias y compilacion cuando el cambio toque codigo.
- Limpiar caches, specs temporales, carpetas `build/` y artefactos no publicables antes de pedir aprobacion.
- No versionar reportes generados, Excels de prueba, imagenes de evidencia, logs, caches ni builds.
- Las pruebas automatizadas versionadas son validas; sus fixtures deben ser anonimos.

## Idioma y documentacion

- `README.md` debe mantenerse en espanol salvo instruccion explicita distinta.
- Las notas publicas de GitHub Release se redactan en ingles.
- `CHANGELOG.md` debe registrar cambios funcionales, sin motivos internos ni politicas sensibles.
- No documentar nombres reales de trabajadores, datos reales del reporte ni referencias internas de operacion.
- En pruebas, documentacion y comentarios versionados usar casos genericos y anonimizados.

## Releases y assets

- El asset obligatorio para actualizacion automatica es `attendance_windows.zip`.
- El instalador para usuario final es `attendance_setup.exe`.
- El paquete portable es `attendance_launcher_portable.zip`.
- Verificar que `dist/attendance_launcher/bundled_assets/attendance_release.json` apunte al tag correcto y a `attendance_windows.zip`.
- Validar simulacion de actualizacion en carpeta temporal; no tocar la instalacion real del equipo para pruebas.

## Reglas de negocio criticas

- No inventar horas bajo ningun caso.
- No asignar eventos por posicion cronologica simple.
- Entrada y salida son eventos rigidos.
- Comida es flexible; su regla operativa es duracion maxima segun jornada.
- Lunes a viernes: comida maxima de 45 minutos.
- Sabado: comida maxima de 30 minutos.
- Domingo: dia no laborable; checadas solo para revision.
- Horas extra no se calculan ni se muestran.
- Una primera checada antes de la referencia de comida puede ser entrada tardia, incluso si es muy tarde.
- Una checada aislada en zona clara de comida no debe convertirse en entrada inventada.
- Checadas cercanas duplicadas deben quedar en auditoria tecnica, no en detalle operativo.
- La zona protegida de comida solo evita falsos positivos de salida; no genera incidencias ni vuelve rigida la comida.

## Salida operativa vs auditoria

- `DETALLE` debe contener solo problemas operativos comprensibles para supervision.
- Mensajes tecnicos como score, ventana usada, normalizacion o razon de clasificacion deben quedar solo en auditoria.
- No ocultar incidencias reales en la capa de negocio para compensar una clasificacion incorrecta.
- Si el clasificador se equivoca, corregir la clasificacion desde origen.
