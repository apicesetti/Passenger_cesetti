# Mapa de asientos PID — Add-in MyGeotab

Add-in para MyGeotab que muestra, en tiempo real, quién va sentado en cada vehículo a partir de las tarjetas RFID leídas por el lector **PID de DIDCOM** (IOX). El conductor se identifica con llave NFC (IOX-NFCREADER) y se toma de `DeviceStatusInfo`.

## Vistas
- **Flota**: tarjetas por vehículo con ocupación, conductor y alertas. Buscador por patente o nombre.
- **Detalle**: plano del vehículo visto desde arriba (camioneta, combi o colectivo). Los asientos se llenan del 1 en adelante a medida que suben los pasajeros: verde = pasajero registrado, amarillo = tarjeta sin registrar, violeta = conductor.
- **Pasajeros**: padrón propio (hex de tarjeta, nombre, legajo/DNI, empresa, turno). Alta, edición, baja, importación desde Excel/CSV y lista de tarjetas leídas sin registrar.
- **Vehículos**: tipo y cantidad de asientos de pasajeros por vehículo.
- **Configuración**: descenso, desbloqueo, regla de bajada, vaciado por ignición, ventana y refresco.

## Datos
| Qué | Dónde |
|---|---|
| Lecturas de tarjetas | `CustomData` (trama `RUID` + UID). Se lee con `GetFeed`. |
| Conductor | `DeviceStatusInfo.driver` → `User` |
| Ignición (opcional) | `StatusData` `DiagnosticIgnitionId` |
| Padrón, vehículos y configuración | `AddInData` con `addInId = ad6jPCvZGTvCRcGDo4yXoiQ` |

Trama de ejemplo: `52 55 49 44 | 00 00 00 8C 9A 8C 13` → `RUID` + UID `8C9A8C13` (se quitan los bytes `00` de relleno).

## Reglas de ascenso y descenso
Iguales al add-in PID:
- Relectura de la misma tarjeta antes del **desbloqueo** (60 s) → se ignora.
- Tarjeta a bordo releída antes del **descenso** (30 min) desde que subió → bajada.
- Modo *PID*: una relectura después del tiempo de descenso cuenta como nuevo ascenso (el pasajero sigue a bordo). Modo *alternar*: toda relectura es bajada.
- Al subir toma el primer asiento libre desde el 1. Si no hay lugar queda “sin asiento” y toma el primero que se libere.

## Instalación
1. GitHub Pages: *Settings → Pages → Deploy from branch → main / (root)*.
2. MyGeotab: *Administración → Sistema → Configuración del sistema → Add-Ins → Nuevo add-in*, pegar `config.json`.
3. Aparece en el menú **Actividad → Mapa de asientos**.

## Estructura
- `asientosPid.html`: página del add-in (solo referencia los archivos; MyGeotab descarta los `<style>` en línea).
- `css/asientosPid.css`: estilos según el sistema de diseño **Zenith** de Geotab (usa sus variables, con respaldo del tema claro).
- `js/logic.js`: reglas de ascenso y descenso y planos de asientos (sin DOM).
- `js/app.js`: vistas, API de MyGeotab y AddInData.

## Pruebas
```
npm install
npm test     # lógica con las tramas reales de CESETTI_TEST + interfaz con API simulada
```
