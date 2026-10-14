# Inventario TI — T&C Asociados

Aplicación de escritorio para Windows que inventaría los equipos de cómputo,
registra a quién se entregan (con sus periféricos) y genera el **Acta de
Entrega FR-GTI-003**.

## Qué hace

| Botón | Resultado |
|---|---|
| ➕ Nuevo Inventario | Lee el hardware del equipo (WMI) y lo registra junto con el colaborador y los periféricos entregados. |
| 📊 Generar Reporte Excel | `reportes/Inventario_TYC_<fecha>.xlsx` con todo el inventario. |
| 📄 Acta de Entrega (1 a 1) | Eliges un colaborador y se genera su acta en `actas/`. |
| 📁 Actas Globales (todos) | Genera el acta de **cada** colaborador con equipo, de una sola vez. |

## Periféricos y acta de entrega

Al registrar un equipo se marcan los periféricos que se entregan (mouse,
teclado, monitor adicional, audífono, multipuerto, adaptador de red, base
refrigerante, conector de base refrigerante, y hasta 2 "otros" libres), cada
uno con su serial opcional. Se guardan en `equipos.perifericos`.

Al generar el acta:

- Los periféricos que tienen casilla en el formato quedan **marcados**.
- Los que no (Monitor adicional y los "otros") marcan **"Otro, Cual?"** con su nombre.
- Los **seriales** van en *Observaciones adicionales*.
- El sistema operativo se escribe en *Sistema Op.* (celda `C28`).

### Detección automática de periféricos

El botón **🔍 Detectar periféricos conectados** del formulario consulta WMI y
pre-marca lo que Windows ve conectado; el técnico revisa y corrige.

| Periférico | Qué se obtiene |
|---|---|
| Monitor externo (ej. Samsung) | Marca, modelo y **serial** (vienen del EDID del monitor). La pantalla del portátil se ignora. |
| Mouse / teclado USB (ej. Logitech) | Marca por el VID del fabricante. El serial solo si el dispositivo lo trae; los combos económicos (con cable o receptor inalámbrico) normalmente no. |
| Adaptador USB-Red | Nombre del adaptador y su **MAC** como identificador. |
| Hub USB (ej. el de 3 puertos del adaptador) | Se marca "Multipuerto"; casi nunca tiene serial. |

Lo conectado no es necesariamente lo que se entrega: por eso es una sugerencia.
Para ajustar los filtros con hardware real, ejecutar en Windows con los
periféricos conectados `python herramientas\diagnostico_perifericos.py` y
revisar `diagnostico_perifericos.txt`.

El mapeo periférico → celda está en `core/actas.py` (`CASILLAS`). La plantilla
se rellena parchando el XML de la hoja y no con openpyxl, porque openpyxl
pierde el vínculo que hace que Excel muestre las casillas como checkbox reales.

## Estructura

```
app.py              punto de entrada
core/
  database.py       SQLite: tablas, migración, guardado
  hardware.py       lectura de hardware (WMI, solo Windows)
  inventario.py     registra colaborador + equipo + periféricos
  perifericos.py    catálogo y (de)serialización de periféricos
  dispositivos.py   detección automática de periféricos (WMI)
  actas.py          generación del FR-GTI-003 desde la plantilla
  reportes.py       exportación a Excel
  rutas.py          rutas válidas en código fuente y en el .exe
gui/                ventanas (customtkinter)
assets/             plantilla_acta.xlsx, background.png
herramientas/       diagnostico_perifericos.py (qué expone Windows)
tests/              pruebas del núcleo (sin GUI ni Windows)
build_exe.bat       genera el .exe
```

## Ejecutar desde el código

```bat
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

## Generar el .exe

En Windows, con Python 3.10+ instalado, doble clic en `build_exe.bat`
(crea el entorno, instala dependencias y ejecuta PyInstaller).
Resultado: `dist\InventarioTI.exe`, un solo archivo.

- Copia el `.exe` a una carpeta **donde se pueda escribir**: junto a él se crean
  `data\` (base de datos), `actas\` y `reportes\`.
- Para conservar un inventario existente, copia su `inventario.db` a `data\`
  junto al `.exe`.
- La plantilla y el fondo viajan dentro del `.exe`. Si pones una
  `assets\plantilla_acta.xlsx` junto al `.exe`, esa tiene prioridad (sirve para
  actualizar el formato sin recompilar).

## Base de datos

SQLite. Si `equipos` ya existía sin la columna `perifericos`, se agrega sola al
abrir la aplicación.

```sql
CREATE TABLE colaboradores (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    tipo_documento TEXT, documento TEXT UNIQUE,
    nombre TEXT NOT NULL, cargo TEXT, area TEXT, sede TEXT, correo TEXT,
    fecha_registro DATETIME DEFAULT (datetime('now','localtime'))
);

CREATE TABLE equipos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    colaborador_id INTEGER,
    hostname TEXT, usuario_windows TEXT, serial TEXT UNIQUE,
    fabricante TEXT, modelo TEXT, procesador TEXT,
    memoria TEXT,       -- JSON: módulos de RAM
    discos TEXT,        -- JSON: discos
    ip TEXT, mac TEXT,
    sistema TEXT,       -- dict de Python como texto: nombre, version, build, arquitectura
    perifericos TEXT,   -- JSON: [{"clave": "mouse", "tipo": "Mouse", "serial": "ABC123", "detalle": "Logitech"}]
    fecha_registro DATETIME DEFAULT (datetime('now','localtime'))
);
```

Dónde queda el archivo: con el `.exe`, en `data\inventario.db` junto al
ejecutable; ejecutando `python app.py`, en `core\data\inventario.db`.

## Pruebas

```bat
python -m unittest discover -s tests -t . -v
```
