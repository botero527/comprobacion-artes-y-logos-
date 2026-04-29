# 📚 Guía Completa para Aprender del Código - Script ARTES

¡Hola bro! Vamos a descomponer este código juntos para que puedas aprender conceptos fundamentales de Python y programación. Este script es muy práctico y te va a enseñar muchas cosas importantes. 🚀

---

## 🎯 Introducción: ¿Qué hace este Script?

Este script automatiza la búsqueda de carpetas llamadas "ARTES" dentro de una estructura de directorios muy específica:

```
RUTA_BASE (\\192.168.2.37\ingenieria\PRODUCCION\AGP PLANOS TECNICOS)
├── Vehículo 1
│   ├── Modelo A
│   │   ├── Versión 1.0
│   │   │   └── ARTES ← (esta carpeta nos interesa)
│   │   └── Versión 2.0
│   └── Modelo B
│       └── ...
└── Vehículo 2
    └── ...
```

**Básicamente:**
1. Escanea una carpeta de red enorme
2. Busca carpetas llamadas "ARTES" dentro de cada versión de cada modelo de cada vehículo
3. Genera un reporte en Excel con todas las rutas encontradas
4. Crea una hoja por cada vehículo para facilitar la lectura

---

## 📦 Importaciones (Líneas 7-10)

```python
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime
```

### 📖 Explicación:

| Módulo | ¿Para qué sirve? | Consejo profesional |
|--------|------------------|---------------------|
| `os` | Manipular archivos y carpetas (crear, eliminar, navegar) | Es fundamental para cualquier script que maneje archivos |
| `openpyxl` | Crear y modificar archivos Excel | Es la librería más popular para Excel en Python |
| `datetime` | Obtener fecha y hora actual | Úsalo para logs y marcas de tiempo |

**💡 Consejo profesional:** 
- Siempre importa solo lo que necesitas (`from módulo import cosa1, cosa2`)
- Esto hace el código más limpio y consume menos memoria
- El alias `as` (ej: `import numpy as np`) se usa para abbreviary nombres largos

---

## ⚙️ Constantes de Configuración (Líneas 12-18)

```python
RUTA_BASE = r"\\192.168.2.37\ingenieria\PRODUCCION\AGP PLANOS TECNICOS"
NOMBRE_CARPETA_ARTE = "ARTES"
ARCHIVO_SALIDA = "Migracion_ARTES_AGP.xlsx"
```

### 📖 Explicación:

| Constante | Valor | Propósito |
|-----------|-------|-----------|
| `RUTA_BASE` | Ruta UNC de red | Punto de inicio del escaneo |
| `NOMBRE_CARPETA_ARTE` | "ARTES" | Nombre de carpeta a buscar |
| `ARCHIVO_SALIDA` | Nombre del Excel | Donde se guardará el resultado |

**💡 Consejo profesional:**
- Las constantes van en MAYÚSCULAS con GUIONES_BAJOS
- La `r"..."` antes del string significa **string raw** (crudo)
- Evita que las barras `\` se interpreten como caracteres especiales
- **Variables de configuración** al inicio facilitan cambiar el comportamiento sin tocar el código

---

## 🎨 Estilos y Colores (Líneas 20-26)

```python
COLOR_HEADER   = "1F3864"
COLOR_TITULO   = "2E75B6"
COLOR_BLANCO   = "FFFFFF"
COLOR_FILA_ALT = "EEF4FF"

thin   = Side(style="thin", color="BBBBBB")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
```

### 📖 Explicación:

Estos son **constantes de estilo** para el Excel. Python usa:
- **RGB hexadecimal**: `#RRGGBB` pero en Excel se omite el `#`
- `Color_HEADER = "1F3864"` = Azul oscuro corporativo
- `Color_TITULO = "2E75B6"` = Azul medio

**El objeto `Border`:**
```python
thin = Side(style="thin", color="BBBBBB")  # Línea delgada gris
border = Border(left=thin, right=thin, top=thin, bottom=thin)  # Borde completo
```

**💡 Consejo profesional:**
- Define colores como constantes al inicio
- Esto facilita mantener un diseño consistente
- Permite cambiar el "tema" de todo el documento en un solo lugar

---

## 🔧 Primera Función: `estilo_header()` (Líneas 29-33)

```python
def estilo_header(cell, bg=COLOR_HEADER):
    cell.font      = Font(name="Arial", bold=True, color=COLOR_BLANCO, size=10)
    cell.fill      = PatternFill("solid", start_color=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border    = border
```

### 📖 ¿Qué hace?
Aplica formato de encabezado a una celda de Excel:
- **Font**: Arial, negrita, blanca, tamaño 10
- **Fill (fondo)**: Color sólido (el que se pase por parámetro)
- **Alignment**: Centrado horizontal y vertical, texto envuelto
- **Border**: Borde completo

### 🧠 Conceptos que enseña:

1. **Parámetros con valor por defecto:**
   ```python
   def estilo_header(cell, bg=COLOR_HEADER):
   ```
   Si no pasas `bg`, usa `COLOR_HEADER` por defecto.

2. **Programación orientada a objetos (POO):**
   - `cell.font`, `cell.fill`, `cell.alignment` son **atributos** del objeto `cell`
   - `Font()`, `PatternFill()`, `Alignment()` son **constructores** que crean objetos

3. **Principio DRY (Don't Repeat Yourself):**
   - En lugar de escribir el mismo formato 10 veces, creas UNA función y la reutilizas

**💡 Consejo profesional:**
- Las funciones pequeñas y especializadas son mejores que las funciones gigantes
- Nombra las funciones con verbos: `estilo_header`, `calcular_total`, `procesar_datos`

---

## 🔧 Segunda Función: `estilo_dato()` (Líneas 36-41)

```python
def estilo_dato(cell, alt=False, center=False):
    cell.font      = Font(name="Arial", size=10)
    cell.alignment = Alignment(horizontal="center" if center else "left", vertical="center")
    cell.border    = border
    if alt:
        cell.fill = PatternFill("solid", start_color=COLOR_FILA_ALT)
```

### 📖 ¿Qué hace?
Aplica formato de datos (texto normal) a una celda:
- Fuente Arial tamaño 10
- Alineación: centrado si `center=True`, sino a la izquierda
- Borde completo
- Si `alt=True`, aplica color de fila alternada (para leggibilidad)

### 🧠 Conceptos que enseña:

1. **Operador ternario (condicional en una línea):**
   ```python
   "center" if center else "left"
   ```
   Es como decir: "Si center es True, usa 'center', sino usa 'left'"

2. **Parámetros opcionales:**
   - `alt=False` significa "por defecto no es fila alternada"
   - `center=False` significa "por defecto no está centrado"

3. **Condicional simple:**
   ```python
   if alt:
       cell.fill = PatternFill(...)
   ```

**💡 Consejo profesional:**
- Usa valores por defecto sensatos (`False`, `0`, `""`, `None`)
- Esto hace que las funciones sean más fáciles de usar

---

## 🔧 Tercera Función: `escanear()` (Líneas 44-93)

Esta es la función **PRINCIPAL** de procesamiento de datos. Es la más importante del script.

```python
def escanear(ruta_base):
    """
    Recorre: Vehiculo > Modelo > Version
    En cada Version busca carpeta ARTES y guarda la ruta exacta UNC.
    Retorna dict: { vehiculo: [ {modelo, version, ruta_arte} ] }
    """
    resultado = {}

    if not os.path.exists(ruta_base):
        print(f"[ERROR] No se puede acceder a: {ruta_base}")
        print("Verifica conexion de red y que la ruta este accesible.")
        return resultado

    vehiculos = sorted([d for d in os.listdir(ruta_base)
                        if os.path.isdir(os.path.join(ruta_base, d))])

    print(f"Vehiculos encontrados: {len(vehiculos)}\n")

    for vehiculo in vehiculos:
        ruta_vehiculo = os.path.join(ruta_base, vehiculo)
        filas = []

        modelos = sorted([d for d in os.listdir(ruta_vehiculo)
                          if os.path.isdir(os.path.join(ruta_vehiculo, d))])

        for modelo in modelos:
            ruta_modelo = os.path.join(ruta_vehiculo, modelo)

            versiones = sorted([d for d in os.listdir(ruta_modelo)
                                if os.path.isdir(os.path.join(ruta_modelo, d))])

            for version in versiones:
                ruta_version = os.path.join(ruta_vehiculo, modelo, version)
                ruta_arte    = os.path.join(ruta_version, NOMBRE_CARPETA_ARTE)

                if os.path.isdir(ruta_arte):
                    filas.append({
                        "modelo":    modelo,
                        "version":   version,
                        "ruta_arte": ruta_arte,
                    })
                    print(f"  [OK] {vehiculo} | {modelo} | {version}")
                    print(f"       {ruta_arte}")
                else:
                    print(f"  [--] {vehiculo} | {modelo} | {version} (sin ARTES)")

        if filas:
            resultado[vehiculo] = filas

    return resultado
```

### 📖 Paso a Paso:

#### Paso 1: Inicializar el resultado
```python
resultado = {}
```
Creamos un diccionario vacío. Estructura: `{vehiculo: [lista de datos]}`

#### Paso 2: Validar que la ruta existe
```python
if not os.path.exists(ruta_base):
    print(f"[ERROR] No se puede acceder a: {ruta_base}")
    print("Verifica conexion de red y que la ruta este accesible.")
    return resultado
```
Si la ruta de red no está disponible, salimos temprano con un error claro.

**Concepto importante:** Esto se llama **"guard clause"** o validación temprana. Es mejor validar y salir rápido que seguir procesando con datos inválidos.

#### Paso 3: Obtener lista de vehículos
```python
vehiculos = sorted([d for d in os.listdir(ruta_base)
                    if os.path.isdir(os.path.join(ruta_base, d))])
```

**Desglosemos esto (list comprehension + filter):**

```python
# Esto es una "list comprehension" con condición
[d for d in os.listdir(ruta_base) if os.path.isdir(os.path.join(ruta_base, d))]

# Equivalente a:
vehiculos = []
for d in os.listdir(ruta_base):
    ruta_completa = os.path.join(ruta_base, d)
    if os.path.isdir(ruta_completa):  # Solo directorios, no archivos
        vehiculos.append(d)
```

**`os.path.join()`**: Combina rutas de forma correcta (agrega `\ entre carpetas)
```python
os.path.join("carpeta1", "carpeta2", "archivo.txt")
# Resultado: "carpeta1\carpeta2\archivo.txt" (en Windows)
```

**`sorted()`**: Ordena alfabéticamente la lista

**💡 Consejo profesional:**
- Las list comprehensions son más rápidas y pythonicas que los bucles for
- Pero si son muy complejas, es mejor usar un for normal para legibilidad

#### Paso 4: Bucle anidado (tres niveles)
```python
for vehiculo in vehiculos:
    for modelo in modelos:
        for version in versiones:
            # procesa cada combinación
```

Este es un patrón clásico de **recorrido de árbol de directorios**. Cada nivel representa una carpeta dentro de otra.

#### Paso 5: Buscar la carpeta ARTES
```python
ruta_arte = os.path.join(ruta_version, NOMBRE_CARPETA_ARTE)

if os.path.isdir(ruta_arte):
    # La carpeta existe, guardamos los datos
    filas.append({
        "modelo":    modelo,
        "version":   version,
        "ruta_arte": ruta_arte,
    })
```

**Estructura de datos - Lista de diccionarios:**
```python
filas = [
    {"modelo": "Corolla", "version": "XRS 2024", "ruta_arte": "..."},
    {"modelo": "Corolla", "version": "XLE 2024", "ruta_arte": "..."},
    {"modelo": "Hilux", "version": "SR5 2024", "ruta_arte": "..."},
]
```

#### Paso 6: Guardar en el diccionario resultado
```python
if filas:  # Solo si hay datos
    resultado[vehiculo] = filas
```

### 🧠 Estructura de datos final:

```python
resultado = {
    "TOYOTA": [
        {"modelo": "Corolla", "version": "XRS", "ruta_arte": "..."},
        {"modelo": "Corolla", "version": "XLE", "ruta_arte": "..."}
    ],
    "HYUNDAI": [
        {"modelo": "Tucson", "version": "Limited", "ruta_arte": "..."}
    ]
}
```

**💡 Consejo profesional:**
- Usa diccionarios para buscar datos rápidamente por clave
- Usa listas cuando el orden importa
- Anida estructuras cuando representas datos jerárquicos (como carpetas)

---

## 🔧 Cuarta Función: `crear_excel()` (Líneas 96-204)

Esta función genera el archivo Excel con formato profesional.

```python
def crear_excel(resultado):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
```

### 📖 Inicializar Excel

```python
wb = openpyxl.Workbook()  # Crear nuevo libro
wb.remove(wb.active)     # Eliminar hoja por defecto "Sheet"
```

**Concepto:** Al crear un Workbook en openpyxl, siempre viene con una hoja llamada "Sheet". La eliminamos porque crearemos las nuestras.

### 📖 Crear hoja maestra (líneas 102-150)

```python
ws_master = wb.create_sheet("TODAS LAS RUTAS (IT)")
```

Esta hoja contiene TODAS las rutas en una sola tabla (útil para el equipo de IT que necesita migrar los archivos).

#### Encabezados合并 celdas:
```python
ws_master.merge_cells("A1:E1")  # Combina celdas A1 a E1
ws_master["A1"] = "MIGRACION ARTES — AGP PLANOS TECNICOS"
```

#### Formateo del título:
```python
ws_master["A1"].font      = Font(name="Arial", size=14, bold=True, color=COLOR_BLANCO)
ws_master["A1"].fill      = PatternFill("solid", start_color=COLOR_HEADER)
ws_master["A1"].alignment = Alignment(horizontal="center", vertical="center")
ws_master.row_dimensions[1].height = 32
```

**Concepto importante - Method Chaining:**
En Python, cada método devuelve un objeto, permitiendo encadenar:
```python
# Esto NO es chaining (es asignar a variables separadas)
celda = ws_master["A1"]
celda.font = Font(...)
celda.alignment = Alignment(...)
```

#### Escribir datos con bucle:
```python
fila = 4
for vehiculo, filas in sorted(resultado.items()):
    for datos in filas:
        alt = (fila % 2 == 0)  # alternar colores
        estilo_dato(ws_master.cell(fila, 1, fila - 3), alt=alt, center=True)
        # ... más celdas
        fila += 1
```

**Concepto - enumerate vs índice manual:**
```python
# Este código usa índice manual (fila += 1)
# También podrías usar enumerate:
for fila_idx, datos in enumerate(filas, start=4):
    # enumerate te da (índice, valor)
```

#### Anchos de columna:
```python
ws_master.column_dimensions["A"].width = 6
ws_master.column_dimensions["B"].width = 25
# ...
```

#### Freeze panes (congelar filas/columnas):
```python
ws_master.freeze_panes = "A4"
```
Esto congela las primeras 3 filas (título y encabezados), muy útil cuando hay muchos datos.

### 📖 Crear hojas por vehículo (líneas 152-197)

```python
for vehiculo, filas in sorted(resultado.items()):
    nombre_hoja = (vehiculo[:31]
                   .replace("/","-").replace("\\","-")
                   .replace("*","").replace("?","")
                   .replace("[","").replace("]","").replace(":",""))
    ws = wb.create_sheet(nombre_hoja)
```

**Limpiar nombre de hoja:**
Excel no permite ciertos caracteres en nombres de hojas:
- `/` → `-`
- `\` → `-`
- `*` → eliminar
- `?` → eliminar
- `:` → eliminar

**Cortar a 31 caracteres:** Excel limita los nombres de hojas a 31 caracteres.

### 📖 Guardar archivo:
```python
wb.save(ARCHIVO_SALIDA)
```

---

## 🚀 Punto de Entrada (Líneas 207-219)

```python
if __name__ == "__main__":
    print("=" * 60)
    print("  REPORTE RUTAS ARTES — AGP PLANOS TECNICOS")
    print("=" * 60)
    print(f"Ruta base: {RUTA_BASE}\n")

    resultado = escanear(RUTA_BASE)

    if not resultado:
        print("\n[!] No se encontraron carpetas ARTES.")
        print("    Verifica conexion de red y ruta.")
    else:
        crear_excel(resultado)
```

### 📖 ¿Qué es `if __name__ == "__main__":`?

Esta condición especial significa:
> "Ejecuta este código SOLO cuando corras el script directamente, NO cuando lo importes como módulo"

**¿Para qué sirve?**
- Permite que el script sea importable por otros scripts
- Pero si lo ejecutas python directamente, hace lo que tiene que hacer
- Es una **mejores prácticas de Python**

**Ejemplo:**
```python
# Si tienes funciones_utiles.py y lo importas desde otro archivo,
# el código dentro de "if __name__ == '__main__':" NO se ejecutará.
# Solo se ejecutará si haces: python funciones_utiles.py
```

---

## 📊 Resumen de Conceptos Aprendidos

| Concepto | Ejemplo en el código | Para qué sirve |
|----------|---------------------|----------------|
| **List comprehension** | `[d for d in os.listdir(ruta) if os.path.isdir(...)]` | Crear listas de forma concisa |
| **Diccionarios** | `{"modelo": "Corolla", "version": "XRS"}` | Almacenar datos con clave-valor |
| **Funciones con parámetros por defecto** | `def estilo_dato(cell, alt=False)` | Hacer código reutilizable |
| **Métodos de strings** | `.replace()`, `.join()`, `.format()` | Manipular texto |
| **Operador ternario** | `"center" if center else "left"` | Condicionales en una línea |
| **F-strings** | `f"[OK] {vehiculo} | {modelo}"` | Interpolación de variables |
| **Rutas con os.path** | `os.path.join()`, `os.path.isdir()` | Manejo de archivos multiplataforma |
| **POO con openpyxl** | `cell.font`, `ws.merge_cells()` | Manipular Excel |
| **Bucle anidado** | `for vehiculo > for modelo > for version` | Recorrer estructuras jerárquicas |
| **Programación defensiva** | `if not os.path.exists(...): return` | Validar antes de procesar |

---

## 💪 Ejercicios para Practicar

### Ejercicio 1: Modifica el script
Cambia el color de las filas alternadas a verde en vez de azul.

### Ejercicio 2: Agrega una función
Crea una función `contar_archivos(ruta_arte)` que cuente cuántos archivos hay dentro de cada carpeta ARTES.

### Ejercicio 3: Mejora el código
Agrega validación para cuando `resultado` está vacío (no mostró nada).

### Ejercicio 4: Experimenta
Cambia la `RUTA_BASE` a una carpeta local y prueba el script.

---

## 🎓 Consejos Finales para Ser Mejor Programador

1. **Lee código de otros**: Este script es buen ejemplo de código limpio
2. **Practica jeden día**: La programación es como el gimnasio
3. **Comenta tu código**: Like that makes el código más entendible
4. **Divide y vencerás**: Funciones pequeñas son mejores que funciones gigantes
5. **Usa nombres descriptivos**: `resultado` > `r`, `escanear` > `fn1`
6. **Aprende a debuguear**: `print()` es tu mejor amigo para ver qué pasa
7. **No temas equivocarte**: Los errores son la mejor forma de aprender

---

¡Espero que esta guía te ayude a entender mejor el código y a mejorar como programador, bro! 🚀

**Nota:** Este script está muy bien escrito y sigue las mejores prácticas de Python. Es un excelente ejemplo para aprender. ¡Felicitaciones al que lo escribió! 🎉
