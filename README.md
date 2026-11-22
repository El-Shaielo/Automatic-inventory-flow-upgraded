## ¿Qué es este proyecto?

Este proyecto es una herramienta de automatización para procesar movimientos de
inventario registrados en un archivo de Excel y generar un reporte organizado
por entradas, salidas, productos y meses.

Incluye una interfaz web sencilla y también puede ejecutarse desde la línea de
comandos.

## ¿Para qué sirve?

El programa permite:

- Filtrar los movimientos de una marca o referencia de ítem.
- Clasificar las entradas y salidas según su tipo.
- Consolidar las cantidades por producto y mes.
- Mostrar totales mensuales y totales generales.
- Descargar un archivo Excel con las hojas **COMPLETO** y **LIMPIO**.
- Validar que el archivo tenga las columnas y los datos requeridos.

## Requisitos

- Windows.
- Python 3.
- Las dependencias indicadas en `requirements.txt`.

Dependencias principales:

- FastAPI y Uvicorn para la aplicación web.
- Pandas y NumPy para procesar los movimientos.
- OpenPyXL y XlsxWriter para leer y crear archivos Excel.
- Python-multipart para recibir formularios y archivos.

## Instalación

1. Abre PowerShell o el Símbolo del sistema.
2. Ubícate en la carpeta del proyecto:

```powershell
cd "C:\ruta\al\proyecto"
```

3. Crea un entorno virtual:

```powershell
py -m venv venv
```

4. Activa el entorno virtual:

```powershell
.\venv\Scripts\Activate.ps1
```

Si PowerShell impide ejecutar el script de activación, también puedes activar el
entorno desde el Símbolo del sistema:

```bat
venv\Scripts\activate.bat
```

5. Instala las dependencias:

```powershell
python -m pip install -r requirements.txt
```

## Uso desde la interfaz web

1. Con el entorno virtual activo, inicia el servidor:

```powershell
python -m uvicorn automatizar_flujo_de_inventario__oficial_mejorado:app --reload
```

2. Abre un navegador y visita:

```text
http://127.0.0.1:8000
```

3. Selecciona un archivo Excel con extensión `.xlsx` o `.xlsm`.
4. Escribe la marca o referencia del ítem que debe coincidir con la columna
   `Referencia del Item`.
5. Pulsa **Generar y descargar reporte**.

El navegador descargará automáticamente el archivo
`Flujo de Inventario <referencia>.xlsx`.

## Uso desde la línea de comandos

Para usar el archivo de entrada predeterminado
`Flujo de inventario (1).xlsx` y la referencia `LECAT`, ejecuta:

```powershell
python automatizar_flujo_de_inventario__oficial_mejorado.py
```

Para indicar otro archivo, referencia o ubicación de salida:

```powershell
python automatizar_flujo_de_inventario__oficial_mejorado.py `
  --input "ruta\archivo.xlsx" `
  --referencia-item "LECAT" `
  --output "ruta\reporte.xlsx"
```

Argumentos disponibles:

| Argumento | Descripción | Valor predeterminado |
|---|---|---|
| `--input` | Ruta del archivo Excel de entrada | `Flujo de inventario (1).xlsx` |
| `--referencia-item` | Referencia que se procesará | `LECAT` |
| `--output` | Ruta del archivo Excel generado | `Flujo de Inventario <referencia>.xlsx` |

## Formato del archivo de entrada

El archivo debe ser Excel (`.xlsx` o `.xlsm`) y contener las siguientes
columnas:

| Columna | Uso |
|---|---|
| `Fecha` | Fecha del movimiento |
| `Producto` | Nombre o descripción del producto |
| `Referencia del Item` | Referencia usada para filtrar los movimientos |
| `Transaccion` | Debe indicar `Entrada` o `Salida` |
| `Cantidad` | Cantidad movida |
| `Notas` | Tipo o descripción del movimiento |

Las fechas vacías o inválidas, las cantidades inválidas y las transacciones
distintas de `Entrada` o `Salida` generan un mensaje de error.

## Resultado

El reporte generado contiene:

- **COMPLETO:** detalle mensual, agrupado por notas y producto, con totales.
- **LIMPIO:** totales mensuales y totales generales por producto.

La clasificación de notas distingue, entre otras categorías, facturas,
comprobantes, devoluciones, remisiones, nuevos ítems y actualizaciones. Las
tildes y mayúsculas no impiden el reconocimiento de las categorías.

## Detener el servidor

Para detener la aplicación web, vuelve a la ventana de PowerShell donde está
ejecutándose Uvicorn y pulsa `Ctrl+C`.
