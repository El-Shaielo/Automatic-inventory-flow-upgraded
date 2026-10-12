# Reporte web de flujo de inventario

El script `automatizar_flujo_de_inventario__oficial_mejorado.py` reúne la
aplicación web FastAPI y toda la lógica de generación del reporte. También se
puede ejecutar directamente desde la consola.

## Ejecución en Windows

Desde la carpeta del proyecto, instala las dependencias y levanta el servidor:

```powershell
.\venv\Scripts\python.exe -m pip install -r requirements.txt
.\venv\Scripts\python.exe -m uvicorn automatizar_flujo_de_inventario__oficial_mejorado:app --reload
```

Abre <http://127.0.0.1:8000> en el navegador.

Para generar el reporte desde la consola:

```powershell
.\venv\Scripts\python.exe .\automatizar_flujo_de_inventario__oficial_mejorado.py --input ".\Flujo de inventario (1).xlsx" --output ".\Reporte_Inventario.xlsx" --referencia-item "LECAT"
```

## Archivo de entrada

Se aceptan archivos `.xlsx` y `.xlsm`. El Excel debe incluir las columnas
`Fecha`, `Producto`, `Referencia del Item`, `Transaccion`, `Cantidad` y
`Notas`. La marca indicada en el formulario se compara con `Referencia del
Item`; el reporte descargado conserva las hojas `COMPLETO` y `LIMPIO` y se
guarda como `Flujo de Inventario <marca>.xlsx`.
