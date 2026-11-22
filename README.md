## What is this project?

This project is an automation tool that processes inventory movements recorded
in an Excel file and generates a report organized by entries, exits, products,
and months.

It includes a simple web interface and can also run from the command line.

## What is it used for?

The program allows you to:

- Filter movements by brand or item reference.
- Classify entries and exits by type.
- Consolidate quantities by product and month.
- Show monthly and overall totals.
- Download an Excel file with the **COMPLETO** and **LIMPIO** worksheets.
- Validate that the file contains the required columns and data.

## Requirements

- Windows.
- Python 3.
- The dependencies listed in `requirements.txt`.

Main dependencies:

- FastAPI and Uvicorn for the web application.
- Pandas and NumPy for processing the movements.
- OpenPyXL and XlsxWriter for reading and creating Excel files.
- Python-multipart for receiving forms and uploaded files.

## Installation

1. Open PowerShell or Command Prompt.
2. Navigate to the project folder:

```powershell
cd "C:\path\to\project"
```

3. Create a virtual environment:

```powershell
py -m venv venv
```

4. Activate the virtual environment:

```powershell
.\venv\Scripts\Activate.ps1
```

If PowerShell prevents the activation script from running, you can activate the
environment from Command Prompt instead:

```bat
venv\Scripts\activate.bat
```

5. Install the dependencies:

```powershell
python -m pip install -r requirements.txt
```

## Using the web interface

1. With the virtual environment active, start the server:

```powershell
python -m uvicorn automatizar_flujo_de_inventario__oficial_mejorado:app --reload
```

2. Open a browser and visit:

```text
http://127.0.0.1:8000
```

3. Select an Excel file with a `.xlsx` or `.xlsm` extension.
4. Enter the brand or item reference that matches the `Referencia del Item`
   column.
5. Click **Generar y descargar reporte**.

The browser will automatically download:
`Flujo de Inventario <reference>.xlsx`.



## Input file format

The file must be an Excel workbook (`.xlsx` or `.xlsm`) and contain the
following columns:

| Column | Use |
|---|---|
| `Fecha` | Movement date |
| `Producto` | Product name or description |
| `Referencia del Item` | Reference used to filter movements |
| `Transaccion` | Must indicate `Entrada` or `Salida` |
| `Cantidad` | Moved quantity |
| `Notas` | Movement type or description |

Empty or invalid dates, invalid quantities, and transactions other than
`Entrada` or `Salida` generate an error message.

## Output

The generated report contains:

- **COMPLETO:** Monthly detail grouped by notes and product, including totals.
- **LIMPIO:** Monthly and overall totals by product.

The notes classification distinguishes categories including invoices,
receipts, returns, remittances, new items, and updates. Accents and uppercase
letters do not prevent category recognition.

## Stop the server

To stop the web application, return to the PowerShell window running Uvicorn
and press `Ctrl+C`.
