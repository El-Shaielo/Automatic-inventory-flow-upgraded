"""Interfaz web para generar el reporte de flujo de inventario."""

import argparse
from io import BytesIO
from pathlib import Path
import re
import unicodedata
from urllib.parse import quote
from zipfile import BadZipFile

import numpy as np
import pandas as pd
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, StreamingResponse
from openpyxl.utils.exceptions import InvalidFileException


app = FastAPI(title="Reporte de flujo de inventario")
HTML_PATH = Path(__file__).resolve().parent / "index.html"
EXTENSIONES_EXCEL_PERMITIDAS = {".xlsx", ".xlsm"}
TIPO_EXCEL = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

MESES_ESPANOL = {
    1: "Enero",
    2: "Febrero",
    3: "Marzo",
    4: "Abril",
    5: "Mayo",
    6: "Junio",
    7: "Julio",
    8: "Agosto",
    9: "Septiembre",
    10: "Octubre",
    11: "Noviembre",
    12: "Diciembre",
}

COLUMNAS_REQUERIDAS = {
    "Fecha",
    "Producto",
    "Referencia del Item",
    "Transaccion",
    "Cantidad",
    "Notas",
}


def nombre_archivo_reporte(marca):
    marca_segura = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", marca).strip(" .")
    if not marca_segura:
        marca_segura = "Marca"
    return f"Flujo de Inventario {marca_segura}.xlsx"


def normalizar_texto(valor):
    """Normaliza texto para clasificar notas con o sin tildes."""
    if pd.isna(valor):
        return ""

    texto = unicodedata.normalize("NFKD", str(valor).casefold())
    return "".join(
        caracter
        for caracter in texto
        if not unicodedata.combining(caracter) and caracter != "\ufffd"
    ).strip()


def cargar_movimientos(ruta_entrada, referencia_item):
    df_inicial = pd.read_excel(ruta_entrada)
    columnas_faltantes = COLUMNAS_REQUERIDAS.difference(df_inicial.columns)
    if columnas_faltantes:
        raise ValueError(
            "Al archivo de entrada le faltan columnas requeridas: "
            + ", ".join(sorted(columnas_faltantes))
        )

    referencia_normalizada = normalizar_texto(referencia_item)
    referencias = df_inicial["Referencia del Item"].map(normalizar_texto)
    df = df_inicial.loc[referencias == referencia_normalizada].copy()
    if df.empty:
        raise ValueError(
            f"No hay movimientos para la referencia del item {referencia_item!r}."
        )

    df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce")
    fechas_invalidas = df["Fecha"].isna()
    if fechas_invalidas.any():
        raise ValueError(
            f"Hay {int(fechas_invalidas.sum())} movimientos con una fecha vacía o inválida."
        )

    cantidad_original = df["Cantidad"].astype("string")
    cantidad_limpia = cantidad_original.str.replace(
        r"\s*-\s*Unidades\s*$", "", case=False, regex=True
    ).str.strip()
    df["Cantidad"] = pd.to_numeric(cantidad_limpia, errors="coerce")
    cantidades_invalidas = cantidad_original.notna() & df["Cantidad"].isna()
    if cantidades_invalidas.any():
        raise ValueError(
            f"Hay {int(cantidades_invalidas.sum())} movimientos con una cantidad inválida."
        )
    df["Cantidad"] = df["Cantidad"].fillna(0)

    df["_transaccion_normalizada"] = df["Transaccion"].map(normalizar_texto)
    transacciones_no_reconocidas = sorted(
        valor
        for valor in df.loc[
            ~df["_transaccion_normalizada"].isin({"entrada", "salida"}),
            "Transaccion",
        ].dropna().unique()
    )
    if transacciones_no_reconocidas:
        raise ValueError(
            "Se encontraron transacciones distintas de Entrada/Salida: "
            + ", ".join(map(str, transacciones_no_reconocidas))
        )

    df["_periodo"] = df["Fecha"].dt.to_period("M")
    df["Mes"] = df["Fecha"].dt.month.map(MESES_ESPANOL)
    df["_nota_normalizada"] = df["Notas"].map(normalizar_texto)
    return df.sort_values("Fecha")


def crear_secciones(df):
    """Separa primero por transacción y luego clasifica las notas de cada grupo."""
    entradas = df.loc[df["_transaccion_normalizada"] == "entrada"].copy()
    salidas = df.loc[df["_transaccion_normalizada"] == "salida"].copy()

    return [
        ("Entradas (general)", entradas),
        (
            "Entradas (entrada)",
            entradas.loc[
                entradas["_nota_normalizada"].str.startswith("entrada")
            ].copy(),
        ),
        (
            "Entradas (remisión)",
            entradas.loc[
                entradas["_nota_normalizada"].str.startswith("remis")
            ].copy(),
        ),
        (
            "Entradas (nuevo ítem)",
            entradas.loc[
                entradas["_nota_normalizada"].str.startswith("nuevo")
            ].copy(),
        ),
        (
            "Entradas (actualización)",
            entradas.loc[
                entradas["_nota_normalizada"].str.startswith("actualiza")
            ].copy(),
        ),
        ("Salidas (general)", salidas),
        (
            "Salidas (facturas)",
            salidas.loc[
                salidas["_nota_normalizada"].str.startswith("fac")
                | salidas["_nota_normalizada"].str.contains(
                    "comprob", regex=False
                )
            ].copy(),
        ),
        (
            "Salidas (factura de venta)",
            salidas.loc[
                salidas["_nota_normalizada"].str.startswith("factura de venta")
            ].copy(),
        ),
        (
            "Salidas (facturación manual)",
            salidas.loc[
                salidas["_nota_normalizada"].str.startswith(
                    "facturacion manual"
                )
            ].copy(),
        ),
        (
            "Salidas (salida)",
            salidas.loc[
                salidas["_nota_normalizada"].str.startswith("sali")
            ].copy(),
        ),
        (
            "Salidas (comprobante)",
            salidas.loc[
                salidas["_nota_normalizada"].str.contains(
                    "comprob", regex=False
                )
            ].copy(),
        ),
        (
            "Salidas (devolución)",
            salidas.loc[
                salidas["_nota_normalizada"].str.startswith("devol")
            ].copy(),
        ),
    ]


def preparar_tabla(seccion, nombre_columna, columnas_productos):
    """Agrupa por nota, producto y mes, conservando la fecha del movimiento."""
    registros = seccion.copy()
    registros["Notas"] = registros["Notas"].fillna("")

    cantidades = (
        registros.groupby(
            ["_periodo", "Notas", "Producto"], dropna=False, sort=True
        )["Cantidad"]
        .agg(np.sum)
        .unstack("Producto", fill_value=0)
    )
    cantidades.columns.name = None
    cantidades = cantidades.reindex(columns=columnas_productos, fill_value=0)

    fechas = (
        registros.groupby(["_periodo", "Notas"], dropna=False, sort=True)["Fecha"]
        .min()
        .rename("Fecha")
    )
    tabla = cantidades.join(fechas).reset_index()
    tabla["Mes"] = tabla["Fecha"].dt.month.map(MESES_ESPANOL)
    tabla["Fecha"] = tabla["Fecha"].dt.strftime("%Y-%m-%d %H:%M:%S")
    tabla = tabla.rename(columns={"Notas": nombre_columna})

    columnas = ["_periodo", nombre_columna, *columnas_productos, "Fecha", "Mes"]
    return tabla[columnas], columnas_productos


def escribir_reporte(secciones, ruta_salida):
    secciones_generales = dict(secciones)
    movimientos_generales = pd.concat(
        [
            secciones_generales["Entradas (general)"],
            secciones_generales["Salidas (general)"],
        ],
        ignore_index=True,
    )
    columnas_productos = sorted(
        movimientos_generales["Producto"].dropna().unique()
    )

    with pd.ExcelWriter(ruta_salida, engine="xlsxwriter") as writer:
        fila_completo = 0
        fila_limpio = 0
        formato_encabezado = writer.book.add_format(
            {
                "bold": True,
                "border": 1,
                "align": "center",
                "valign": "top",
            }
        )

        def escribir_tabla_con_encabezado(tabla, hoja, fila):
            hoja_excel = writer.sheets[hoja]
            hoja_excel.write_row(
                fila,
                0,
                tabla.columns.tolist(),
                formato_encabezado,
            )
            tabla.to_excel(
                writer,
                sheet_name=hoja,
                startrow=fila + 1,
                index=False,
                header=False,
            )

        for nombre_columna, seccion in secciones:
            if seccion.empty:
                continue

            tabla, _ = preparar_tabla(
                seccion, nombre_columna, columnas_productos
            )
            if not columnas_productos:
                continue

            for _, grupo in tabla.groupby("_periodo", sort=True):
                detalle = grupo.drop(columns="_periodo").copy()
                total = {
                    columna: detalle[columna].sum()
                    for columna in columnas_productos
                }
                total[nombre_columna] = "Total"
                total["Fecha"] = ""
                total["Mes"] = ""
                detalle = pd.concat(
                    [detalle, pd.DataFrame([total], columns=detalle.columns)],
                    ignore_index=True,
                ).replace({np.nan: ""})

                titulo_mes = pd.DataFrame([[f"Mes: {grupo['Mes'].iloc[0]}"]])
                titulo_mes.to_excel(
                    writer,
                    sheet_name="COMPLETO",
                    startrow=fila_completo,
                    index=False,
                    header=False,
                )
                escribir_tabla_con_encabezado(
                    detalle, "COMPLETO", fila_completo + 1
                )
                fila_completo += len(detalle) + 4

                total_limpio = pd.DataFrame(
                    [
                        {
                            nombre_columna: "Total",
                            **{
                                columna: grupo[columna].sum()
                                for columna in columnas_productos
                            },
                        }
                    ]
                )
                total_limpio = total_limpio[
                    [nombre_columna, *columnas_productos]
                ]

                titulo_mes.to_excel(
                    writer,
                    sheet_name="LIMPIO",
                    startrow=fila_limpio,
                    index=False,
                    header=False,
                )
                escribir_tabla_con_encabezado(
                    total_limpio, "LIMPIO", fila_limpio + 1
                )
                fila_limpio += len(total_limpio) + 3

            totales_anuales = pd.DataFrame(
                [
                    {
                        nombre_columna: "Total",
                        **{
                            columna: tabla[columna].sum()
                            for columna in columnas_productos
                        },
                    }
                ]
            )
            totales_anuales = totales_anuales[
                [nombre_columna, *columnas_productos]
            ]

            formato_negrita = writer.book.add_format(
                {"bold": True, "font_size": 11}
            )
            writer.sheets["LIMPIO"].write(
                fila_limpio, 0, "Todos los meses", formato_negrita
            )
            escribir_tabla_con_encabezado(
                totales_anuales, "LIMPIO", fila_limpio + 1
            )
            fila_limpio += len(totales_anuales) + 5


@app.get("/", include_in_schema=False)
def inicio():
    return FileResponse(HTML_PATH)


@app.post("/generar")
def generar(
    archivo: UploadFile = File(...),
    marca: str = Form(...),
):
    extension = Path(archivo.filename or "").suffix.lower()
    if extension not in EXTENSIONES_EXCEL_PERMITIDAS:
        raise HTTPException(
            status_code=400,
            detail="Selecciona un archivo Excel con extensión .xlsx o .xlsm.",
        )

    marca = marca.strip()
    if not marca:
        raise HTTPException(
            status_code=400,
            detail="Escribe la marca o referencia del ítem que deseas reportar.",
        )

    try:
        movimientos = cargar_movimientos(archivo.file, marca)
    except (ValueError, BadZipFile, InvalidFileException) as error:
        raise HTTPException(status_code=400, detail=str(error)) from error

    reporte = BytesIO()
    escribir_reporte(crear_secciones(movimientos), reporte)
    reporte.seek(0)
    nombre_archivo = nombre_archivo_reporte(marca)
    return StreamingResponse(
        reporte,
        media_type=TIPO_EXCEL,
        headers={
            "Content-Disposition": (
                f"attachment; filename*=UTF-8''{quote(nombre_archivo)}"
            )
        },
    )


def main():
    parser = argparse.ArgumentParser(
        description="Genera el reporte mensual de flujo de inventario."
    )
    carpeta_script = Path(__file__).resolve().parent
    parser.add_argument(
        "--input",
        type=Path,
        default=carpeta_script / "Flujo de inventario (1).xlsx",
        help="Ruta del archivo Excel de entrada.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        help="Ruta del archivo Excel de salida (por defecto usa la marca).",
    )
    parser.add_argument(
        "--referencia-item",
        default="LECAT",
        help="Referencia del item que se incluirá en el reporte.",
    )
    argumentos = parser.parse_args()

    if not argumentos.input.is_file():
        raise FileNotFoundError(
            f"No se encontró el archivo de entrada: {argumentos.input}"
        )

    movimientos = cargar_movimientos(
        argumentos.input, argumentos.referencia_item
    )
    secciones = crear_secciones(movimientos)
    ruta_salida = argumentos.output or (
        carpeta_script / nombre_archivo_reporte(argumentos.referencia_item)
    )
    escribir_reporte(secciones, ruta_salida)

    cantidad_entradas = int(
        movimientos["_transaccion_normalizada"].eq("entrada").sum()
    )
    cantidad_salidas = int(
        movimientos["_transaccion_normalizada"].eq("salida").sum()
    )
    print(
        f"Reporte generado: {ruta_salida} "
        f"({cantidad_entradas} entradas, {cantidad_salidas} salidas)."
    )


if __name__ == "__main__":
    main()
