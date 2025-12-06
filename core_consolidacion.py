# core_consolidacion.py

import os
import re
from typing import List, Dict, Any, Optional, Tuple

import pandas as pd

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule

COLUMNA_CANTIDAD = "Cantidad"

COLUMNAS_DESCRIPTIVAS = ["Marca", "Descripcion", "Departamento", "SubFamilia", "Familia"]

# Prioridades base (se pueden sobreescribir desde la GUI)
PRIORIDAD_DEPARTAMENTOS_DEFAULT: Dict[str, int] = {
    "BEBIDAS SIN ALCOHOL": 100,
    "ADITIVOS PARA LAVADOS": 100,
    "ALMACEN": 100,
    "ACEITES": 90,
    "BEBIDAS": 50,
    "LIMPIEZA Y CUIDADO PERSONAL": 30,
    "LIMPIEZA Y CUIDADO": 30,
    "DESAYUNO": 40,
    "ARROZ": 30,
    "ENLATADOS": 30,
    "ALIM VARIOS": 30,
}

CATEGORIAS_ESPECIALES: List[str] = [
    "ELECTRO",
    "ELECTRODOMESTICOS",
    "FERRETERIA",
    "RODADOS",
]

MESES_ES: List[str] = [
    "ENERO",
    "FEBRERO",
    "MARZO",
    "ABRIL",
    "MAYO",
    "JUNIO",
    "JULIO",
    "AGOSTO",
    "SEPTIEMBRE",
    "OCTUBRE",
    "NOVIEMBRE",
    "DICIEMBRE",
]

MAPA_MES: Dict[str, int] = {m: i + 1 for i, m in enumerate(MESES_ES)}

# --------- Explicaciones para fila 2 en cada hoja --------- #

EXPLICACIONES_CONSOLIDADO: Dict[str, str] = {
    "IdArticulo": "Código único del producto",
    "Marca": "Nombre del fabricante",
    "Descripcion": "Nombre comercial del producto",
    "Departamento": "Categoría principal de venta",
    "SubFamilia": "Subcategoría del producto",
    "Familia": "Clasificación comercial",
    "TOTAL CORRIENTES": "Unidades vendidas totales en sucursal Corrientes",
    "TOTAL HIPER": "Unidades vendidas totales en sucursal Hipermercado",
    "TOTAL CONSOLIDADO": "Sumatoria Corrientes + Hipermercado (período completo)",
}

EXPLICACIONES_RANKING: Dict[str, str] = {
    "IdArticulo": "Código único del producto",
    "Marca": "Nombre del fabricante",
    "Descripcion": "Nombre comercial del producto",
    "Total Vendido": "Unidades totales vendidas (todas las sucursales y meses)",
}

EXPLICACIONES_POR_SUCURSAL: Dict[str, str] = {
    "SUCURSAL": "Nombre de la sucursal",
    "TOTAL": "Unidades vendidas en todo el período",
}

EXPLICACIONES_MATRIZ_DEPTO: Dict[str, str] = {
    "Departamento": "Categoría principal de venta",
    "CORRIENTES": "Unidades del departamento en Corrientes (período)",
    "HIPER": "Unidades del departamento en Hipermercado (período)",
    "TOTAL": "Sumatoria Corrientes + Hiper (período)",
}

EXPLICACIONES_EVOLUCION: Dict[str, str] = {
    "Departamento": "Categoría principal de venta",
}

EXPLICACIONES_ESTACIONALIDAD: Dict[str, str] = {
    "IdArticulo": "Código único del producto",
    "Marca": "Nombre del fabricante",
    "Descripcion": "Nombre comercial del producto",
    "Departamento": "Categoría principal de venta",
    "SubFamilia": "Subcategoría del producto",
    "Familia": "Clasificación comercial",
    "VENTA_TOTAL_ANUAL": "Ventas totales en el período",
    "PROMEDIO_MENSUAL": "Venta promedio por mes",
    "MES_PICO": "Mes con mayor venta",
    "VENTA_PICO": "Unidades en mes pico",
    "INDICE_PICO": "Mes pico ÷ promedio (%)",
    # Los índices por mes se generan dinámicamente: "MES ÷ promedio (%)"
}

# ---------------------------------------------------------- #


def normalizar_mes(mes_str: str) -> str:
    mes_up = mes_str.strip().upper()
    for m in MESES_ES:
        if m in mes_up:
            return m
    raise ValueError(f"No se pudo normalizar mes desde: {mes_str}")


def parsear_nombre_archivo(nombre: str) -> Optional[Tuple[str, int, Optional[str]]]:
    """
    Devuelve (mes_str, anio_int, sucursal_str_o_None) o None si no se puede parsear.
    Espera algo tipo: '3. MARZO 2025 CORRIENTES.xlsx'
    """
    base = os.path.splitext(os.path.basename(nombre))[0]
    base_up = base.upper().replace("  ", " ")

    # Buscar sucursal conocida
    sucursal: Optional[str] = None
    for s in ["HIPER", "CORRIENTES"]:
        if s in base_up:
            sucursal = s
            base_up = base_up.replace(s, "").strip()
            break

    # Mes y año
    m = re.search(r"(\d{1,2})\s*\.?\s+([A-ZÁÉÍÓÚÑ ]+)\s+(\d{4})", base_up)
    if not m:
        return None

    _, mes_txt, anio_txt = m.groups()
    try:
        mes_norm = normalizar_mes(mes_txt)
    except ValueError:
        return None

    anio = int(anio_txt)
    return mes_norm, anio, sucursal


def consolidar_datos(
    archivos_info: List[Dict[str, Any]],
    prioridades_depto: Optional[Dict[str, int]] = None,
) -> pd.DataFrame:
    """
    archivos_info: lista de diccionarios:
        "ruta": str,
        "mes": "MARZO",
        "anio": 2025,
        "sucursal": "HIPER"

    prioridades_depto: dict opcional para sobreescribir PRIORIDAD_DEPARTAMENTOS_DEFAULT
    """
    if prioridades_depto is None:
        prioridades = PRIORIDAD_DEPARTAMENTOS_DEFAULT.copy()
    else:
        prioridades = prioridades_depto

    todos: List[pd.DataFrame] = []

    for info in archivos_info:
        ruta = info["ruta"]
        mes = info["mes"]
        anio = info["anio"]
        sucursal = info["sucursal"]

        if not os.path.exists(ruta):
            continue

        df = pd.read_excel(ruta, sheet_name=0)
        if COLUMNA_CANTIDAD not in df.columns or "IdArticulo" not in df.columns:
            continue

        columnas_existentes = [c for c in COLUMNAS_DESCRIPTIVAS if c in df.columns]
        columnas_a_usar = ["IdArticulo"] + columnas_existentes + [COLUMNA_CANTIDAD]
        df_f = df[columnas_a_usar].copy()

        df_f["MES"] = f"{mes} {anio}"
        df_f["SUCURSAL"] = sucursal

        # Limpieza básica
        for col in COLUMNAS_DESCRIPTIVAS:
            if col in df_f.columns:
                df_f[col] = df_f[col].astype(str).str.strip()

        # Normalización de Departamento (se puede extender)
        if "Departamento" in df_f.columns:
            df_f["Departamento"] = df_f["Departamento"].str.upper()
            df_f.loc[df_f["Departamento"] == "ACEITES", "Departamento"] = "ALMACEN"
            df_f.loc[
                df_f["Departamento"] == "HIGIENE PERSONAL", "Departamento"
            ] = "LIMPIEZA Y CUIDADO"

        todos.append(df_f)

    if not todos:
        raise ValueError("No se pudo leer ningún archivo válido.")

    df_all = pd.concat(todos, ignore_index=True)

    # Consolidar por prioridad de departamento
    df_all["PRIORIDAD"] = df_all["Departamento"].map(prioridades).fillna(0)
    df_all["CLAVE_PRODUCTO"] = (
        df_all["IdArticulo"].astype(str)
        + "|"
        + df_all.get("Marca", "").astype(str)
        + "|"
        + df_all.get("Descripcion", "").astype(str)
        + "|"
        + df_all.get("SubFamilia", "").astype(str)
        + "|"
        + df_all.get("Familia", "").astype(str)
    )

    idx_max = df_all.groupby("CLAVE_PRODUCTO")["PRIORIDAD"].idxmax()
    dept_final = df_all.loc[idx_max, ["CLAVE_PRODUCTO", "Departamento"]].drop_duplicates(
        "CLAVE_PRODUCTO"
    )

    df_all = df_all.drop(columns=["Departamento"])
    df_all = df_all.merge(
        dept_final.rename(columns={"Departamento": "Departamento"}),
        on="CLAVE_PRODUCTO",
        how="left",
    )

    df_all = df_all.drop(columns=["PRIORIDAD", "CLAVE_PRODUCTO"])

    # Agrupar final
    df_all = df_all.groupby(
        [
            "IdArticulo",
            "Marca",
            "Descripcion",
            "Departamento",
            "SubFamilia",
            "Familia",
            "MES",
            "SUCURSAL",
        ],
        as_index=False,
    )[COLUMNA_CANTIDAD].sum()

    # Redondeo
    df_all[COLUMNA_CANTIDAD] = df_all[COLUMNA_CANTIDAD].apply(
        lambda x: int(x + 0.5) if x >= 0 else int(x - 0.5)
    )

    return df_all


def _orden_mes_clave(mes_ano: str) -> Tuple[int, int]:
    # mes_ano: "MARZO 2025"
    partes = mes_ano.split()
    if len(partes) != 2:
        return (9999, 99)
    mes_txt, anio_txt = partes
    anio = int(anio_txt)
    mes_num = MAPA_MES.get(mes_txt, 99)
    return (anio, mes_num)


# --------- Helpers de estilo para Excel --------- #


def _col_letra(idx: int) -> str:
    """Convierte índice de columna (1-based) a letra estilo Excel."""
    result = ""
    while idx:
        idx, rem = divmod(idx - 1, 26)
        result = chr(65 + rem) + result
    return result


def _aplicar_estilos_fila2_generica(
    ws, encabezados: List[str], explicaciones_dict: Dict[str, str], num_filas_datos: int
) -> None:
    """
    Aplica estilos a la fila 2 (explicaciones) para cualquier hoja:
    - Fondo gris claro
    - Texto itálico tamaño 9
    - Bordes finos
    - Alineación centrada con wrap_text
    """
    fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
    font = Font(italic=True, size=9)
    alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    for col_idx, encabezado in enumerate(encabezados, start=1):
        cell = ws.cell(row=2, column=col_idx)

        # 1) Explicación fija por nombre exacto
        valor = explicaciones_dict.get(encabezado)

        # 2) Patrones dinámicos

        # Consolidado / Categorías Especiales: "MARZO 2025_CORRIENTES"
        if valor is None and "_" in encabezado and " " in encabezado:
            mes_anio, suc = encabezado.split("_", 1)
            valor = f"Unidades {mes_anio} en {suc}"

        # Por Sucursal / Evolución: columnas de meses "MARZO 2025"
        if valor is None and any(m in encabezado for m in MESES_ES):
            if encabezados and encabezados[0] == "Departamento":
                valor = f"Unidades del departamento en {encabezado}"
            else:
                valor = f"Unidades vendidas en {encabezado}"

        # Estacionalidad: índices por mes "INDICE_MARZO 2025"
        if (
            valor is None
            and encabezado.upper().startswith("INDICE_")
            and encabezado != "INDICE_PICO"
        ):
            mes_txt = encabezado.replace("INDICE_", "")
            valor = f"{mes_txt} ÷ promedio (%)"

        if valor is None:
            valor = encabezado

        cell.value = valor
        cell.fill = fill
        cell.font = font
        cell.alignment = alignment
        cell.border = border

    # Alto de la fila 2
    ws.row_dimensions[2].height = 30

    # Bordes y alineación en filas de datos (3 en adelante)
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    for row in range(3, num_filas_datos + 3):
        for col in range(1, len(encabezados) + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = thin_border

            # Heurística simple: primeras columnas texto, resto números
            if col <= 6 or isinstance(cell.value, str):
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="right", vertical="center")


def _aplicar_estilos_estacionalidad(ws, num_filas_datos: int) -> None:
    """
    Estilos específicos para la hoja Estacionalidad:
    - Fila 2 con explicaciones
    - Formato porcentaje 0.0% en todas las columnas INDICE_*
    - Formato condicional verde/amarillo/rojo en INDICE_PICO
    """
    encabezados = [c.value for c in ws[1]]
    _aplicar_estilos_fila2_generica(
        ws, encabezados, EXPLICACIONES_ESTACIONALIDAD, num_filas_datos
    )

    # Columnas de índices (INDICE_PICO + INDICE_MES...)
    col_indices: List[int] = []
    col_pico_idx: Optional[int] = None

    for idx, nombre in enumerate(encabezados, start=1):
        if not nombre:
            continue
        if nombre.startswith("INDICE_"):
            col_indices.append(idx)
        if nombre == "INDICE_PICO":
            col_pico_idx = idx

    # Formato porcentaje 0.0% en todas las columnas de índice
    for col_idx in col_indices:
        for row in range(3, num_filas_datos + 3):
            cell = ws.cell(row=row, column=col_idx)
            cell.number_format = "0.0%"

    # Formato condicional solo sobre INDICE_PICO
    if col_pico_idx is None:
        return

    col_letra = _col_letra(col_pico_idx)
    rango = f"{col_letra}3:{col_letra}{num_filas_datos + 2}"

    verde_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    verde_font = Font(bold=True, color="006100")
    regla_verde = CellIsRule(
        operator="greaterThan", formula=["1.10"], fill=verde_fill, font=verde_font
    )

    amarillo_fill = PatternFill(
        start_color="FFEB9C", end_color="FFEB9C", fill_type="solid"
    )
    amarillo_font = Font(bold=True, color="9C6500")
    regla_amarillo = CellIsRule(
        operator="between", formula=["0.90", "1.10"], fill=amarillo_fill, font=amarillo_font
    )

    rojo_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    rojo_font = Font(bold=True, color="9C0006")
    regla_rojo = CellIsRule(
        operator="lessThan", formula=["0.90"], fill=rojo_fill, font=rojo_font
    )

    ws.conditional_formatting.add(rango, regla_verde)
    ws.conditional_formatting.add(rango, regla_amarillo)
    ws.conditional_formatting.add(rango, regla_rojo)


# ---------------------------------------------------------- #


def generar_reportes(
    df: pd.DataFrame,
    ruta_salida: str,
    columnas_consolidado: Optional[List[str]] = None,
    habilitar_ranking: bool = True,
    habilitar_por_sucursal: bool = True,
    habilitar_matriz: bool = True,
    habilitar_evolucion: bool = True,
    habilitar_especiales: bool = True,
    filtros_especiales: Optional[Dict[str, List[str]]] = None,
    habilitar_estacionalidad: bool = False,
) -> None:
    """
    Genera reportes en un solo Excel, con opciones:
    - columnas_consolidado: nombres de columnas en el orden deseado.
    - habilitar_*: booleans para crear o no cada hoja adicional.
    - filtros_especiales: dict opcional {"departamentos": [...], "marcas": [...]}.
    - habilitar_estacionalidad: si True, genera hoja con índice estacional.
    """
    meses_ordenados = sorted(df["MES"].unique(), key=_orden_mes_clave)
    sucursales = sorted(df["SUCURSAL"].dropna().unique())

    with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
        # 1) Consolidado (siempre se genera)
        df_temp = df.copy()
        df_temp["MES_SUC"] = df_temp["MES"] + "_" + df_temp["SUCURSAL"]

        idx_cols = [
            "IdArticulo",
            "Marca",
            "Descripcion",
            "Departamento",
            "SubFamilia",
            "Familia",
        ]

        df_pivot = df_temp.pivot_table(
            index=idx_cols,
            columns="MES_SUC",
            values=COLUMNA_CANTIDAD,
            aggfunc="sum",
            fill_value=0,
        ).reset_index()

        for col in df_pivot.columns:
            if col not in idx_cols:
                df_pivot[col] = df_pivot[col].astype(int)

        cols_def = idx_cols.copy()

        # Totales por mes (sin sucursal)
        for mes in meses_ordenados:
            cols_mes = [c for c in df_pivot.columns if c.startswith(mes + "_")]
            if not cols_mes:
                continue
            df_pivot[mes] = df_pivot[cols_mes].sum(axis=1).astype(int)
            cols_def.append(mes)

        # Totales por sucursal y consolidado
        total_cols: List[str] = []
        for suc in sucursales:
            cols_suc = [c for c in df_pivot.columns if c.endswith("_" + suc)]
            if not cols_suc:
                continue
            col_total = f"TOTAL {suc.upper()}"
            df_pivot[col_total] = df_pivot[cols_suc].sum(axis=1).astype(int)
            total_cols.append(col_total)

        if total_cols:
            df_pivot["TOTAL CONSOLIDADO"] = df_pivot[total_cols].sum(axis=1).astype(int)
            cols_def.extend(total_cols + ["TOTAL CONSOLIDADO"])

        # Orden personalizado desde la GUI
        if columnas_consolidado:
            orden = [c for c in columnas_consolidado if c in df_pivot.columns]
            extras = [c for c in df_pivot.columns if c not in orden]
            cols_finales = orden + extras
        else:
            cols_finales = cols_def

        df_final = df_pivot[cols_finales].sort_values("IdArticulo").reset_index(
            drop=True
        )
        df_final.to_excel(writer, sheet_name="Consolidado", index=False)

        # 2) Ranking de Ventas
        if habilitar_ranking:
            ranking = df.groupby(["IdArticulo", "Marca", "Descripcion"]).agg(
                {COLUMNA_CANTIDAD: "sum"}
            ).reset_index()
            ranking = ranking.sort_values(
                COLUMNA_CANTIDAD, ascending=False
            ).reset_index(drop=True)
            ranking.rename(columns={COLUMNA_CANTIDAD: "Total Vendido"}, inplace=True)
            ranking.to_excel(writer, sheet_name="Ranking de Ventas", index=False)

        # 3) Por Sucursal
        if habilitar_por_sucursal:
            por_suc = df.pivot_table(
                index="SUCURSAL",
                columns="MES",
                values=COLUMNA_CANTIDAD,
                aggfunc="sum",
                fill_value=0,
            ).reset_index()
            cols_tot = [m for m in meses_ordenados if m in por_suc.columns]
            if cols_tot:
                por_suc["TOTAL"] = por_suc[cols_tot].sum(axis=1)
            por_suc.to_excel(writer, sheet_name="Por Sucursal", index=False)

        # 4) Matriz (Departamento x Sucursal)
        if habilitar_matriz:
            matriz = df.pivot_table(
                index="Departamento",
                columns="SUCURSAL",
                values=COLUMNA_CANTIDAD,
                aggfunc="sum",
                fill_value=0,
            )
            matriz["TOTAL"] = matriz.sum(axis=1)
            matriz = matriz.sort_values("TOTAL", ascending=False)
            matriz.to_excel(writer, sheet_name="Matriz")

        # 5) Evolución Mensual
        if habilitar_evolucion:
            evol = df.pivot_table(
                index="Departamento",
                columns="MES",
                values=COLUMNA_CANTIDAD,
                aggfunc="sum",
                fill_value=0,
            )
            cols_evol = [m for m in meses_ordenados if m in evol.columns]
            evol = evol[cols_evol]
            evol.to_excel(writer, sheet_name="Evolución Mensual")

        # 6) Categorías Especiales con filtros opcionales
        if habilitar_especiales:
            df_espec = df.copy()

            if filtros_especiales:
                deps = filtros_especiales.get("departamentos") or []
                marcas = filtros_especiales.get("marcas") or []
                if deps:
                    deps_up = [d.upper().strip() for d in deps]
                    df_espec = df_espec[
                        df_espec["Departamento"].str.upper().isin(deps_up)
                    ]
                if marcas:
                    marcas_up = [m.upper().strip() for m in marcas]
                    df_espec = df_espec[
                        df_espec["Marca"].str.upper().isin(marcas_up)
                    ]
            else:
                df_espec = df_espec[
                    df_espec["Departamento"].str.upper().isin(CATEGORIAS_ESPECIALES)
                ]

            if not df_espec.empty:
                tmp = df_espec.copy()
                tmp["MES_SUC"] = tmp["MES"] + "_" + tmp["SUCURSAL"]

                idx_cols_espec = [
                    "IdArticulo",
                    "Marca",
                    "Descripcion",
                    "Departamento",
                    "SubFamilia",
                    "Familia",
                ]

                piv = tmp.pivot_table(
                    index=idx_cols_espec,
                    columns="MES_SUC",
                    values=COLUMNA_CANTIDAD,
                    aggfunc="sum",
                    fill_value=0,
                ).reset_index()

                for col in piv.columns:
                    if col not in idx_cols_espec:
                        piv[col] = piv[col].astype(int)

                cols_espec_final = idx_cols_espec.copy()

                for mes in meses_ordenados:
                    cols_mes = [c for c in piv.columns if c.startswith(mes + "_")]
                    if not cols_mes:
                        continue
                    piv[mes] = piv[cols_mes].sum(axis=1).astype(int)
                    cols_espec_final.append(mes)

                total_cols_espec: List[str] = []
                for suc in sucursales:
                    cols_suc = [c for c in piv.columns if c.endswith("_" + suc)]
                    if not cols_suc:
                        continue
                    col_total = f"TOTAL {suc.upper()}"
                    piv[col_total] = piv[cols_suc].sum(axis=1).astype(int)
                    total_cols_espec.append(col_total)

                if total_cols_espec:
                    piv["TOTAL CONSOLIDADO"] = piv[total_cols_espec].sum(
                        axis=1
                    ).astype(int)
                    cols_espec_final.extend(total_cols_espec + ["TOTAL CONSOLIDADO"])

                piv_final = piv[cols_espec_final].sort_values(
                    "IdArticulo"
                ).reset_index(drop=True)

                piv_final.to_excel(
                    writer, sheet_name="Categorias Especiales", index=False
                )

        # 7) Estacionalidad (índice estacional y mes pico por producto)
        if habilitar_estacionalidad:
            df_est = df.copy()
            idx_cols_est = [
                "IdArticulo",
                "Marca",
                "Descripcion",
                "Departamento",
                "SubFamilia",
                "Familia",
            ]

            grp = df_est.groupby(idx_cols_est + ["MES"], as_index=False)[
                COLUMNA_CANTIDAD
            ].sum()

            if not grp.empty:
                piv_est = grp.pivot_table(
                    index=idx_cols_est,
                    columns="MES",
                    values=COLUMNA_CANTIDAD,
                    aggfunc="sum",
                    fill_value=0,
                ).reset_index()

                mes_cols = [m for m in meses_ordenados if m in piv_est.columns]
                if mes_cols:
                    piv_est["VENTA_TOTAL_ANUAL"] = piv_est[mes_cols].sum(axis=1)
                    n_meses = len(mes_cols)
                    piv_est["PROMEDIO_MENSUAL"] = (
                        piv_est["VENTA_TOTAL_ANUAL"] / n_meses
                    )
                    piv_est["PROMEDIO_MENSUAL"] = piv_est[
                        "PROMEDIO_MENSUAL"
                    ].replace(0, pd.NA)

                    piv_est["MES_PICO"] = piv_est[mes_cols].idxmax(axis=1)

                    def _venta_pico(row: pd.Series) -> float:
                        mes = row["MES_PICO"]
                        if pd.isna(mes):
                            return 0.0
                        return float(row.get(mes, 0.0))

                    piv_est["VENTA_PICO"] = piv_est.apply(_venta_pico, axis=1)

                    # Índice del mes pico
                    piv_est["INDICE_PICO"] = (
                        piv_est["VENTA_PICO"] / piv_est["PROMEDIO_MENSUAL"]
                    )
                    piv_est["INDICE_PICO"] = piv_est["INDICE_PICO"].fillna(0).round(3)

                    # Índices por mes (INDICE_MARZO 2025, etc.)
                    for mes in mes_cols:
                        nombre_ind = f"INDICE_{mes}"
                        piv_est[nombre_ind] = (
                            piv_est[mes] / piv_est["PROMEDIO_MENSUAL"]
                        )
                        piv_est[nombre_ind] = piv_est[nombre_ind].fillna(0).round(3)

                    cols_out = (
                        idx_cols_est
                        + [
                            "VENTA_TOTAL_ANUAL",
                            "PROMEDIO_MENSUAL",
                            "MES_PICO",
                            "VENTA_PICO",
                            "INDICE_PICO",
                        ]
                        + [f"INDICE_{m}" for m in mes_cols]
                    )

                    est_df = piv_est[cols_out]
                    est_df.to_excel(writer, sheet_name="Estacionalidad", index=False)

    # === Después de escribir el Excel: aplicar estilos con openpyxl ===
    wb = load_workbook(ruta_salida)

    # 1) Consolidado
    if "Consolidado" in wb.sheetnames:
        ws = wb["Consolidado"]
        num_filas = max(ws.max_row - 2, 0)
        encabezados = [c.value for c in ws[1]]
        _aplicar_estilos_fila2_generica(
            ws, encabezados, EXPLICACIONES_CONSOLIDADO, num_filas
        )

    # 2) Ranking de Ventas
    if "Ranking de Ventas" in wb.sheetnames:
        ws = wb["Ranking de Ventas"]
        num_filas = max(ws.max_row - 2, 0)
        encabezados = [c.value for c in ws[1]]
        _aplicar_estilos_fila2_generica(
            ws, encabezados, EXPLICACIONES_RANKING, num_filas
        )

    # 3) Por Sucursal
    if "Por Sucursal" in wb.sheetnames:
        ws = wb["Por Sucursal"]
        num_filas = max(ws.max_row - 2, 0)
        encabezados = [c.value for c in ws[1]]
        _aplicar_estilos_fila2_generica(
            ws, encabezados, EXPLICACIONES_POR_SUCURSAL, num_filas
        )

    # 4) Matriz por Departamento
    if "Matriz" in wb.sheetnames:
        ws = wb["Matriz"]
        num_filas = max(ws.max_row - 2, 0)
        encabezados = [c.value for c in ws[1]]
        _aplicar_estilos_fila2_generica(
            ws, encabezados, EXPLICACIONES_MATRIZ_DEPTO, num_filas
        )

    # 5) Evolución Mensual
    if "Evolución Mensual" in wb.sheetnames:
        ws = wb["Evolución Mensual"]
        num_filas = max(ws.max_row - 2, 0)
        encabezados = [c.value for c in ws[1]]
        _aplicar_estilos_fila2_generica(
            ws, encabezados, EXPLICACIONES_EVOLUCION, num_filas
        )

    # 6) Categorías Especiales
    if "Categorias Especiales" in wb.sheetnames:
        ws = wb["Categorias Especiales"]
        num_filas = max(ws.max_row - 2, 0)
        encabezados = [c.value for c in ws[1]]
        _aplicar_estilos_fila2_generica(
            ws, encabezados, EXPLICACIONES_CONSOLIDADO, num_filas
        )

    # 7) Estacionalidad
    if "Estacionalidad" in wb.sheetnames:
        ws = wb["Estacionalidad"]
        num_filas = max(ws.max_row - 2, 0)
        _aplicar_estilos_estacionalidad(ws, num_filas)

    wb.save(ruta_salida)
