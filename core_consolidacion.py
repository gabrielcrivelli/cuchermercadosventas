# core_consolidacion.py

import os
import re
import pandas as pd

from openpyxl import load_workbook
from openpyxl.styles import (
    PatternFill,
    Font,
    Alignment,
    Border,
    Side,
)
from openpyxl.formatting.rule import CellIsRule

COLUMNA_CANTIDAD = "Cantidad"

COLUMNAS_DESCRIPTIVAS = ["Marca", "Descripcion", "Departamento", "SubFamilia", "Familia"]

# Prioridades base (se pueden sobreescribir desde la GUI)
PRIORIDAD_DEPARTAMENTOS_DEFAULT = {
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

CATEGORIAS_ESPECIALES = ["ELECTRO", "ELECTRODOMESTICOS", "FERRETERIA", "RODADOS"]

MESES_ES = [
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

MAPA_MES = {m: i + 1 for i, m in enumerate(MESES_ES)}

# --------- Explicaciones para fila 2 en cada hoja --------- #

EXPLICACIONES_CONSOLIDADO = {
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

EXPLICACIONES_RANKING = {
    "IdArticulo": "Código único del producto",
    "Marca": "Nombre del fabricante",
    "Descripcion": "Nombre comercial del producto",
    "Total Vendido": "Unidades totales vendidas (todas las sucursales y meses)",
}

EXPLICACIONES_POR_SUCURSAL = {
    "SUCURSAL": "Nombre de la sucursal",
    "TOTAL": "Unidades vendidas en todo el período",
}

EXPLICACIONES_MATRIZ_DEPTO = {
    "Departamento": "Categoría principal de venta",
    "CORRIENTES": "Unidades del departamento en sucursal Corrientes (período)",
    "HIPER": "Unidades del departamento en sucursal Hipermercado (período)",
    "TOTAL": "Sumatoria Corrientes + Hiper (período)",
}

EXPLICACIONES_EVOLUCION = {
    "Departamento": "Categoría principal de venta",
}

EXPLICACIONES_ESTACIONALIDAD = {
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
    # Índices por mes se generan dinámicamente: "[MES] ÷ promedio (%)"
}

# ---------------------------------------------------------- #


def normalizar_mes(mes_str: str) -> str:
    mes_up = mes_str.strip().upper()
    for m in MESES_ES:
        if m in mes_up:
            return m
    raise ValueError(f"No se pudo normalizar mes desde: {mes_str}")


def parsear_nombre_archivo(nombre: str):
    """
    Devuelve (mes_str, anio_int, sucursal_str_o_None) o None si no se puede parsear.
    Espera algo tipo: '3. MARZO 2025 CORRIENTES.xlsx'
    """
    base = os.path.splitext(os.path.basename(nombre))[0]
    base_up = base.upper().replace("  ", " ")

    # Buscar sucursal conocida
    sucursal = None
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


def consolidar_datos(archivos_info, prioridades_depto=None):
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

    todos = []

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

    df = pd.concat(todos, ignore_index=True)

    # Consolidar por prioridad de departamento
    df["PRIORIDAD"] = df["Departamento"].map(prioridades).fillna(0)
    df["CLAVE_PRODUCTO"] = (
        df["IdArticulo"].astype(str)
        + "|"
        + df.get("Marca", "").astype(str)
        + "|"
        + df.get("Descripcion", "").astype(str)
        + "|"
        + df.get("SubFamilia", "").astype(str)
        + "|"
        + df.get("Familia", "").astype(str)
    )

    idx_max = df.groupby("CLAVE_PRODUCTO")["PRIORIDAD"].idxmax()
    dept_final = df.loc[idx_max, ["CLAVE_PRODUCTO", "Departamento"]].drop_duplicates(
        "CLAVE_PRODUCTO"
    )

    df = df.drop(columns=["Departamento"])
    df = df.merge(
        dept_final.rename(columns={"Departamento": "Departamento"}),
        on="CLAVE_PRODUCTO",
        how="left",
    )

    df = df.drop(columns=["PRIORIDAD", "CLAVE_PRODUCTO"])

    # Agrupar final
    df = df.groupby(
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
    df[COLUMNA_CANTIDAD] = df[COLUMNA_CANTIDAD].apply(
        lambda x: int(x + 0.5) if x >= 0 else int(x - 0.5)
    )

    return df


def _orden_mes_clave(mes_ano: str):
    # mes_ano: "MARZO 2025"
    partes = mes_ano.split()
    if len(partes) != 2:
        return (9999, 99)
    mes_txt, anio_txt = partes
    anio = int(anio_txt)
    mes_num = MAPA_MES.get(mes_txt, 99)
    return (anio, mes_num)


# --------- Estilos genéricos para fila 2 --------- #


def _aplicar_estilos_fila2_generica(ws, encabezados, explicaciones_dict, num_filas_datos):
    """
    Aplica estilos profesionales a fila 2 (explicaciones) para cualquier hoja.
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

    # Fila 2
    for col_idx, encabezado in enumerate(encabezados, start=1):
        cell = ws.cell(row=2, column=col_idx)

        # Lógica especial según patrón de nombre
        valor = explicaciones_dict.get(encabezado)

        # Consolidado / Categorías Especiales: MES_AAAA_SUCURSAL
        if valor is None and "_" in encabezado and " " in encabezado:
            # ejemplo: "MARZO 2025_CORRIENTES"
            partes = encabezado.split("_", 1)
            mes_anio = partes[0]
            suc = partes[1]
            valor = f"Unidades {mes_anio} en {suc}"

        # Por Sucursal: columnas de meses
        if valor is None and encabezado in MESES_ES:
            valor = f"Unidades vendidas en {encabezado.title()}"

        # Evolución Mensual: meses en columnas
        if valor is None and any(m in encabezado for m in MESES_ES):
            valor = f"Unidades del departamento en {encabezado}"

        # Estacionalidad: índices por mes
        if (
            valor is None
            and encabezado.upper().startswith("INDICE_")
            and encabezado != "INDICE_PICO"
        ):
            mes = encabezado.replace("INDICE_", "")
            valor = f"{mes.title()} ÷ promedio (%)"

        if valor is None:
            valor = encabezado

        cell.value = valor
        cell.fill = fill
        cell.font = font
        cell.alignment = alignment
        cell.border = border

    # Alto fila 2
    ws.row_dimensions[2].height = 30

    # Bordes datos
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

            # Alineación texto vs números (heurística simple)
            if col <= 6 or isinstance(cell.value, str):
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="right", vertical="center")


def _aplicar_estilos_estacionalidad(ws, num_filas_datos):
    """
    Aplica:
    - Fila 2 con explicaciones y estilos
    - Formato porcentaje en columnas de índices
    - Formato condicional verde/amarillo/rojo sobre INDICE_PICO
    """
    encabezados = [cell.value for cell in ws[1]]
    _aplicar_estilos_fila2_generica(
        ws, encabezados, EXPLICACIONES_ESTACIONALIDAD, num_filas_datos
    )

    # Columnas de índices (INDICE_PICO + INDICE_ENERO..INDICE_DICIEMBRE)
    col_indices = []
    for idx, nombre in enumerate(encabezados, start=1):
        if nombre and nombre.startswith("INDICE_"):
            col_indices.append(idx)

    # Formato porcentaje 0.0%
    for col_idx in col_indices:
        for row in range(3, num_filas_datos + 3):
            cell = ws.cell(row=row, column=col_idx)
            cell.number_format = "0.0%"

    # Formato condicional sobre INDICE_PICO (si existe)
    col_pico = None
    for idx, nombre in enumerate(encabezados, start=1):
        if nombre == "INDICE_PICO":
            col_pico = idx
            break

    if col_pico is None:
        return

    rango = f"{_col_letra(col_pico)}3:{_col_letra(col_pico)}{num_filas_datos + 2}"

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
        operator="between",
        formula=["1.0", "1.10"],  # 100%–110%
        fill=amarillo_fill,
        font=amarillo_font,
    )

    rojo_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    rojo_font = Font(bold=True, color="9C0006")
    regla_rojo = CellIsRule(
        operator="lessThan", formula=["1.0"], fill=rojo_fill, font=rojo_font
    )

    ws.conditional_formatting.add(rango, regla_verde)
    ws.conditional_formatting.add(rango, regla_amarillo)
    ws.conditional_formatting.add(rango, regla_rojo)


def _col_letra(idx):
    """Convierte índice de columna (1-based) a letra estilo Excel."""
    result = ""
    while idx:
        idx, rem = divmod(idx - 1, 26)
        result = chr(65 + rem) + result
    return result


# ---------------------------------------------------------- #


def generar_reportes(
    df,
    ruta_salida,
    columnas_consolidado=None,
    habilitar_ranking=True,
    habilitar_por_sucursal=True,
    habilitar_matriz=True,
    habilitar_evolucion=True,
    habilitar_especiales=True,
    filtros_especiales=None,
    habilitar_estacionalidad=False,
):
    """
    Genera reportes en un solo Excel, con opciones.
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

        # Columnas por defecto
        cols_def = idx_cols.copy()

        for mes in meses_ordenados:
            cols_mes = [c for c in df_pivot.columns if c.startswith(mes + "_")]
            if not cols_mes:
                continue
            df_pivot[mes] = df_pivot[cols_mes].sum(axis=1).astype(int)
            cols_def.append(mes)

        total_cols = []
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

        # Aplicar orden personalizado si viene desde GUI
        if columnas_consolidado:
            orden = [c for c in columnas_consolidado if c in df_pivot.columns]
            extras = [c for c in df_pivot.columns if c not in orden]
            cols_finales = orden + extras
        else:
            cols_finales = cols_def

        df_final = (
            df_pivot[cols_finales].sort_values("IdArticulo").reset_index(drop=True)
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

                total_cols_espec = []
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
                if not mes_cols:
                    # sin meses no tiene sentido la hoja
                    pass
                else:
                    piv_est["VENTA_TOTAL_ANUAL"] = piv_est[mes_cols].sum(axis=1)
                    n_meses = len(mes_cols)
                    piv_est["PROMEDIO_MENSUAL"] = (
                        piv_est["VENTA_TOTAL_ANUAL"] / n_meses
                    )
                    piv_est["PROMEDIO_MENSUAL"] = piv_est[
                        "PROMEDIO_MENSUAL"
                    ].replace(0, pd.NA)

                    piv_est["MES_PICO"] = piv_est[mes_cols].idxmax(axis=1)

                    def _venta_pico(row):
                        mes = row["MES_PICO"]
                        if pd.isna(mes):
                            return 0
                        return row.get(mes, 0)

                    piv_est["VENTA_PICO"] = piv_est.apply(_venta_pico, axis=1)

                    piv_est["INDICE_PICO"] = (
                        piv_est["VENTA_PICO"] / piv_est["PROMEDIO_MENSUAL"]
                    )
                    piv_est["INDICE_PICO"] = piv_est["INDICE_PICO"].fillna(0).round(3)

                    # Índices por mes (INDICE_ENERO, etc.)
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

        # Fin with: se guarda el archivo aquí

    # -------- Estilos en todas las hojas (usando openpyxl) -------- #

    wb = load_workbook(ruta_salida)

    # Consolidado
    if "Consolidado" in wb.sheetnames:
        ws = wb["Consolidado"]
        num_filas = ws.max_row - 2
        encabezados = [c.value for c in ws[1]]
        _aplicar_estilos_fila2_generica(
            ws, encabezados, EXPLICACIONES_CONSOLIDADO, num_filas
        )

    # Ranking de Ventas
    if "Ranking de Ventas" in wb.sheetnames:
        ws = wb["Ranking de Ventas"]
        num_filas = ws.max_row - 2
        encabezados = [c.value for c in ws[1]]
        _aplicar_estilos_fila2_generica(
            ws, encabezados, EXPLICACIONES_RANKING, num_filas
        )

    # Por Sucursal
    if "Por Sucursal" in wb.sheetnames:
        ws = wb["Por Sucursal"]
        num_filas = ws.max_row - 2
        encabezados = [c.value for c in ws[1]]
        _aplicar_estilos_fila2_generica(
            ws, encabezados, EXPLICACIONES_POR_SUCURSAL, num_filas
        )

    # Matriz (Departamento x Sucursal)
    if "Matriz" in wb.sheetnames:
        ws = wb["Matriz"]
        num_filas = ws.max_row - 2
        encabezados = [c.value for c in ws[1]]
        _aplicar_estilos_fila2_generica(
        ws, encabezados, EXPLICACIONES_MATRIZ_DEPTO, num_filas
        )

    # Evolución Mensual
    if "Evolución Mensual" in wb.sheetnames:
        ws = wb["Evolución Mensual"]
        num_filas = ws.max_row - 2
        encabezados = [c.value for c in ws[1]]
        _aplicar_estilos_fila2_generica(
            ws, encabezados, EXPLICACIONES_EVOLUCION, num_filas
        )

    # Categorías Especiales
    if "Categorias Especiales" in wb.sheetnames:
        ws = wb["Categorias Especiales"]
        num_filas = ws.max_row - 2
        encabezados = [c.value for c in ws[1]]
        _aplicar_estilos_fila2_generica(
            ws, encabezados, EXPLICACIONES_CONSOLIDADO, num_filas
        )

    # Estacionalidad
    if "Estacionalidad" in wb.sheetnames:
        ws = wb["Estacionalidad"]
        num_filas = ws.max_row - 2
        _aplicar_estilos_estacionalidad(ws, num_filas)

    wb.save(ruta_salida)
