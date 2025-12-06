# core_consolidacion.py

import os
import re
import pandas as pd
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side, numbers
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
    "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
    "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE",
]

MAPA_MES = {m: i + 1 for i, m in enumerate(MESES_ES)}

# ============================================================================
# DICCIONARIOS DE EXPLICACIONES POR HOJA
# ============================================================================

EXPLICACIONES_CONSOLIDADO = {
    'IdArticulo': 'Código único del producto',
    'Marca': 'Nombre del fabricante',
    'Descripcion': 'Nombre comercial del producto',
    'Departamento': 'Categoría principal de venta',
    'SubFamilia': 'Subcategoría del producto',
    'Familia': 'Clasificación comercial',
}

EXPLICACIONES_RANKING = {
    'IdArticulo': 'Código único del producto',
    'Marca': 'Nombre del fabricante',
    'Descripcion': 'Nombre comercial del producto',
    'Total Vendido': 'Unidades totales vendidas (suma todas sucursales y meses)',
}

EXPLICACIONES_POR_SUCURSAL = {
    'SUCURSAL': 'Nombre de la sucursal',
    'TOTAL': 'Unidades vendidas en todo el período',
}

EXPLICACIONES_MATRIZ_DEPTO = {
    'Departamento': 'Categoría principal de venta',
    'CORRIENTES': 'Unidades del departamento en Corrientes (período)',
    'HIPER': 'Unidades del departamento en Hipermercado (período)',
    'TOTAL': 'Sumatoria Corrientes + Hiper (período)',
}

EXPLICACIONES_EVOLUCION = {
    'Departamento': 'Categoría principal de venta',
}

EXPLICACIONES_ESTACIONALIDAD = {
    'IdArticulo': 'Código único del producto',
    'Marca': 'Nombre del fabricante',
    'VENTA_TOTAL_ANUAL': 'Ventas totales en el período',
    'PROMEDIO_MENSUAL': 'Venta promedio por mes',
    'MES_PICO': 'Mes con mayor venta',
    'VENTA_PICO': 'Unidades en mes pico',
    'INDICE_PICO': 'Mes pico ÷ promedio (%)',
}


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
    base_up = base.upper().replace(" ", " ")

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

        # Normalización de Departamento
        if "Departamento" in df_f.columns:
            df_f["Departamento"] = df_f["Departamento"].str.upper()
            df_f.loc[df_f["Departamento"] == "ACEITES", "Departamento"] = "ALMACEN"
            df_f.loc[df_f["Departamento"] == "HIGIENE PERSONAL", "Departamento"] = "LIMPIEZA Y CUIDADO"

        todos.append(df_f)

    if not todos:
        raise ValueError("No se pudo leer ningún archivo válido.")

    df = pd.concat(todos, ignore_index=True)

    # Consolidar por prioridad de departamento
    df["PRIORIDAD"] = df["Departamento"].map(prioridades).fillna(0)
    df["CLAVE_PRODUCTO"] = (
        df["IdArticulo"].astype(str)
        + "|" + df.get("Marca", "").astype(str)
        + "|" + df.get("Descripcion", "").astype(str)
        + "|" + df.get("SubFamilia", "").astype(str)
        + "|" + df.get("Familia", "").astype(str)
    )

    idx_max = df.groupby("CLAVE_PRODUCTO")["PRIORIDAD"].idxmax()
    dept_final = df.loc[idx_max, ["CLAVE_PRODUCTO", "Departamento"]].drop_duplicates("CLAVE_PRODUCTO")

    df = df.drop(columns=["Departamento"])
    df = df.merge(
        dept_final.rename(columns={"Departamento": "Departamento"}),
        on="CLAVE_PRODUCTO",
        how="left",
    )

    df = df.drop(columns=["PRIORIDAD", "CLAVE_PRODUCTO"])

    # Agrupar final
    df = df.groupby(
        ["IdArticulo", "Marca", "Descripcion", "Departamento", "SubFamilia", "Familia", "MES", "SUCURSAL"],
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


# ============================================================================
# FUNCIÓN GENÉRICA PARA APLICAR ESTILOS A FILA 2
# ============================================================================

def _aplicar_estilos_fila2_generica(ws, encabezados, explicaciones_dict, num_filas_datos, meses_para_explicar=None):
    """
    Aplica estilos profesionales a fila 2 (explicaciones) para cualquier hoja.
    
    Args:
        ws: Worksheet de openpyxl
        encabezados: Lista de nombres de columnas (row 1)
        explicaciones_dict: Dict {nombre_columna: explicación}
        num_filas_datos: Cantidad de filas de datos
        meses_para_explicar: Lista de meses para generar explicaciones automáticas
    """
    # Estilo fila 2
    fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
    font = Font(italic=True, size=9)
    alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # Llenar row 2
    for col_idx, encabezado in enumerate(encabezados, start=1):
        cell = ws.cell(row=2, column=col_idx)
        
        # Buscar en diccionario
        if encabezado in explicaciones_dict:
            cell.value = explicaciones_dict[encabezado]
        # Si es un mes, generar explicación automática
        elif meses_para_explicar and any(mes in encabezado for mes in meses_para_explicar):
            cell.value = f"Unidades vendidas en {encabezado}"
        else:
            cell.value = encabezado
        
        cell.fill = fill
        cell.font = font
        cell.alignment = alignment
        cell.border = border

    # Ajustar alto row 2
    ws.row_dimensions[2].height = 30

    # Aplicar bordes a datos
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    for row in range(3, num_filas_datos + 3):
        for col in range(1, len(encabezados) + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = thin_border


def _aplicar_formato_porcentaje_estacionalidad(ws, meses_nombres):
    """
    Aplica formato de porcentaje (0.0%) a las columnas de índices en Estacionalidad.
    Además aplica formato condicional 3 colores a INDICE_PICO.
    """
    # Encontrar columnas de índices
    encabezados = [cell.value for cell in ws[1]]
    
    # Aplicar porcentaje a INDICE_PICO, INDICE_ENERO, etc.
    for col_idx, encabezado in enumerate(encabezados, start=1):
        if encabezado and "INDICE_" in str(encabezado):
            # Aplicar formato porcentaje 0.0%
            for row in range(3, ws.max_row + 1):
                cell = ws.cell(row=row, column=col_idx)
                cell.number_format = '0.0%'
    
    # Aplicar formato condicional a INDICE_PICO
    if 'INDICE_PICO' in encabezados:
        col_indice_pico = encabezados.index('INDICE_PICO') + 1
        
        # Rango: desde row 3 hasta el final
        rango = f"{chr(64 + col_indice_pico)}3:{chr(64 + col_indice_pico)}{ws.max_row}"
        
        # Verde (>110%) = #C6EFCE
        green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        green_font = Font(color="006100", bold=True)
        green_rule = CellIsRule(operator='greaterThan', formula=['1.1'], fill=green_fill, font=green_font)
        ws.conditional_formatting.add(rango, green_rule)
        
        # Rojo (<90%) = #FFC7CE
        red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        red_font = Font(color="9C0006", bold=True)
        red_rule = CellIsRule(operator='lessThan', formula=['0.9'], fill=red_fill, font=red_font)
        ws.conditional_formatting.add(rango, red_rule)
        
        # Amarillo (90-110%) = #FFEB9C (aplicado por defecto, es lo que queda)


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
    Genera reportes en un solo Excel, con opciones:
    - columnas_consolidado: lista de nombres de columnas en el orden deseado
    - habilitar_*: booleans para crear o no cada hoja adicional
    - filtros_especiales: dict opcional {"departamentos": [...], "marcas": [...]}
    - habilitar_estacionalidad: si True, genera hoja con índice estacional
    """
    meses_ordenados = sorted(df["MES"].unique(), key=_orden_mes_clave)
    sucursales = sorted(df["SUCURSAL"].dropna().unique())

    with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
        
        # ============================================================================
        # 1) CONSOLIDADO (siempre se genera)
        # ============================================================================
        df_temp = df.copy()
        df_temp["MES_SUC"] = df_temp["MES"] + "_" + df_temp["SUCURSAL"]

        idx_cols = ["IdArticulo", "Marca", "Descripcion", "Departamento", "SubFamilia", "Familia"]

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

        # Construir columnas por defecto
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

        # Aplicar orden personalizado si existe
        if columnas_consolidado:
            orden = [c for c in columnas_consolidado if c in df_pivot.columns]
            extras = [c for c in df_pivot.columns if c not in orden]
            cols_finales = orden + extras
        else:
            cols_finales = cols_def

        dffinal = df_pivot[cols_finales].sort_values("IdArticulo").reset_index(drop=True)
        dffinal.to_excel(writer, sheet_name="Consolidado", index=False)
        
        # Aplicar estilos fila 2 al Consolidado
        _aplicar_estilos_fila2_generica(
            writer.book["Consolidado"],
            dffinal.columns.tolist(),
            EXPLICACIONES_CONSOLIDADO,
            len(dffinal),
            meses_para_explicar=MESES_ES
        )

        # ============================================================================
        # 2) RANKING DE VENTAS
        # ============================================================================
        if habilitar_ranking:
            ranking = df.groupby(["IdArticulo", "Marca", "Descripcion"]).agg(
                {COLUMNA_CANTIDAD: "sum"}
            ).reset_index()
            ranking = ranking.sort_values(COLUMNA_CANTIDAD, ascending=False).reset_index(drop=True)
            ranking.rename(columns={COLUMNA_CANTIDAD: "Total Vendido"}, inplace=True)
            ranking.to_excel(writer, sheet_name="Ranking de Ventas", index=False)
            
            # Aplicar estilos fila 2
            _aplicar_estilos_fila2_generica(
                writer.book["Ranking de Ventas"],
                ranking.columns.tolist(),
                EXPLICACIONES_RANKING,
                len(ranking)
            )

        # ============================================================================
        # 3) POR SUCURSAL
        # ============================================================================
        if habilitar_por_sucursal:
            porsuc = df.pivot_table(
                index="SUCURSAL",
                columns="MES",
                values=COLUMNA_CANTIDAD,
                aggfunc="sum",
                fill_value=0,
            ).reset_index()

            cols_tot = [m for m in meses_ordenados if m in porsuc.columns]
            if cols_tot:
                porsuc["TOTAL"] = porsuc[cols_tot].sum(axis=1)

            porsuc.to_excel(writer, sheet_name="Por Sucursal", index=False)
            
            # Aplicar estilos fila 2
            _aplicar_estilos_fila2_generica(
                writer.book["Por Sucursal"],
                porsuc.columns.tolist(),
                EXPLICACIONES_POR_SUCURSAL,
                len(porsuc),
                meses_para_explicar=MESES_ES
            )

        # ============================================================================
        # 4) MATRIZ DEPARTAMENTO x SUCURSAL
        # ============================================================================
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

            matriz_reset = matriz.reset_index()
            matriz_reset.to_excel(writer, sheet_name="Matriz", index=False)
            
            # Aplicar estilos fila 2
            _aplicar_estilos_fila2_generica(
                writer.book["Matriz"],
                matriz_reset.columns.tolist(),
                EXPLICACIONES_MATRIZ_DEPTO,
                len(matriz_reset)
            )

        # ============================================================================
        # 5) EVOLUCIÓN MENSUAL
        # ============================================================================
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
            
            evol_reset = evol.reset_index()
            evol_reset.to_excel(writer, sheet_name="Evolución Mensual", index=False)
            
            # Aplicar estilos fila 2
            _aplicar_estilos_fila2_generica(
                writer.book["Evolución Mensual"],
                evol_reset.columns.tolist(),
                EXPLICACIONES_EVOLUCION,
                len(evol_reset),
                meses_para_explicar=MESES_ES
            )

        # ============================================================================
        # 6) CATEGORÍAS ESPECIALES
        # ============================================================================
        if habilitar_especiales:
            dfespec = df.copy()

            # Filtros opcionales
            if filtros_especiales:
                deps = filtros_especiales.get("departamentos") or []
                marcas = filtros_especiales.get("marcas") or []

                if deps:
                    deps_up = [d.upper().strip() for d in deps]
                    dfespec = dfespec[dfespec["Departamento"].str.upper().isin(deps_up)]

                if marcas:
                    marcas_up = [m.upper().strip() for m in marcas]
                    dfespec = dfespec[dfespec["Marca"].str.upper().isin(marcas_up)]
            else:
                dfespec = dfespec[dfespec["Departamento"].str.upper().isin(CATEGORIAS_ESPECIALES)]

            if not dfespec.empty:
                tmp = dfespec.copy()
                tmp["MES_SUC"] = tmp["MES"] + "_" + tmp["SUCURSAL"]

                idx_cols_espec = ["IdArticulo", "Marca", "Descripcion", "Departamento", "SubFamilia", "Familia"]

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
                    piv["TOTAL CONSOLIDADO"] = piv[total_cols_espec].sum(axis=1).astype(int)
                    cols_espec_final.extend(total_cols_espec + ["TOTAL CONSOLIDADO"])

                piv_final = piv[cols_espec_final].sort_values("IdArticulo").reset_index(drop=True)
                piv_final.to_excel(writer, sheet_name="Categorias Especiales", index=False)
                
                # Aplicar estilos fila 2 (igual que Consolidado)
                _aplicar_estilos_fila2_generica(
                    writer.book["Categorias Especiales"],
                    piv_final.columns.tolist(),
                    EXPLICACIONES_CONSOLIDADO,
                    len(piv_final),
                    meses_para_explicar=MESES_ES
                )

        # ============================================================================
        # 7) ESTACIONALIDAD
        # ============================================================================
        if habilitar_estacionalidad:
            dfest = df.copy()
            idx_cols_est = ["IdArticulo", "Marca", "Descripcion", "Departamento", "SubFamilia", "Familia"]

            grp = dfest.groupby(idx_cols_est + ["MES"], as_index=False)[COLUMNA_CANTIDAD].sum()

            if not grp.empty:
                pivest = grp.pivot_table(
                    index=idx_cols_est,
                    columns="MES",
                    values=COLUMNA_CANTIDAD,
                    aggfunc="sum",
                    fill_value=0,
                ).reset_index()

                # Columnas de meses
                mes_cols = [m for m in meses_ordenados if m in pivest.columns]

                if mes_cols:
                    # Total anual y promedio
                    pivest["VENTA_TOTAL_ANUAL"] = pivest[mes_cols].sum(axis=1)
                    n_meses = len(mes_cols)
                    pivest["PROMEDIO_MENSUAL"] = pivest["VENTA_TOTAL_ANUAL"] / n_meses
                    pivest["PROMEDIO_MENSUAL"] = pivest["PROMEDIO_MENSUAL"].replace(0, pd.NA)

                    # Mes pico
                    pivest["MES_PICO"] = pivest[mes_cols].idxmax(axis=1)

                    def venta_pico(row):
                        mes = row["MES_PICO"]
                        if pd.isna(mes):
                            return 0
                        return row.get(mes, 0)

                    pivest["VENTA_PICO"] = pivest.apply(venta_pico, axis=1)

                    # Índice pico
                    pivest["INDICE_PICO"] = pivest["VENTA_PICO"] / pivest["PROMEDIO_MENSUAL"]
                    pivest["INDICE_PICO"] = pivest["INDICE_PICO"].fillna(0).round(2)

                    # Índices por mes
                    for mes in mes_cols:
                        col_indice = f"INDICE_{mes}"
                        pivest[col_indice] = pivest[mes] / pivest["PROMEDIO_MENSUAL"]
                        pivest[col_indice] = pivest[col_indice].fillna(0).round(2)

                    # Orden de columnas
                    cols_out = (
                        idx_cols_est +
                        ["VENTA_TOTAL_ANUAL", "PROMEDIO_MENSUAL", "MES_PICO", "VENTA_PICO", "INDICE_PICO"] +
                        [f"INDICE_{mes}" for mes in mes_cols]
                    )

                    pivest_final = pivest[cols_out].sort_values("IdArticulo").reset_index(drop=True)
                    pivest_final.to_excel(writer, sheet_name="Estacionalidad", index=False)

                    # Aplicar estilos fila 2
                    explicaciones_est = EXPLICACIONES_ESTACIONALIDAD.copy()
                    for mes in mes_cols:
                        explicaciones_est[f"INDICE_{mes}"] = f"{mes} ÷ promedio (%)"

                    _aplicar_estilos_fila2_generica(
                        writer.book["Estacionalidad"],
                        pivest_final.columns.tolist(),
                        explicaciones_est,
                        len(pivest_final)
                    )

                    # Aplicar formato porcentaje y condicional
                    _aplicar_formato_porcentaje_estacionalidad(
                        writer.book["Estacionalidad"],
                        mes_cols
                    )
