# core_consolidacion.py

import os
import re
import pandas as pd
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

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

# Explicaciones para columnas de Estacionalidad
EXPLICACIONES_ESTACIONALIDAD = {
    "Venta total año": "Ventas totales en el período",
    "Venta promedio mensual": "Venta promedio por mes",
    "Mes de pico": "Mes con mayor venta",
    "Venta en mes pico": "Unidades en mes pico",
    "Índice estacional pico": "Mes pico / promedio (%)",
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
        - "ruta": str,
        - "mes": "MARZO",
        - "anio": 2025,
        - "sucursal": "HIPER"
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


def _aplicar_estilos_estacionalidad(ws, num_filas_datos):
    """
    Aplica estilos a la hoja de estacionalidad:
    - Encabezados con fondo gris y texto blanco
    - Fila de explicaciones debajo de encabezados
    - Formato condicional de colores según ÍNDICE_PICO
    """
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

    # Estilos
    header_fill = PatternFill(start_color="4F4F4F", end_color="4F4F4F", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    explanation_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
    explanation_font = Font(italic=True, size=9, color="000000")
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    # Fila 1: Encabezados
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border

    # Fila 2: Explicaciones
    explicaciones_map = {
        "IdArticulo": "Código",
        "Marca": "Marca",
        "Descripcion": "Descripción",
        "Departamento": "Departamento",
        "SubFamilia": "Subfamilia",
        "Familia": "Familia",
        "Venta total año": "Ventas totales en período",
        "Venta promedio mensual": "Venta promedio por mes",
        "Mes de pico": "Mes con mayor venta",
        "Venta en mes pico": "Unidades en mes pico",
        "Índice estacional pico": "Mes pico / promedio (%)",
    }

    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=2, column=col)
        header_cell = ws.cell(row=1, column=col)
        header_text = header_cell.value
        explicacion = explicaciones_map.get(str(header_text), "")
        cell.value = explicacion
        cell.fill = explanation_fill
        cell.font = explanation_font
        cell.alignment = center_align
        cell.border = thin_border

    # Formato condicional para columna ÍNDICE_PICO (porcentaje)
    # Buscar columna con ese nombre
    indice_col = None
    for col in range(1, ws.max_column + 1):
        if ws.cell(row=1, column=col).value == "Índice estacional pico":
            indice_col = col
            break

    if indice_col:
        green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")  # Verde claro
        green_font = Font(color="006100", bold=True)
        yellow_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")  # Amarillo claro
        yellow_font = Font(color="9C6500", bold=True)
        red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")  # Rojo claro
        red_font = Font(color="9C0006", bold=True)

        # Aplicar colores según rango (filas 3 en adelante, omitiendo fila de explicaciones)
        for row in range(3, num_filas_datos + 3):
            cell = ws.cell(row=row, column=indice_col)
            if cell.value is not None:
                try:
                    valor = float(str(cell.value).replace("%", ""))
                    if valor > 110:  # Mayor a 110%
                        cell.fill = green_fill
                        cell.font = green_font
                    elif valor < 90:  # Menor a 90%
                        cell.fill = red_fill
                        cell.font = red_font
                    else:  # Entre 90% y 110%
                        cell.fill = yellow_fill
                        cell.font = yellow_font
                except (ValueError, AttributeError):
                    pass

    # Ajustar ancho de columnas
    ws.row_dimensions[1].height = 25
    ws.row_dimensions[2].height = 40
    for col in range(1, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(col)].width = 18


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
      (se usan solo las que existan; el resto se ignora).
    - habilitar_*: booleans para crear o no cada hoja adicional.
    - filtros_especiales: dict opcional {"departamentos": [...], "marcas": [...]}
      para filtrar la hoja de Categorías Especiales.
    - habilitar_estacionalidad: si True, genera hoja con índice estacional,
      índices por mes y mes pico por producto.
    """
    meses_ordenados = sorted(df["MES"].unique(), key=_orden_mes_clave)
    sucursales = sorted(df["SUCURSAL"].dropna().unique())

    with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
        # 1) CONSOLIDADO (siempre se genera)
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

        # Aplicar columnas personalizadas si se proporcionan
        if columnas_consolidado:
            cols_def = [c for c in columnas_consolidado if c in df_pivot.columns]

        df_pivot = df_pivot[cols_def]
        df_pivot.to_excel(writer, sheet_name="Consolidado", index=False)

        # 2) ESTACIONALIDAD (si está habilitada)
        if habilitar_estacionalidad:
            df_estacionalidad = _generar_estacionalidad(df, meses_ordenados)
            df_estacionalidad.to_excel(writer, sheet_name="Estacionalidad", index=False, startrow=2)

            # Aplicar estilos a la hoja de estacionalidad
            ws_est = writer.sheets["Estacionalidad"]
            _aplicar_estilos_estacionalidad(ws_est, len(df_estacionalidad))

        # 3) RANKING (si está habilitado)
        if habilitar_ranking:
            df_ranking = df_pivot[idx_cols + ["TOTAL CONSOLIDADO"]].copy()
            df_ranking = df_ranking.sort_values("TOTAL CONSOLIDADO", ascending=False).head(50)
            df_ranking.to_excel(writer, sheet_name="Top 50", index=False)

        # 4) POR SUCURSAL (si está habilitado)
        if habilitar_por_sucursal:
            for suc in sucursales:
                df_suc = df[df["SUCURSAL"] == suc].copy()
                df_suc_pivot = df_suc.pivot_table(
                    index=idx_cols,
                    columns="MES",
                    values=COLUMNA_CANTIDAD,
                    aggfunc="sum",
                    fill_value=0,
                ).reset_index()

                for col in df_suc_pivot.columns:
                    if col not in idx_cols:
                        df_suc_pivot[col] = df_suc_pivot[col].astype(int)

                cols_suc = idx_cols + [m for m in meses_ordenados if m in df_suc_pivot.columns]
                df_suc_pivot = df_suc_pivot[cols_suc]

                sheet_name = f"Sucursal {suc[:10]}"
                df_suc_pivot.to_excel(writer, sheet_name=sheet_name, index=False)

        # 5) MATRIZ (si está habilitada)
        if habilitar_matriz:
            df_matriz = df_pivot[["Departamento", "TOTAL CONSOLIDADO"]].copy()
            df_matriz = df_matriz.groupby("Departamento", as_index=False)["TOTAL CONSOLIDADO"].sum()
            df_matriz = df_matriz.sort_values("TOTAL CONSOLIDADO", ascending=False)
            df_matriz.to_excel(writer, sheet_name="Matriz Departamentos", index=False)

        # 6) EVOLUCIÓN (si está habilitada)
        if habilitar_evolucion:
            df_evo = df.groupby("MES", as_index=False)[COLUMNA_CANTIDAD].sum()
            df_evo = df_evo.sort_values("MES", key=lambda x: x.map(_orden_mes_clave))
            df_evo.to_excel(writer, sheet_name="Evolución Mensual", index=False)

        # 7) CATEGORÍAS ESPECIALES (si está habilitada)
        if habilitar_especiales:
            filtros = filtros_especiales or {}
            depto_filtro = filtros.get("departamentos", CATEGORIAS_ESPECIALES)
            df_esp = df_pivot[df_pivot["Departamento"].str.contains("|".join(depto_filtro), case=False, na=False)]
            if not df_esp.empty:
                df_esp.to_excel(writer, sheet_name="Categorías Especiales", index=False)


def _generar_estacionalidad(df, meses_ordenados):
    """
    Genera tabla de estacionalidad con:
    - Venta total anual
    - Promedio mensual
    - Mes de pico
    - Venta en mes pico
    - Índice estacional pico (en %)
    - Índices por mes (en %)
    """
    idx_cols = ["IdArticulo", "Marca", "Descripcion", "Departamento", "SubFamilia", "Familia"]

    # Crear tabla base consolidada
    df_cons = df.pivot_table(
        index=idx_cols,
        columns="MES",
        values="Cantidad",
        aggfunc="sum",
        fill_value=0,
    ).reset_index()

    # Calcular venta total anual (suma de todos los meses)
    cols_meses = [m for m in meses_ordenados if m in df_cons.columns]
    df_cons["Venta total año"] = df_cons[cols_meses].sum(axis=1).astype(int)

    # Calcular promedio mensual (total / número de meses con datos)
    num_meses = len(cols_meses)
    df_cons["Venta promedio mensual"] = (df_cons["Venta total año"] / num_meses).round(2)

    # Encontrar mes de pico (máximo de ventas)
    def obtener_mes_pico(row):
        valores_meses = [(m, row[m]) for m in cols_meses if pd.notna(row[m])]
        if not valores_meses:
            return None
        mes_pico, _ = max(valores_meses, key=lambda x: x[1])
        return mes_pico

    df_cons["Mes de pico"] = df_cons.apply(obtener_mes_pico, axis=1)

    # Venta en mes pico
    def obtener_venta_pico(row):
        mes_pico = row["Mes de pico"]
        if mes_pico and mes_pico in df_cons.columns:
            return int(row[mes_pico])
        return 0

    df_cons["Venta en mes pico"] = df_cons.apply(obtener_venta_pico, axis=1)

    # Índice estacional pico (como %)
    def calcular_indice_pico(row):
        if row["Venta promedio mensual"] > 0:
            indice = (row["Venta en mes pico"] / row["Venta promedio mensual"]) * 100
            return f"{indice:.1f}%"
        return "0%"

    df_cons["Índice estacional pico"] = df_cons.apply(calcular_indice_pico, axis=1)

    # Agregar índices por mes (en %)
    for mes in cols_meses:
        def calcular_indice_mes(row, m=mes):
            if row["Venta promedio mensual"] > 0:
                indice = (row[m] / row["Venta promedio mensual"]) * 100
                return f"{indice:.1f}%"
            return "0%"

        col_nombre = f"Índice {mes}"
        df_cons[col_nombre] = df_cons.apply(calcular_indice_mes, axis=1)

    # Seleccionar columnas finales (descriptivas + métricas + índices)
    cols_finales = idx_cols + [
        "Venta total año",
        "Venta promedio mensual",
        "Mes de pico",
        "Venta en mes pico",
        "Índice estacional pico",
    ]

    # Agregar índices por mes al final
    cols_finales.extend([f"Índice {mes}" for mes in cols_meses])

    df_estacionalidad = df_cons[cols_finales].copy()

    return df_estacionalidad
