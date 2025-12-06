import os
from datetime import datetime
from tempfile import NamedTemporaryFile
from io import BytesIO

import streamlit as st
import pandas as pd

# ===================== CORE LOGIC (inline) =====================
# Se incluye la lógica de consolidación directamente para evitar imports problemáticos

COLUMNA_CANTIDAD = "Cantidad"
COLUMNAS_DESCRIPTIVAS = ["Marca", "Descripcion", "Departamento", "SubFamilia", "Familia"]

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

COLUMNAS_DISPONIBLES = [
    "IdArticulo", "Marca", "Descripcion", "Departamento", "SubFamilia", "Familia",
    "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
    "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE",
    "TOTAL CORRIENTES", "TOTAL HIPER", "TOTAL CONSOLIDADO",
]


def normalizar_mes(mes_str: str) -> str:
    mes_up = mes_str.strip().upper()
    for m in MESES_ES:
        if m in mes_up:
            return m
    raise ValueError(f"No se pudo normalizar mes desde: {mes_str}")


def parsear_nombre_archivo(nombre: str):
    import re
    base = os.path.splitext(os.path.basename(nombre))[0]
    base_up = base.upper().replace("  ", " ")

    sucursal = None
    for s in ["HIPER", "CORRIENTES"]:
        if s in base_up:
            sucursal = s
            base_up = base_up.replace(s, "").strip()
            break

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


def consolidar_datos_from_bytes(archivos_info):
    """
    archivos_info: lista de dicts con:
        - "bytes": bytes del archivo
        - "nombre": nombre original
        - "mes": str
        - "anio": int
        - "sucursal": str
    """
    todos = []

    for info in archivos_info:
        file_bytes = info["bytes"]
        mes = info["mes"]
        anio = info["anio"]
        sucursal = info["sucursal"]

        try:
            df = pd.read_excel(BytesIO(file_bytes), sheet_name=0)
        except Exception:
            continue

        if COLUMNA_CANTIDAD not in df.columns or "IdArticulo" not in df.columns:
            continue

        columnas_existentes = [c for c in COLUMNAS_DESCRIPTIVAS if c in df.columns]
        columnas_a_usar = ["IdArticulo"] + columnas_existentes + [COLUMNA_CANTIDAD]

        df_f = df[columnas_a_usar].copy()
        df_f["MES"] = f"{mes} {anio}"
        df_f["SUCURSAL"] = sucursal

        for col in COLUMNAS_DESCRIPTIVAS:
            if col in df_f.columns:
                df_f[col] = df_f[col].astype(str).str.strip()

        if "Departamento" in df_f.columns:
            df_f["Departamento"] = df_f["Departamento"].str.upper()
            df_f.loc[df_f["Departamento"] == "ACEITES", "Departamento"] = "ALMACEN"
            df_f.loc[df_f["Departamento"] == "HIGIENE PERSONAL", "Departamento"] = "LIMPIEZA Y CUIDADO"

        todos.append(df_f)

    if not todos:
        raise ValueError("No se pudo leer ningún archivo válido.")

    df = pd.concat(todos, ignore_index=True)

    prioridades = PRIORIDAD_DEPARTAMENTOS_DEFAULT.copy()
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
    df = df.merge(dept_final, on="CLAVE_PRODUCTO", how="left")
    df = df.drop(columns=["PRIORIDAD", "CLAVE_PRODUCTO"])

    df = df.groupby(
        ["IdArticulo", "Marca", "Descripcion", "Departamento", "SubFamilia", "Familia", "MES", "SUCURSAL"],
        as_index=False,
    )[COLUMNA_CANTIDAD].sum()

    df[COLUMNA_CANTIDAD] = df[COLUMNA_CANTIDAD].apply(
        lambda x: int(x + 0.5) if x >= 0 else int(x - 0.5)
    )

    return df


def _orden_mes_clave(mes_ano: str):
    partes = mes_ano.split()
    if len(partes) != 2:
        return (9999, 99)
    mes_txt, anio_txt = partes
    anio = int(anio_txt)
    mes_num = MAPA_MES.get(mes_txt, 99)
    return (anio, mes_num)


def generar_reportes_bytes(
    df,
    columnas_consolidado=None,
    habilitar_ranking=True,
    habilitar_por_sucursal=True,
    habilitar_matriz=True,
    habilitar_evolucion=True,
    habilitar_especiales=True,
    filtros_especiales=None,
    habilitar_estacionalidad=False,
):
    """Genera el Excel en memoria y retorna bytes."""
    output = BytesIO()
    meses_ordenados = sorted(df["MES"].unique(), key=_orden_mes_clave)
    sucursales = sorted(df["SUCURSAL"].dropna().unique())

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # 1) Consolidado
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

        if columnas_consolidado:
            orden = [c for c in columnas_consolidado if c in df_pivot.columns]
            extras = [c for c in df_pivot.columns if c not in orden]
            cols_finales = orden + extras
        else:
            cols_finales = cols_def

        df_final = df_pivot[cols_finales].sort_values("IdArticulo").reset_index(drop=True)
        df_final.to_excel(writer, sheet_name="Consolidado", index=False)

        # 2) Ranking
        if habilitar_ranking:
            ranking = df.groupby(["IdArticulo", "Marca", "Descripcion"]).agg(
                {COLUMNA_CANTIDAD: "sum"}
            ).reset_index()
            ranking = ranking.sort_values(COLUMNA_CANTIDAD, ascending=False).reset_index(drop=True)
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

        # 4) Matriz
        if habilitar_matriz:
            matriz = df.pivot_table(
                index="Departamento",
                columns="SUCURSAL",
                values=COLUMNA_CANTIDAD,
                aggfunc="sum",
                fill_value=0,
            )
            matriz["TOTAL"] = matriz.sum(axis=1)
            matriz.sort_values("TOTAL", ascending=False).to_excel(writer, sheet_name="Matriz")

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

        # 6) Categorías Especiales
        if habilitar_especiales:
            df_espec = df.copy()

            if filtros_especiales:
                deps = filtros_especiales.get("departamentos") or []
                marcas = filtros_especiales.get("marcas") or []

                if deps:
                    deps_up = [d.upper().strip() for d in deps]
                    df_espec = df_espec[df_espec["Departamento"].str.upper().isin(deps_up)]

                if marcas:
                    marcas_up = [m.upper().strip() for m in marcas]
                    df_espec = df_espec[df_espec["Marca"].str.upper().isin(marcas_up)]
            else:
                df_espec = df_espec[
                    df_espec["Departamento"].str.upper().isin(CATEGORIAS_ESPECIALES)
                ]

            if not df_espec.empty:
                tmp = df_espec.copy()
                tmp["MES_SUC"] = tmp["MES"] + "_" + tmp["SUCURSAL"]

                idx_cols_espec = [
                    "IdArticulo", "Marca", "Descripcion",
                    "Departamento", "SubFamilia", "Familia",
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
                    piv["TOTAL CONSOLIDADO"] = piv[total_cols_espec].sum(axis=1).astype(int)
                    cols_espec_final.extend(total_cols_espec + ["TOTAL CONSOLIDADO"])

                piv_final = piv[cols_espec_final].sort_values("IdArticulo").reset_index(drop=True)
                piv_final.to_excel(writer, sheet_name="Categorias Especiales", index=False)

        # 7) Estacionalidad
        if habilitar_estacionalidad:
            df_est = df.copy()

            idx_cols_est = [
                "IdArticulo", "Marca", "Descripcion",
                "Departamento", "SubFamilia", "Familia",
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
                    piv_est["PROMEDIO_MENSUAL"] = piv_est["VENTA_TOTAL_ANUAL"] / n_meses
                    piv_est["PROMEDIO_MENSUAL"] = piv_est["PROMEDIO_MENSUAL"].replace(0, pd.NA)

                    piv_est["MES_PICO"] = piv_est[mes_cols].idxmax(axis=1)

                    def _venta_pico(row):
                        mes = row["MES_PICO"]
                        if pd.isna(mes):
                            return 0
                        return row.get(mes, 0)

                    piv_est["VENTA_PICO"] = piv_est.apply(_venta_pico, axis=1)

                    piv_est["INDICE_PICO"] = piv_est["VENTA_PICO"] / piv_est["PROMEDIO_MENSUAL"]
                    piv_est["INDICE_PICO"] = piv_est["INDICE_PICO"].fillna(0).round(2)

                    cols_out = idx_cols_est + [
                        "VENTA_TOTAL_ANUAL",
                        "PROMEDIO_MENSUAL",
                        "MES_PICO",
                        "VENTA_PICO",
                        "INDICE_PICO",
                    ]

                    piv_est[cols_out].to_excel(writer, sheet_name="Estacionalidad", index=False)

    output.seek(0)
    return output.getvalue()


# ===================== STREAMLIT UI =====================

st.set_page_config(
    page_title="Consolidador de Ventas - Cucher Mercados",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Consolidador de Ventas Mensuales")
st.caption("Cucher Mercados - Resistencia, Chaco")

st.markdown("---")

# Parámetros superiores
col1, col2 = st.columns([1, 4])

with col1:
    anio = st.number_input(
        "Año",
        min_value=2000,
        max_value=2030,
        value=datetime.now().year,
        step=1,
    )

with col2:
    st.write("**Meses a procesar:**")
    meses_sel = st.multiselect(
        "Meses",
        options=MESES_ES,
        default=MESES_ES,
        label_visibility="collapsed",
    )

st.markdown("---")

# Carga de archivos
st.subheader("📁 Archivos Excel")
st.caption("Formato esperado: `N. MES AÑO SUCURSAL.xlsx` (ej: `3. MARZO 2025 CORRIENTES.xlsx`)")

uploaded_files = st.file_uploader(
    "Subir archivos Excel",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
)

archivos_info = []
tabla_archivos = []

if uploaded_files:
    for uf in uploaded_files:
        file_bytes = uf.getbuffer()

        info = {
            "bytes": bytes(file_bytes),
            "nombre": uf.name,
            "mes": None,
            "anio": None,
            "sucursal": None,
        }
        parsed = parsear_nombre_archivo(uf.name)
        estado = "✅ OK"

        if parsed:
            mes, anio_arch, sucursal = parsed
            info["mes"] = mes
            info["anio"] = anio_arch
            info["sucursal"] = sucursal if sucursal else "DESCONOCIDA"
            if anio_arch != anio:
                estado = f"⚠️ AÑO {anio_arch}≠{anio}"
        else:
            estado = "❌ No detectado"

        archivos_info.append(info)
        tabla_archivos.append({
            "Archivo": uf.name,
            "MES": info["mes"] or "",
            "AÑO": info["anio"] or "",
            "SUCURSAL": info["sucursal"] or "",
            "ESTADO": estado,
        })

if tabla_archivos:
    st.dataframe(pd.DataFrame(tabla_archivos), use_container_width=True, hide_index=True)

st.markdown("---")

# Configuración
st.subheader("⚙️ Configurar Excel Resultado")

col_cfg1, col_cfg2, col_cfg3 = st.columns([3, 2, 3])

with col_cfg1:
    st.markdown("**Columnas del consolidado**")
    columnas_sel = st.multiselect(
        "Columnas (orden de salida)",
        options=COLUMNAS_DISPONIBLES,
        default=COLUMNAS_DISPONIBLES,
        label_visibility="collapsed",
    )

with col_cfg2:
    st.markdown("**Hojas a generar**")
    habilitar_ranking = st.checkbox("Ranking de Ventas", value=True)
    habilitar_por_sucursal = st.checkbox("Ventas por Sucursal", value=True)
    habilitar_matriz = st.checkbox("Matriz por Departamento", value=True)
    habilitar_evolucion = st.checkbox("Evolución Mensual", value=True)
    habilitar_especiales = st.checkbox("Categorías Especiales", value=True)
    habilitar_estacionalidad = st.checkbox("Estacionalidad", value=True)

with col_cfg3:
    st.markdown("**Filtros Categorías Especiales**")
    deps_txt = st.text_input("Departamentos (separados por coma):", value="")
    marcas_txt = st.text_input("Marcas (separadas por coma):", value="")

st.markdown("---")

# Procesar
if st.button("🚀 Procesar", type="primary"):
    try:
        if not archivos_info:
            st.error("Debes subir al menos un archivo Excel.")
            st.stop()

        if not meses_sel:
            st.error("Debes seleccionar al menos un mes.")
            st.stop()

        archivos_filtrados = [
            a for a in archivos_info
            if a["anio"] == anio and a["mes"] in meses_sel
        ]

        if not archivos_filtrados:
            st.error("No hay archivos que coincidan con el año y meses seleccionados.")
            st.stop()

        with st.spinner("Consolidando datos..."):
            df = consolidar_datos_from_bytes(archivos_filtrados)

        if not columnas_sel:
            st.error("Debes seleccionar al menos una columna.")
            st.stop()

        deps = [d.strip() for d in deps_txt.split(",") if d.strip()]
        marcas = [m.strip() for m in marcas_txt.split(",") if m.strip()]
        filtros_especiales = (
            {"departamentos": deps, "marcas": marcas} if (deps or marcas) else None
        )

        with st.spinner("Generando reportes..."):
            excel_bytes = generar_reportes_bytes(
                df,
                columnas_consolidado=columnas_sel,
                habilitar_ranking=habilitar_ranking,
                habilitar_por_sucursal=habilitar_por_sucursal,
                habilitar_matriz=habilitar_matriz,
                habilitar_evolucion=habilitar_evolucion,
                habilitar_especiales=habilitar_especiales,
                filtros_especiales=filtros_especiales,
                habilitar_estacionalidad=habilitar_estacionalidad,
            )

        ts = datetime.now().strftime("%Y%m%d_%H%M")
        filename = f"Ventas_Consolidadas_{ts}.xlsx"

        st.success("✅ Proceso completado!")

        st.download_button(
            label="⬇️ Descargar Excel Consolidado",
            data=excel_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"Error: {e}")
