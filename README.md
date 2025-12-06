# 📊 Consolidador de Ventas - Cucher Mercados v2.3

> Sistema avanzado de consolidación y análisis de ventas con 7 hojas Excel profesionales, fila 2 con explicaciones, porcentajes automáticos e índices estacionales.

**Versión**: 2.3  
**Última actualización**: Diciembre 2025  
**Desarrollado para**: Cucher Mercados - Resistencia, Chaco, Argentina

---

## ✨ Novedades v2.3

### 🎯 Lo Nuevo Ahora
1. **Fila 2 con Explicaciones en TODAS las 7 hojas**
   - Explicaciones profesionales automáticas
   - Estilos uniformes: gris, itálico, centrado
   - Generación automática para columnas de meses

2. **Formato Porcentaje en Estacionalidad**
   - Muestra `113.0%` en lugar de `1.13`
   - Aplicado a INDICE_PICO e INDICE_<MES>
   - Formato dinámico según datos disponibles

3. **Colores Condicionales 3 Niveles**
   - 🟢 Verde: >110% (muy estacional)
   - 🟡 Amarillo: 90-110% (normal)
   - 🔴 Rojo: <90% (bajo pico)

4. **Índices Mensuales Nuevos**
   - INDICE_ENERO, INDICE_FEBRERO, ..., INDICE_DICIEMBRE
   - Cálculo: Mes ÷ Promedio Mensual
   - Formato porcentaje automático

---

## 📋 Las 7 Hojas

| # | Hoja | Propósito | Fila 2 | Colores |
|---|------|----------|--------|---------|
| 1️⃣ | **Consolidado** | Matriz producto × mes × sucursal | ✅ Sí | ❌ No |
| 2️⃣ | **Ranking de Ventas** | Top productos ordenados | ✅ Sí | ❌ No |
| 3️⃣ | **Por Sucursal** | Resumen por sucursal y mes | ✅ Sí | ❌ No |
| 4️⃣ | **Matriz** | Depto × Sucursal | ✅ Sí | ❌ No |
| 5️⃣ | **Evolución Mensual** | Depto × Mes (tendencias) | ✅ Sí | ❌ No |
| 6️⃣ | **Categorías Especiales** | Filtrado por depto/marca | ✅ Sí | ❌ No |
| 7️⃣ | **Estacionalidad** | Índices y mes pico | ✅ Sí | ✅ Sí |

---

## 🚀 Quick Start

### 1. Instalar Dependencias
```bash
pip install -r requirements.txt
```

Requiere:
- `pandas >= 1.3.0`
- `openpyxl >= 3.0.0` (para estilos y colores)

### 2. Usar en Python
```python
from core_consolidacion import consolidar_datos, generar_reportes

# Paso 1: Consolidar datos desde archivos Excel
archivos_info = [
    {
        "ruta": "datos/Marzo_2025_HIPER.xlsx",
        "mes": "MARZO",
        "anio": 2025,
        "sucursal": "HIPER"
    },
    {
        "ruta": "datos/Marzo_2025_CORRIENTES.xlsx",
        "mes": "MARZO",
        "anio": 2025,
        "sucursal": "CORRIENTES"
    },
    # ... más archivos
]

df = consolidar_datos(archivos_info)

# Paso 2: Generar reportes (todas las hojas)
generar_reportes(
    df,
    ruta_salida="Ventas_Consolidadas_Final.xlsx",
    habilitar_ranking=True,
    habilitar_por_sucursal=True,
    habilitar_matriz=True,
    habilitar_evolucion=True,
    habilitar_especiales=True,
    habilitar_estacionalidad=True  # Nuevo en v2.3
)
```

### 3. Ejecutar GUI
```bash
python ventas_consolidator_gui.py
```

---

## 📊 Ejemplo de Salida

### Consolidado (Fila 2)
```
Row 1: IdArticulo | Marca | Descripcion | ... | TOTAL CORRIENTES | TOTAL HIPER
Row 2: Código único del producto | Nombre del fabricante | Nombre comercial | ... | Unidades en Corrientes | Unidades en Hipermercado
Row 3: 7 | Cañuelas | Ac Cañuelas | ... | 21705 | 16373
```

### Estacionalidad (Con Porcentajes)
```
Row 1: IdArticulo | Marca | VENTA_TOTAL_ANUAL | PROMEDIO_MENSUAL | MES_PICO | INDICE_PICO | INDICE_ENERO | ... | INDICE_DICIEMBRE
Row 2: Código único | Fabricante | Total anual | Promedio por mes | Mes pico | Mes pico ÷ promedio (%) | Enero ÷ promedio (%) | ... | Diciembre ÷ promedio (%)
Row 3: 7 | Cañuelas | 38.078 | 3.173 | MARZO | 113.0% 🟢 | 32.1% 🔴 | ... | 112.8% 🟢
```

---

## 🔧 Archivos del Proyecto

```
├── core_consolidacion.py          # Core principal (ACTUALIZADO v2.3)
├── ventas_consolidator_gui.py     # Interfaz gráfica
├── requirements.txt               # Dependencias
├── README.md                      # Este archivo
├── CAMBIOS.md                     # Detalles de cambios v2.3
├── commit_guide.md                # Guía para commits en GitHub
│
├── scripts/                       # Scripts auxiliares
│   ├── commit.sh                  # Commit automático
│   ├── push.sh                    # Push a GitHub
│   ├── status.sh                  # Ver estado repo
│   └── setup.sh                   # Setup inicial
│
└── docs/
    ├── api.md                     # Documentación API
    └── ejemplos.md                # Ejemplos de uso
```

---

## 🎯 Cambios v2.3 - Detalles Técnicos

### Nueva Función: `_aplicar_estilos_fila2_generica()`

Aplica estilos uniformes a la fila 2 en cualquier hoja:

```python
def _aplicar_estilos_fila2_generica(ws, encabezados, explicaciones_dict, 
                                   num_filas_datos, meses_para_explicar=None):
    """
    Args:
        ws: Worksheet de openpyxl
        encabezados: Lista de nombres de columnas (row 1)
        explicaciones_dict: Dict {nombre_columna: explicación}
        num_filas_datos: Cantidad de filas de datos
        meses_para_explicar: Lista de meses (ej: MESES_ES)
    """
```

**Características:**
- Fondo gris (#D3D3D3)
- Texto itálico, tamaño 9, centrado
- Ajuste de texto automático
- Alto de fila 30px
- Bordes finos en datos
- Explicaciones automáticas para meses

### Nueva Función: `_aplicar_formato_porcentaje_estacionalidad()`

Aplica formato porcentaje y colores a Estacionalidad:

```python
def _aplicar_formato_porcentaje_estacionalidad(ws, meses_nombres):
    """
    Aplica:
    - Formato 0.0% a INDICE_PICO e INDICE_<MES>
    - Verde: > 110%
    - Amarillo: 90-110%
    - Rojo: < 90%
    """
```

### Diccionarios de Explicaciones

```python
# Consolidado + Categorías Especiales
EXPLICACIONES_CONSOLIDADO = {
    'IdArticulo': 'Código único del producto',
    'Marca': 'Nombre del fabricante',
    # ... 6 columnas base
}

# Ranking
EXPLICACIONES_RANKING = {
    'IdArticulo': 'Código único del producto',
    'Marca': 'Nombre del fabricante',
    'Descripcion': 'Nombre comercial del producto',
    'Total Vendido': 'Unidades totales vendidas (suma todas sucursales y meses)',
}

# Por Sucursal
EXPLICACIONES_POR_SUCURSAL = {
    'SUCURSAL': 'Nombre de la sucursal',
    'TOTAL': 'Unidades vendidas en todo el período',
}

# Matriz Departamento
EXPLICACIONES_MATRIZ_DEPTO = {
    'Departamento': 'Categoría principal de venta',
    'CORRIENTES': 'Unidades del departamento en Corrientes (período)',
    'HIPER': 'Unidades del departamento en Hipermercado (período)',
    'TOTAL': 'Sumatoria Corrientes + Hiper (período)',
}

# Evolución Mensual
EXPLICACIONES_EVOLUCION = {
    'Departamento': 'Categoría principal de venta',
}

# Estacionalidad
EXPLICACIONES_ESTACIONALIDAD = {
    'IdArticulo': 'Código único del producto',
    'Marca': 'Nombre del fabricante',
    'VENTA_TOTAL_ANUAL': 'Ventas totales en el período',
    'PROMEDIO_MENSUAL': 'Venta promedio por mes',
    'MES_PICO': 'Mes con mayor venta',
    'VENTA_PICO': 'Unidades en mes pico',
    'INDICE_PICO': 'Mes pico ÷ promedio (%)',
}
```

---

## 🔄 Flujo de Datos

```
Archivos Excel (mes/sucursal)
        ↓
consolidar_datos()
        ↓
DataFrame consolidado (IdArticulo + MES + SUCURSAL + Cantidad)
        ↓
generar_reportes()
        ├─ Consolidado (pivot IdArticulo × MES_SUC)
        ├─ Ranking (sort por Total Vendido)
        ├─ Por Sucursal (pivot SUCURSAL × MES)
        ├─ Matriz (pivot Departamento × SUCURSAL)
        ├─ Evolución Mensual (pivot Departamento × MES)
        ├─ Categorías Especiales (filtrado + pivot)
        └─ Estacionalidad (índices + colores)
        ↓
Excel único con 7 hojas (todas formateadas)
```

---

## ✅ Checklist Pre-Deploy

- [ ] Instalar `openpyxl >= 3.0` (para CellIsRule)
- [ ] Probar generación con datos reales
- [ ] Verificar fila 2 en todas las 7 hojas
- [ ] Verificar porcentajes en Estacionalidad (113.0%)
- [ ] Verificar colores en INDICE_PICO
- [ ] Verificar INDICE_ENERO, FEBRERO, etc. existen
- [ ] Sin errores en logs
- [ ] Archivo Excel abre correctamente en Office/Calc

---

## 📦 Commit en GitHub

### Opción 1: Script Automático
```bash
chmod +x commit.sh
./commit.sh
```

### Opción 2: Manual
```bash
git add core_consolidacion.py CAMBIOS.md
git commit -m "feat: implementar fila 2 + porcentajes + índices mensuales

- Nueva función _aplicar_estilos_fila2_generica() para estilos uniformes
- Diccionarios de explicaciones para cada hoja
- Formato porcentaje (0.0%) en índices
- Colores condicionales 3 niveles en INDICE_PICO
- Índices mensuales INDICE_ENERO hasta INDICE_DICIEMBRE
- Fila 2 con explicaciones en 7 hojas"

git push origin main
```

---

## 🐛 Troubleshooting

### Error: `ModuleNotFoundError: No module named 'openpyxl'`
```bash
pip install openpyxl>=3.0
```

### Error: `CellIsRule not found`
- Asegúrate de tener openpyxl >= 3.0.0
- `pip show openpyxl` para verificar versión

### Los porcentajes no se muestran
- Verificar que `number_format = '0.0%'` se aplica
- Abrir Excel y recalcular: F9

### Colores no aparecen
- Verificar que `CellIsRule` está importado
- Recalcular condicionales en Excel

---

## 📊 Estadísticas de Código

```
core_consolidacion.py:
  • Líneas: 450+
  • Funciones: 8
  • Diccionarios: 6
  • Imports: pandas, openpyxl
  • Complejidad: Media (bien documentado)
```

---

## 🎓 Recursos

- [Documentación openpyxl](https://openpyxl.readthedocs.io/)
- [Pandas Pivot Table](https://pandas.pydata.org/docs/reference/api/pandas.pivot_table.html)
- [Git Documentation](https://git-scm.com/doc)

---

## 📞 Soporte

Para issues o consultas:
1. Revisar CAMBIOS.md
2. Revisar commit_guide.md
3. Ejecutar tests
4. Contactar equipo de desarrollo

---

## 📄 Licencia

Uso interno - Cucher Mercados S.A.

---

**Hecho con ❤️ para Cucher Mercados**  
*Resistencia, Chaco, Argentina*  
*Diciembre 2025*
