# README de Cambios - Versión 2.3

## 🎯 Cambios Principales

### ✅ Fila 2 con Explicaciones
Ahora **todas las 7 hojas** tienen:
- ✔️ Fila 2 con explicaciones descriptivas
- ✔️ Fondo gris (#D3D3D3)
- ✔️ Texto itálico, tamaño 9
- ✔️ Centrado y ajuste de texto
- ✔️ Alto automático (30px)

### ✅ Porcentajes en Estacionalidad
- Formato `0.0%` en todas las columnas de índices
- Ejemplos: 1.13 → 113.0%, 0.95 → 95.0%

### ✅ Índices por Mes
Nuevas columnas en Estacionalidad:
- `INDICE_ENERO`
- `INDICE_FEBRERO`
- ... hasta `INDICE_DICIEMBRE`

### ✅ Colores Condicionales en INDICE_PICO
- 🟢 **VERDE (>110%)**: Producto muy estacional
- 🟡 **AMARILLO (90-110%)**: Distribución normal
- 🔴 **ROJO (<90%)**: Bajo pico

---

## 📝 Hojas Afectadas

### 1️⃣ Consolidado
- Explicaciones para: IdArticulo, Marca, Descripcion, Departamento, SubFamilia, Familia
- Meses se explican automáticamente

### 2️⃣ Ranking de Ventas
- Explicaciones para: IdArticulo, Marca, Descripcion, Total Vendido

### 3️⃣ Por Sucursal
- Explicaciones para: SUCURSAL, TOTAL
- Meses se explican automáticamente

### 4️⃣ Matriz por Departamento
- Explicaciones para: Departamento, CORRIENTES, HIPER, TOTAL

### 5️⃣ Evolución Mensual
- Explicaciones para: Departamento
- Meses se explican automáticamente

### 6️⃣ Categorías Especiales
- Explicaciones iguales a Consolidado
- Meses se explican automáticamente

### 7️⃣ Estacionalidad
- Explicaciones para todas las columnas
- Formato porcentaje en índices
- Colores condicionales en INDICE_PICO
- Índices por mes: INDICE_ENERO, INDICE_FEBRERO, etc.

---

## 🔧 Cambios Técnicos

### Nueva Función: `_aplicar_estilos_fila2_generica()`
```python
def _aplicar_estilos_fila2_generica(ws, encabezados, explicaciones_dict, 
                                   num_filas_datos, meses_para_explicar=None)
```
- Aplica estilos uniformes a la fila 2
- Genera explicaciones automáticas para meses
- Reutilizable en todas las hojas

### Nueva Función: `_aplicar_formato_porcentaje_estacionalidad()`
```python
def _aplicar_formato_porcentaje_estacionalidad(ws, meses_nombres)
```
- Aplica formato `0.0%` a columnas de índices
- Aplica formato condicional 3 colores

### Diccionarios Nuevos
- `EXPLICACIONES_CONSOLIDADO`
- `EXPLICACIONES_RANKING`
- `EXPLICACIONES_POR_SUCURSAL`
- `EXPLICACIONES_MATRIZ_DEPTO`
- `EXPLICACIONES_EVOLUCION`
- `EXPLICACIONES_ESTACIONALIDAD`

---

## 📦 Archivos Generados

```
core_consolidacion.py   (principal - 400+ líneas de lógica)
commit.sh              (script para hacer commit)
push.sh                (script para hacer push)
status.sh              (script para ver estado)
setup.sh               (script para setup inicial)
CAMBIOS.md             (este archivo)
```

---

## 🚀 Cómo Usar

### 1. Reemplazar el archivo actual
```bash
# Respalda el original
cp core_consolidacion.py core_consolidacion.py.bak

# Copia el nuevo
cp core_consolidacion_v2.3.py core_consolidacion.py
```

### 2. Commitear cambios
```bash
# Opción 1: Script automático
chmod +x commit.sh
./commit.sh

# Opción 2: Manual
git add core_consolidacion.py
git commit -m "feat: agregar fila 2 con explicaciones + porcentajes + índices mensuales"
```

### 3. Hacer push a GitHub
```bash
chmod +x push.sh
./push.sh

# O manual:
git push origin main
```

### 4. Ver estado
```bash
chmod +x status.sh
./status.sh
```

---

## ✅ Testing

Antes de hacer push, verifica en Excel que:

- [ ] Fila 2 tiene explicaciones en todas las hojas
- [ ] Fila 2 tiene fondo gris
- [ ] Estacionalidad muestra porcentajes (113.0%, 95.0%, etc.)
- [ ] INDICE_PICO tiene colores: verde/amarillo/rojo
- [ ] Todos los INDICE_* están en formato porcentaje
- [ ] No hay errores al generar el Excel

---

## 🐛 Notas Importantes

1. **openpyxl >= 3.0** requerido para CellIsRule
2. Los colores son automáticos (no requieren configuración manual)
3. La fila 2 se genera automáticamente en cada hoja
4. Los índices por mes se generan solo si existen datos de ese mes

---

## 📊 Ejemplo de Salida

### Fila 2 en Consolidado
```
Row 2: "Código único del producto" | "Nombre del fabricante" | ... | "Nombre comercial del producto"
```

### Estacionalidad con Porcentajes
```
INDICE_PICO: 113.0% 🟢
INDICE_ENERO: 45.2% 🔴
INDICE_MARZO: 112.7% 🟢
```

---

## 📞 Soporte

Para problemas:
1. Verifica requirements.txt (openpyxl, pandas)
2. Revisa los logs en `debug.log` si existe
3. Asegúrate de usar Python 3.8+

---

**Versión**: 2.3  
**Fecha**: Diciembre 2025  
**Desarrollado para**: Cucher Mercados - Resistencia, Chaco
