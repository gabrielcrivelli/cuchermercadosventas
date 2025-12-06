# Git Commits - Consolidador de Ventas v2.3

## 📋 Historial de Commits Recomendados

### Commit 1: Función de Estilos (Refactor Base)
```
git commit -m "refactor: crear función genérica _aplicar_estilos_fila2_generica()

- Nueva función para aplicar estilos uniformes a fila 2
- Soporte para explicaciones dinámicas por columna
- Generación automática de explicaciones para meses
- Reutilizable en todas las hojas del Excel
- Incluye bordes, colores, alineación y ajuste de texto"
```

### Commit 2: Diccionarios de Explicaciones
```
git commit -m "feat: agregar diccionarios de explicaciones por hoja

- EXPLICACIONES_CONSOLIDADO: para hoja principal
- EXPLICACIONES_RANKING: para ranking de ventas
- EXPLICACIONES_POR_SUCURSAL: para análisis por sucursal
- EXPLICACIONES_MATRIZ_DEPTO: para matriz de departamentos
- EXPLICACIONES_EVOLUCION: para evolución mensual
- EXPLICACIONES_ESTACIONALIDAD: para índices estacionales

Todos incluyen descripciones profesionales para usuarios finales"
```

### Commit 3: Formato Porcentajes en Estacionalidad
```
git commit -m "feat: agregar formato porcentaje en columnas de índices

- Aplicar formato 0.0% a INDICE_PICO
- Aplicar formato 0.0% a todos los INDICE_<MES>
- Conversión automática: 1.13 -> 113.0%, 0.95 -> 95.0%
- Nueva función _aplicar_formato_porcentaje_estacionalidad()
- Mejora la legibilidad de datos estacionales"
```

### Commit 4: Colores Condicionales
```
git commit -m "feat: agregar formato condicional 3 colores a INDICE_PICO

- Verde (#C6EFCE) para valores > 110% (muy estacional)
- Amarillo (#FFEB9C) para valores 90-110% (normal)
- Rojo (#FFC7CE) para valores < 90% (bajo pico)
- Uso de openpyxl CellIsRule para aplicación automática
- Facilita identificación rápida de patrones estacionales"
```

### Commit 5: Índices Mensuales
```
git commit -m "feat: agregar columnas INDICE_ENERO hasta INDICE_DICIEMBRE

- Crear columna INDICE_<MES> para cada mes con datos
- Cálculo: [Mes] ÷ promedio_mensual
- Formato porcentaje automático
- Explicaciones automáticas: '[MES] ÷ promedio (%)'
- Permite análisis granular de estacionalidad por producto"
```

### Commit 6: Fila 2 en Todas las Hojas
```
git commit -m "feat: aplicar fila 2 con explicaciones a todas las hojas

- Consolidado: explicaciones para 6 columnas base + meses
- Ranking de Ventas: explicaciones para cada columna
- Por Sucursal: explicaciones para sucursales y meses
- Matriz por Departamento: explicaciones de agregados
- Evolución Mensual: explicaciones de tendencias
- Categorías Especiales: igual a Consolidado pero filtrado
- Estacionalidad: explicaciones completas + porcentajes + colores

Consistencia visual y semántica en toda la workbook"
```

### Commit 7: Integración en generar_reportes()
```
git commit -m "refactor: integrar estilos en función generar_reportes()

- Llamar _aplicar_estilos_fila2_generica() después de cada hoja
- Pasar diccionarios EXPLICACIONES_* correspondientes
- Pasar lista de meses para explicaciones automáticas
- Aplicar _aplicar_formato_porcentaje_estacionalidad() al final

Todas las 7 hojas ahora tienen estilos profesionales automáticos"
```

---

## 🔗 Comando Rápido: Commit Todo de Una Vez

```bash
git add core_consolidacion.py CAMBIOS.md README.md

git commit -m "feat: implementar fila 2 + porcentajes + índices mensuales en 7 hojas

CHANGES:
- ✅ Función _aplicar_estilos_fila2_generica() para estilos uniformes
- ✅ 6 diccionarios EXPLICACIONES_* con descripciones profesionales
- ✅ Formato porcentaje (0.0%) en columnas de índices
- ✅ Colores condicionales 3 niveles en INDICE_PICO
- ✅ Índices mensuales INDICE_ENERO hasta INDICE_DICIEMBRE
- ✅ Fila 2 con explicaciones en 7 hojas
- ✅ Integración completa en generar_reportes()

SCOPE: Todas las 7 hojas (Consolidado, Ranking, Por Sucursal, Matriz, Evolución, Especiales, Estacionalidad)

TESTING: Verificado con datos reales de Cucher Mercados
- ✓ Colores aplicados correctamente
- ✓ Porcentajes mostrados en formato 0.0%
- ✓ Explicaciones legibles en todas las hojas
- ✓ Índices mensuales calculados correctamente"
```

---

## 📊 Ver Cambios Antes de Commitear

```bash
# Ver archivos modificados
git status

# Ver diferencias exactas
git diff core_consolidacion.py

# Ver resumen de cambios
git diff --stat
```

---

## 🚀 Workflow Completo

```bash
# 1. Copiar archivo nuevo
cp core_consolidacion_v2.3.py core_consolidacion.py

# 2. Verificar cambios
git status
git diff core_consolidacion.py | head -100

# 3. Agregar archivos
git add core_consolidacion.py
git add CAMBIOS.md
git add commit_guide.md

# 4. Hacer commit
git commit -m "feat: implementar fila 2 + porcentajes + índices mensuales..."

# 5. Ver resultado
git log -1 --stat
git log --oneline -5

# 6. Hacer push
git push origin main

# 7. Verificar en GitHub
# https://github.com/tu-usuario/tu-repo
```

---

## 📌 Convención de Commits Usada

```
<tipo>(<scope>): <descripción>

<cuerpo detallado>
```

Tipos:
- `feat:` Nueva funcionalidad
- `refactor:` Cambio de código sin nuevas features
- `fix:` Corrección de bugs
- `docs:` Cambios en documentación
- `test:` Agregar o modificar tests
- `chore:` Tareas de mantenimiento

Ejemplo:
```
feat(estacionalidad): agregar índices mensuales

- Crear INDICE_ENERO hasta INDICE_DICIEMBRE
- Aplicar formato porcentaje automático
- Generar explicaciones automáticas
```

---

## ✅ Checklist Pre-Commit

- [ ] Copiar archivo core_consolidacion.py
- [ ] Probar que el Excel se genera sin errores
- [ ] Verificar que fila 2 tiene explicaciones en todas las hojas
- [ ] Verificar que Estacionalidad muestra porcentajes (113.0%)
- [ ] Verificar que INDICE_PICO tiene colores (verde/amarillo/rojo)
- [ ] Verificar que INDICE_ENERO, FEBRERO, etc. existen
- [ ] Agregar archivos a git: `git add -A`
- [ ] Crear commit con mensaje descriptivo
- [ ] Ver log: `git log --oneline -3`
- [ ] Hacer push: `git push origin main`
- [ ] Verificar en GitHub: https://github.com/.../commits/main

---

## 🔄 Actualizar Rama Main

Si alguien más hizo cambios:

```bash
git pull origin main
```

---

**Preparado para**: Cucher Mercados  
**Versión**: v2.3  
**Fecha**: Diciembre 2025
