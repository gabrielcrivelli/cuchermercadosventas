#!/bin/bash

# Script de instalación y configuración inicial del repositorio
# Uso: ./setup.sh

echo "==================================================================="
echo "Cucher Mercados - Setup Inicial del Repositorio"
echo "==================================================================="
echo ""

# 1. Crear .gitignore si no existe
if [ ! -f .gitignore ]; then
    echo "📝 Creando .gitignore..."
    cat > .gitignore << 'EOF'
# Archivos Python
__pycache__/
*.py[cod]
*$py.class
*.so
.Python
build/
develop-eggs/
dist/
downloads/
eggs/
.eggs/
lib/
lib64/
parts/
sdist/
var/
wheels/
*.egg-info/
.installed.cfg
*.egg

# Virtual environments
venv/
env/
ENV/
env.bak/
venv.bak/

# IDE
.vscode/
.idea/
*.swp
*.swo
*~
.DS_Store

# Logs
*.log

# Archivos Excel (locales)
Ventas_Consolidadas*.xlsx
output.xlsx
test_*.xlsx

# Archivos CSV
*.csv

# Directorio temp
temp/
tmp/
.tmp/
EOF
    echo "✅ .gitignore creado"
else
    echo "✓ .gitignore ya existe"
fi

# 2. Inicializar git si no lo está
if [ ! -d .git ]; then
    echo ""
    echo "🔧 Inicializando repositorio git..."
    git init
    echo "✅ Repositorio inicializado"
else
    echo "✓ Repositorio git ya existe"
fi

# 3. Configurar usuario
echo ""
echo "👤 Configurando usuario git..."
echo "Ingresa tu nombre (ej: Cucher): "
read -r nombre
echo "Ingresa tu email (ej: cucher@mail.com): "
read -r email

git config --global user.name "$nombre"
git config --global user.email "$email"

echo "✅ Usuario configurado: $nombre <$email>"

# 4. Crear archivos de configuración
if [ ! -f .github/workflows/python-app.yml ]; then
    echo ""
    echo "📦 Creando workflows de GitHub Actions..."
    mkdir -p .github/workflows
    cat > .github/workflows/python-app.yml << 'EOF'
name: Python App

on:
  push:
    branches: [ main ]
  pull_request:
    branches: [ main ]

jobs:
  build:
    runs-on: ubuntu-latest
    steps:
    - uses: actions/checkout@v2
    - name: Set up Python
      uses: actions/setup-python@v2
      with:
        python-version: 3.9
    - name: Install dependencies
      run: |
        python -m pip install --upgrade pip
        pip install -r requirements.txt
    - name: Test with pytest
      run: |
        pip install pytest
        pytest
EOF
    echo "✅ Workflow creado"
fi

# 5. Hacer primer commit
echo ""
echo "📝 Creando primer commit..."

git add -A
git commit -m "chore: setup inicial del repositorio" 2>/dev/null

if [ $? -eq 0 ]; then
    echo "✅ Primer commit creado"
else
    echo "ℹ️  No hay cambios para commitear"
fi

# 6. Mostrar resumen
echo ""
echo "==================================================================="
echo "✅ Setup completado"
echo "==================================================================="
echo ""
echo "📋 Próximos pasos:"
echo "1. Agregar remoto: git remote add origin https://github.com/tu-usuario/tu-repo.git"
echo "2. Hacer primer push: git push -u origin main"
echo "3. Hacer cambios y commitear: ./commit.sh"
echo "4. Hacer push a GitHub: ./push.sh"
echo "5. Ver estado: ./status.sh"
echo ""
