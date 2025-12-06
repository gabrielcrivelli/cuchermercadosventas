#!/bin/bash

# Script para ver el estado completo del repositorio
# Uso: ./status.sh

echo "==================================================================="
echo "Cucher Mercados - Estado del Repositorio"
echo "==================================================================="
echo ""

# 1. Información general
echo "📊 Información del Repositorio"
echo "--------------------------------"
echo "Rama actual: $(git branch --show-current)"
echo "Usuario git: $(git config user.name)"
echo "Email: $(git config user.email)"
echo ""

# 2. Estado de los archivos
echo "📁 Estado de los Archivos"
echo "--------------------------------"
git status

echo ""
echo "📝 Último commit"
echo "--------------------------------"
git log -1 --pretty=format:"%h - %an (%ar): %s"
echo ""

echo ""
echo "📈 Commit log (últimos 5)"
echo "--------------------------------"
git log --oneline -5

echo ""
echo "🔗 Remoto (Origin)"
echo "--------------------------------"
git remote -v

echo ""
echo "✅ Reporte completado"
