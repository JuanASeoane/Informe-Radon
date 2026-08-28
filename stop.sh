#!/data/data/com.termux/files/usr/bin/bash
# ============================================================
# stop.sh - Detiene la app de Detectores de Radón si está en marcha
# ============================================================
if pkill -f "streamlit run app.py" 2>/dev/null; then
    echo "App detenida."
else
    echo "La app no estaba en marcha."
fi
read -p "Pulsa Enter para cerrar..."
