#!/data/data/com.termux/files/usr/bin/bash
# ============================================================
# start.sh - Arranca (o reabre) la app de Detectores de Radón
# ============================================================
APP_DIR="$HOME/radon"
PORT=8501
URL="http://localhost:$PORT"

cd "$APP_DIR" 2>/dev/null || {
    echo "No se encuentra $APP_DIR. Ejecuta primero setup.sh"
    read -p "Pulsa Enter para salir..."
    exit 1
}

if curl -s -o /dev/null "$URL"; then
    echo "La app ya estaba en marcha en $URL"
else
    echo "Arrancando la app..."
    nohup streamlit run app.py --server.headless true --server.port "$PORT" \
        > "$APP_DIR/streamlit.log" 2>&1 &

    echo "Esperando a que arranque..."
    for i in $(seq 1 40); do
        sleep 1
        if curl -s -o /dev/null "$URL"; then
            echo "¡Lista!"
            break
        fi
    done
fi

# Abrir el navegador automáticamente si está disponible termux-api;
# si no, se muestra la URL para abrirla a mano.
if command -v termux-open-url >/dev/null 2>&1; then
    termux-open-url "$URL"
else
    echo "Abre esta dirección en Chrome: $URL"
fi
