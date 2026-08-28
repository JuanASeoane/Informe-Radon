#!/data/data/com.termux/files/usr/bin/bash
# ============================================================
# setup.sh - Instalación de la app "Detectores de Radón" en Termux
# ============================================================
# Ejecuta esto UNA SOLA VEZ la primera vez. Deja app.py,
# requirements.txt, start.sh y stop.sh en la carpeta Descargas
# del móvil antes de lanzarlo.
#
# Uso:
#   bash setup.sh
# ============================================================
set -e

echo "== 1/6  Actualizando paquetes de Termux =="
pkg update -y && pkg upgrade -y

echo "== 2/6  Instalando Python y librerías necesarias =="
pkg install -y python libjpeg-turbo zlib freetype git termux-api

echo "== 3/6  Dando acceso a la carpeta Descargas del móvil =="
termux-setup-storage
sleep 2

APP_DIR="$HOME/radon"
mkdir -p "$APP_DIR" "$APP_DIR/.streamlit"
DOWNLOADS="$HOME/storage/downloads"

echo "== 4/6  Copiando los archivos de la app =="
for f in app.py requirements.txt start.sh stop.sh favicon.png fondo_app.jpg logo_laboratorio_default.png; do
    if [ -f "$DOWNLOADS/$f" ]; then
        cp "$DOWNLOADS/$f" "$APP_DIR/"
        echo "  Copiado: $f"
    else
        echo "  AVISO: no se encontró $DOWNLOADS/$f"
        echo "         Cópialo manualmente a $APP_DIR y vuelve a ejecutar setup.sh"
    fi
done
# Módulo del informe de resultados completo (Word/PDF oficial): coloca en
# Descargas las carpetas "utils_informe" y "assets_informe" tal cual,
# con todos sus archivos dentro, antes de ejecutar este script.
for carpeta in utils_informe assets_informe; do
    if [ -d "$DOWNLOADS/$carpeta" ]; then
        cp -r "$DOWNLOADS/$carpeta" "$APP_DIR/"
        echo "  Copiada carpeta: $carpeta"
    else
        echo "  AVISO: no se encontró la carpeta $DOWNLOADS/$carpeta"
        echo "         El botón \"Informe de resultados completo\" no funcionará sin ella."
    fi
done
if [ -f "$DOWNLOADS/config.toml" ]; then
    cp "$DOWNLOADS/config.toml" "$APP_DIR/.streamlit/"
    echo "  Copiado: .streamlit/config.toml (tema naranja)"
else
    echo "  AVISO: no se encontró $DOWNLOADS/config.toml"
    echo "         Cópialo manualmente a $APP_DIR/.streamlit/ para el tema naranja"
fi
chmod +x "$APP_DIR/start.sh" "$APP_DIR/stop.sh" 2>/dev/null || true

cd "$APP_DIR"
echo "== 5/6  Instalando dependencias de Python (puede tardar varios minutos) =="
pip install --upgrade pip
pip install -r requirements.txt

echo "== 6/6  Creando acceso directo para Termux:Widget =="
mkdir -p "$HOME/.shortcuts"
cat > "$HOME/.shortcuts/Radon.sh" << 'EOF'
#!/data/data/com.termux/files/usr/bin/bash
bash "$HOME/radon/start.sh"
EOF
cat > "$HOME/.shortcuts/Radon (detener).sh" << 'EOF'
#!/data/data/com.termux/files/usr/bin/bash
bash "$HOME/radon/stop.sh"
EOF
chmod +x "$HOME/.shortcuts/Radon.sh" "$HOME/.shortcuts/Radon (detener).sh"

echo
echo "============================================================"
echo " ¡Instalación completada!"
echo "============================================================"
echo " Para arrancar la app ahora mismo, ejecuta:"
echo "     bash $APP_DIR/start.sh"
echo
echo " Para tener un icono en la pantalla de inicio del móvil:"
echo "   1. Instala la app 'Termux:Widget' desde F-Droid."
echo "   2. Mantén pulsado en la pantalla de inicio > Widgets >"
echo "      Termux:Widget > añádelo a la pantalla."
echo "   3. En el widget aparecerán 'Radon' (arrancar/abrir) y"
echo "      'Radon (detener)' (parar el servidor)."
echo "============================================================"
