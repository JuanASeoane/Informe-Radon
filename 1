# Generador de informes de mediciones de radón

Aplicación Streamlit que, a partir del Excel de detectores de radón, genera
automáticamente el **"Informe de resultados de medicións de radon no centro
de traballo"** siguiendo el modelo oficial (UPRL / Servizo Galego de Saúde)
y la documentación de referencia del **Consello de Seguridade Nuclear (CSN)**.

## ¿Qué hace la app?

1. Carga un Excel con la hoja `Detectores` (columnas fijas, ver más abajo).
2. Permite elegir el centro de trabajo sobre el que generar el informe.
3. Solicita los datos que no vienen en el Excel (xerencia, CIF, dirección,
   superficies, información sobre trabajadores, etc.).
4. Muestra una tabla editable de resultados (para añadir, por ejemplo, la
   incertidumbre expandida y K de cada detector).
5. Genera el informe en **Word (.docx)** con:
   - Cabecera institucional con logotipos (opcionales, subibles desde la app).
   - Tabla de resultados con los valores superiores a 300 Bq/m³ resaltados
     en negrita y rojo, como exige la plantilla.
   - Conclusión generada automáticamente según si se supera o no el nivel
     de referencia (editable manualmente si se prefiere).
   - Apartado de documentación de referencia (normativa CSN, RD 732/2019,
     Ley de Prevención de Riesgos Laborales, etc.).
   - Listado de anexos (I a IV).
6. Botón de descarga del `.docx` generado.

## Formato de Excel esperado

Hoja llamada `Detectores` (o la primera hoja del libro) con estas columnas:

| Columna | Descripción |
|---|---|
| Centro | Entidad/empresa que realiza la medición |
| Área | Hospital / centro de salud (se usa por defecto para agrupar el informe) |
| ID | Identificador interno de la fila |
| Planta | Planta del edificio |
| Sala | Nombre de la sala |
| Código Sala | Código de la zona de muestreo |
| Personas trabajando en la sala | — |
| Puestos en la sala | Puestos de trabajo asociados |
| Código | Código del detector |
| Fecha de colocación | Fecha de inicio de exposición |
| Fecha de retirada real | Fecha de fin de exposición |
| Resultado medición Bq/m³ | Concentración de radón medida |
| Plano / Foto situación / Foto detector | Referencias documentales (no usadas en el cuerpo del informe) |

Si tu Excel usa otras columnas, ajusta `utils/excel_parser.py`.

## Estructura del proyecto

```
.
├── app.py                     # Interfaz Streamlit
├── utils/
│   ├── excel_parser.py        # Lectura y normalización del Excel
│   └── docx_generator.py      # Construcción del informe Word
├── requirements.txt
└── .streamlit/
    └── config.toml            # Tema opcional
```

## Ejecutar en local

```bash
python -m venv .venv
source .venv/bin/activate        # En Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

La app se abrirá en `http://localhost:8501`.

## Desplegar en Streamlit Community Cloud (vía GitHub)

1. **Crea un repositorio en GitHub** y sube este proyecto:

   ```bash
   git init
   git add .
   git commit -m "Primera versión del generador de informes de radón"
   git branch -M main
   git remote add origin https://github.com/<tu-usuario>/<tu-repo>.git
   git push -u origin main
   ```

2. Entra en **[share.streamlit.io](https://share.streamlit.io)** (Streamlit
   Community Cloud) con tu cuenta de GitHub.
3. Pulsa **"New app"**, selecciona el repositorio, la rama (`main`) y el
   fichero principal (`app.py`).
4. Pulsa **"Deploy"**. Streamlit instalará automáticamente las dependencias
   de `requirements.txt` y publicará la app en una URL tipo
   `https://<tu-app>.streamlit.app`.
5. Cada `git push` a `main` actualiza la app desplegada automáticamente.

### Notas para el despliegue

- No subas datos reales de pacientes/trabajadores ni Excels con información
  sensible al repositorio; la app no guarda datos, todo se procesa en memoria
  durante la sesión del usuario.
- Si el repositorio es público, cualquiera con la URL puede usar la app. Para
  restringir el acceso, usa un repositorio privado (Streamlit Cloud lo admite
  en el plan gratuito con límite de colaboradores) o activa la autenticación
  de Streamlit Cloud (Google/SAML según el plan).

## Personalización

- **Logotipos institucionales**: se suben desde la propia app (no se
  incluyen en el repositorio para evitar problemas de derechos de imagen).
  Si prefieres logotipos fijos, colócalos en `assets/` y ajusta
  `utils/docx_generator.py` (`_build_header`) para cargarlos por defecto.
- **Normativa de referencia**: la lista de documentos del apartado 9 está en
  `utils/docx_generator.py` (`REFERENCIAS`). Actualízala si cambia la
  normativa aplicable.
- **Nivel de referencia**: actualmente 300 Bq/m³ (`NIVEL_REFERENCIA_BQ_M3`
  en `utils/docx_generator.py`), conforme al RD 732/2019 y la Guía de
  Seguridad del CSN 11.4.
