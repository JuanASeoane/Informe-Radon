# Generador de informes de mediciones de radón

Aplicación Streamlit que, a partir del Excel de detectores de radón, genera
automáticamente el **"Informe de resultados de medicións de radon no centro
de traballo"** siguiendo el modelo oficial (UPRL / Servizo Galego de Saúde)
y la documentación de referencia del **Consello de Seguridade Nuclear (CSN)**.

## ¿Qué hace la app?

1. Carga un Excel con las hojas `Detectores`, `Planos` y `Categorías
   profesionales` (ver formato más abajo).
2. Permite elegir el centro de trabajo sobre el que generar el informe
   (columna `Centro`).
3. **Autorrellena** los campos del formulario a partir del propio Excel:
   - Xerencia y CIF ← hoja `Planos` (Empresa / CIF).
   - Servizo/Unidade mostrexada ← columna `Área`.
   - Dirección, técnico y fecha ← metadatos de la hoja `Detectores`.
   - Descripción de los puestos de trabajo ← salas medidas (columna `Sala`).
   - N.º de trabajadores ← hoja `Categorías profesionales` (total y desglose
     por categoría).

   Todos los campos autorrellenados son editables antes de generar el informe.
4. Muestra la tabla de resultados del punto 7 **semicubierta**: código de
   zona, código de detector, fechas y puesto de trabajo ya vienen del Excel;
   el **resultado (Bq/m³) y la incertidumbre** se completan a mano en la
   propia tabla editable, o automáticamente subiendo un **Excel de
   resultados del laboratorio** (con una columna `Código` que haga de clave
   de cruce con el detector).
5. Genera el informe en **Word (.docx)** con:
   - Cabecera institucional con un único logotipo a todo el ancho (opcional,
     subible desde la app) y más separación respecto al cuerpo del informe.
   - Tabla de resultados con los valores superiores a 300 Bq/m³ resaltados
     en negrita y rojo, como exige la plantilla.
   - Conclusión generada automáticamente según si se supera o no el nivel
     de referencia (editable manualmente si se prefiere).
   - Apartado de documentación de referencia (normativa CSN, RD 732/2019,
     Ley de Prevención de Riesgos Laborales, etc.).
   - Listado de anexos (I a IV).
6. Botón de descarga del `.docx` generado.

## Formato de Excel esperado

El libro debe tener tres hojas:

### Hoja `Detectores`

- **Fila 1**: metadatos sueltos del centro (Centro, Área, "Técnico" + nombre,
  "Fecha" + fecha, "Dirección" + dirección). No es necesario que estén
  alineados con las columnas de la tabla; la app busca cada etiqueta
  ("Técnico", "Fecha", "Dirección") y toma el valor que tenga al lado.
- **Fila 2**: cabecera real de la tabla de datos.
- **Fila 3 en adelante**: una fila por detector, con columnas como:

  | Columna | Descripción | Uso en el informe |
  |---|---|---|
  | Centro | Centro de trabajo | Punto 1 · CENTRO |
  | Área | Servicio/unidad muestreada | Punto 1 · SERVIZO/UNIDADE MOSTREXADA |
  | Sala | Nombre de la sala | Punto 3 · listado de puestos de trabajo |
  | Código de la sala | Código de zona de muestreo | Punto 7 · 1ª columna |
  | Profesionales en la sala | Puesto/s de trabajo asociados | Punto 7 · última columna |
  | Código | Código del detector | Punto 7 · 2ª columna |
  | Fecha de colocación | Inicio de exposición | Punto 7 |
  | Fecha de retirada real | Fin de exposición | Punto 7 |
  | Resultado (Bq/m³) | Concentración medida (puede venir vacía) | Punto 7 |
  | Incertidumbre | Incertidumbre expandida y K (puede venir vacía) | Punto 7 |

### Hoja `Planos`

- Fila 1: `Empresa` | `<nombre de la xerencia/empresa>`
- Fila 2: `CIF` | `<CIF>`
- Filas siguientes: tabla `Nombre` / `Imagen` con los planos de cada planta
  (contenido informativo para el Anexo II; no se inserta automáticamente en
  el cuerpo del informe).

### Hoja `Categorías profesionales`

| Categoría profesional | Nº personas expuestas |
|---|---|
| Médico | 2 |
| Enfermería | 2 |
| ... | ... |

Se usa para autorrellenar el número total de trabajadores y su desglose en
el punto 3 del informe.

### Excel de resultados del laboratorio (opcional)

Cualquier Excel con, al menos, una columna `Código` (con el mismo código de
detector que en la hoja `Detectores`) y columnas cuyo nombre empiece por
`Resultado` y/o `Incertidumbre`. Al subirlo, la app cruza los valores por
código de detector y completa la tabla del punto 7 automáticamente.

Si tu Excel usa columnas distintas, ajusta `utils/excel_parser.py`.

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
