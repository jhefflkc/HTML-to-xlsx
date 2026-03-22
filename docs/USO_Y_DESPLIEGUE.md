# HTML to XLSX

Aplicación web local para arrastrar archivos HTML de Matrícula UNI, extraer la tabla de horarios y descargar un archivo XLSX con el mismo formato del ejemplo entregado.

## Cómo usarla

1. Abre `index.html` en tu navegador.
2. Arrastra uno o varios archivos `.html` / `.htm` al área punteada.
3. Revisa la vista previa de la tabla extraída.
4. Haz clic en **Descargar XLSX**.

## Qué hace

- Busca en el HTML la tabla con las columnas:
  - `Cod`
  - `Secc`
  - `Curso`
  - `Tipo`
  - `Horario`
  - `Aula`
  - `Docente`
- Extrae las filas del `tbody`.
- Genera un archivo `.xlsx` sin depender de librerías externas.
- Mantiene el formato general del archivo de ejemplo:
  - hoja principal `Horarios`
  - hoja secundaria `Fuente`
  - encabezado azul
  - rango tabular con autofiltro

## Notas

- Todo el procesamiento ocurre localmente en el navegador.
- Si cargas varios HTML, las filas se consolidan en un solo Excel y la hoja `Fuente` lista los archivos usados.


## Publicarla desde GitHub Actions

1. Sube estos archivos a tu repositorio en GitHub.
2. En **Settings > Pages**, selecciona **Build and deployment source: GitHub Actions**.
3. Ejecuta el workflow **Deploy web to GitHub Pages** desde la pestaña **Actions** o haz push a `main`, `master` o `work`.
4. Cuando termine, GitHub mostrará la URL pública de la web en el job de deploy.

La configuración del workflow está en `.github/workflows/deploy-pages.yml`.
