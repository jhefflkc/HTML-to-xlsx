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
