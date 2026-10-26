# Mis Recetas — planificador semanal

Aplicación web estática para cargar un Excel de recetas, filtrar comidas y organizar un menú semanal. El contenido del archivo se procesa localmente en el navegador; no se envía a un backend.

## Requisitos
- Node.js 20 o posterior recomendado
- Un Excel `.xlsx` o `.xls` con una sola hoja (también admite CSV)

## Encabezados esperados
`Comida`, `Categoría`, `Temporada`, `Ingredientes`, `Calorías`, `Proteína (g)`, `Carbohidratos (g)`, `Grasas (g)`, `Notas`, `Link a receta`

Las columnas principales para cargar son `Comida`, `Categoría`, `Temporada`, `Ingredientes` y `Calorías`. Las demás pueden estar vacías, aunque se recomienda usar los encabezados indicados para que todos los filtros funcionen. Los nombres de encabezado no distinguen entre mayúsculas y minúsculas ni entre tildes. La app lee la primera hoja del archivo.

## Ejecutar en local
```bash
npm install
npm run dev
```
Abre la URL local que muestre Vite.

## Publicar en Netlify
1. Descomprime el ZIP y sube la carpeta a un repositorio Git (GitHub, GitLab o Bitbucket).
2. En Netlify, selecciona **Add new site → Import an existing project** y conecta el repositorio.
3. Configura:
   - Build command: `npm run build`
   - Publish directory: `dist`
4. Pulsa **Deploy**.

Netlify también puede desplegarlo desde su CLI. El archivo `netlify.toml` ya incluye el comando y el directorio de publicación.

## Funcionalidad
- Filtros combinables por categoría, temporada, calorías máximas, proteína mínima, carbohidratos máximos, grasas máximas, notas y texto en nombre/ingredientes.
- Menú de lunes a domingo, con desayuno, comida y cena. Se puede añadir más de una receta a cada franja.
- Navegación entre semanas.
- Resumen de nutrición calculado como suma del menú dividido entre 7 (promedio diario por semana).
- Enlaces a recetas si `Link a receta` contiene una URL `http://` o `https://`.
- Exportación del menú seleccionado como CSV compatible con Excel.
- Datos de ejemplo para probar sin subir archivo.

## Limitaciones de esta primera versión
- El menú se mantiene en la memoria de la pestaña y se pierde al recargar o cerrar la página.
- No se guardan archivos ni menús en una cuenta o base de datos.
- Los valores nutricionales son los del Excel; la aplicación no los verifica.
- Los filtros nutricionales tratan las celdas vacías como cero.
