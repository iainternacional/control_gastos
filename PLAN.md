# Control de Gastos — Plan del Proyecto

Proyecto para llevar el control de gastos en pareja: fácil de usar desde el celular, gratis, basado en Google.

## Decisiones

- **Plataforma**: Google Sheets (base de datos) + Google Apps Script (web app propia con HTML Service).
- **Manejo del dinero**: un solo fondo común, sin división de cuentas ni registro de quién pagó.
- **Categorías** (editables luego desde el Sheet): Mercado, Transporte, Vivienda/Servicios, Salud, Entretenimiento, Ropa, Deudas/Préstamos, Otros.
- **Funciones incluidas**: presupuesto mensual por categoría, gráficas y resumen mensual.
- **Fuera de alcance por ahora**: notificaciones/recordatorios automáticos, división de gastos entre los dos.

## Arquitectura

Un **Google Sheet** como base de datos, con 3 pestañas:

- `Gastos`: Fecha, Categoría, Descripción, Monto
- `Categorías`: lista editable de categorías
- `Presupuesto`: Categoría + límite mensual

Un **Google Apps Script** vinculado al Sheet, que sirve una mini web app (HTML Service) con dos pantallas:

1. **Registrar gasto**: formulario simple (fecha por defecto hoy, categoría en desplegable, descripción, monto), pensado para llenarse en pocos segundos desde el celular.
2. **Resumen**: gastado vs presupuesto por categoría, total del mes, y una gráfica (Google Charts, sin costo ni API key).

El código se desarrolla localmente con `clasp` y se sube al proyecto de Apps Script (`clasp push`), en lugar de escribirse desde el editor web.

## Acceso desde el celular

- El script se despliega como "Web App" (ejecuta como el dueño, acceso "cualquiera con el link" o restringido a las dos cuentas de Google). La pareja no necesita permisos sobre el Sheet, solo el link de la web app.
- Cada uno abre el link en Chrome del celular y usa "Agregar a pantalla de inicio" para que quede como ícono de app.

## Pasos de construcción

1. Crear el Google Sheet con las 3 pestañas y datos base de categorías/presupuestos.
2. Crear el proyecto Apps Script (vinculado al Sheet) localmente con `clasp create`.
3. Escribir `doGet()` + funciones de servidor: `addExpense()`, `getResumen()`, `getPresupuesto()`.
4. Construir el HTML/CSS/JS del formulario y del resumen (mobile-first, botones grandes).
5. Agregar gráfica de gastos por categoría/mes con Google Charts.
6. `clasp push` + desplegar como Web App, obtener la URL.
7. Probar en ambos celulares y agregar a pantalla de inicio.
