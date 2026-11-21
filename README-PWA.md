# Panel de Seguimiento — PWA

Aplicación web para darle seguimiento a la **merma pendiente**, las **transferencias pendientes** y el **stock bloqueado** de las sucursales de La Paz, BCS, a partir de los archivos Excel que llegan por correo.

Los datos viven **solo en el dispositivo** donde se usa (localStorage). Nada se sube a ningún servidor.

## Qué hace

- **Alimentar los paneles:** sube los archivos .xlsx y la app detecta sola el tipo por sus columnas:
  - **Merma:** total pendiente, por sucursal (monto, registros, entrada más antigua, días, nivel CRÍTICO / SEGUIMIENTO / CONTROLADO y acción sugerida), distribución por antigüedad (0-3, 4-7, 8-15, 16-30, +30 días) y evolución por fecha de entrada.
  - **Transferencias (formato DESTINO):** documentos, piezas embarcadas y pendientes, antigüedad, por sucursal origen y destino, documento más antiguo. Acepta fechas como fecha de Excel y como texto "DD.MM.AAAA".
  - **Transferencias (formato SAP):** importe total, por sucursal origen, antigüedad y documento más antiguo.
  - **Bloqueado (formato CONCENTRADO):** importe total (suma de V. Stock bloq., cruzada con la hoja RESUMEN), importe por centro, unidades bloqueadas, por Jerarquía 2 y top de artículos.
  - **Bloqueado (formato original):** importe neto, impacto absoluto, por sucursal, por departamento, top de artículos y stock positivo vs negativo.
- **Cortes:** cada archivo importado crea un corte con su fecha (la app sugiere la fecha del nombre del archivo, si no la fecha más reciente del archivo, si no hoy; se puede cambiar antes de importar). El historial queda en la pestaña **Cargas** y un corte se puede borrar si se subió un archivo equivocado. Se guardan hasta 12 cortes por tipo.
- **Seguimiento:** en cada renglón de sucursal o documento se puede marcar el estado (Pendiente / En seguimiento / Liberado / Escalado) y escribir una nota. Se guarda en el dispositivo y se mantiene entre cortes.
- **Exportar a Excel:** botones para descargar el **Reporte completo** (Resumen + panel y detallado de los tres reportes) o cada reporte por separado (hoja "Panel" y hoja "Detallado" con todas las filas del corte más reciente).
- Si un archivo no se reconoce, la app muestra qué columnas espera cada formato.

## Publicar en GitHub Pages

1. Crea un repositorio público en GitHub, por ejemplo `Panel-de-Seguimiento`.
2. Sube todos los archivos de esta carpeta a la rama `main`, en la raíz del repositorio.
3. En el repositorio ve a **Settings → Pages**.
4. En **Source** elige **Deploy from a branch**, rama **main**, carpeta **/ (root)**. Guarda.
5. Espera 1-2 minutos. La liga queda así:
   `https://TU_USUARIO.github.io/Panel-de-Seguimiento/`

## Instalar en el iPhone

1. Abre la liga en **Safari** (tiene que ser Safari, no Chrome).
2. Toca el botón **Compartir** (el cuadro con la flecha hacia arriba).
3. Baja y toca **"Agregar a pantalla de inicio"**.
4. Toca **Agregar**. El icono "Seguimiento" queda en tu pantalla de inicio y abre como app.

> La primera vez ábrela con internet para que se descargue. Después los paneles ya cargados se ven sin conexión, pero para leer archivos nuevos sí necesitas internet (la librería de Excel se carga desde internet).

## Archivos

- `index.html` — la página principal.
- `app.js` — detección, lectura de Excel, paneles, seguimiento y exportación.
- `styles.css` — diseño (móvil primero).
- `manifest.webmanifest` — datos de la app instalada.
- `sw.js` — caché para uso sin conexión.
- `icon.svg` / `icon-512.png` — icono de la app.
