# Guia-ASIN-ATM

Extension de Chrome para consultar guias de envio en Amazon Seller y registrar ASIN + cantidad en Excel Online.

## Arquitectura base

- MV3 con `background.js` (service worker) para defaults y logs.
- `popup/` con UI principal (tabs, modo asistido y directo).
- `content/` con content scripts base para Amazon Seller y Excel Online (placeholders).
- Persistencia local con `chrome.storage.local`.

## Estructura

- `extension/manifest.json`
- `extension/background.js`
- `extension/popup/popup.html`
- `extension/popup/popup.css`
- `extension/popup/popup.js`
- `extension/content/amazon_seller.js`
- `extension/content/excel_online.js`
- `extension/_locales/es/messages.json`
- `extension/_locales/en_US/messages.json`

## Estado actual

- UI lista en popup con tabs: Configuracion, Validacion, Ejecucion y Logs.
- Soporte de idioma (Espanol e Ingles US) y seleccion de region Amazon.
- Captura de seleccion para Documento 1 y Documento 2 con aprobacion previa.
- Validacion opcional contra datos existentes del Documento 2.
- Pipeline basico: captura guias, consulta Amazon, prepara salida y pega en Excel.
