# Manual de Usuario — Guia Extractor (Amazon Seller)

Este manual explica **paso a paso** como usar la extension para buscar guias en Amazon Seller y generar un TSV listo para pegar en Excel Online.

---

## 1. Objetivo de la extension

La extension automatiza el proceso de:

1. Tomar una lista de **Tracking ID / Order ID** que pegaste.
2. Buscar cada guia en **Manage Orders** en Seller Central.
3. Extraer **ASIN, Cantidad y Tracking**.
4. Entregar un resultado **TSV sin encabezados** listo para pegar en Excel Online.

---

## 2. Requisitos previos

Antes de usar la extension, asegúrate de:

- Tener abierta una sesion activa en Seller Central.
- Estar en la pagina **Manage Orders**.
- Mantener la pestaña de Seller Central activa durante la busqueda.

---

## 3. Instalacion rapida

1. Abre `chrome://extensions`.
2. Activa **Developer mode**.
3. Click en **Load unpacked**.
4. Selecciona la carpeta: `C:\Users\Public\GitHub\Guia-Asin-ATM\extension`.
5. Verifica que la extension aparezca como habilitada.

---

## 4. Interfaz del popup

El popup contiene 4 secciones principales:

### 4.1 Guias de entrada

- Aqui pegas la lista de Tracking ID u Order ID.
- La extension muestra:
  - **Validas**
  - **Duplicadas**
  - **Invalidas**

### 4.2 Busqueda en Amazon

Botones principales:

- **Buscar en Amazon**: inicia el proceso.
- **Detener**: cancela la busqueda en curso.

Tambien muestra:

- Estado de busqueda
- Progreso `Procesadas X / Y`

### 4.3 Resultado (TSV)

Muestra los resultados correctos en formato TSV:

```
ASIN<TAB>CANTIDAD<TAB>TRACKING
```

Ejemplo:

```
B01ARKFXVS	1	7967181390
B0D61GYNW0	2	702-5626297-9927461
```

### 4.4 Revision manual

Se listan las guias con problemas:

- multiples ASIN
- cantidad no detectada
- no encontrado
- error de busqueda

Formato TSV:

```
GUIA<TAB>ASIN<TAB>CANTIDAD<TAB>MOTIVO
```

---

## 5. Flujo completo (paso a paso)

### Paso 1: Abre Manage Orders

1. Entra a Seller Central.
2. Ve a **Orders > Manage Orders**.
3. Asegurate de que la pagina esta cargada completamente.

### Paso 2: Prepara tus guias

Ejemplo de guias validas:

```
7966941710
702-5626297-9927461
7966803095
```

Puedes pegarlas una por linea o separadas por espacios/comas.

### Paso 3: Pega las guias en el popup

1. Abre el popup de la extension.
2. Pega la lista en **Guias de entrada**.
3. Verifica los contadores de validas/duplicadas/invalidas.

### Paso 4: Ejecuta la busqueda

1. Click en **Buscar en Amazon**.
2. Espera a que complete todas las guias.
3. No cambies de pestaña mientras corre.

### Paso 5: Copia el resultado

1. Copia el contenido de **Resultado (TSV)**.
2. Pegalo en Excel Online. Cada columna se separa automaticamente.

### Paso 6: Revisa manualmente si es necesario

Si alguna guia aparece en **Revision manual**, revisala directamente en Seller Central.

---

## 6. Como decide el filtro (Order ID vs Tracking ID)

- Si la guia contiene guiones (`-`), se asume **Order ID**.
- Si es numerica, se asume **Tracking ID**.

Esto permite usar ambos formatos en una misma lista.

---

## 7. Preguntas frecuentes

### 7.1 ¿Por que algunas guias no regresan resultado?

Posibles causas:

- La guia no existe en el rango de fechas seleccionado.
- El pedido pertenece a otra cuenta/marketplace.
- No se cargo correctamente la pagina.

### 7.2 ¿Por que aparece “Cantidad no detectada”?

Algunas veces la interfaz de Amazon no muestra la cantidad de forma clara.
Estas guias se envian a revision manual.

### 7.3 ¿La extension funciona en otros idiomas?

Si. Reconoce tanto etiquetas en **Ingles** como en **Español**.

---

## 8. Recomendaciones

- Mantener el rango de fechas de **Manage Orders** bien ajustado.
- Evitar usar otras pestañas durante la busqueda.
- Si hay muchas guias, dividir en bloques de 50–100.

---

## 9. Contacto / soporte

Si necesitas ajustar el comportamiento, solicita cambios en:

- Reglas de validacion de guias
- Formato de salida TSV
- Identificacion de ASIN o cantidad

