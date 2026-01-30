# Guia Extractor (Amazon Seller)

Extension de Chrome para **pegar una lista de guias (Tracking ID / Order ID)**, buscarlas en **Amazon Seller > Manage Orders**, y generar un **TSV sin encabezados** con `ASIN \t CANTIDAD \t TRACKING` listo para pegar en Excel Online.

## ¿Que problema resuelve?

Automatiza la consulta manual de ordenes en Seller Central cuando tienes muchas guias. La extension:

- Lee una lista de guias que el usuario pega en el popup.
- Busca cada guia en **Manage Orders** usando el filtro correcto:
  - Si contiene guiones → **Order ID**
  - Si es numerica → **Tracking ID**
- Extrae **ASIN, Cantidad y Tracking** de los resultados.
- Genera salida TSV para Excel y un reporte de revision manual.

---

## Requisitos

- Google Chrome (o Chromium) con extensiones habilitadas.
- Sesion iniciada en **Seller Central**.
- Abrir la pagina **Manage Orders** y mantenerla activa durante la busqueda.

---

## Instalacion (modo desarrollador)

1. Abre `chrome://extensions`.
2. Activa **Developer mode**.
3. Click en **Load unpacked**.
4. Selecciona la carpeta: `C:\Users\Public\GitHub\Guia-Asin-ATM\extension`.
5. Asegura que la extension aparezca habilitada.

---

## Flujo de uso (resumen)

1. Abre **Seller Central** y entra a **Manage Orders**.
2. Abre el popup de la extension.
3. Pega la columna de guias en **Guias de entrada**.
4. Click en **Buscar en Amazon**.
5. Copia el resultado TSV y pegalo en Excel.
6. Revisa la seccion **Revision manual** si hay guias con multiples ASIN o cantidad faltante.

---

## Formato de entrada

Puedes pegar guias en cualquiera de estos formatos:

- Una guia por linea:
  ```
  7966941710
  702-5626297-9927461
  7966803095
  ```
- Varias guias separadas por espacios, tabulaciones o comas.

La extension normaliza automaticamente:

- **Order ID** con formato `###-#######-#######`
- **Tracking ID** numericos entre 8 y 14 digitos

---

## Salida TSV (sin encabezados)

El resultado principal se entrega **sin encabezados** en el formato:

```
ASIN <TAB> CANTIDAD <TAB> TRACKING
```

Ejemplo:

```
B01ARKFXVS	1	7967181390
B0D61GYNW0	2	702-5626297-9927461
```


---

## Revision manual

Se agrega a revision manual si:

- Hay **varios ASIN** para una misma guia.
- No se puede detectar la cantidad.
- No se encontro resultado.
- Ocurre un error durante la busqueda.

Formato TSV de revision manual:

```
GUIA <TAB> ASIN <TAB> CANTIDAD <TAB> MOTIVO
```

---

## Consejos de uso

- Mantener **solo una pestaña** abierta de Seller Central para evitar que la extension seleccione el contexto equivocado.
- No cambiar de pestaña mientras la extension esta buscando.
- Si una guia no aparece, intenta ampliar el rango de fechas en **Manage Orders**.

---

## Problemas comunes

**1) “No se encontro el buscador”**  
Verifica que estas en **Manage Orders** y que la sesion esta activa.

**2) No devuelve resultados para guias existentes**  
Revisa el filtro y el rango de fechas en Seller Central.

**3) Copia siempre el mismo producto**  
Normalmente ocurre si la busqueda no termina de refrescar. Espera a que cargue antes de iniciar otra busqueda.

---

## Licencia

Uso interno.
