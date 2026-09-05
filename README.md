# ðŸ¥ Asistente de Recepciones PDF para Excel (Office Web Add-in)

Complemento oficial de Excel para procesar automÃ¡ticamente facturas, remisiones, certificados de calidad y registros sanitarios en PDF e insertarlos en la hoja **`recepciones_2026`** de tu libro de Excel.

---

## ðŸš€ CaracterÃ­sticas
* **Drag & Drop**: Arrastra cualquier PDF o remisiÃ³n directamente al panel lateral dentro de Excel.
* **ExtracciÃ³n de 31 Campos**:
  * Clave, DescripciÃ³n, Lote, Fecha de FabricaciÃ³n, Fecha de Caducidad.
  * Registro Sanitario COFEPRIS, Marca, Fabricante, PaÃ­s de Origen.
  * Folio SAT / Factura, RemisiÃ³n, Precios Unitarios, IVA, Totales en Pesos.
  * Folio de Entrada / SIIA, Carta de Canje, Observaciones.
* **InserciÃ³n con un Clic**: Vuelca las filas extraÃ­das de manera segura mediante la API nativa de Office.js (`Excel.run`).
* **BotÃ³n de Copia RÃ¡pida**: Copia la tabla tabulada lista para pegar con `Ctrl + V`.

---

## ðŸ“¦ Estructura del Proyecto
```
excel-recepciones-addin/
â”œâ”€â”€ manifest.xml                # Manifiesto de configuraciÃ³n de Office Add-in
â”œâ”€â”€ package.json                # ConfiguraciÃ³n de dependencias y scripts
â”œâ”€â”€ README.md                   # DocumentaciÃ³n de instalaciÃ³n y uso
â”œâ”€â”€ assets/                     # Iconos para la cinta de opciones de Excel
â”‚   â”œâ”€â”€ icon-16.png
â”‚   â”œâ”€â”€ icon-32.png
â”‚   â””â”€â”€ icon-80.png
â””â”€â”€ src/
    â””â”€â”€ taskpane/
        â”œâ”€â”€ taskpane.html       # Interfaz visual del panel lateral
        â”œâ”€â”€ taskpane.css        # Estilos modernos Fluent UI
        â””â”€â”€ taskpane.js         # LÃ³gica de extracciÃ³n de PDF e inserciÃ³n en Excel
```

---

## ðŸ› ï¸ CÃ³mo Instalar y Probar en Excel

### 1. Iniciar el servidor local
```bash
npx serve -l 3000
```

### 2. Cargar el Manifiesto en Excel (Sideloading)
1. Abre tu archivo `recepciones_2026.xlsx` en Microsoft Excel.
2. Ve a la pestaÃ±a **Insertar** > **Mis complementos** (o *Obtener complementos*).
3. Haz clic en **Cargar mi complemento** (o *Administrar complementos compartidos*).
4. Selecciona el archivo `manifest.xml` de esta carpeta.
5. VerÃ¡s aparecer una nueva pestaÃ±a **ABISALUD** en la cinta superior con el botÃ³n **Procesar PDF**.