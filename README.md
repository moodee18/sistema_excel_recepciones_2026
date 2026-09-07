# 🏥 Asistente de Recepciones PDF y Escaneos para Excel (Office Web Add-in)

Complemento oficial de Excel para procesar automáticamente facturas, remisiones, certificados de calidad y registros sanitarios en PDF (**digitales o escaneados**) e insertarlos con precisión en la hoja **`recepciones_2026`** de tu libro de Excel.

---

## 🚀 Nuevas Capacidades para PDFs Escaneados

* **Detección Automática de Escaneos**: Detecta si el archivo es un PDF escaneado (imágenes sin capa de texto) y activa automáticamente el motor visual.
* **Motor IA Gemini Vision (Recomendado)**:
  * Renderizado HD de páginas en Canvas.
  * Extracción inteligente mediante modelos multimodales (`gemini-2.5-flash` o `gemini-1.5-flash`).
  * Reconocimiento de sellos, tablas complejas, códigos de barras, números manuscritos y notas de remisión.
* **Motor OCR Local Tesseract.js v5**:
  * 100% offline y en el navegador, con diccionario en español (`spa+eng`).
  * Preprocesamiento de imagen con filtros de contraste y binarización para escaneos borrosos o con sombras.
* **Editor Interactivo de 31 Columnas**:
  * Edición en vivo de cabecera (Folio, Proveedor, RFC, Contrato, Orden de Suministro, Fechas).
  * Edición de partidas (Clave ###.###.####, Descripción, Cantidad, Lote, Caducidad, Fabricación, Registro Sanitario SSA, Precios y Totales).
  * Recálculo automático de montos e IVA.
* **Galería de Miniaturas**:
  * Permite visualizar las páginas escaneadas analizadas para comparar con los datos extraídos.
* **Inserción con 1 Clic**:
  * Inserción directa en la hoja `recepciones_2026` mediante Office.js.
  * Botón de copiado tabulado (TSV) para pegar con `Ctrl + V`.

---

## 📦 Estructura del Proyecto
```
excel-recepciones-addin/
├── manifest.xml                # Manifiesto de configuración de Office Add-in
├── package.json                # Configuración de dependencias y scripts
├── README.md                   # Documentación de instalación y uso
├── assets/                     # Iconos para la cinta de opciones de Excel
│   ├── icon-16.png
│   ├── icon-32.png
│   └── icon-80.png
└── src/
    └── taskpane/
        ├── taskpane.html       # Interfaz visual del panel lateral con soporte OCR / IA
        ├── taskpane.css        # Estilos modernos Fluent UI y diseño responsivo
        └── taskpane.js         # Motores Gemini Vision, Tesseract OCR y Office.js
```

---

## 🛠️ Cómo Iniciar y Usar

### 1. Iniciar el servidor local (o abrir en navegador)
```bash
npx serve -l 3000
```
O simplemente abrir `src/taskpane/taskpane.html` en cualquier navegador web moderno (Edge, Chrome, Firefox).

### 2. Configurar la API Key de Gemini (Opcional para Modo IA)
1. Haz clic en el icono ⚙️ en la esquina superior derecha del panel.
2. Pega tu API Key de Google Gemini.
3. Haz clic en **Probar Conexión** y luego en **Guardar Configuración**.

### 3. Cargar el Manifiesto en Excel (Sideloading)
1. Abre tu archivo `recepciones_2026.xlsx` en Microsoft Excel.
2. Ve a la pestaña **Insertar** > **Mis complementos** (o *Obtener complementos*).
3. Haz clic en **Cargar mi complemento** (o *Administrar complementos compartidos*).
4. Selecciona el archivo `manifest.xml` de esta carpeta.
5. Verás aparecer una nueva pestaña **ABISALUD** en la cinta superior con el botón **Procesar PDF**.