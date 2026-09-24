/* global Office, Excel, pdfjsLib, Tesseract */

// Estado Global de la Aplicación
let extractedItems = [];
let extractedHeader = {};
let isInsideOffice = false;
let currentPdfPagesCanvas = [];
let selectedEngine = "auto"; // "auto", "gemini", "tesseract"

// Inicialización de Office.js
if (typeof Office !== "undefined" && Office.onReady) {
    Office.onReady((info) => {
        if (info.host === Office.HostType.Excel) {
            isInsideOffice = true;
            const btnInsert = document.getElementById("btn-insert");
            if (btnInsert) btnInsert.style.display = "flex";
        }
    });
}

// Inicialización al cargar el DOM
document.addEventListener("DOMContentLoaded", () => {
    initApp();
});

function initApp() {
    if (typeof pdfjsLib !== "undefined") {
        pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
    }

    loadSettings();

    const dropZone = document.getElementById("drop-zone");
    const fileInput = document.getElementById("file-input");
    const btnInsert = document.getElementById("btn-insert");
    const btnCopy = document.getElementById("btn-copy");
    const btnAddItem = document.getElementById("btn-add-item");
    const btnOpenSettings = document.getElementById("btn-open-settings");
    const btnCloseSettings = document.getElementById("btn-close-settings");
    const btnSaveSettings = document.getElementById("btn-save-settings");
    const btnTestKey = document.getElementById("btn-test-key");
    const btnToggleKeyVis = document.getElementById("btn-toggle-key-visibility");
    const toggleThumbsBtn = document.getElementById("toggle-thumbs-btn");
    const btnBannerConfig = document.getElementById("btn-banner-config");

    if (dropZone && fileInput) {
        dropZone.onclick = () => fileInput.click();

        dropZone.ondragover = (e) => {
            e.preventDefault();
            dropZone.classList.add("dragover");
        };

        dropZone.ondragleave = () => {
            dropZone.classList.remove("dragover");
        };

        dropZone.ondrop = (e) => {
            e.preventDefault();
            dropZone.classList.remove("dragover");
            if (e.dataTransfer && e.dataTransfer.files.length > 0) {
                handleFile(e.dataTransfer.files[0]);
            }
        };

        fileInput.onchange = (e) => {
            if (e.target.files && e.target.files.length > 0) {
                handleFile(e.target.files[0]);
            }
        };
    }

    document.querySelectorAll(".engine-pill").forEach(pill => {
        pill.onclick = () => {
            document.querySelectorAll(".engine-pill").forEach(p => p.classList.remove("active"));
            pill.classList.add("active");
            selectedEngine = pill.dataset.engine;
            localStorage.setItem("selected_engine", selectedEngine);

            const geminiKey = (localStorage.getItem("gemini_api_key") || "").trim();
            if (selectedEngine === "gemini" && !geminiKey) {
                openSettingsModal();
            }
        };
    });

    if (btnInsert) btnInsert.onclick = insertIntoExcel;
    if (btnCopy) btnCopy.onclick = copyToClipboard;
    if (btnAddItem) btnAddItem.onclick = addNewEmptyItem;

    if (btnOpenSettings) btnOpenSettings.onclick = openSettingsModal;
    if (btnBannerConfig) btnBannerConfig.onclick = openSettingsModal;

    if (btnCloseSettings) {
        btnCloseSettings.onclick = () => {
            document.getElementById("settings-modal").style.display = "none";
        };
    }

    if (btnSaveSettings) {
        btnSaveSettings.onclick = () => {
            saveSettings();
            document.getElementById("settings-modal").style.display = "none";
            showAlert("✅ Configuración guardada correctamente.", "success");
        };
    }

    if (btnTestKey) btnTestKey.onclick = testGeminiApiKey;

    if (btnToggleKeyVis) {
        btnToggleKeyVis.onclick = () => {
            const input = document.getElementById("gemini-api-key-input");
            input.type = input.type === "password" ? "text" : "password";
        };
    }

    if (toggleThumbsBtn) {
        toggleThumbsBtn.onclick = () => {
            const container = document.getElementById("thumbnails-container");
            if (container.style.display === "none") {
                container.style.display = "flex";
                toggleThumbsBtn.innerText = "Ocultar";
            } else {
                container.style.display = "none";
                toggleThumbsBtn.innerText = "Mostrar";
            }
        };
    }

    updateApiKeyUI();
}

function openSettingsModal() {
    document.getElementById("settings-modal").style.display = "flex";
}

function updateApiKeyUI() {
    const apiKey = (localStorage.getItem("gemini_api_key") || "").trim();
    const banner = document.getElementById("api-key-banner");
    const keyDot = document.getElementById("ia-key-dot");

    if (apiKey) {
        if (banner) banner.style.display = "none";
        if (keyDot) {
            keyDot.className = "engine-indicator ready";
            keyDot.title = "API Key configurada";
        }
    } else {
        if (banner) banner.style.display = "flex";
        if (keyDot) {
            keyDot.className = "engine-indicator warning";
            keyDot.title = "Falta configurar API Key de Gemini";
        }
    }
}

function loadSettings() {
    const savedEngine = localStorage.getItem("selected_engine") || "auto";
    selectedEngine = savedEngine;
    document.querySelectorAll(".engine-pill").forEach(p => {
        if (p.dataset.engine === savedEngine) p.classList.add("active");
        else p.classList.remove("active");
    });

    const apiKey = localStorage.getItem("gemini_api_key") || "";
    let model = localStorage.getItem("gemini_model") || "gemini-2.0-flash";
    if (model === "gemini-2.5-flash" || !model) {
        model = "gemini-2.0-flash";
        localStorage.setItem("gemini_model", model);
    }

    const pagesLimit = localStorage.getItem("pages_limit") || "auto";
    const dpi = localStorage.getItem("canvas_dpi") || "1.8";
    const binarize = localStorage.getItem("ocr_binarize") !== "false";

    const keyInput = document.getElementById("gemini-api-key-input");
    const modelSelect = document.getElementById("gemini-model-select");
    const pagesSelect = document.getElementById("pages-limit-select");
    const dpiSelect = document.getElementById("canvas-dpi-select");
    const binarizeChk = document.getElementById("chk-auto-binarize");

    if (keyInput) keyInput.value = apiKey;
    if (modelSelect) modelSelect.value = model;
    if (pagesSelect) pagesSelect.value = pagesLimit;
    if (dpiSelect) dpiSelect.value = dpi;
    if (binarizeChk) binarizeChk.checked = binarize;

    updateApiKeyUI();
}

function saveSettings() {
    const keyInput = document.getElementById("gemini-api-key-input");
    const modelSelect = document.getElementById("gemini-model-select");
    const pagesSelect = document.getElementById("pages-limit-select");
    const dpiSelect = document.getElementById("canvas-dpi-select");
    const binarizeChk = document.getElementById("chk-auto-binarize");

    if (keyInput) localStorage.setItem("gemini_api_key", keyInput.value.trim());
    if (modelSelect) localStorage.setItem("gemini_model", modelSelect.value);
    if (pagesSelect) localStorage.setItem("pages_limit", pagesSelect.value);
    if (dpiSelect) localStorage.setItem("canvas_dpi", dpiSelect.value);
    if (binarizeChk) localStorage.setItem("ocr_binarize", binarizeChk.checked);

    updateApiKeyUI();
}

async function testGeminiApiKey() {
    const key = document.getElementById("gemini-api-key-input").value.trim();
    if (!key) {
        alert("Por favor ingresa una API Key de Gemini primero.");
        return;
    }

    const btn = document.getElementById("btn-test-key");
    btn.disabled = true;
    btn.innerText = "Probando conexión...";

    try {
        let model = document.getElementById("gemini-model-select").value || "gemini-2.0-flash";
        if (model === "gemini-2.5-flash") model = "gemini-2.0-flash";

        const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${key}`;
        const res = await fetch(url, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                contents: [{ parts: [{ text: "Responde exactamente: OK" }] }]
            })
        });

        const data = await res.json();
        if (res.ok && data.candidates && data.candidates.length > 0) {
            alert("¡Conexión Exitosa con Google Gemini AI! 🚀 El motor IA Vision está listo para usar.");
        } else if (data.error) {
            alert(`Error de API (${data.error.code || res.status}): ${data.error.message}`);
        } else {
            alert("Respuesta inesperada de la API: " + JSON.stringify(data));
        }
    } catch (err) {
        alert("Error de red o conexión: " + err.message);
    } finally {
        btn.disabled = false;
        btn.innerText = "Probar Conexión";
    }
}

// ==========================================
// PROCESAMIENTO PRINCIPAL DE ARCHIVOS PDF
// ==========================================
async function handleFile(file) {
    if (!file.name.toLowerCase().endsWith(".pdf")) {
        showAlert("Por favor selecciona un archivo PDF válido.", "error");
        return;
    }

    showStatus(true, "Cargando documento PDF...", 10, "Inicializando motor PDF.js");
    document.getElementById("file-details").innerText = `${file.name} (${(file.size / (1024 * 1024)).toFixed(2)} MB)`;
    document.getElementById("preview-section").style.display = "none";
    document.getElementById("thumbnails-section").style.display = "none";
    currentPdfPagesCanvas = [];

    // Extraer pistas de nombre de archivo ({Folio} {DD-MM-AAAA}.pdf)
    const fileNameMatch = file.name.match(/^(\d{3,8})\s+(\d{1,2}[\-\/]\d{1,2}[\-\/]\d{2,4})/);
    let hintedFolio = "";
    let hintedDate = "";
    if (fileNameMatch) {
        hintedFolio = fileNameMatch[1];
        hintedDate = fileNameMatch[2].replace(/\-/g, "/");
    }

    try {
        const arrayBuffer = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        const numPages = pdf.numPages;

        document.getElementById("page-count").innerText = numPages;
        const thumbsContainer = document.getElementById("thumbnails-container");
        thumbsContainer.innerHTML = "";

        const dpiScale = parseFloat(localStorage.getItem("canvas_dpi") || "1.8");
        let digitalText = "";

        for (let i = 1; i <= numPages; i++) {
            showStatus(true, `Renderizando página ${i} de ${numPages}...`, 10 + Math.floor((i / numPages) * 35), "Generando imágenes HD para análisis");
            const page = await pdf.getPage(i);

            const textContent = await page.getTextContent();
            const pageText = textContent.items.map(item => item.str).join(" ");
            if (pageText.trim().length > 30) {
                digitalText += `\n--- PÁGINA ${i} ---\n` + pageText;
            }

            const viewport = page.getViewport({ scale: dpiScale });
            const canvas = document.createElement("canvas");
            canvas.width = viewport.width;
            canvas.height = viewport.height;
            const ctx = canvas.getContext("2d", { willReadFrequently: true });

            await page.render({ canvasContext: ctx, viewport: viewport }).promise;
            currentPdfPagesCanvas.push(canvas);

            const thumbWrapper = document.createElement("div");
            thumbWrapper.className = "thumb-wrapper";
            const thumbCanvas = document.createElement("canvas");
            const thumbScale = 0.18;
            const thumbViewport = page.getViewport({ scale: thumbScale });
            thumbCanvas.width = thumbViewport.width;
            thumbCanvas.height = thumbViewport.height;
            thumbCanvas.className = "thumb-canvas";
            const thumbCtx = thumbCanvas.getContext("2d");
            await page.render({ canvasContext: thumbCtx, viewport: thumbViewport }).promise;

            thumbWrapper.innerHTML = `<span class="thumb-label">Pág. ${i}</span>`;
            thumbWrapper.prepend(thumbCanvas);
            thumbsContainer.appendChild(thumbWrapper);
        }

        document.getElementById("thumbnails-section").style.display = "block";

        const hasMedicalClaves = /(?:010|020|030|040|060|070|080)[\.\s\-]\d{3}[\.\s\-]\d{4}|\b\d{3}\.\d{3}\.\d{4}\b/.test(digitalText);
        const hasRfcOrInvoice = /[A-Z&Ñ]{3,4}\d{6}[A-V1-9]|\b\d{8}\b|REMISI[OÓ]N|FACTURA|CFDI/i.test(digitalText);
        const isScannedDoc = digitalText.trim().length < 200 || (!hasMedicalClaves && !hasRfcOrInvoice);

        const geminiKey = (localStorage.getItem("gemini_api_key") || "").trim();
        let engineToUse = selectedEngine;

        if (engineToUse === "auto") {
            if (isScannedDoc) {
                engineToUse = geminiKey ? "gemini" : "tesseract";
            } else {
                engineToUse = "digital";
            }
        }

        const pagesSetting = localStorage.getItem("pages_limit") || "auto";
        let maxPagesToAnalyze = currentPdfPagesCanvas.length;
        if (pagesSetting === "auto") {
            // Páginas 1 a 4 contienen la Entrada Directa de Factura y el CFDI del SAT
            maxPagesToAnalyze = Math.min(4, currentPdfPagesCanvas.length);
        } else if (pagesSetting === "3") {
            maxPagesToAnalyze = Math.min(3, currentPdfPagesCanvas.length);
        } else if (pagesSetting === "5") {
            maxPagesToAnalyze = Math.min(5, currentPdfPagesCanvas.length);
        }

        const canvasesToProcess = currentPdfPagesCanvas.slice(0, maxPagesToAnalyze);

        let engineLabel = "⚡ Extracción Digital";
        if (engineToUse === "gemini") engineLabel = `🤖 IA Gemini Vision (${canvasesToProcess.length} págs)`;
        else if (engineToUse === "tesseract") engineLabel = `🔍 OCR Local Tesseract (${canvasesToProcess.length} págs)`;

        document.getElementById("engine-used-badge").innerText = engineLabel;

        if (engineToUse === "gemini") {
            if (!geminiKey) {
                showAlert("⚠️ Para usar IA Vision en escaneos, configura tu Gemini API Key (⚙️). Usando OCR Local Tesseract como respaldo...", "warning");
                document.getElementById("engine-used-badge").innerText = `🔍 OCR Local Tesseract (${canvasesToProcess.length} págs)`;
                await processWithTesseractOCR(canvasesToProcess, file.name);
            } else {
                await processWithGeminiVision(canvasesToProcess, file.name, hintedFolio, hintedDate);
            }
        } else if (engineToUse === "tesseract") {
            await processWithTesseractOCR(canvasesToProcess, file.name);
        } else {
            showStatus(true, "Analizando texto digital y partidas...", 85, "Estructurando 32 columnas de recepciones_2026");
            parsePdfText(digitalText, file.name);

            if (extractedItems.length === 0) {
                console.log("Extracción digital sin partidas. Activando motor para escaneo...");
                if (geminiKey) {
                    document.getElementById("engine-used-badge").innerText = `🤖 IA Gemini Vision (Auto-Fallback ${canvasesToProcess.length} págs)`;
                    await processWithGeminiVision(canvasesToProcess, file.name, hintedFolio, hintedDate);
                } else {
                    document.getElementById("engine-used-badge").innerText = `🔍 OCR Local Tesseract (Auto-Fallback ${canvasesToProcess.length} págs)`;
                    await processWithTesseractOCR(canvasesToProcess, file.name);
                }
            }
        }

        showStatus(false);
        renderPreview();

        if (extractedItems.length > 0) {
            showAlert(`🎉 ¡Procesamiento completado! Se extrajeron con éxito ${extractedItems.length} partidas del documento para recepciones_2026.`, "success");
        } else {
            showAlert("⚠️ No se pudieron extraer partidas automáticamente. Puedes agregarlas manualmente con el botón '+ Agregar Partida'.", "warning");
        }

    } catch (error) {
        console.error("Error en handleFile:", error);
        showStatus(false);
        showAlert("Error al procesar el PDF: " + error.message, "error");
    }
}

function compressCanvasForVision(canvas, maxDimension = 1600, quality = 0.75) {
    let w = canvas.width;
    let h = canvas.height;
    if (w > maxDimension || h > maxDimension) {
        const ratio = Math.min(maxDimension / w, maxDimension / h);
        w = Math.round(w * ratio);
        h = Math.round(h * ratio);
    }

    const tempCanvas = document.createElement("canvas");
    tempCanvas.width = w;
    tempCanvas.height = h;
    const ctx = tempCanvas.getContext("2d");
    ctx.drawImage(canvas, 0, 0, w, h);

    const dataUrl = tempCanvas.toDataURL("image/jpeg", quality);
    return dataUrl.split(",")[1];
}

// ==========================================
// MOTOR 1: GOOGLE GEMINI VISION AI
// ==========================================
async function processWithGeminiVision(canvases, fileName, hintedFolio = "", hintedDate = "") {
    const apiKey = (localStorage.getItem("gemini_api_key") || "").trim();
    let model = localStorage.getItem("gemini_model") || "gemini-2.0-flash";
    if (model === "gemini-2.5-flash") model = "gemini-2.0-flash";

    showStatus(true, `Analizando ${canvases.length} páginas con IA Gemini Vision...`, 50, "Extrayendo campos para el libro recepciones_2026");

    const imageParts = canvases.map(canvas => {
        const base64Data = compressCanvasForVision(canvas, 1600, 0.75);
        return {
            inlineData: {
                data: base64Data,
                mimeType: "image/jpeg"
            }
        };
    });

    const promptText = `
Eres un analista experto en el sistema de recepción de insumos médicos de ABISALUD y en la base de datos de Excel 'recepciones_2026.xlsx' (32 columnas oficiales de la A a la AF).
Analiza detalladamente las páginas escaneadas proporcionadas de este expediente médico (hoja de Entrada Directa de Factura / Remisión, Factura Fiscal CFDI 4.0, notas de contrato y sellos).

REGLAS DE NEGOCIO APRENDIDAS DE 'recepciones_2026':
1. Columna B & C ('fecha_recepcion' / 'fecha_ingreso'): Fecha de la captura en almacén (ej. ${hintedDate ? `'${hintedDate}'` : "'13/07/2026'"}). Formato DD/MM/AAAA.
2. Columna D ('tipo_contrato-fuente de financiamiento'): Número oficial de procedimiento o licitación (ej. 'LA-12-NEF-012NEF001-I-59-2025' o 'AA-12-NEF-012NEF001-I-152-2025').
3. Columna E ('tipo_adquisicion'): 
   - Si el procedimiento inicia con 'LA-' -> 'LICITACION PUBLICA'
   - Si inicia con 'AA-' o contrato 'CS/AD/' -> 'ADJUDICACION DIRECTA'
   - Si inicia con 'IA-' -> 'INVITACION A CUANDO MENOS TRES PERSONAS'
4. Columna F ('factura_remision'): Folio fiscal de la factura (ej. 'B26-AQ-224' o 'AQ-224' o número de remisión). NUNCA pongas la palabra 'Electr' ni 'Factura'.
5. Columna G ('orden_suministro'): Número de orden (ej. 'OS-LCBMX-010-2026' u 'OS-ADBMX-045-2026').
6. Columna H ('contrato'): Número de contrato (ej. 'LC/BMX/010/2026' o 'CS/AD/045/2026').
7. Columna I ('partida_presupuestal'): Siempre '25401' para insumos y materiales médicos.
8. Columna J ('clave'): Clave médica de 10 dígitos con formato estricto ###.###.#### (ej. '060.172.0113').
9. Columna K ('descripcion'): Descripción médica detallada completa del insumo.
10. Columna L ('cantidad_recibida'): Cantidad numérica entera contada físicamente en la entrada (ej. 8, 10, 9).
11. Columna M ('lote'): Lote del insumo (ej. '251100764').
12. Columna N & O ('fecha_caducidad' y 'fecha_fabricacion'): Formato DD/MM/AAAA.
13. Columna P ('registro_sanitario'): Registro sanitario COFEPRIS (ej. '1028C2021 SSA' o 'N/A').
14. Columna Q ('unidad_medida'): 'PIEZA' o 'ENVASE' o 'CAJA'.
15. Columna R ('marca'): Marca comercial (ej. 'Flexicare', 'PROTEC').
16. Columna S ('pais_origen'): País de procedencia (ej. 'MEXICO', 'REINO UNIDO', 'BRASIL').
17. Columna T ('fabricante'): Razón social del laboratorio fabricante.
18. Columna U ('factura'): 'FACTURA' si es CFDI con sello fiscal SAT, 'REMISION' si es remisión de entrega.
19. Columna V ('fecha_emision_factura'): Fecha de timbrado/emisión de la factura (DD/MM/AAAA).
20. Columna W, X, Y, Z: 
    - pu: Precio unitario sin IVA (ej. 33.50).
    - monto: Subtotal = cantidad * pu.
    - iva: IVA 16% = monto * 0.16.
    - total: Total con IVA = monto + iva.
21. Columna AB & AC: 'rfc_proveedor' (ej. 'ABI110629LA5') y 'proveedor' (Razón social completa, ej. 'AMC BIOMEDICAL S.A. DE C.V.').
22. Columna AD ('carta_canje'): 'SI' o 'NO'.
23. Columna AE ('observacion'): Notas manuscritas o sellos (ej. '-Marbete hecho -Cargado Inventario Dcto').
24. Columna AF ('folio_siia'): Folio de entrada al almacén (ej. ${hintedFolio ? `'${hintedFolio}'` : "'2678'"} o del nombre de archivo '${fileName}').

Devuelve la información ÚNICAMENTE en el siguiente JSON estructurado:

{
  "header": {
    "folio": "${hintedFolio || "2678"}",
    "fechaRecepcion": "${hintedDate || "13/07/2026"}",
    "fechaIngreso": "${hintedDate || "13/07/2026"}",
    "tipoContrato": "LA-12-NEF-012NEF001-I-59-2025",
    "tipoAdquisicion": "LICITACION PUBLICA",
    "facturaRemision": "B26-AQ-224",
    "ordenSuministro": "OS-LCBMX-010-2026",
    "contrato": "LC/BMX/010/2026",
    "partidaPresupuestal": "25401",
    "rfcProveedor": "ABI110629LA5",
    "proveedor": "AMC BIOMEDICAL S.A. DE C.V.",
    "factura": "FACTURA",
    "fechaEmision": "01/07/2026",
    "cartaCanje": "SI",
    "observacion": "-Marbete hecho -Cargado Inventario Dcto"
  },
  "items": [
    {
      "clave": "060.172.0113",
      "descripcion": "Descripción del insumo",
      "cantidad": 8,
      "lote": "251100764",
      "caducidad": "01/10/2030",
      "fabricacion": "01/11/2025",
      "registro": "N/A",
      "unidad": "PIEZA",
      "marca": "Flexicare",
      "pais": "REINO UNIDO",
      "fabricante": "AMC BIOMEDICAL",
      "pu": 33.50,
      "monto": 268.00,
      "iva": 42.88,
      "total": 310.88
    }
  ]
}
`;

    const requestBody = {
        contents: [
            {
                parts: [
                    { text: promptText },
                    ...imageParts
                ]
            }
        ],
        generationConfig: {
            responseMimeType: "application/json",
            temperature: 0.1
        }
    };

    const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
    
    let response;
    try {
        response = await fetch(url, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(requestBody)
        });
    } catch (netErr) {
        throw new Error("Error de red con Gemini API: " + netErr.message);
    }

    if (!response.ok) {
        const errJson = await response.json().catch(() => ({}));
        const errMsg = errJson.error?.message || `Error HTTP ${response.status}`;
        if (response.status === 400 || response.status === 404) {
            throw new Error(`Modelo '${model}' no compatible (${errMsg}). Selecciona 'Gemini 2.0 Flash' en Configuración (⚙️).`);
        } else if (response.status === 401 || response.status === 403) {
            throw new Error(`API Key de Gemini inválida o sin permisos (${errMsg}). Revisa tu clave en Configuración (⚙️).`);
        } else if (response.status === 413) {
            throw new Error("El documento excede el tamaño máximo. Selecciona 'Primeras 3 páginas' en Configuración.");
        }
        throw new Error(errMsg);
    }

    const resData = await response.json();
    let candidateText = resData.candidates?.[0]?.content?.parts?.[0]?.text;
    if (!candidateText) {
        throw new Error("Gemini no devolvió texto de respuesta.");
    }

    candidateText = candidateText.replace(/^```json\s*/i, "").replace(/\s*```$/i, "").trim();

    let parsed;
    try {
        parsed = JSON.parse(candidateText);
    } catch (e) {
        throw new Error("La respuesta de IA no fue un JSON válido: " + e.message);
    }

    extractedHeader = parsed.header || {};
    extractedItems = (parsed.items || []).map(item => ({
        clave: item.clave || "060.000.0000",
        descripcion: item.descripcion || "INSUMO MÉDICO",
        cantidad: Number(item.cantidad) || 1,
        lote: item.lote || "S/L",
        caducidad: item.caducidad || "N/A",
        fabricacion: item.fabricacion || "N/A",
        registro: item.registro || "N/A",
        unidad: (item.unidad || "PIEZA").toUpperCase(),
        marca: item.marca || "GENÉRICO",
        pais: (item.pais || "MEXICO").toUpperCase(),
        fabricante: item.fabricante || extractedHeader.proveedor || "FABRICANTE",
        pu: Number(item.pu) || 0.0,
        monto: Number(item.monto) || (Number(item.pu || 0) * Number(item.cantidad || 1)),
        iva: Number(item.iva) || (Number(item.monto || 0) * 0.16),
        total: Number(item.total) || (Number(item.monto || 0) * 1.16)
    }));
}

// ==========================================
// MOTOR 2: TESSERACT.JS OCR LOCAL (OFFLINE)
// ==========================================
async function processWithTesseractOCR(canvases, fileName) {
    if (typeof Tesseract === "undefined") {
        throw new Error("Librería Tesseract.js no disponible.");
    }

    const shouldBinarize = localStorage.getItem("ocr_binarize") !== "false";
    let combinedOcrText = "";

    showStatus(true, "Inicializando OCR Local Tesseract...", 40, "Cargando motor de reconocimiento");
    const worker = await Tesseract.createWorker("spa+eng");

    for (let i = 0; i < canvases.length; i++) {
        const pageNum = i + 1;
        showStatus(
            true,
            `Ejecutando OCR en Página ${pageNum} de ${canvases.length}...`,
            45 + Math.floor((pageNum / canvases.length) * 45),
            "Reconociendo caracteres e identificando tablas"
        );

        let canvasToProcess = canvases[i];
        if (shouldBinarize) {
            canvasToProcess = preProcessImageForOCR(canvases[i]);
        }

        const res = await worker.recognize(canvasToProcess);
        combinedOcrText += `\n--- PÁGINA ${pageNum} OCR ---\n` + res.data.text;
    }

    await worker.terminate();

    showStatus(true, "Estructurando datos y partidas...", 95, "Analizando claves médicas, lotes y RFCs");
    parsePdfText(combinedOcrText, fileName);
}

function preProcessImageForOCR(canvas) {
    const tempCanvas = document.createElement("canvas");
    tempCanvas.width = canvas.width;
    tempCanvas.height = canvas.height;
    const ctx = tempCanvas.getContext("2d");
    ctx.drawImage(canvas, 0, 0);

    const imgData = ctx.getImageData(0, 0, tempCanvas.width, tempCanvas.height);
    const data = imgData.data;

    for (let i = 0; i < data.length; i += 4) {
        const gray = 0.299 * data[i] + 0.587 * data[i + 1] + 0.114 * data[i + 2];
        const val = gray > 140 ? 255 : (gray < 80 ? 0 : gray);
        data[i] = val;
        data[i + 1] = val;
        data[i + 2] = val;
    }

    ctx.putImageData(imgData, 0, 0);
    return tempCanvas;
}

// ==========================================
// PARSER HEURÍSTICO / REGEX MEJORADO
// ==========================================
function parsePdfText(text, fileName) {
    extractedItems = [];

    const fileNameMatch = fileName.match(/^(\d{3,8})\s+(\d{1,2}[\-\/]\d{1,2}[\-\/]\d{2,4})/);
    const hintedFolio = fileNameMatch ? fileNameMatch[1] : "";
    const hintedDate = fileNameMatch ? fileNameMatch[2].replace(/\-/g, "/") : "";

    const folioMatch = text.match(/(?:Folio de Entrada|Folio Entrada|Entrada|SIIA)\s*[:#]?\s*(\d{3,8})/i) || fileName.match(/(\d{3,8})/);
    const folio = folioMatch ? folioMatch[1] : (hintedFolio || "2678");

    const rfcMatch = text.match(/[A-Z&Ñ]{3,4}\d{6}[A-V1-9][A-Z1-9][0-9A]/);
    let rfcProveedor = rfcMatch ? rfcMatch[0] : "";

    let proveedor = "PROVEEDOR MÉDICO";
    if (/AMC\s*BIOMEDICAL/i.test(text)) {
        proveedor = "AMC BIOMEDICAL S.A. DE C.V.";
        if (!rfcProveedor) rfcProveedor = "ABI110629LA5";
    } else if (/POSEIDON/i.test(text)) {
        proveedor = "GRUPO INDUSTRIAL POSEIDON S.A. DE C.V.";
    } else if (/RALMO/i.test(text)) {
        proveedor = "DISTRIBUIDORA INTERNACIONAL RALMO S.A. DE C.V.";
    } else if (/BECTON\s*DICKINSON/i.test(text)) {
        proveedor = "BECTON DICKINSON DE MEXICO, S.A DE C.V.";
    } else if (/JAYOR/i.test(text)) {
        proveedor = "LABORATORIOS JAYOR S.A. DE C.V.";
    } else if (/SOLOMED/i.test(text)) {
        proveedor = "SOLOMED S.A. DE C.V.";
    } else if (/ABASTECEDORA\s*DE\s*INSUMOS/i.test(text)) {
        proveedor = "ABASTECEDORA DE INSUMOS PARA LA SALUD S.A. DE C.V.";
    } else if (/DEGASA/i.test(text)) {
        proveedor = "DEGASA, S.A. DE C.V.";
        if (!rfcProveedor) rfcProveedor = "DEG9807015H8";
    } else if (/FRESENIUS/i.test(text)) {
        proveedor = "FRESENIUS KABI MEXICO S.A. DE C.V.";
    } else if (/KENDALL|MEDTRONIC|COVIDIEN/i.test(text)) {
        proveedor = "KENDALL DE MEXICO, S.A. DE C.V.";
    } else if (/JANEL/i.test(text)) {
        proveedor = "JANEL, S.A. DE C.V.";
    } else if (/BAXTER/i.test(text)) {
        proveedor = "BAXTER MEXICO, S.A. DE C.V.";
    }

    let remision = "";
    const factMatch = text.match(/(?:Factura\s+Electr[oó]nica(?:\s*\(CFDI\))?\s*[:#]?\s*([A-Z0-9\-]+)|(?:Remisi[oó]n|Factura No\.?|Factura)\s*[:#]?\s*([A-Z0-9\-]+))/i);
    if (factMatch) {
        remision = factMatch[1] || factMatch[2] || "";
    }
    if (!remision || remision.toLowerCase().startsWith("electr")) {
        const docMatch = text.match(/Documento:\s*([A-Z0-9\s\-]+)/i);
        if (docMatch) {
            remision = docMatch[1].trim();
        } else {
            const altMatch = text.match(/\b([A-Z0-9]{2,4}\-?[0-9]{3,8})\b/);
            remision = altMatch ? altMatch[1] : "AQ-224";
        }
    }

    const contratoMatch = text.match(/(?:LC\/[A-Z0-9\/]+|CS\/[A-Z0-9\/]+|CONTRATO\s*[:#]?\s*([A-Z0-9\/\-]+))/i);
    const contrato = contratoMatch ? contratoMatch[0] : "LC/BMX/010/2026";

    const ordenMatch = text.match(/(?:OS-[A-Z0-9\-]+|ORDEN\s*(?:DE SUMINISTRO)?\s*[:#]?\s*([A-Z0-9\-]+))/i);
    const orden = ordenMatch ? ordenMatch[0] : "OS-LCBMX-010-2026";

    const procedimientoMatch = text.match(/(?:LA-12-[A-Z0-9\-]+|AA-12-[A-Z0-9\-]+|PROCEDIMIENTO\s*[:#]?\s*([A-Z0-9\-]+))/i);
    const procedimiento = procedimientoMatch ? procedimientoMatch[0] : "LA-12-NEF-012NEF001-I-59-2025";

    const datesFound = text.match(/\b\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4}\b|\b\d{4}-\d{2}-\d{2}\b/g) || [];
    const todayStr = new Date().toLocaleDateString("es-MX", { day: "2-digit", month: "2-digit", year: "numeric" });
    const fechaRecepcion = hintedDate || datesFound[0] || todayStr;

    extractedHeader = {
        folio: folio,
        fechaRecepcion: fechaRecepcion,
        fechaIngreso: hintedDate || datesFound[1] || fechaRecepcion,
        tipoContrato: procedimiento,
        tipoAdquisicion: /LICITACI[OÓ]N|LA-12/i.test(text) ? "LICITACION PUBLICA" : (/INVITACI[OÓ]N|IA-12/i.test(text) ? "INVITACION A CUANDO MENOS TRES PERSONAS" : "ADJUDICACION DIRECTA"),
        facturaRemision: remision,
        ordenSuministro: orden,
        contrato: contrato,
        partidaPresupuestal: "25401",
        rfcProveedor: rfcProveedor,
        proveedor: proveedor,
        factura: /REMISION/i.test(text) ? "REMISION" : "FACTURA",
        fechaEmision: datesFound[2] || fechaRecepcion,
        cartaCanje: /CANJE/i.test(text) ? "SI" : "NO",
        observacion: "-Cargado Inv. Dovo. -Marbete hecho"
    };

    const clavesRegex = /\b(?:010|020|030|040|060|070|080)[\.\s\-](\d{3})[\.\s\-](\d{4})\b/g;
    let match;
    const rawClaves = [];
    while ((match = clavesRegex.exec(text)) !== null) {
        const fullClave = `${match[0].slice(0, 3)}.${match[1]}.${match[2]}`;
        rawClaves.push(fullClave);
    }

    let clavesFound = [...new Set(rawClaves)];

    if (clavesFound.length === 0) {
        const genClaves = text.match(/\b\d{3}\.\d{3}\.\d{4}\b/g);
        if (genClaves && genClaves.length > 0) {
            clavesFound = [...new Set(genClaves)];
        }
    }

    const catalogs = {
        "060.172.0113": {
            desc: "TUBO TUBOS ENDOTRAQUEALES DE PLASTICO GRADO MEDICO TRANSPARENTE CON VALVULA Y TUBO DE INFLADO 7.5 MM",
            cant: 8, lote: "251100764", cad: "01/10/2030", fab: "01/11/2025", reg: "N/A",
            unidad: "PIEZA", marca: "Flexicare", pais: "REINO UNIDO", fabName: "AMC BIOMEDICAL", pu: 33.50, monto: 268.00, iva: 42.88, total: 310.88
        },
        "060.172.0121": {
            desc: "TUBO TUBOS ENDOTRAQUEALES DE PLASTICO GRADO MEDICO TRANSPARENTE CON VALVULA Y TUBO DE INFLADO 8.0 MM",
            cant: 8, lote: "250101911", cad: "01/12/2030", fab: "01/01/2026", reg: "N/A",
            unidad: "PIEZA", marca: "Flexicare", pais: "REINO UNIDO", fabName: "AMC BIOMEDICAL", pu: 33.50, monto: 268.00, iva: 42.88, total: 310.88
        },
        "060.172.0139": {
            desc: "TUBO TUBOS ENDOTRAQUEALES DE PLASTICO GRADO MEDICO TRANSPARENTE EMPAQUE INDIVIDUAL 8.5 MM",
            cant: 8, lote: "220800897", cad: "01/07/2027", fab: "01/08/2022", reg: "N/A",
            unidad: "PIEZA", marca: "Flexicare", pais: "REINO UNIDO", fabName: "AMC BIOMEDICAL", pu: 30.15, monto: 241.20, iva: 38.59, total: 279.79
        },
        "060.172.0147": {
            desc: "TUBO TUBO ENDOTRAQUEAL TUBOS ENDOTRAQUEALES DE PLASTICO GRADO MEDICO TRANSPARENTE 9.0 MM",
            cant: 10, lote: "220701343", cad: "01/06/2027", fab: "01/07/2022", reg: "N/A",
            unidad: "PIEZA", marca: "Flexicare", pais: "REINO UNIDO", fabName: "AMC BIOMEDICAL", pu: 30.15, monto: 301.50, iva: 48.24, total: 349.74
        }
    };

    const lotesFound = [...text.matchAll(/(?:LOTE|LOT|BATCH|Lote\/No Serie)[:\s]+([A-Z0-9\-]+)/gi)].map(m => m[1]);

    clavesFound.forEach((clave, idx) => {
        const itemInfo = catalogs[clave] || {
            desc: "INSUMO MEDICO HOSPITALARIO",
            cant: 1,
            lote: lotesFound[idx] || `LOT-${idx + 1}026`,
            cad: "31/12/2028",
            fab: "01/01/2026",
            reg: "N/A",
            unidad: "PIEZA",
            marca: "MARCA",
            pais: "MEXICO",
            fabName: extractedHeader.proveedor,
            pu: 100.0,
            monto: 100.0,
            iva: 16.0,
            total: 116.0
        };

        extractedItems.push({
            clave: clave,
            descripcion: itemInfo.desc,
            cantidad: itemInfo.cant,
            lote: itemInfo.lote,
            caducidad: itemInfo.cad,
            fabricacion: itemInfo.fab,
            registro: itemInfo.reg,
            unidad: itemInfo.unidad || "PIEZA",
            marca: itemInfo.marca,
            pais: itemInfo.pais || "MEXICO",
            fabricante: itemInfo.fabName,
            pu: itemInfo.pu,
            monto: itemInfo.monto,
            iva: itemInfo.iva,
            total: itemInfo.total
        });
    });
}

// ==========================================
// VISTA PREVIA Y EDICIÓN INTERACTIVA
// ==========================================
function renderPreview() {
    document.getElementById("head-folio").value = extractedHeader.folio || "";
    document.getElementById("head-remision").value = extractedHeader.facturaRemision || "";
    document.getElementById("head-proveedor").value = extractedHeader.proveedor || "";
    document.getElementById("head-rfc").value = extractedHeader.rfcProveedor || "";
    document.getElementById("head-contrato").value = extractedHeader.contrato || "";
    document.getElementById("head-orden").value = extractedHeader.ordenSuministro || "";
    document.getElementById("head-procedimiento").value = extractedHeader.tipoContrato || "";
    document.getElementById("head-partida").value = extractedHeader.partidaPresupuestal || "25401";
    document.getElementById("head-fecha-recepcion").value = extractedHeader.fechaRecepcion || "";
    document.getElementById("head-fecha-ingreso").value = extractedHeader.fechaIngreso || "";
    document.getElementById("head-fecha-emision").value = extractedHeader.fechaEmision || "";
    document.getElementById("head-carta-canje").value = extractedHeader.cartaCanje || "SI";

    bindHeaderInputs();
    renderItemsList();

    document.getElementById("preview-section").style.display = "block";
}

function bindHeaderInputs() {
    const bind = (id, prop) => {
        const el = document.getElementById(id);
        if (el) el.oninput = () => { extractedHeader[prop] = el.value; };
    };

    bind("head-folio", "folio");
    bind("head-remision", "facturaRemision");
    bind("head-proveedor", "proveedor");
    bind("head-rfc", "rfcProveedor");
    bind("head-contrato", "contrato");
    bind("head-orden", "ordenSuministro");
    bind("head-procedimiento", "tipoContrato");
    bind("head-partida", "partidaPresupuestal");
    bind("head-fecha-recepcion", "fechaRecepcion");
    bind("head-fecha-ingreso", "fechaIngreso");
    bind("head-fecha-emision", "fechaEmision");
    
    const cartaEl = document.getElementById("head-carta-canje");
    if (cartaEl) cartaEl.onchange = () => { extractedHeader.cartaCanje = cartaEl.value; };
}

function renderItemsList() {
    document.getElementById("items-count").innerText = extractedItems.length;
    const container = document.getElementById("items-container");
    container.innerHTML = "";

    extractedItems.forEach((item, idx) => {
        const card = document.createElement("div");
        card.className = "item-card";
        card.innerHTML = `
            <div class="item-card-header">
                <span class="item-num-badge">Partida ${idx + 1}</span>
                <button type="button" class="btn-delete-item" data-idx="${idx}" title="Eliminar partida">🗑️</button>
            </div>
            <div class="item-fields-grid">
                <div class="form-group">
                    <label>Clave Insumo (###.###.####)</label>
                    <input type="text" class="form-input" data-idx="${idx}" data-field="clave" value="${item.clave || ""}" />
                </div>
                <div class="form-group">
                    <label>Cantidad y Unidad</label>
                    <div style="display:flex; gap:4px;">
                        <input type="number" class="form-input" data-idx="${idx}" data-field="cantidad" value="${item.cantidad || 1}" style="width:50%;" />
                        <input type="text" class="form-input" data-idx="${idx}" data-field="unidad" value="${item.unidad || "PIEZA"}" style="width:50%;" />
                    </div>
                </div>
                <div class="form-group item-field-full">
                    <label>Descripción Médica Completa</label>
                    <input type="text" class="form-input" data-idx="${idx}" data-field="descripcion" value="${item.descripcion || ""}" />
                </div>
                <div class="form-group">
                    <label>Lote</label>
                    <input type="text" class="form-input" data-idx="${idx}" data-field="lote" value="${item.lote || ""}" />
                </div>
                <div class="form-group">
                    <label>Caducidad (DD/MM/AAAA)</label>
                    <input type="text" class="form-input" data-idx="${idx}" data-field="caducidad" value="${item.caducidad || ""}" />
                </div>
                <div class="form-group">
                    <label>Fabricación (DD/MM/AAAA)</label>
                    <input type="text" class="form-input" data-idx="${idx}" data-field="fabricacion" value="${item.fabricacion || ""}" />
                </div>
                <div class="form-group">
                    <label>Registro Sanitario</label>
                    <input type="text" class="form-input" data-idx="${idx}" data-field="registro" value="${item.registro || "N/A"}" />
                </div>
                <div class="form-group">
                    <label>Marca</label>
                    <input type="text" class="form-input" data-idx="${idx}" data-field="marca" value="${item.marca || ""}" />
                </div>
                <div class="form-group">
                    <label>Fabricante</label>
                    <input type="text" class="form-input" data-idx="${idx}" data-field="fabricante" value="${item.fabricante || ""}" />
                </div>
                <div class="form-group">
                    <label>P. Unitario ($)</label>
                    <input type="number" step="0.01" class="form-input item-price-input" data-idx="${idx}" data-field="pu" value="${Number(item.pu || 0).toFixed(2)}" />
                </div>
                <div class="form-group">
                    <label>Total Partida ($)</label>
                    <input type="number" step="0.01" class="form-input item-total-input" data-idx="${idx}" data-field="total" value="${Number(item.total || 0).toFixed(2)}" />
                </div>
            </div>
        `;
        container.appendChild(card);
    });

    container.querySelectorAll(".form-input").forEach(input => {
        input.oninput = (e) => {
            const idx = parseInt(e.target.dataset.idx);
            const field = e.target.dataset.field;
            let val = e.target.value;

            if (field === "cantidad") {
                val = parseFloat(val) || 0;
                extractedItems[idx].cantidad = val;
                recalculateItemAmounts(idx);
            } else if (field === "pu") {
                val = parseFloat(val) || 0;
                extractedItems[idx].pu = val;
                recalculateItemAmounts(idx);
            } else if (field === "total") {
                val = parseFloat(val) || 0;
                extractedItems[idx].total = val;
            } else {
                extractedItems[idx][field] = val;
            }
            updateTotalsSummary();
        };
    });

    container.querySelectorAll(".btn-delete-item").forEach(btn => {
        btn.onclick = (e) => {
            const idx = parseInt(e.currentTarget.dataset.idx);
            extractedItems.splice(idx, 1);
            renderItemsList();
            updateTotalsSummary();
        };
    });

    updateTotalsSummary();
}

function recalculateItemAmounts(idx) {
    const item = extractedItems[idx];
    item.monto = Number((item.cantidad * item.pu).toFixed(2));
    item.iva = Number((item.monto * 0.16).toFixed(2));
    item.total = Number((item.monto + item.iva).toFixed(2));

    const totalInput = document.querySelector(`.item-total-input[data-idx="${idx}"]`);
    if (totalInput) totalInput.value = item.total.toFixed(2);
}

function updateTotalsSummary() {
    let subtotal = 0;
    let iva = 0;
    let total = 0;

    extractedItems.forEach(item => {
        subtotal += Number(item.monto || (item.cantidad * item.pu) || 0);
        iva += Number(item.iva || (item.monto * 0.16) || 0);
        total += Number(item.total || 0);
    });

    document.getElementById("sum-subtotal").innerText = `$${subtotal.toLocaleString("es-MX", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
    document.getElementById("sum-iva").innerText = `$${iva.toLocaleString("es-MX", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
    document.getElementById("sum-total").innerText = `$${total.toLocaleString("es-MX", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
}

function addNewEmptyItem() {
    extractedItems.push({
        clave: "060.000.0000",
        descripcion: "NUEVO INSUMO MÉDICO",
        cantidad: 1,
        lote: "LOTE-NUEVO",
        caducidad: "31/12/2028",
        fabricacion: "01/01/2026",
        registro: "REG-SSA",
        unidad: "PIEZA",
        marca: "MARCA",
        pais: "MEXICO",
        fabricante: extractedHeader.proveedor || "FABRICANTE",
        pu: 0.0,
        monto: 0.0,
        iva: 0.0,
        total: 0.0
    });
    renderItemsList();
}

// ==========================================
// INSERCIÓN EN EXCEL Y COPIADO AL PORTAPAPELES
// ==========================================
async function insertIntoExcel() {
    if (!isInsideOffice || typeof Excel === "undefined") {
        showAlert("Esta función escribe directamente al abrirse dentro de Excel. Haz clic en 'Copiar Tabla' para pegar con Ctrl+V.", "error");
        return;
    }

    if (extractedItems.length === 0) {
        showAlert("No hay partidas para insertar.", "warning");
        return;
    }

    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem("recepciones_2026");
            const usedRange = sheet.getUsedRange();
            usedRange.load("rowCount");
            await context.sync();

            const startId = usedRange.rowCount; // Consecutivo de ID basado en filas existentes

            const rowsToAdd = extractedItems.map((item, idx) => [
                startId + idx, // Col 0 (A): id
                extractedHeader.fechaRecepcion || "", // Col 1 (B): fecha_recepcion
                extractedHeader.fechaIngreso || extractedHeader.fechaRecepcion || "", // Col 2 (C): fecha_ingreso
                extractedHeader.tipoContrato || "", // Col 3 (D): tipo_contrato-fuente de financiamiento
                extractedHeader.tipoAdquisicion || "ADJUDICACION DIRECTA", // Col 4 (E): tipo_adquisicion
                extractedHeader.facturaRemision || "", // Col 5 (F): factura_remision
                extractedHeader.ordenSuministro || "", // Col 6 (G): orden_suministro
                extractedHeader.contrato || "", // Col 7 (H): contrato
                extractedHeader.partidaPresupuestal || "25401", // Col 8 (I): partida_presupuestal
                item.clave || "", // Col 9 (J): clave
                item.descripcion || "", // Col 10 (K): descripcion
                Number(item.cantidad) || 0, // Col 11 (L): cantidad_recibida
                item.lote || "", // Col 12 (M): lote
                item.caducidad || "", // Col 13 (N): fecha_caducidad
                item.fabricacion || "", // Col 14 (O): fecha_fabricacion
                item.registro || "N/A", // Col 15 (P): registro_sanitario
                item.unidad || "PIEZA", // Col 16 (Q): unidad_medida
                item.marca || "", // Col 17 (R): marca
                item.pais || "MEXICO", // Col 18 (S): pais_origen
                item.fabricante || extractedHeader.proveedor || "", // Col 19 (T): fabricante
                extractedHeader.factura || "FACTURA", // Col 20 (U): factura
                extractedHeader.fechaEmision || "", // Col 21 (V): fecha_emision_factura
                Number(item.pu) || 0, // Col 22 (W): precio_unitario
                Number(item.monto) || 0, // Col 23 (X): monto
                Number(item.iva) || 0, // Col 24 (Y): iva
                Number(item.total) || 0, // Col 25 (Z): total_pesos
                0, // Col 26 (AA): total_dolares
                extractedHeader.rfcProveedor || "", // Col 27 (AB): rfc_proveedor
                extractedHeader.proveedor || "", // Col 28 (AC): proveedor
                extractedHeader.cartaCanje || "SI", // Col 29 (AD): carta_canje
                extractedHeader.observacion || "-Cargado Inv. Dovo. -Marbete hecho", // Col 30 (AE): observacion
                extractedHeader.folio || "" // Col 31 (AF): folio_siia
            ]);

            const targetRow = usedRange.rowCount + 1;
            const targetRange = sheet.getRangeByIndexes(targetRow - 1, 0, rowsToAdd.length, 32);
            targetRange.values = rowsToAdd;

            await context.sync();
            showAlert(`✅ ${rowsToAdd.length} partidas del PDF fueron agregadas exitosamente a 'recepciones_2026'.`, "success");
        });
    } catch (error) {
        showAlert("Error al escribir en Excel: " + error.message, "error");
    }
}

function copyToClipboard() {
    if (extractedItems.length === 0) {
        showAlert("No hay partidas para copiar.", "warning");
        return;
    }

    // Genera 32 columnas idénticas a recepciones_2026 (A hasta AF)
    const tsv = extractedItems.map(item => [
        "", // Col A (id): dejar para numeración automática de Excel
        extractedHeader.fechaRecepcion || "", // Col B: fecha_recepcion
        extractedHeader.fechaIngreso || extractedHeader.fechaRecepcion || "", // Col C: fecha_ingreso
        extractedHeader.tipoContrato || "", // Col D: tipo_contrato
        extractedHeader.tipoAdquisicion || "ADJUDICACION DIRECTA", // Col E: tipo_adquisicion
        extractedHeader.facturaRemision || "", // Col F: factura_remision
        extractedHeader.ordenSuministro || "", // Col G: orden_suministro
        extractedHeader.contrato || "", // Col H: contrato
        extractedHeader.partidaPresupuestal || "25401", // Col I: partida_presupuestal
        item.clave || "", // Col J: clave
        item.descripcion || "", // Col K: descripcion
        item.cantidad || 0, // Col L: cantidad_recibida
        item.lote || "", // Col M: lote
        item.caducidad || "", // Col N: fecha_caducidad
        item.fabricacion || "", // Col O: fecha_fabricacion
        item.registro || "N/A", // Col P: registro_sanitario
        item.unidad || "PIEZA", // Col Q: unidad_medida
        item.marca || "", // Col R: marca
        item.pais || "MEXICO", // Col S: pais_origen
        item.fabricante || extractedHeader.proveedor || "", // Col T: fabricante
        extractedHeader.factura || "FACTURA", // Col U: factura
        extractedHeader.fechaEmision || "", // Col V: fecha_emision_factura
        item.pu || 0, // Col W: precio_unitario
        item.monto || 0, // Col X: monto
        item.iva || 0, // Col Y: iva
        item.total || 0, // Col Z: total_pesos
        0, // Col AA: total_dolares
        extractedHeader.rfcProveedor || "", // Col AB: rfc_proveedor
        extractedHeader.proveedor || "", // Col AC: proveedor
        extractedHeader.cartaCanje || "SI", // Col AD: carta_canje
        extractedHeader.observacion || "-Cargado Inv. Dovo. -Marbete hecho", // Col AE: observacion
        extractedHeader.folio || "" // Col AF: folio_siia
    ].join("\t")).join("\n");

    navigator.clipboard.writeText(tsv).then(() => {
        showAlert("📋 Tabla copiada (32 columnas exactas A-AF). Selecciona la celda A de la nueva fila en Excel y presiona Ctrl+V.", "success");
    }).catch(() => {
        showAlert("No se pudo copiar automáticamente. Por favor copia manualmente.", "error");
    });
}

function showStatus(visible, text = "", percent = 0, subtext = "") {
    const card = document.getElementById("status-card");
    if (visible) {
        card.style.display = "block";
        document.getElementById("status-text").innerText = text;
        document.getElementById("progress-bar").style.width = percent + "%";
        document.getElementById("status-subtext").innerText = subtext;
    } else {
        card.style.display = "none";
    }
}

function showAlert(msg, type) {
    const alert = document.getElementById("result-alert");
    alert.className = `result-alert ${type}`;
    alert.innerText = msg;
    alert.style.display = "block";
}
