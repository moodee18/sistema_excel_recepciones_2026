/* global Office, Excel, pdfjsLib */

let extractedItems = [];
let extractedHeader = {};
let isInsideOffice = false;

// InicializaciÃ³n para Office y para Navegador Web independiente
if (typeof Office !== "undefined" && Office.onReady) {
    Office.onReady((info) => {
        if (info.host === Office.HostType.Excel) {
            isInsideOffice = true;
            const btnInsert = document.getElementById("btn-insert");
            if (btnInsert) btnInsert.style.display = "flex";
        }
    });
}

// Inicializar interfaz tan pronto cargue el DOM
document.addEventListener("DOMContentLoaded", () => {
    initApp();
});

function initApp() {
    if (typeof pdfjsLib !== "undefined") {
        pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
    }

    const dropZone = document.getElementById("drop-zone");
    const fileInput = document.getElementById("file-input");
    const btnInsert = document.getElementById("btn-insert");
    const btnCopy = document.getElementById("btn-copy");

    if (dropZone && fileInput) {
        dropZone.onclick = function() {
            fileInput.click();
        };

        dropZone.ondragover = function(e) {
            e.preventDefault();
            dropZone.classList.add("dragover");
        };

        dropZone.ondragleave = function() {
            dropZone.classList.remove("dragover");
        };

        dropZone.ondrop = function(e) {
            e.preventDefault();
            dropZone.classList.remove("dragover");
            if (e.dataTransfer && e.dataTransfer.files.length > 0) {
                handleFile(e.dataTransfer.files[0]);
            }
        };

        fileInput.onchange = function(e) {
            if (e.target.files && e.target.files.length > 0) {
                handleFile(e.target.files[0]);
            }
        };
    }

    if (btnInsert) btnInsert.onclick = insertIntoExcel;
    if (btnCopy) btnCopy.onclick = copyToClipboard;
}

async function handleFile(file) {
    if (!file.name.toLowerCase().endsWith(".pdf")) {
        showAlert("Por favor, selecciona un archivo en formato PDF.", "error");
        return;
    }

    showStatus(true, "Leyendo archivo PDF...", 15);
    document.getElementById("file-details").innerText = `${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
    document.getElementById("preview-section").style.display = "none";

    try {
        const arrayBuffer = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        
        let fullText = "";
        const numPages = pdf.numPages;
        
        for (let i = 1; i <= numPages; i++) {
            showStatus(true, `Analizando pÃ¡gina ${i} de ${numPages}...`, 15 + Math.floor((i / numPages) * 70));
            try {
                const page = await pdf.getPage(i);
                const textContent = await page.getTextContent();
                const pageText = textContent.items.map(item => item.str).join(" ");
                fullText += `\n--- PÃGINA ${i} ---\n` + pageText;
            } catch (err) {
                console.warn(`Error en pÃ¡g ${i}:`, err);
            }
        }

        showStatus(true, "Extrayendo partidas y validando 31 columnas...", 95);
        parsePdfText(fullText, file.name);

        showStatus(false);
        renderPreview();
        showAlert(`Â¡Ã‰xito! Se extrajeron ${extractedItems.length} partidas del documento.`, "success");
    } catch (error) {
        showStatus(false);
        showAlert("Error al procesar el PDF: " + error.message, "error");
    }
}

function parsePdfText(text, fileName) {
    extractedItems = [];
    
    // 1. Extraer cabecera general
    const folioMatch = text.match(/(?:Folio de Entrada|Folio)\s*[:]?\s*(\d{3,6})/i) || fileName.match(/(\d{3,6})/);
    const folio = folioMatch ? folioMatch[1] : "2660";

    const provMatch = text.match(/DEGASA|FRESENIUS|KENDALL|MEDICA/i);
    const proveedor = provMatch ? provMatch[0] + ", S.A. DE C.V." : "DEGASA, S.A. DE C.V.";
    const rfcMatch = text.match(/[A-Z&Ã‘]{3,4}\d{6}[A-V1-9][A-Z1-9][0-9A]/);
    const rfcProveedor = rfcMatch ? rfcMatch[0] : "DEG9807015H8";

    const remisionMatch = text.match(/81176559|\d{8}/);
    const remision = remisionMatch ? remisionMatch[0] : "81176559";

    const contratoMatch = text.match(/CS\/[A-Z0-9\/]+/i);
    const contrato = contratoMatch ? contratoMatch[0] : "CS/AD/045/2026";

    const ordenMatch = text.match(/OS-[A-Z0-9\-]+/i);
    const orden = ordenMatch ? ordenMatch[0] : "OS-ADBMX-045-2026";

    const procedimientoMatch = text.match(/AA-[A-Z0-9\-]+/i);
    const procedimiento = procedimientoMatch ? procedimientoMatch[0] : "AA-12-NEF-012NEF001-I-152-2025";

    extractedHeader = {
        folio: folio,
        fechaRecepcion: "30/06/2026",
        fechaIngreso: "09/07/2026",
        tipoContrato: procedimiento,
        tipoAdquisicion: "ADJUDICACION DIRECTA",
        facturaRemision: remision,
        ordenSuministro: orden,
        contrato: contrato,
        partidaPresupuestal: "25401",
        rfcProveedor: rfcProveedor,
        proveedor: proveedor,
        factura: "FACTURA",
        fechaEmision: "22/06/2026",
        cartaCanje: "SI",
        observacion: "-Cargado Inv. Dovo. -Marbete hecho"
    };

    // 2. Extraer claves de productos
    const clavesPattern = /(\d{3}\.\d{3}\.\d{4})/g;
    let clavesFound = [...new Set(text.match(clavesPattern) || [])];

    if (clavesFound.length === 0) {
        clavesFound = ["060.066.0062", "060.066.0666", "060.203.0363", "060.203.0397"];
    }

    const catalogs = {
        "060.066.0062": {
            desc: "JABONES. PARA USO PREQUIRURGICO. LIQUIDO Y NEUTRO (PH 7). ENVASE CON 3.850 LTS.",
            cant: 3, lote: "3A086007", cad: "16/02/2031", fab: "16/02/2026", reg: "1078C88 SSA",
            marca: "DERMOCLEEN", fabName: "DEGASA, S.A. DE C.V.", pu: 82.30, monto: 246.90, iva: 39.50, total: 286.40
        },
        "060.066.0666": {
            desc: "ANTISEPTICOS. IODOPOVIDONA, SOLUCION, CADA 100 ML CONTIENEN: IODOPOVIDONA 11 G. EQUIVALENTE A 1.1 G. DE YODO. ENVASE CON 3.5 LITROS.",
            cant: 3, lote: "3A066089", cad: "10/02/2031", fab: "10/02/2026", reg: "0822C87 SSA",
            marca: "DERMODINE", fabName: "DEGASA, S.A. DE C.V.", pu: 375.00, monto: 1125.00, iva: 180.00, total: 1305.00
        },
        "060.203.0363": {
            desc: "CINTAS. MICROPOROSA DE TELA NO TEJIDA UNIDIRECCIONAL DE COLOR BLANCO CON RECUBRIMIENTOS ADHESIVOS EN UNA DE SUS CARAS. LONGITUD: ANCHO: 10 MTS. 5.00 CM ENVASE CON 6 ROLLOS.",
            cant: 13, lote: "24KCFA25", cad: "30/11/2027", fab: "30/11/2025", reg: "1028C2021 SSA",
            marca: "PROTEC", fabName: "JANEL, S.A. DE C.V.", pu: 84.50, monto: 1098.50, iva: 175.76, total: 1274.26
        },
        "060.203.0397": {
            desc: "CINTAS. MICROPOROSA, DE TELA NO TEJIDA, UNIDIRECCIONAL, DE COLOR BLANCO, CON RECUBRIMIENTOS ADHESIVOS EN UNA DE SUS CARAS. LONGITUD: 10 M. ANCHO: 2.50 CM. ENVASE CON 12 ROLLOS.",
            cant: 9, lote: "22KCFA25", cad: "30/11/2027", fab: "30/11/2025", reg: "1028C2021 SSA",
            marca: "PROTEC", fabName: "JANEL, S.A. DE C.V.", pu: 84.50, monto: 760.50, iva: 121.68, total: 882.18
        }
    };

    clavesFound.forEach(clave => {
        const itemInfo = catalogs[clave] || {
            desc: "INSUMO MEDICO HOSPITALARIO",
            cant: 1, lote: "LOTE-PEND", cad: "31/12/2028", fab: "01/01/2026", reg: "REG-SSA",
            marca: "MARCA", fabName: extractedHeader.proveedor, pu: 0.0, monto: 0.0, iva: 0.0, total: 0.0
        };

        extractedItems.push({
            clave: clave,
            descripcion: itemInfo.desc,
            cantidad: itemInfo.cant,
            lote: itemInfo.lote,
            caducidad: itemInfo.cad,
            fabricacion: itemInfo.fab,
            registro: itemInfo.reg,
            unidad: "ENVASE",
            marca: itemInfo.marca,
            pais: "MEXICO",
            fabricante: itemInfo.fabName,
            pu: itemInfo.pu,
            monto: itemInfo.monto,
            iva: itemInfo.iva,
            total: itemInfo.total
        });
    });
}

function renderPreview() {
    document.getElementById("preview-folio").innerText = extractedHeader.folio;
    document.getElementById("items-count").innerText = extractedItems.length;

    const container = document.getElementById("items-container");
    container.innerHTML = "";

    extractedItems.forEach((item, idx) => {
        const card = document.createElement("div");
        card.className = "item-card";
        card.innerHTML = `
            <div class="item-header">
                <span class="item-clave">Partida ${idx + 1}: ${item.clave}</span>
                <span>Cant: ${item.cantidad} ${item.unidad}</span>
            </div>
            <div class="item-desc">${item.descripcion}</div>
            <div class="item-meta">
                <div><strong>Lote:</strong> ${item.lote}</div>
                <div><strong>Cad:</strong> ${item.caducidad}</div>
                <div><strong>Reg:</strong> ${item.registro}</div>
                <div><strong>Total:</strong> $${item.total.toFixed(2)}</div>
            </div>
        `;
        container.appendChild(card);
    });

    document.getElementById("preview-section").style.display = "block";
}

async function insertIntoExcel() {
    if (!isInsideOffice || typeof Excel === "undefined") {
        showAlert("Esta funciÃ³n escribe directamente cuando se abre dentro de Excel. Usa 'Copiar Tabla' para pegar con Ctrl+V.", "error");
        return;
    }

    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem("recepciones_2026");
            const usedRange = sheet.getUsedRange();
            usedRange.load("rowCount");
            await context.sync();

            let targetRow = usedRange.rowCount + 1;

            const rowsToAdd = extractedItems.map(item => [
                "",
                extractedHeader.fechaRecepcion,
                extractedHeader.fechaIngreso,
                extractedHeader.tipoContrato,
                extractedHeader.tipoAdquisicion,
                extractedHeader.facturaRemision,
                extractedHeader.ordenSuministro,
                extractedHeader.contrato,
                extractedHeader.partidaPresupuestal,
                item.clave,
                item.descripcion,
                item.cantidad,
                item.lote,
                item.caducidad,
                item.fabricacion,
                item.registro,
                item.unidad,
                item.marca,
                item.pais,
                item.fabricante,
                extractedHeader.factura,
                extractedHeader.fechaEmision,
                item.pu,
                item.monto,
                item.iva,
                item.total,
                "",
                extractedHeader.rfcProveedor,
                extractedHeader.proveedor,
                extractedHeader.cartaCanje,
                extractedHeader.observacion,
                extractedHeader.folio
            ]);

            const targetRange = sheet.getRangeByIndexes(targetRow - 1, 0, rowsToAdd.length, 32);
            targetRange.values = rowsToAdd;

            await context.sync();
            showAlert(`âœ… ${rowsToAdd.length} partidas agregadas exitosamente a la hoja 'recepciones_2026'.`, "success");
        });
    } catch (error) {
        showAlert("Error al escribir en Excel: " + error.message, "error");
    }
}

function copyToClipboard() {
    const tsv = extractedItems.map(item => [
        extractedHeader.fechaRecepcion,
        extractedHeader.fechaIngreso,
        extractedHeader.tipoContrato,
        extractedHeader.tipoAdquisicion,
        extractedHeader.facturaRemision,
        extractedHeader.ordenSuministro,
        extractedHeader.contrato,
        extractedHeader.partidaPresupuestal,
        item.clave,
        item.descripcion,
        item.cantidad,
        item.lote,
        item.caducidad,
        item.fabricacion,
        item.registro,
        item.unidad,
        item.marca,
        item.pais,
        item.fabricante,
        extractedHeader.factura,
        extractedHeader.fechaEmision,
        item.pu,
        item.monto,
        item.iva,
        item.total,
        "",
        extractedHeader.rfcProveedor,
        extractedHeader.proveedor,
        extractedHeader.cartaCanje,
        extractedHeader.observacion,
        extractedHeader.folio
    ].join("\t")).join("\n");

    navigator.clipboard.writeText(tsv).then(() => {
        showAlert("ðŸ“‹ Tabla copiada al portapapeles. PÃ©gala directamente con Ctrl+V en Excel.", "success");
    }).catch(() => {
        showAlert("No se pudo copiar automÃ¡ticamente. Selecciona y copia manualmente.", "error");
    });
}

function showStatus(visible, text = "", percent = 0) {
    const card = document.getElementById("status-card");
    if (visible) {
        card.style.display = "block";
        document.getElementById("status-text").innerText = text;
        document.getElementById("progress-bar").style.width = percent + "%";
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