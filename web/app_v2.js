/**
 * AQUASHIELD · ExportDesk - Lógica Principal V2.1
 * Auditoría Comex | Legalización DUS | Control de BL
 *
 * Fixes v2.1 (Auditoría Técnica):
 *   C-03: Deduplicar BLs al agregar (evita acumulación de archivos repetidos)
 *   C-04: Rechazar archivos no reconocidos en identifyExcel (no asumir BL)
 *   C-05: escHtml() — sanitizar innerHTML en renderResults y renderDUS (XSS)
 *   A-03: saveToMemory usa includes() directo en vez de regex de emojis
 *   A-04: Reemplazar sentinel __ninguno__ por bandera booleana
 *   A-06: Capturar QuotaExceededError en localStorage
 *   M-02: Consolidar parseExcelDate/excelDateToJS en excelSerialToDate()
 *   M-03: getEstadoGeneral separa ⚠️ Error Fecha de ✅ Procesado
 *   M-05: procesarDUS() solo se ejecuta si hay datos DUS cargados
 */

const state = {
    sap: null,
    bls: [],
    maestros: {
        asignacion: null,
        procesados: null,
        analistas: null
    },
    dus: null,
    results: [],
    dusResults: [],
    currentModule: 'auditoria',
    chart: null,
    fechaHoy: new Date(), // Fecha real del sistema
    searchQuery: '',
    sortState:    { col: null, dir: 'asc' },
    sortStateDUS: { col: null, dir: 'asc' },
    activeGrupo: null   // null = Todos, 'congelado', 'fresco'
};

// UI Elements
const dropZones = {
    smart: document.getElementById('drop-smart'),
    sap: document.getElementById('drop-sap'),
    bl: document.getElementById('drop-bl'),
    maestro: document.getElementById('drop-maestro'),
    procesados: document.getElementById('drop-procesados'),
    analistas: document.getElementById('drop-analistas'),
    dus: document.getElementById('drop-dus')
};

const runBtn = document.getElementById('run-audit-btn');
const exportBtn = document.getElementById('export-excel-btn');
const exportBtnDus = document.getElementById('export-excel-btn-dus');

// --- Initialization ---
document.addEventListener('DOMContentLoaded', () => {
    setupDropZones();
    initChart();
    
    runBtn.addEventListener('click', runAudit);
    const runDusBtn = document.getElementById('run-dus-btn');
    if (runDusBtn) runDusBtn.addEventListener('click', runDUS);
    exportBtn.addEventListener('click', exportToExcel);
    if (exportBtnDus) exportBtnDus.addEventListener('click', exportToExcel);

    // Ordenamiento por columnas (click en encabezados)
    document.querySelectorAll('th.sortable').forEach(th => {
        th.addEventListener('click', () => {
            const col = th.dataset.col;
            const tableId = th.closest('table').id;
            const stateKey = tableId === 'dus-table' ? 'sortStateDUS' : 'sortState';
            const current = state[stateKey];
            const newDir  = (current.col === col && current.dir === 'asc') ? 'desc' : 'asc';
            state[stateKey] = { col, dir: newDir };
            // Actualizar indicadores visuales
            document.querySelectorAll(`#${tableId} th.sortable`).forEach(t => {
                t.classList.remove('sort-asc', 'sort-desc');
            });
            th.classList.add(newDir === 'asc' ? 'sort-asc' : 'sort-desc');
            // Re-renderizar
            if (tableId === 'dus-table') {
                renderDUS(getSelectedValues('dus-multi-select'));
            } else {
                renderResults(getSelectedValues('status-multi-select'));
            }
        });
    });

    // Multi-Select Logic
    setupMultiSelect('status-multi-select', 'ms-trigger', 'ms-label', (selected) => {
        renderResults(selected);
        updateExportButtonText(selected);
    });

    setupMultiSelect('dus-multi-select', 'ms-dus-trigger', 'ms-dus-label', (selected) => {
        renderDUS(selected);
        updateExportButtonText(selected);
    });

    // Tab Logic
    const tabBtns = document.querySelectorAll('.tab-btn');
    tabBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            const module = btn.getAttribute('data-module');
            switchTab(module);
        });
    });

    // Aplicar estado inicial de widgets según módulo por defecto (auditoria)
    switchTab('auditoria');

    // Botones de Memoria - toggle vista
    document.getElementById('memory-toggle-btn')?.addEventListener('click', () => toggleMemoryView('auditoria'));
    document.getElementById('memory-toggle-btn-dus')?.addEventListener('click', () => toggleMemoryView('dus'));

    // Menú tres puntos - toggle dropdown
    ['auditoria', 'dus'].forEach(mod => {
        const menuBtn = document.getElementById(`memory-menu-btn-${mod}`);
        const dropdown = document.getElementById(`memory-dropdown-${mod}`);
        if (menuBtn && dropdown) {
            menuBtn.addEventListener('click', (e) => {
                e.stopPropagation();
                // cerrar todos los otros dropdowns
                document.querySelectorAll('.memory-menu-dropdown').forEach(d => {
                    if (d !== dropdown) d.classList.remove('open');
                });
                dropdown.classList.toggle('open');
            });
        }
    });

    // Cerrar menú al hacer clic fuera (ignorar clicks dentro del dropdown)
    document.addEventListener('click', (e) => {
        if (!e.target.closest('.memory-menu-dropdown') && !e.target.closest('.memory-menu-trigger')) {
            document.querySelectorAll('.memory-menu-dropdown').forEach(d => d.classList.remove('open'));
        }
    });

    // Cargar contadores de memoria al inicio
    updateMemoryUI();

    // Directorio de Equipo - modal
    document.getElementById('team-btn')?.addEventListener('click', openTeamModal);
    document.getElementById('team-modal-close')?.addEventListener('click', closeTeamModal);
    document.getElementById('team-modal-overlay')?.addEventListener('click', (e) => {
        if (e.target.id === 'team-modal-overlay') closeTeamModal();
    });
    document.getElementById('team-add-btn')?.addEventListener('click', addTeamMember);
    // Enter en inputs
    ['team-input-code', 'team-input-name'].forEach(id => {
        document.getElementById(id)?.addEventListener('keydown', e => { if (e.key === 'Enter') addTeamMember(); });
    });

    // Gráfico de Analistas
    seedTeamDirectory(); // precarga el equipo por defecto
    initAnalystChart();

    // Filtros del gráfico de analistas
    document.getElementById('analyst-month-filter')?.addEventListener('change', refreshAnalystChart);

    // Multi-select analistas
    setupAnalystMultiSelect();

    // Status filters del gráfico de analistas
    document.querySelectorAll('.analyst-status-filter').forEach(cb => {
        cb.addEventListener('change', refreshAnalystChart);
    });

    console.log("AQUASHIELD · ExportDesk cargado correctamente");
});

function setupMultiSelect(containerId, triggerId, labelId, onChange) {
    const container = document.getElementById(containerId);
    const trigger = document.getElementById(triggerId);
    if (!container || !trigger) return;

    // Trigger ONCLICK robusto y libre de conflictos
    trigger.onclick = (e) => {
        e.stopPropagation();
        const isOpen = container.classList.contains('open');
        // Cerrar todos primero
        document.querySelectorAll('.multi-select').forEach(ms => ms.classList.remove('open'));
        // Si no estaba abierto, abrirlo ahora
        if (!isOpen) container.classList.add('open');
    };

    // Actualizamos los checkboxes (esto ocurre cada vez que se regeneran los datos)
    const checkboxes = container.querySelectorAll('.ms-option input');
    checkboxes.forEach(cb => {
        cb.onchange = () => {
            const checked = Array.from(container.querySelectorAll('input:checked')).map(i => i.value);
            updateMultiSelectLabel(checked, labelId);
            onChange(checked);
        };
    });
}

// Cierre global al hacer clic fuera de cualquier multi-select
document.addEventListener('click', (e) => {
    if (!e.target.closest('.multi-select')) {
        document.querySelectorAll('.multi-select').forEach(ms => ms.classList.remove('open'));
    }
});

function switchTab(module) {
    state.currentModule = module;
    
    // UI Updates - Tabs
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.toggle('active', b.getAttribute('data-module') === module));
    
    const isAuditoria = module === 'auditoria';
    document.getElementById('auditoria-results-container').style.display = isAuditoria ? 'block' : 'none';
    document.getElementById('dus-results-container').style.display = isAuditoria ? 'none' : 'block';
    
    document.querySelector('[data-for="auditoria"]').style.display = isAuditoria ? 'block' : 'none';
    document.querySelector('[data-for="dus"]').style.display = isAuditoria ? 'none' : 'block';

    // --- Mostrar/Ocultar widgets según módulo ---
    // Auditoría Comex: SAP + BLs + Asignación + Procesados + Analistas (sin DUS)
    // Legalización DUS: SAP + Analistas + DUS (sin BLs, sin Asignación, sin Procesados)
    // NOTA: código unificado en un solo bloque — el primer bloque antiguo era redundante
    // y tenía un ternario muerto (.closest('.drop-zone') siempre verdadero para IDs de drop-zone)
    ['drop-bl', 'drop-maestro', 'drop-procesados'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.style.display = isAuditoria ? '' : 'none';
    });
    const elDUS = document.getElementById('drop-dus');
    if (elDUS) elDUS.style.display = isAuditoria ? 'none' : '';

    // Actualizar texto de zona inteligente
    const smartInfo = document.getElementById('smart-info');
    if (smartInfo) {
        smartInfo.innerText = isAuditoria
            ? 'Auto-detectará DUS, SAP, BLs, y Maestros'
            : 'Auto-detectará SAP, Analista Facturación y DUS';
    }

    // Botón ejecutar: mostrar el correcto según módulo
    const runDusBtnEl = document.getElementById('run-dus-btn');
    if (runBtn)      { runBtn.style.display      = isAuditoria ? '' : 'none'; runBtn.disabled = true; runBtn.classList.remove('glow'); }
    if (runDusBtnEl) { runDusBtnEl.style.display = isAuditoria ? 'none' : ''; runDusBtnEl.disabled = true; runDusBtnEl.classList.remove('glow'); }
    checkReadyState();

    const selected = isAuditoria ? getSelectedValues('status-multi-select') : getSelectedValues('dus-multi-select');
    updateExportButtonText(selected);
    
    // Update Chart for Current Module
    if (isAuditoria) {
        updateChart(state.results);
        // Mostrar widgets exclusivos de Auditoría
        const treemapW = document.getElementById('treemap-chart-widget');
        if (treemapW) treemapW.style.display = '';
    } else {
        updateChartForDUS(state.dusResults);
    }
    
    updateStats();

    // Actualizar gráfico de analistas según módulo activo
    refreshAnalystFilters();
    refreshAnalystChart();

    // Mostrar/ocultar filtro "Sin Zarpe" segun modulo
    const szFilter = document.getElementById('analyst-sinzarpe-filter');
    if (szFilter) szFilter.style.display = isAuditoria ? 'none' : 'flex';

    // Widget DUS Observaciones: visible solo en modulo DUS
    const obsWidget = document.getElementById('dus-obs-widget');
    if (obsWidget) {
        obsWidget.style.display = isAuditoria ? 'none' : '';
        if (!isAuditoria && !obsWidget._obsSetup) {
            setupObsWidget();
            obsWidget._obsSetup = true;
        }
    }

    // KPI Panel: actualizar según módulo activo
    renderKPIPanel(isAuditoria ? state.results : state.dusResults);
}

function updateDynamicFilters(results, containerId, triggerId, labelId, renderFn, module) {
    const container = document.getElementById(containerId);
    if (!container) return;

    const optionsContainer = container.querySelector('.multi-select-options');
    if (!optionsContainer) return;

    // Extraer estados únicos presentes en los resultados
    // Para Auditoría: agrupar por ESTADO GENERAL en lugar de ESTATUS_FINAL individual
    const isAuditMod = (module === 'auditoria');
    let states;
    if (isAuditMod) {
        // Generar opciones únicas de ESTADO GENERAL
        const egSet = new Set(results.map(r => getEstadoGeneral(String(r.ESTATUS_FINAL)).label));
        states = Array.from(egSet).sort();
    } else {
        states = Array.from(new Set(results.map(r => String(r.ESTATUS_FINAL).trim()))).sort();
    }
    
    if (states.length === 0) {
        optionsContainer.innerHTML = '<div style="padding:10px; opacity:0.5; font-size:0.8rem">No hay datos</div>';
        return;
    }

    optionsContainer.innerHTML = states.map(s => {
        return `
            <label class="ms-option">
                <input type="checkbox" value="${s}" checked> 
                <span>${s}</span>
            </label>
        `;
    }).join('');

    // Re-setup events
    setupMultiSelect(containerId, triggerId, labelId, (selected) => {
        renderFn(selected);
        updateExportButtonText(selected);
        updateMultiSelectLabel(selected, labelId);
    });

    // Reset label
    updateMultiSelectLabel(states, labelId);
}

function updateMultiSelectLabel(selected, labelId = 'ms-label') {
    const label = document.getElementById(labelId);
    if (!label) return;
    
    if (selected.length === 0) {
        label.innerText = "🔍 Ninguno";
    } else {
        const container = label.closest('.multi-select');
        const total = container.querySelectorAll('.ms-option').length;
        
        if (selected.length === total) {
            label.innerText = labelId === 'ms-label' ? "🔍 Todos los Estados" : "🔍 Todos los DUS";
        } else if (selected.length === 1) {
            label.innerText = `🔍 ${selected[0]}`;
        } else {
            label.innerText = `🔍 ${selected.length} seleccionados`;
        }
    }
}

// --- Drop Zone Logic ---
function setupDropZones() {
    Object.entries(dropZones).forEach(([key, zone]) => {
        // Prevenir comportamiento por defecto
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            zone.addEventListener(eventName, e => {
                e.preventDefault();
                e.stopPropagation();
            }, false);
        });

        zone.addEventListener('dragover', () => {
            zone.classList.add('drag-over');
        });

        zone.addEventListener('dragleave', () => {
            zone.classList.remove('drag-over');
        });

        zone.addEventListener('drop', (e) => {
            zone.classList.remove('drag-over');
            const files = Array.from(e.dataTransfer.files);
            handleFiles(key, files, zone);
        });

        zone.addEventListener('click', () => {
            const input = document.createElement('input');
            input.type = 'file';
            input.multiple = zone.hasAttribute('multi');
            input.onchange = (e) => {
                const files = Array.from(e.target.files);
                handleFiles(key, files, zone);
            };
            input.click();
        });
    });
}

async function handleFiles(type, files, zone) {
    const infoDiv = zone ? zone.querySelector('.file-info') : null;
    if (!files || files.length === 0) return;

    try {
        if (type === 'smart') {
            if (infoDiv) infoDiv.innerText = `Procesando ${files.length} archivos...`;
            let successCount = 0;
            for (const file of files) {
                const data = await readExcel(file);
                const detectedType = identifyExcel(data, file.name);
                if (!detectedType) {
                    // C-04: archivo no reconocido — mostrar aviso en la zona inteligente
                    // P-04: escHtml en file.name para evitar XSS con nombres de archivo maliciosos
                    if (infoDiv) infoDiv.innerHTML += `<br><span style="color:#f97316">⚠️ "${escHtml(file.name)}" no reconocido — arrástralo a la zona correcta</span>`;
                    continue;
                }
                const targetZone = dropZones[detectedType];
                assignDataToType(detectedType, data, targetZone, file.name);
                successCount++;
            }
            if (infoDiv) infoDiv.innerHTML = `<span style="color:#f97316">✅ ${successCount} Archivos Autofilados</span>`;
            zone.style.borderColor = '#f97316';
            zone.style.background = 'rgba(249, 115, 22, 0.05)';
        } else {
            for (const file of files) {
                if (infoDiv) infoDiv.innerText = `Leyendo ${file.name}...`;
                const data = await readExcel(file);
                assignDataToType(type, data, zone, file.name);
            }
        }
    } catch (err) {
        console.error("Error procesando Excel:", err);
        if (infoDiv) infoDiv.innerText = `❌ Error: ${err.message}`;
        if (zone) zone.style.borderColor = '#f87171';
    }
    
    checkReadyState();
}

function identifyExcel(data, filename = "") {
    const fn = String(filename).toLowerCase();
    
    // 1. Detección Primaria por Nombre de Archivo
    if (fn.includes('export') || fn.includes('flujo') || fn.includes('sap.xlsx') || fn.includes('1.export')) return 'sap';
    if (fn.includes('dus') || fn.includes('6.dus')) return 'dus';
    if (fn.includes('asignacion') || fn.includes('asignación') || fn.includes('maestro clientes')) return 'maestro';
    if (fn.includes('analista')) return 'analistas';
    if (fn.includes('procesados') || fn.includes('excepciones')) return 'procesados';
    if (fn.includes('bill') || fn.includes('bls') || fn.includes('lading')) return 'bl';

    // 2. Detección Secundaria por Columnas (Si el nombre es genérico)
    if (data && data.length > 0) {
        for (let i = 0; i < Math.min(data.length, 5); i++) {
            const keys = Object.values(data[i]).map(v => String(v).toLowerCase());
            const hasMatch = (col) => keys.some(v => v.includes(col.toLowerCase()));
            
            if (hasMatch('texto breve de material') || hasMatch('clase de movimiento') || hasMatch('nº doc.compras')) return 'sap';
            if (hasMatch('estado dus') || hasMatch('aprob.dus')) return 'dus';
            if (hasMatch('analista asignado') || (hasMatch('cliente') && hasMatch('responsable'))) return 'maestro';
            if (hasMatch('factura') && hasMatch('creado por')) return 'analistas';
            if (hasMatch('número solicitante') && hasMatch('centro')) return 'analistas';
            if (hasMatch('numero solicitante') && hasMatch('centro')) return 'analistas';
            if ((hasMatch('estado auditoría') && hasMatch('motivo')) || (hasMatch('pedido flujo') && hasMatch('motivo'))) return 'procesados';
            if (hasMatch('bl') || hasMatch('bill') || hasMatch('lading') || hasMatch('pedido')) return 'bl';
        }
    }
    
    // C-04: Rechazar archivos no reconocidos — NO asumir BL.
    // Devolver null para que assignDataToType lo ignore y pida al usuario arrastrar manualmente.
    console.warn(`[identifyExcel] Archivo no reconocido: "${filename}". Se ignorará.`);
    return null;
}

function assignDataToType(type, data, zone, filename = "") {
    const infoDiv = zone ? zone.querySelector('.file-info') : null;
    // C-04: ignorar archivos que no pudieron identificarse
    // P-04: escHtml en filename para evitar XSS
    if (!type) {
        if (infoDiv) infoDiv.innerHTML = `<span style="color:#f97316">⚠️ No reconocido: ${escHtml(filename)}.<br>Arrástralo a la zona correcta.</span>`;
        if (zone) zone.style.borderColor = '#f97316';
        return;
    }
    if (type === 'sap') {
        state.sap = data;
        if(infoDiv) { infoDiv.innerHTML = `<span style="color:#4ade80">✅ SAP cargado</span>`; zone.style.borderColor = '#4ade80'; zone.style.background = 'rgba(74, 222, 128, 0.05)'; }
    } else if (type === 'bl') {
        state.bls.push(...data);
        // C-03: Deduplicar inmediatamente — si el mismo pedido aparece más de una vez,
        // la última versión gana (útil cuando se sube un BL corregido en la misma sesión)
        const blDedup = new Map();
        state.bls.forEach(row => {
            const key = cleanPedido(row.Pedido);
            if (key) blDedup.set(key, row);
        });
        state.bls = Array.from(blDedup.values());
        if(infoDiv) { infoDiv.innerHTML = `<span style="color:#4ade80">✅ ${state.bls.length} registros BL únicos</span>`; zone.style.borderColor = '#4ade80'; zone.style.background = 'rgba(74, 222, 128, 0.05)'; }
    } else if (type === 'maestro') {
        state.maestros.asignacion = data;
        if(infoDiv) { infoDiv.innerHTML = `<span style="color:#4ade80">✅ Maestro Clientes OK</span>`; zone.style.borderColor = '#4ade80'; zone.style.background = 'rgba(74, 222, 128, 0.05)'; }
    } else if (type === 'procesados') {
        state.maestros.procesados = data;
        if(infoDiv) { infoDiv.innerHTML = `<span style="color:#4ade80">✅ Maestro Excepciones OK</span>`; zone.style.borderColor = '#4ade80'; zone.style.background = 'rgba(74, 222, 128, 0.05)'; }
    } else if (type === 'analistas') {
        state.maestros.analistas = data;
        if(infoDiv) { infoDiv.innerHTML = `<span style="color:#4ade80">✅ Analistas SAP OK</span>`; zone.style.borderColor = '#4ade80'; zone.style.background = 'rgba(74, 222, 128, 0.05)'; }
    } else if (type === 'dus') {
        state.dus = data;
        if(infoDiv) { infoDiv.innerHTML = `<span style="color:#4ade80">✅ DUS cargado OK</span>`; zone.style.borderColor = '#4ade80'; zone.style.background = 'rgba(74, 222, 128, 0.05)'; }
    }
}


function readExcel(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array', cellDates: true });
                const sheetName = workbook.SheetNames[0];
                const json = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" });
                resolve(json);
            } catch (err) { reject(err); }
        };
        reader.onerror = (err) => reject(err);
        reader.readAsArrayBuffer(file);
    });
}

function checkReadyState() {
    const module = state.currentModule;
    const dusBtnEl = document.getElementById('run-dus-btn');

    if (module === 'dus') {
        // DUS solo necesita: SAP + Analistas.
        const ready = !!(state.sap && state.maestros.analistas);
        if (dusBtnEl) {
            dusBtnEl.disabled = !ready;
            dusBtnEl.classList.toggle('glow', ready);
        }
    } else {
        // Auditoría Comex necesita: SAP + al menos 1 BL
        const ready = !!(state.sap && state.bls.length > 0);
        runBtn.disabled = !ready;
        runBtn.classList.toggle('glow', ready);
    }
}

function showLoading(msg = "Procesando...") {
    const overlay = document.getElementById('loading-overlay');
    const details = document.getElementById('loading-details');
    overlay.style.display = 'flex';
    details.innerText = msg;
}

function hideLoading() {
    document.getElementById('loading-overlay').style.display = 'none';
}

function runAudit() {
    // Redirigir al motor DUS si estamos en ese módulo
    if (state.currentModule === 'dus') {
        if (!state.sap || !state.maestros.analistas) {
            alert('Para DUS necesitas cargar: SAP (export.XLSX) y Analista Facturación (SAP).');
            return;
        }
        runDUS();
        return;
    }
    if (!state.sap || state.bls.length === 0) return;
    
    showLoading("Cruzando datos y calculando KPIs...");
    
    // Timeout para permitir que el navegador dibuje el "Cargando"
    setTimeout(() => {
        try {
            const sapData = state.sap;
            const blData = state.bls;
            const asignacion = state.maestros.asignacion || [];
            const procesados = state.maestros.procesados || [];

            // P-01: TODAS las Maps deben tener guard `if (key)` antes de .set()
            // Sin esto, cleanPedido("") retorna "" y se crea una entrada fantasma
            // que matchea cualquier otro pedido que también sea "" → cruces incorrectos
            const blMap = new Map();
            blData.forEach(row => {
                const key = cleanPedido(row.Pedido);
                if (key) blMap.set(key, row);
            });

            const manualMap = new Map();
            procesados.forEach(row => {
                const key = cleanPedido(row['PEDIDO FLUJO']);
                // Buscar motivo en múltiples nombres de columna posibles
                const motivo = row.MOTIVO || row.Motivo || row.OBSERVACION || row.Observacion
                    || row.OBSERVACIONES || row.Observaciones
                    || row.NOTA || row.Nota || row.COMENTARIO || row.Comentario
                    || row['DESCRIPCION'] || row['Descripcion']
                    || "Validado sin motivo";
                if (key) manualMap.set(key, motivo);
            });

            const asignacionMap = new Map();
            const terrestreSet = new Set();
            if (state.maestros.asignacion) {
                state.maestros.asignacion.forEach(row => {
                    if (row.Cliente) {
                        const key = normalizeText(row.Cliente);
                        asignacionMap.set(key, row.Responsable || row.Analista);
                        if ((row.Responsable || row.Analista || "").toLowerCase().includes('terrestre')) {
                            terrestreSet.add(key);
                        }
                    }
                });
            }

            const analistaMap = new Map();
            const analistaPedidoMap = new Map();
            const sapUserMap = new Map(); // ID -> Nombre

            if (state.maestros.analistas) {
                state.maestros.analistas.forEach(row => {
                    const factura = row['N° Factura'] || row['Factura'] || row['Folio'] || row['Número Solicitante'] || row['Numero Solicitante'];
                    const pedido = row['N° Pedido'] || row['N Pedido'] || row['Pedido'] || row['Orden'];
                    const creador = row['Creado por'] || row['Analista'] || row['Nombre de usuario'] || row['Centro'];
                    const userId = row['Usuario'] || row['SAP User'] || row['Centro'];

                    // P-01: guard de llave vacía en Maps de analistas
                    const facKey = cleanPedido(factura);
                    const pedKey = cleanPedido(pedido);
                    if (facKey && creador) analistaMap.set(facKey, creador);
                    if (pedKey && creador) analistaPedidoMap.set(pedKey, creador);
                    
                    // Si encontramos un ID y un nombre, lo guardamos para traducir en DUS
                    if (userId && creador && String(userId).includes('020')) {
                        sapUserMap.set(String(userId).trim(), creador);
                    }
                });
            }
            
            // También podemos sacar mapeos de SAP general si hay IDs
            if (state.sap) {
                state.sap.forEach(row => {
                    const id = row['Usuario'] || row['ID'];
                    const nombre = row['Nombre de usuario'] || row['Creado por'];
                    if (id && nombre && String(id).includes('020')) {
                        sapUserMap.set(String(id).trim(), nombre);
                    }
                });
            }

            state.sapUserMap = sapUserMap; // Persistir para uso en procesarDUS

            const results = [];
            const processedOrders = new Set();

            // Pre-scan: pedidos con al menos una factura válida en el reporte SAP.
            // Permite priorizar la fila "facturado" sobre la fila "pendiente" del mismo pedido.
            const pedidosConFactura = new Set();
            sapData.forEach(row => {
                const ped = cleanPedido(row['NPedido'] || row['N°Pedido']);
                const fac = cleanValue(row['° Factura'] || row['N Factura'] || row['Folio Factura'] || row['Factura'] || row['Folio']);
                if (ped && fac && fac !== '0') pedidosConFactura.add(ped);
            });

            sapData.forEach(row => {
                const descMatch = (row['Descripcin'] || row['Descripción'] || "");
                const desc = String(descMatch);
                if (!desc.includes('@') || desc.includes('#')) return;

                const pedidoOriginal = row['NPedido'] || row['N°Pedido'];
                const pedidoClean = cleanPedido(pedidoOriginal);
                
                const folioFacturaRaw = row['N° Factura'] || row['N Factura'] || row['Folio Factura'] || row['Factura'] || row['Folio'] || row['Doc.facturación'] || row['Documento de facturación'] || row['Doc. fact.'] || row['Documento'];
                const folioKeyRaw = folioFacturaRaw ? cleanPedido(folioFacturaRaw) : "NO_FAC";
                
                // DE-DUPLICACIÓN INTELIGENTE: Ignorar combinaciones idénticas de Pedido+Factura
                // Esto permite que si un pedido tiene 2 facturas diferentes, ambas salgan en el reporte.
                const dedupeKey = `${pedidoClean}_${folioKeyRaw}`;
                if (processedOrders.has(dedupeKey)) return;
                processedOrders.add(dedupeKey);

                const cliente = row['Nombre Solicitante'];
                const clienteKey = normalizeText(cliente);
                
                const blMatch = blMatchSearch(pedidoClean, blMap);
                const motivoManual = manualMap.get(pedidoClean);

                const fechaFactura = parseExcelDate(row['Fecha Factura']);
                const folioKey = folioKeyRaw === "NO_FAC" ? null : folioKeyRaw;
                const fechaBL = blMatch ? parseExcelDate(blMatch.FechaCreacion) : null;

                const semaforo = calculateSemaforo(fechaFactura);
                
                // Lógica de Responsable
                let responsable = "Sin Asignar";
                const sapUserFact = folioKey ? analistaMap.get(folioKey) : null;
                const sapUserPed = analistaPedidoMap.get(pedidoClean);
                let fallbackAnalista = asignacionMap.get(clienteKey);
                
                if (!fallbackAnalista) {
                    // Búsqueda difusa si no hay match directo
                    for (const [key, val] of asignacionMap) {
                        if (clienteKey.includes(key) || key.includes(clienteKey)) {
                            fallbackAnalista = val;
                            break;
                        }
                    }
                }

                if (terrestreSet.has(clienteKey)) {
                    responsable = "🚚 Procesado Terrestre";
                } else if (sapUserFact) {
                    responsable = sapUserFact;
                } else if (sapUserPed) {
                    responsable = sapUserPed;
                } else if (fallbackAnalista) {
                    responsable = fallbackAnalista;
                }

                // Cálculo de Estatus y Demora
                let demora = (fechaFactura && !isNaN(fechaFactura)) ? Math.floor((state.fechaHoy - fechaFactura) / (1000 * 60 * 60 * 24)) : 0;
                if (demora > 10000) demora = 0; 
                
                const tGestion = (fechaBL && fechaFactura) ? Math.floor((fechaBL - fechaFactura) / (1000 * 60 * 60 * 24)) : null;

                let estatus = "❌ Pendiente";
                if (!folioFacturaRaw || String(folioFacturaRaw).trim() === "" || !fechaFactura) {
                    // Si este pedido tiene factura en otra fila del mismo reporte, omitir el duplicado pendiente
                    if (pedidosConFactura.has(pedidoClean)) return;
                    estatus = "📦 Pendiente Facturación";
                    demora = 0;
                } else if (semaforo === "💀 Rojo Crítico") {
                    estatus = "⚠️ Error Fecha";
                } else if (blMatch) {
                    estatus = "✅ Procesado";
                } else if (responsable.includes('Terrestre')) {
                    estatus = "🚚 Procesado Terrestre";
                } else if (motivoManual) {
                    estatus = "⚠️ Validado Manualmente";
                }

                results.push({
                    ESTATUS_FINAL: estatus,
                    DEMORA: demora,
                    T_GESTIÓN: tGestion,
                    PEDIDO: pedidoClean,
                    CLIENTE: cliente,
                    RESPONSABLE: responsable,
                    FECHA_FACTURA: fechaFactura && !isNaN(fechaFactura) ? fechaFactura : null,
                    FECHA_BL: fechaBL,
                    MOTIVO: estatus === '⚠️ Validado Manualmente' ? (motivoManual || '') : '',
                    BL: blMatch ? blMatch.BL : ""
                });
            });

            state.results = results;
            // M-05: solo procesar DUS si hay datos cargados (evita trabajo innecesario)
            state.dusResults = state.dus ? procesarDUS() : [];
            
            // Actualizar Filtros Dinámicos
            updateDynamicFilters(state.results, 'status-multi-select', 'ms-trigger', 'ms-label', renderResults, 'auditoria');
            updateDynamicFilters(state.dusResults, 'dus-multi-select', 'ms-dus-trigger', 'ms-dus-label', renderDUS, 'dus');
            
            const selectedAuditoria = Array.from(new Set(results.map(r => r.ESTATUS_FINAL)));
            const selectedDUS = Array.from(new Set(state.dusResults.map(r => r.ESTATUS_FINAL)));
            
            renderResults(selectedAuditoria);
            renderDUS(selectedDUS);
            
            updateStats();
            if (state.currentModule === 'auditoria') {
                updateChart(results);
            } else {
                updateChartForDUS(state.dusResults);
            }
            // Plotly: actualizar tendencia y heatmap con los datos frescos
            updateTrendChart();
            // KPI Panel: actualizar indicadores
            renderKPIPanel(state.currentModule === 'dus' ? state.dusResults : state.results);
            if(exportBtn) exportBtn.disabled = false;
            if(exportBtnDus) exportBtnDus.disabled = false;
            // Auto-guardar en memoria los registros procesados/legalizados
            saveToMemory('auditoria', state.results);
            saveToMemory('dus', state.dusResults);
            updateMemoryUI();
        } catch (err) {
            console.error(err);
            alert("Error en la auditoría: " + err.message);
        } finally {
            hideLoading();
        }
    }, 100);
}

/**
 * Motor independiente para Legalización DUS.
 * Solo requiere: SAP (export) + Analista Facturación + (opcional) Asignación Clientes + DUS.
 */
function runDUS() {
    showLoading("Procesando Legalización DUS...");
    setTimeout(() => {
        try {
            state.dusResults = procesarDUS();

            updateDynamicFilters(state.dusResults, 'dus-multi-select', 'ms-dus-trigger', 'ms-dus-label', renderDUS, 'dus');
            const selectedDUS = Array.from(new Set(state.dusResults.map(r => r.ESTATUS_FINAL)));
            renderDUS(selectedDUS);
            updateStats();
            updateChartForDUS(state.dusResults);
            renderKPIPanel(state.dusResults);
            if (exportBtnDus) exportBtnDus.disabled = false;
            // Auto-guardar en memoria los registros legalizados
            saveToMemory('dus', state.dusResults);
            updateMemoryUI();
        } catch (err) {
            console.error(err);
            alert("Error en DUS: " + err.message);
        } finally {
            hideLoading();
        }
    }, 100);
}

// =============================================================
// MOTOR DE MEMORIA PERSISTENTE (localStorage)
// Guarda pedidos procesados/legalizados entre sesiones.
// =============================================================

const MEMORY_KEYS = {
    auditoria: 'accomex_memory_auditoria',
    dus:       'accomex_memory_dus'
};

// Estados que califican para ser guardados en memoria
const MEMORY_STATUSES = {
    auditoria: ['✅ Procesado', '⚠️ Validado', '🚚 Procesado Terrestre'],
    dus:       ['✅ Legalizado']
};

let memoryViewActive = { auditoria: false, dus: false };

/** Carga la memoria del módulo indicado. Retorna un objeto { pedido: {...} } */
function loadMemory(module) {
    try {
        const mem = JSON.parse(localStorage.getItem(MEMORY_KEYS[module]) || '{}');
        // OBS-3: migración automática de memoria DUS pre-P-06
        // Las entradas viejas tenían PEDIDO = rawReferencia (ej: "REF-4500012345/4500012346")
        // Las nuevas usan PEDIDO = pedidoClean (ej: "4500012345").
        // Si encontramos llaves no-numéricas en la memoria DUS, re-keyeamos automáticamente.
        if (module === 'dus' && !mem._migrated_p06) {
            let changed = false;
            const newMem = {};
            Object.entries(mem).forEach(([key, val]) => {
                // Si la llave contiene letras o caracteres no numéricos, es formato viejo
                if (/[^0-9]/.test(key)) {
                    const digits = key.match(/\d{7,10}/);
                    const cleanKey = digits ? digits[0] : key;
                    // Solo migrar si no colisiona con una entrada ya existente
                    if (!mem[cleanKey] && !newMem[cleanKey]) {
                        val.PEDIDO_RAW = val.PEDIDO_RAW || val.PEDIDO || key;
                        val.PEDIDO = cleanKey;
                        newMem[cleanKey] = val;
                        changed = true;
                    } else {
                        newMem[key] = val; // mantener si colisiona
                    }
                } else {
                    newMem[key] = val;
                }
            });
            if (changed) {
                newMem._migrated_p06 = true;
                try { localStorage.setItem(MEMORY_KEYS[module], JSON.stringify(newMem)); }
                catch { /* silenciar si QuotaExceeded */ }
                // Eliminar flag interno antes de retornar (no es un pedido)
                delete newMem._migrated_p06;
                return newMem;
            }
            // Marcar como migrado aunque no haya cambios (evitar recorrer cada vez)
            mem._migrated_p06 = true;
            try { localStorage.setItem(MEMORY_KEYS[module], JSON.stringify(mem)); }
            catch { /* silenciar */ }
        }
        // Limpiar flag interno del objeto retornado (no contaminar conteos)
        delete mem._migrated_p06;
        return mem;
    } catch { return {}; }
}

/** Guarda/fusiona nuevos registros procesados en la memoria */
function saveToMemory(module, results) {
    const mem = loadMemory(module);
    const validStatuses = MEMORY_STATUSES[module];
    let added = 0;

    results.forEach(r => {
        const status = r.ESTATUS_FINAL || '';
        // A-03: comparación directa por includes() — evita regex de emojis multi-codepoint
        // que puede no funcionar correctamente en todos los navegadores/SO
        const isFinal = validStatuses.some(s => status.includes(s));
        if (!isFinal) return;

        const key = String(r.PEDIDO || '').trim();
        if (!key) return;

        mem[key] = {
            ...r,
            FECHA_FACTURA: (r.FECHA_FACTURA instanceof Date)
                ? r.FECHA_FACTURA.toISOString()
                : (r.FECHA_FACTURA || ''),
            FECHA_BL: (r.FECHA_BL instanceof Date)
                ? r.FECHA_BL.toISOString()
                : (r.FECHA_BL || ''),
            savedAt: new Date().toISOString().slice(0, 10)
        };
        added++;
    });

    // A-06: capturar QuotaExceededError — localStorage tiene ~5-10 MB de límite
    try {
        // OBS-3: preservar flag de migración para que loadMemory no recorra cada vez
        if (module === 'dus') mem._migrated_p06 = true;
        localStorage.setItem(MEMORY_KEYS[module], JSON.stringify(mem));
    } catch (e) {
        if (e.name === 'QuotaExceededError' || e.code === 22) {
            console.warn('⚠️ Memoria llena (localStorage). Exporta el historial y límpialo.');
            const banner = document.getElementById('memory-warning-banner');
            if (banner) { banner.style.display = 'flex'; }
            else { alert('⚠️ La memoria histórica está llena. Exporta y limpia el historial para continuar guardando.'); }
        }
    }
    // OBS-4: limpiar el flag interno antes de contar (no es un pedido real)
    delete mem._migrated_p06;
    return Object.keys(mem).length;
}

/** Actualiza el contador del botón Memoria y muestra/oculta el botón de limpiar */
function updateMemoryUI() {
    ['auditoria', 'dus'].forEach(mod => {
        const mem = loadMemory(mod);
        const count = Object.keys(mem).length;
        const countEl = document.getElementById(`memory-count-${mod}`);
        const clearBtn = document.getElementById(mod === 'auditoria' ? 'memory-clear-btn' : 'memory-clear-btn-dus');
        if (countEl) countEl.textContent = count;
        if (clearBtn) clearBtn.style.display = count > 0 ? 'flex' : 'none';
    });
}

/** Limpia la memoria del módulo actual con confirmación */
/** Exporta la memoria del módulo a un archivo Excel (también sirve como respaldo) */
function exportMemory(module) {
    const mem = loadMemory(module);
    const records = Object.values(mem);
    if (records.length === 0) {
        alert('No hay registros en memoria para exportar.');
        return;
    }

    const isAuditoria = module === 'auditoria';
    const sheetName = isAuditoria ? 'Memoria Auditoría' : 'Memoria DUS';

    const wsData = isAuditoria
        ? records.map(r => ({
            'N° PEDIDO':      r.PEDIDO || '',
            'ESTATUS FINAL':  r.ESTATUS_FINAL || '',
            'CLIENTE':        r.CLIENTE || '',
            'RESPONSABLE':    r.RESPONSABLE || '',
            'FECHA FAC.':     r.FECHA_FACTURA ? new Date(r.FECHA_FACTURA) : '',
            'FECHA BL':       r.FECHA_BL ? new Date(r.FECHA_BL) : '',
            'MOTIVO':         r.MOTIVO || '',
            'GUARDADO EL':    r.savedAt || ''
          }))
        : records.map(r => ({
            'N° PEDIDO':      r.PEDIDO || '',
            'ESTATUS FINAL':  r.ESTATUS_FINAL || '',
            'CONSIGNATARIO':  r.CONSIGNATARIO || '',
            'RESPONSABLE':    r.RESPONSABLE || '',
            'FECHA FACTURA':  r.FECHA_FACTURA ? new Date(r.FECHA_FACTURA) : '',
            'FACTURA SAP':    r.FACTURA_SAP || '',
            'GUARDADO EL':    r.savedAt || ''
          }));

    const ws = XLSX.utils.json_to_sheet(wsData, { cellDates: true });

    // Formato de fecha a columnas con "FECHA"
    if (ws['!ref']) {
        const range = XLSX.utils.decode_range(ws['!ref']);
        for (let C = range.s.c; C <= range.e.c; C++) {
            const hCell = ws[XLSX.utils.encode_cell({c: C, r: 0})];
            if (hCell && /fecha/i.test(String(hCell.v || ''))) {
                for (let R = 1; R <= range.e.r; R++) {
                    const ref = XLSX.utils.encode_cell({c: C, r: R});
                    if (ws[ref] && ws[ref].v instanceof Date) {
                        ws[ref].t = 'd';
                        ws[ref].z = 'DD-MM-YYYY';
                    }
                }
            }
        }
        // Estilo de cabecera
        for (let C = range.s.c; C <= range.e.c; C++) {
            const ref = XLSX.utils.encode_cell({c: C, r: 0});
            if (ws[ref]) {
                ws[ref].s = {
                    fill: { fgColor: { rgb: '1E293B' } },
                    font: { color: { rgb: 'FFFFFF' }, bold: true }
                };
            }
        }
    }

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    const fecha = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `RESPALDO_${sheetName.replace(' ', '_')}_${fecha}.xlsx`);
}

function clearMemory(module) {
    const mem = loadMemory(module);
    const count = Object.keys(mem).length;
    if (count === 0) { alert('La memoria ya está vacía.'); return; }
    if (!confirm(`¿Eliminar los ${count} registros de memoria de ${module === 'dus' ? 'Legalización DUS' : 'Auditoría Comex'}?\n\nEsta acción no se puede deshacer.`)) return;
    localStorage.removeItem(MEMORY_KEYS[module]);
    memoryViewActive[module] = false;
    updateMemoryUI();
    // Re-renderizar sin memoria
    if (module === 'auditoria') {
        const btn = document.getElementById('memory-toggle-btn');
        if (btn) btn.classList.remove('active');
        renderResults(getSelectedValues('status-multi-select'));
    } else {
        const btn = document.getElementById('memory-toggle-btn-dus');
        if (btn) btn.classList.remove('active');
        renderDUS(getSelectedValues('dus-multi-select'));
    }
}

/** Toggle: muestra u oculta los registros de memoria junto a los actuales */
function toggleMemoryView(module) {
    memoryViewActive[module] = !memoryViewActive[module];
    const btnId = module === 'auditoria' ? 'memory-toggle-btn' : 'memory-toggle-btn-dus';
    const btn = document.getElementById(btnId);
    if (btn) btn.classList.toggle('active', memoryViewActive[module]);
    // Re-renderizar con o sin memoria
    if (module === 'auditoria') {
        renderResults(getSelectedValues('status-multi-select'));
    } else {
        renderDUS(getSelectedValues('dus-multi-select'));
    }
}

/** Devuelve registros de memoria que NO están en los resultados actuales */
function getMemoryOnlyRecords(module, currentResults) {
    if (!memoryViewActive[module]) return [];
    const mem = loadMemory(module);
    const currentKeys = new Set(currentResults.map(r => String(r.PEDIDO || '').trim()));
    return Object.values(mem)
        .filter(r => !currentKeys.has(String(r.PEDIDO || '').trim()))
        .map(r => ({
            ...r,
            _fromMemory: true,
            FECHA_FACTURA: r.FECHA_FACTURA ? new Date(r.FECHA_FACTURA) : null,
            FECHA_BL: r.FECHA_BL ? new Date(r.FECHA_BL) : null
        }));
}

// C-05: Función de escape HTML para sanitizar datos de Excel antes de inyectarlos en el DOM.
// Previene XSS si una celda del Excel contiene HTML o scripts.
function escHtml(str) {
    return String(str == null ? '' : str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function normalizeText(text) {
    if (!text) return "";
    return text.toString().toLowerCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "") // Quita acentos
        .replace(/[^a-z0-9]/g, "");     // Quita símbolos y espacios (trim innecesario: ya no quedan espacios)
}

/**
 * Limpia pedidos y facturas: elimina ceros a la izquierda y espacios.
 */
function cleanValue(val) {
    if (val === null || val === undefined) return "";
    return String(val).trim().replace(/^0+/, '').split('.')[0];
}

// Búsqueda robusta en el mapa de BLs
function blMatchSearch(pedidoClean, blMap) {
    return blMap.get(pedidoClean);
}

// M-02: función canónica consolidada para convertir fechas Excel → JS Date.
// Reemplaza tanto parseExcelDate como excelDateToJS (que eran duplicadas con pequeñas diferencias).
function excelSerialToDate(val) {
    if (!val) return null;
    if (val instanceof Date) return val;
    if (typeof val === 'number') return new Date(Math.round((val - 25569) * 86400 * 1000));
    const d = new Date(val);
    return isNaN(d.getTime()) ? null : d;
}
// Alias de compatibilidad para que el código existente siga funcionando sin cambios masivos
const parseExcelDate = excelSerialToDate;

function calculateSemaforo(fecha) {
    if (!fecha) return "⚪ N/A";
    // P-05: usar state.fechaHoy en vez de new Date() para consistencia con DEMORA.
    const hoy = new Date(state.fechaHoy);
    hoy.setHours(0, 0, 0, 0);
    // FIX: normalizar fecha de factura a medianoche antes de comparar.
    // Sin esto, una factura de hoy con hora > 00:00 se detectaba como "futura".
    // Error Fecha solo debe dispararse para MAÑANA en adelante (estrictamente futuro).
    const fechaNorm = new Date(fecha);
    fechaNorm.setHours(0, 0, 0, 0);
    if (fechaNorm > hoy) return "💀 Rojo Crítico";
    const dias = Math.floor((hoy - fechaNorm) / (1000 * 60 * 60 * 24));
    if (dias <= 3) return "🟢 Verde";
    if (dias <= 7) return "🟠 Naranja";
    return "🔴 Rojo";
}

function renderResults(selectedStatuses = []) {
    const tbody = document.querySelector('#results-table tbody');
    const emptyState = document.getElementById('table-empty-state');
    let data = state.results;

    // Agregar registros de memoria que no están en los resultados actuales
    const memOnly = getMemoryOnlyRecords('auditoria', data);
    const allData = [...data, ...memOnly];

    if (allData.length > 0) emptyState.style.display = 'none';

    let filtered = allData;
    if (selectedStatuses.length > 0) {
        const totalOptions = document.querySelectorAll('#status-multi-select .ms-option').length;
        if (selectedStatuses.length < totalOptions) {
            // Filtrar por ESTADO GENERAL (agrupado), no por ESTATUS_FINAL individual
            filtered = filtered.filter(res =>
                selectedStatuses.includes(getEstadoGeneral(res.ESTATUS_FINAL).label)
            );
        }
    }

    // Aplicar búsqueda de texto
    filtered = applySearch(filtered, state.searchQuery, false);
    // Aplicar filtro de grupo
    const gFilter = getGrupoAnalysts();
    if (gFilter) filtered = filtered.filter(r => gFilter.has(String(r.RESPONSABLE || '').toUpperCase().trim()));
    // Aplicar orden de columna
    const ss = state.sortState;
    filtered = sortResults(filtered, ss.col, ss.dir);

    tbody.innerHTML = filtered.map(res => {
        const eg = getEstadoGeneral(res.ESTATUS_FINAL);
        // OBS-6: escHtml para fechas string (evita XSS si Excel envía texto en vez de fecha)
        const fFac = (res.FECHA_FACTURA instanceof Date) ? res.FECHA_FACTURA.toLocaleDateString() : escHtml(res.FECHA_FACTURA || '-');
        const fBL  = (res.FECHA_BL instanceof Date) ? res.FECHA_BL.toLocaleDateString() : escHtml(res.FECHA_BL || '-');
        // C-05: escHtml() sanitiza datos de Excel (evita XSS si una celda contiene HTML)
        const ePed  = escHtml(res.PEDIDO);
        const eCli  = escHtml(res.CLIENTE);
        const eResp = escHtml(res.RESPONSABLE);
        const eMot  = escHtml(res.MOTIVO);
        const eEst  = escHtml(res.ESTATUS_FINAL);
        return `
        <tr class="${res._fromMemory ? 'memory-row' : ''}">
            <td><span class="badge ${eg.badgeClass}">${eg.label}</span></td>
            <td title="${eEst}"><span class="badge ${getBadgeClass(res.ESTATUS_FINAL)}">${eEst}</span></td>
            <td>${res.DEMORA ?? '-'}</td>
            <td>${res.T_GESTIÓN !== null && res.T_GESTIÓN !== undefined ? res.T_GESTIÓN : '-'}</td>
            <td title="${ePed}">${ePed}</td>
            <td title="${eCli}">${eCli || '-'}</td>
            <td title="${eResp}">${eResp}</td>
            <td style="white-space:nowrap">${fFac}</td>
            <td style="white-space:nowrap">${fBL}</td>
            <td title="${eMot}">${eMot || '-'}</td>
        </tr>`;
    }).join('');
    document.getElementById('table-empty-state').style.display = filtered.length ? 'none' : 'flex';
    // Actualizar contador de búsqueda
    const countEl = document.getElementById('search-count');
    if (countEl) countEl.textContent = state.searchQuery ? `${filtered.length} resultado${filtered.length !== 1 ? 's' : ''}` : '';
}

/** Clasifica ESTADO_GENERAL a partir del ESTATUS_FINAL */
function getEstadoGeneral(estatusFinal) {
    const s = String(estatusFinal || '');
    // M-03: Error Fecha PRIMERO — antes de verificar '⚠' para no clasificarlo como Procesado
    // '⚠️ Error Fecha' contiene '⚠' pero NO debe contarse como Procesado
    if (s.includes('Error') || s.includes('💀')) {
        return { label: '❌ Error Fecha', badgeClass: 'badge-critical' };
    }
    // Procesados: ✅ Procesado, 🚚 Terrestre, ⚠️ Validado Manualmente
    if (s.includes('✅') || s.includes('🚚') || s.includes('Validado')) {
        return { label: '✅ Procesado', badgeClass: 'badge-green' };
    }
    return { label: '❌ Pendiente', badgeClass: 'badge-red' };
}

// ── Utilidades de búsqueda y orden ──────────────────────────

/** Filtra un array de resultados por texto libre */
function applySearch(arr, query, isDUS = false) {
    if (!query || !query.trim()) return arr;
    const q = query.toLowerCase().trim();
    return arr.filter(r => {
        const fields = isDUS
            ? [r.PEDIDO, r.PEDIDO_RAW, r.CONSIGNATARIO, r.RESPONSABLE, r.ESTATUS_FINAL, r.ESTADO_DUS_RAW]
            : [r.PEDIDO, r.CLIENTE, r.RESPONSABLE, r.ESTATUS_FINAL, r.MOTIVO];
        return fields.some(f => f && String(f).toLowerCase().includes(q));
    });
}

/** Ordena un array de resultados por columna y dirección */
function sortResults(arr, col, dir) {
    if (!col) return arr;
    return [...arr].sort((a, b) => {
        let av, bv;
        switch (col) {
            case 'estadoGeneral': av = (a._isDUS ? getDUSEstadoGeneral(a.ESTATUS_FINAL).label : getEstadoGeneral(a.ESTATUS_FINAL).label); bv = (b._isDUS ? getDUSEstadoGeneral(b.ESTATUS_FINAL).label : getEstadoGeneral(b.ESTATUS_FINAL).label); break;
            case 'detalle':       av = String(a.ESTATUS_FINAL); bv = String(b.ESTATUS_FINAL); break;
            case 'demora':        av = a.DEMORA ?? -Infinity; bv = b.DEMORA ?? -Infinity; break;
            case 'tGestion':      av = a.T_GESTIÓN ?? -Infinity; bv = b.T_GESTIÓN ?? -Infinity; break;
            case 'pedido':        av = String(a.PEDIDO || ''); bv = String(b.PEDIDO || ''); break;
            case 'cliente':       av = String(a.CLIENTE || ''); bv = String(b.CLIENTE || ''); break;
            case 'consignatario': av = String(a.CONSIGNATARIO || ''); bv = String(b.CONSIGNATARIO || ''); break;
            case 'responsable':   av = String(a.RESPONSABLE || ''); bv = String(b.RESPONSABLE || ''); break;
            case 'fechaFac':      av = a.FECHA_FACTURA || a.FECHA_FACTURA_DATE || 0; bv = b.FECHA_FACTURA || b.FECHA_FACTURA_DATE || 0; break;
            case 'fechaBL':       av = a.FECHA_BL || 0; bv = b.FECHA_BL || 0; break;
            default:              return 0;
        }
        if (av < bv) return dir === 'asc' ? -1 : 1;
        if (av > bv) return dir === 'asc' ?  1 : -1;
        return 0;
    });
}

/** Handler del buscador en tiempo real */
function onSearchInput(value) {
    state.searchQuery = value;
    const clearBtn = document.getElementById('search-clear-btn');
    if (clearBtn) clearBtn.style.display = value ? 'flex' : 'none';
    const input = document.getElementById('table-search');
    if (input && input.value !== value) input.value = value;
    if (state.currentModule === 'auditoria') {
        renderResults(getSelectedValues('status-multi-select'));
        updateExportButtonText(getSelectedValues('status-multi-select'));
    } else {
        renderDUS(getSelectedValues('dus-multi-select'));
        updateExportButtonText(getSelectedValues('dus-multi-select'));
    }
}

function getBadgeClass(estatus) {
    if (estatus.includes('✅')) return 'badge-green';
    if (estatus.includes('Validado')) return 'badge-green'; // Manual validation also green
    if (estatus.includes('🚚')) return 'badge-blue';
    if (estatus.includes('📦')) return 'badge-yellow'; // Facturación
    if (estatus.includes('IVV')) return 'badge-orange'; // IVV specific orange
    if (estatus.includes('❌')) return 'badge-red';
    if (estatus.includes('Error')) return 'badge-critical';
    return 'badge-orange';
}

function updateStats() {
    const isAuditoria = state.currentModule === 'auditoria';
    let activeData = isAuditoria ? state.results : state.dusResults;

    // Aplicar filtro de grupo a los datos activos
    const gFilter = getGrupoAnalysts();
    if (gFilter && activeData) {
        activeData = activeData.filter(r => gFilter.has(String(r.RESPONSABLE || '').toUpperCase().trim()));
    }

    // Total procesados: suma memoria del módulo activo (fuente de verdad histórica)
    const mem = loadMemory(isAuditoria ? 'auditoria' : 'dus');
    const memTotal = Object.keys(mem).length;

    let processed = 0, pending = 0;
    const data = activeData || [];
    const total = data.length;

    if (isAuditoria) {
        // getEstadoGeneral ya agrupa terrestres y validados en '✅ Procesado'
        // NO restar por separado — eso generaba doble descuento y pending < real
        processed = data.filter(r => getEstadoGeneral(r.ESTATUS_FINAL).label === '✅ Procesado').length;
        pending   = data.filter(r => getEstadoGeneral(r.ESTATUS_FINAL).label === '❌ Pendiente').length;

        document.getElementById('stat-pending-label').innerText = 'Pendientes Comex';
        document.getElementById('stat-pending-icon').setAttribute('data-lucide', 'clock');
        document.getElementById('stat-pending-desc').innerText = 'Sin BL o factura';
    } else {
        processed = data.filter(r => getDUSEstadoGeneral(r.ESTATUS_FINAL).label === '✅ Legalizado').length;
        pending   = data.filter(r => !['✅ Legalizado', '❌ Anulado'].includes(getDUSEstadoGeneral(r.ESTATUS_FINAL).label)).length;

        document.getElementById('stat-pending-label').innerText = 'Pendientes DUS';
        document.getElementById('stat-pending-icon').setAttribute('data-lucide', 'file-warning');
        document.getElementById('stat-pending-desc').innerText = 'DUS no legalizados';
    }

    // KPI 1: Total Procesados = memoria acumulada (histórico)
    const statTotalEl = document.getElementById('stat-total');
    if (statTotalEl) {
        statTotalEl.innerText = memTotal;
        const subEl = statTotalEl.closest('.stat-card')?.querySelector('.stat-sub');
        if (subEl) subEl.innerText = processed > 0 ? `+${processed} esta sesión` : 'Desde historial';
    }

    // KPI 2: Pendientes (valor filtrado por grupo activo)
    document.getElementById('stat-pending').innerText = pending > 0 ? pending : (total > 0 ? 0 : '-');

    // KPI 2 sub: Desglose por grupo — siempre desde datos COMPLETOS para mostrar ambos grupos
    const gruposEl = document.getElementById('stat-pending-grupos');
    if (gruposEl) {
        const fullData = (isAuditoria ? state.results : state.dusResults) || [];
        // Misma lógica exacta que el número grande
        const isPendingFn = isAuditoria
            ? r => getEstadoGeneral(r.ESTATUS_FINAL).label === '❌ Pendiente'
            : r => !['✅ Legalizado', '❌ Anulado'].includes(getDUSEstadoGeneral(r.ESTATUS_FINAL).label);

        // Helper: códigos de un grupo específico
        const getGroupCodes = (grupo) => {
            const dir = loadTeamDirectory();
            const s = new Set();
            Object.entries(dir).forEach(([code, entry]) => {
                const g = typeof entry === 'string' ? null : (entry.grupo || null);
                if (g === grupo) s.add(code.toUpperCase().trim());
            });
            return s;
        };

        const cCodes = getGroupCodes('congelado');
        const fCodes = getGroupCodes('fresco');
        const pendingC = cCodes.size > 0 ? fullData.filter(r => isPendingFn(r) && cCodes.has(String(r.RESPONSABLE || '').toUpperCase().trim())).length : null;
        const pendingF = fCodes.size > 0 ? fullData.filter(r => isPendingFn(r) && fCodes.has(String(r.RESPONSABLE || '').toUpperCase().trim())).length : null;

        if (pendingC !== null || pendingF !== null) {
            gruposEl.innerHTML =
                (pendingC !== null ? `<span class="grupo-pill congelado" title="Pendientes Congelado">❄️ ${pendingC}</span>` : '') +
                (pendingF !== null ? `<span class="grupo-pill fresco" title="Pendientes Fresco">🌿 ${pendingF}</span>` : '');
        } else {
            gruposEl.innerHTML = '';
        }
    }

    // KPI 3: % Cumplimiento
    const pct = total > 0 ? ((processed / total) * 100).toFixed(1) + '%' : '-%';
    const effLabelEl = document.getElementById('stat-efficiency-label');
    const effValEl   = document.getElementById('stat-efficiency');
    const effDescEl  = document.getElementById('stat-efficiency-desc');
    if (effLabelEl) effLabelEl.innerText = isAuditoria ? '% Cumplimiento' : '% Legalización';
    if (effValEl)   effValEl.innerText   = pct;
    if (effDescEl)  effDescEl.innerText  = total > 0 ? `${processed} de ${total} ${isAuditoria ? 'procesados' : 'legalizados'}` : 'de pedidos procesados';

    // Color del KPI 3 según % (verde/amarillo/rojo)
    const effCard = effValEl?.closest('.stat-card');
    if (effCard && total > 0) {
        const pctNum = processed / total;
        effCard.style.borderColor = pctNum >= 0.8 ? 'rgba(74,222,128,0.3)'
            : pctNum >= 0.5 ? 'rgba(251,191,36,0.3)'
            : 'rgba(248,113,113,0.3)';
    }

    if (window.lucide) lucide.createIcons();
}


// =============================================================
// PLOTLY.JS — GRÁFICOS INTERACTIVOS
// =============================================================

/** Layout base Plotly: fondo transparente, texto claro, sin ejes blancos */
const PLOTLY_DARK_LAYOUT = {
    paper_bgcolor: 'rgba(0,0,0,0)',
    plot_bgcolor: 'rgba(0,0,0,0)',
    font: { family: 'Inter, sans-serif', color: 'rgba(255,255,255,0.75)', size: 12 },
    margin: { t: 30, r: 20, b: 40, l: 50 }
};
const PLOTLY_CONFIG = {
    responsive: true,
    displaylogo: false,
    modeBarButtonsToRemove: ['lasso2d', 'select2d', 'autoScale2d']
};

function initChart() {
    // Plotly se inicializa en el primer render — no necesita canvas
}

/** Gauge semicircular de cumplimiento (reemplaza doughnut) */
function updateChart(results) {
    const el = document.getElementById('gaugeChart');
    if (!el) return;
    el.style.display = '';
    const wf = document.getElementById('waterfallChart');
    if (wf) wf.style.display = 'none';

    if (!results || results.length === 0) {
        Plotly.purge(el);
        return;
    }

    const total = results.length;
    const proc = results.filter(r => r.ESTATUS_FINAL.includes('✅') || r.ESTATUS_FINAL.includes('🚚')).length;
    const pct = Math.round((proc / total) * 100);

    const data = [{
        type: 'indicator',
        mode: 'gauge+number+delta',
        value: pct,
        number: { suffix: '%', font: { size: 48, color: '#e2e8f0' } },
        gauge: {
            axis: { range: [0, 100], tickwidth: 1, tickcolor: 'rgba(255,255,255,0.2)', dtick: 25 },
            bar: { color: pct >= 80 ? '#10b981' : pct >= 50 ? '#f59e0b' : '#ef4444', thickness: 0.3 },
            bgcolor: 'rgba(255,255,255,0.05)',
            borderwidth: 0,
            steps: [
                { range: [0, 40], color: 'rgba(239,68,68,0.12)' },
                { range: [40, 70], color: 'rgba(245,158,11,0.12)' },
                { range: [70, 100], color: 'rgba(16,185,129,0.12)' }
            ],
            threshold: {
                line: { color: '#818cf8', width: 3 },
                thickness: 0.8,
                value: pct
            }
        }
    }];

    const layout = {
        ...PLOTLY_DARK_LAYOUT,
        height: 280,
        margin: { t: 20, r: 30, b: 10, l: 30 },
        annotations: [{
            text: `<b>${proc}</b> de <b>${total}</b> procesados`,
            x: 0.5, y: -0.05,
            showarrow: false,
            font: { size: 13, color: 'rgba(255,255,255,0.5)' }
        }]
    };

    Plotly.react(el, data, layout, PLOTLY_CONFIG);
    updateTrendChart();
    updateHeatmapChart(results);
    updateTreemapChart(results);
    renderKPIPanel(results);
}

/** Waterfall DUS — pipeline visual del flujo de legalización */
function updateChartForDUS(results) {
    const wf = document.getElementById('waterfallChart');
    const gauge = document.getElementById('gaugeChart');
    if (!wf) return;
    wf.style.display = '';
    if (gauge) gauge.style.display = 'none';

    if (!results || results.length === 0) {
        Plotly.purge(wf);
        return;
    }

    // Agrupar por estado
    const counts = {};
    results.forEach(r => {
        const st = r.ESTATUS_FINAL || 'Desconocido';
        counts[st] = (counts[st] || 0) + 1;
    });

    // Ordenar: legalizados primero, luego por cantidad descendente
    const sorted = Object.entries(counts).sort((a, b) => {
        if (a[0].includes('✅')) return -1;
        if (b[0].includes('✅')) return 1;
        return b[1] - a[1];
    });

    const labels = sorted.map(([k]) => k.length > 25 ? k.slice(0, 22) + '…' : k);
    const values = sorted.map(([, v]) => v);
    const colors = sorted.map(([k]) => {
        if (k.includes('✅')) return '#10b981';
        if (k.includes('IVV') || k.includes('Sin Zarpe')) return '#f59e0b';
        if (k.includes('Anulad')) return '#64748b';
        return '#ef4444';
    });

    const data = [{
        type: 'bar',
        x: labels,
        y: values,
        marker: {
            color: colors,
            line: { color: 'rgba(255,255,255,0.1)', width: 1 }
        },
        text: values.map(v => String(v)),
        textposition: 'outside',
        textfont: { color: 'rgba(255,255,255,0.7)', size: 12 },
        hovertemplate: '<b>%{x}</b><br>Pedidos: %{y}<extra></extra>'
    }];

    const layout = {
        ...PLOTLY_DARK_LAYOUT,
        height: 300,
        xaxis: {
            tickfont: { size: 10, color: 'rgba(255,255,255,0.6)' },
            tickangle: -20,
            gridcolor: 'rgba(255,255,255,0.05)'
        },
        yaxis: {
            gridcolor: 'rgba(255,255,255,0.06)',
            tickfont: { color: 'rgba(255,255,255,0.5)' }
        },
        showlegend: false,
        bargap: 0.3
    };

    Plotly.react(wf, data, layout, PLOTLY_CONFIG);
    // Ocultar gráficos exclusivos de Auditoría cuando estamos en DUS
    const treemapW = document.getElementById('treemap-chart-widget');
    if (treemapW) treemapW.style.display = 'none';
}

/** Tendencia mensual — línea de procesados vs pendientes desde memoria */
function updateTrendChart() {
    const el = document.getElementById('trendChart');
    const emptyEl = document.getElementById('trend-empty');
    if (!el) return;

    // Combinar datos de sesión actual + memoria
    const mem = loadMemory(state.currentModule === 'dus' ? 'dus' : 'auditoria');
    const currentData = state.currentModule === 'dus' ? state.dusResults : state.results;

    // Construir mapa mes → {procesados, pendientes}
    const monthly = {};

    // Desde memoria
    Object.entries(mem).forEach(([key, val]) => {
        if (key.startsWith('_')) return;
        const fecha = val.FECHA_GUARDADO || val.fecha;
        if (!fecha) return;
        const mes = String(fecha).slice(0, 7); // YYYY-MM
        if (!monthly[mes]) monthly[mes] = { proc: 0, pend: 0 };
        const est = val.ESTATUS_FINAL || '';
        if (est.includes('✅') || est.includes('🚚') || est.includes('Legalizado')) {
            monthly[mes].proc++;
        } else {
            monthly[mes].pend++;
        }
    });

    // Desde sesión actual
    const hoy = new Date().toISOString().slice(0, 7);
    if (currentData && currentData.length > 0) {
        if (!monthly[hoy]) monthly[hoy] = { proc: 0, pend: 0 };
        currentData.forEach(r => {
            const est = r.ESTATUS_FINAL || '';
            if (est.includes('✅') || est.includes('🚚') || est.includes('Legalizado')) {
                monthly[hoy].proc++;
            } else {
                monthly[hoy].pend++;
            }
        });
    }

    const meses = Object.keys(monthly).sort();
    if (meses.length === 0) {
        Plotly.purge(el);
        if (emptyEl) emptyEl.style.display = 'flex';
        return;
    }
    if (emptyEl) emptyEl.style.display = 'none';

    const procData = meses.map(m => monthly[m].proc);
    const pendData = meses.map(m => monthly[m].pend);
    const labels = meses.map(m => {
        const [y, mo] = m.split('-');
        return new Date(y, mo - 1).toLocaleDateString('es-CL', { month: 'short', year: '2-digit' });
    });

    const data = [
        {
            type: 'scatter', mode: 'lines+markers',
            name: 'Procesados',
            x: labels, y: procData,
            line: { color: '#10b981', width: 3, shape: 'spline' },
            marker: { size: 8, color: '#10b981' },
            fill: 'tozeroy',
            fillcolor: 'rgba(16,185,129,0.08)'
        },
        {
            type: 'scatter', mode: 'lines+markers',
            name: 'Pendientes',
            x: labels, y: pendData,
            line: { color: '#ef4444', width: 3, shape: 'spline', dash: 'dot' },
            marker: { size: 8, color: '#ef4444', symbol: 'diamond' }
        }
    ];

    const layout = {
        ...PLOTLY_DARK_LAYOUT,
        height: 280,
        showlegend: true,
        legend: { x: 0, y: 1.15, orientation: 'h', font: { size: 11 } },
        xaxis: { gridcolor: 'rgba(255,255,255,0.06)', tickfont: { size: 11 } },
        yaxis: { gridcolor: 'rgba(255,255,255,0.06)', tickfont: { size: 11 }, title: 'Pedidos' }
    };

    Plotly.react(el, data, layout, PLOTLY_CONFIG);
}

/** Heatmap de demora promedio: analista (Y) × semáforo (X) */
function updateHeatmapChart(results) {
    const el = document.getElementById('heatmapChart');
    const emptyEl = document.getElementById('heatmap-empty');
    if (!el) return;

    if (!results || results.length === 0) {
        Plotly.purge(el);
        if (emptyEl) emptyEl.style.display = 'flex';
        return;
    }
    if (emptyEl) emptyEl.style.display = 'none';

    // Agrupar por responsable × rango de demora
    const ranges = ['0-3 días', '4-7 días', '8-14 días', '15+ días'];
    const analystData = {};

    results.forEach(r => {
        const resp = r.RESPONSABLE || 'Sin Asignar';
        const demora = parseInt(r.DEMORA) || 0;
        if (!analystData[resp]) analystData[resp] = [0, 0, 0, 0];
        if (demora <= 3)      analystData[resp][0]++;
        else if (demora <= 7) analystData[resp][1]++;
        else if (demora <= 14) analystData[resp][2]++;
        else                  analystData[resp][3]++;
    });

    // Solo mostrar top 15 analistas con más pedidos
    const analysts = Object.entries(analystData)
        .sort((a, b) => b[1].reduce((s, v) => s + v, 0) - a[1].reduce((s, v) => s + v, 0))
        .slice(0, 15);

    if (analysts.length === 0) {
        Plotly.purge(el);
        if (emptyEl) emptyEl.style.display = 'flex';
        return;
    }

    const yLabels = analysts.map(([name]) => {
        const resolved = typeof resolveAnalystName === 'function' ? resolveAnalystName(name) : name;
        return resolved.length > 18 ? resolved.slice(0, 15) + '…' : resolved;
    });
    const zData = analysts.map(([, counts]) => counts);

    const data = [{
        type: 'heatmap',
        x: ranges,
        y: yLabels,
        z: zData,
        colorscale: [
            [0, 'rgba(16,185,129,0.15)'],
            [0.33, 'rgba(245,158,11,0.4)'],
            [0.66, 'rgba(249,115,22,0.6)'],
            [1, 'rgba(239,68,68,0.85)']
        ],
        hovertemplate: '<b>%{y}</b><br>%{x}: %{z} pedidos<extra></extra>',
        showscale: true,
        colorbar: {
            title: 'Pedidos',
            titlefont: { size: 11, color: 'rgba(255,255,255,0.6)' },
            tickfont: { color: 'rgba(255,255,255,0.5)' },
            len: 0.8
        }
    }];

    const h = Math.max(260, analysts.length * 32 + 80);
    const layout = {
        ...PLOTLY_DARK_LAYOUT,
        height: h,
        xaxis: { tickfont: { size: 11 }, side: 'top' },
        yaxis: { tickfont: { size: 10 }, autorange: 'reversed' },
        margin: { t: 50, r: 100, b: 20, l: 120 }
    };

    Plotly.react(el, data, layout, PLOTLY_CONFIG);
}

/** Treemap de distribución por cliente (solo Auditoría) */
function updateTreemapChart(results) {
    const el = document.getElementById('treemapChart');
    const emptyEl = document.getElementById('treemap-empty');
    const widget = document.getElementById('treemap-chart-widget');
    if (!el) return;

    // Solo visible en Auditoría
    if (state.currentModule !== 'auditoria') {
        if (widget) widget.style.display = 'none';
        return;
    }
    if (widget) widget.style.display = '';

    if (!results || results.length === 0) {
        Plotly.purge(el);
        if (emptyEl) emptyEl.style.display = 'flex';
        return;
    }
    if (emptyEl) emptyEl.style.display = 'none';

    // Agrupar por cliente
    const clientMap = {};
    results.forEach(r => {
        const cliente = r.CLIENTE || 'Sin Cliente';
        if (!clientMap[cliente]) clientMap[cliente] = { total: 0, proc: 0 };
        clientMap[cliente].total++;
        if (r.ESTATUS_FINAL.includes('✅') || r.ESTATUS_FINAL.includes('🚚')) {
            clientMap[cliente].proc++;
        }
    });

    // Top 20 clientes por volumen
    const top = Object.entries(clientMap)
        .sort((a, b) => b[1].total - a[1].total)
        .slice(0, 20);

    const labels = ['Clientes'];
    const parents = [''];
    const values = [0];
    const colors = [0];
    const texts = [''];

    top.forEach(([name, { total, proc }]) => {
        const pct = Math.round((proc / total) * 100);
        const shortName = name.length > 22 ? name.slice(0, 19) + '…' : name;
        labels.push(shortName);
        parents.push('Clientes');
        values.push(total);
        colors.push(pct);
        texts.push(`${total} pedidos · ${pct}% procesado`);
    });

    const data = [{
        type: 'treemap',
        labels: labels,
        parents: parents,
        values: values,
        text: texts,
        textinfo: 'label+text',
        textfont: { size: 11 },
        marker: {
            colors: colors,
            colorscale: [
                [0, '#ef4444'],
                [0.5, '#f59e0b'],
                [1, '#10b981']
            ],
            line: { width: 1, color: 'rgba(15,23,42,0.6)' },
            showscale: true,
            colorbar: {
                title: '% Cumpl.',
                titlefont: { size: 10, color: 'rgba(255,255,255,0.6)' },
                tickfont: { color: 'rgba(255,255,255,0.5)' },
                len: 0.6,
                ticksuffix: '%'
            }
        },
        hovertemplate: '<b>%{label}</b><br>%{text}<extra></extra>',
        pathbar: { visible: false }
    }];

    const layout = {
        ...PLOTLY_DARK_LAYOUT,
        height: 350,
        margin: { t: 10, r: 10, b: 10, l: 10 }
    };

    Plotly.react(el, data, layout, PLOTLY_CONFIG);
}

function cleanPedido(val) {
    if (!val) return "";
    // M-02 related: mismo criterio que limpiar_pedido en Python
    // Eliminar parte decimal, limpiar ceros iniciales
    let s = String(val).split('.')[0].trim();
    while (s.startsWith('0')) s = s.substring(1).trim();
    // A-01 JS: si queda vacío (era "0" o "000"), retornar "" para que no se use como llave
    return s; // "" será ignorado por los Map.set() que hacen guard: if (key)
}

/**
 * Buscador inteligente de columnas por nombre o posición.
 */
function getVal(row, aliases, indexFallback = -1) {
    if (!row) return null;
    
    // 1. Intento por Nombres (Ignora mayúsculas y caracteres especiales)
    for (const alias of aliases) {
        if (row[alias] !== undefined && row[alias] !== null && String(row[alias]).trim() !== "") {
            return row[alias];
        }
        
        // Búsqueda difusa (ej: Normaliza "N°Pedido" a "npedido")
        const cleanAlias = normalizeText(alias);
        for (const key in row) {
            if (normalizeText(key) === cleanAlias && row[key] !== undefined && row[key] !== null) {
                return row[key];
            }
        }
    }

    // 2. Intento por Índice (Fallback si los nombres fallan)
    if (indexFallback >= 0) {
        if (Array.isArray(row) && row[indexFallback] !== undefined) return row[indexFallback];
        const keys = Object.keys(row);
        if (keys[indexFallback] !== undefined) return row[keys[indexFallback]];
    }

    return null;
}

// M-02: excelDateToJS ahora es un alias de excelSerialToDate (función canónica definida arriba)
const excelDateToJS = excelSerialToDate;

function procesarDUS() {
    if (!state.dus || state.dus.length === 0) return [];


    // 1. Mapa: Factura -> Analista (Desde "5.Analista que facturo.XLSX")
    //    Estructura: Col1=Factura, Col2=Fecha factura, Col3=Creado por
    const analistaFactMap = new Map();
    if (state.maestros.analistas) {
        state.maestros.analistas.forEach(row => {
            const fac = cleanValue(getVal(row, ['Factura', 'Folio'], 0));
            const usr = getVal(row, ['Creado por', 'Nombre de usuario', 'Analista'], 2);
            if (fac && usr) analistaFactMap.set(fac, usr);
        });
    }

    // 2. Mapa: Pedido -> Factura (Desde "1.export.XLSX")
    //    Col4=N°Pedido (índice 3), Col12=N° Factura (índice 11)
    const sapBridgeMap = new Map();
    const sapDateMap = new Map();
    if (state.sap) {
        state.sap.forEach(row => {
            const ped = cleanValue(getVal(row, ['N°Pedido', 'NPedido'], 3));
            const fac = cleanValue(getVal(row, ['N° Factura', 'Folio Factura'], 11));
            const fechaRaw = getVal(row, ['Fecha Factura'], 14);
            if (ped && fac) {
                sapBridgeMap.set(ped, fac);
                if (fechaRaw) sapDateMap.set(ped, excelDateToJS(fechaRaw));
            }
        });
    }

    // 3. PROCESAR CADA LÍNEA DEL DUS
    //    Col1=id(0), Col2=Referencia(1), Col3=Consignatario(2), Col4=Estado DUS(3), Col5=Vía Transporte(4), Col6=Fecha Zarpe(5)?
    return state.dus.map((row, idx) => {
        const rawReferencia = String(getVal(row, ['Referencia'], 1) || '');
        const consignatarioStr = getVal(row, ['Consignatario'], 2) || 'N/A';
        const estadoDUS = String(getVal(row, ['Estado DUS'], 3) || 'Pendiente');
        const viaTransporte = getVal(row, ['Vía Transporte', 'Via Transporte', 'Via de Transporte'], 4) || '';

        // Fecha zarpe (columna 5 si existe)
        const zarpeRaw = getVal(row, ['Fecha Zarpe', 'Zarpe', 'F. Zarpe'], 5);
        const fechaZarpe = zarpeRaw ? excelDateToJS(zarpeRaw) : null;
        const zarpeEsFutura = fechaZarpe instanceof Date && !isNaN(fechaZarpe) && fechaZarpe > new Date();
        const sinZarpe = !zarpeRaw || zarpeEsFutura;

        // CLASIFICACIÓN EXPLÍCITA por valor de "Estado DUS"
        let estatusFinal;

        // ✅ Legalizado: DUS completamente cerrado aduaneramente
        if (estadoDUS.includes('Finalizado') || estadoDUS.includes('(Legalizado)')) {
            estatusFinal = '✅ Legalizado';

        // ✅ Legalizado: Pend. Presentación (IVV) = legalizado aduaneramente, solo falta trámite tributario IVV
        } else if (estadoDUS.includes('IVV')) {
            estatusFinal = '✅ Legalizado';

        // ⚠️ Sin Zarpe: nave aún no ha zarpado
        } else if (
            estadoDUS.includes('Pend. Ingreso') ||
            estadoDUS.includes('Zarpe Nave') ||
            estadoDUS.includes('Pend. Zarpe')
        ) {
            estatusFinal = '⚠️ Sin Zarpe';

        // ❌ Anulado: DUS cancelado
        } else if (estadoDUS.includes('Anulado')) {
            estatusFinal = '❌ Anulado';

        // ❌ Pendiente: Pend. Aceptación, Pend. Presentación DUS AT, Pend. Legalización, Rechazado, etc.
        } else {
            estatusFinal = '❌ Pendiente Legalización';
        }

        // --- Sub-estatus SLA: refinar Legalizado/Pendiente con deadline ---
        const deadlineDate = fechaFacturaDate ? calcularDeadlineDUS(fechaFacturaDate) : null;
        if (deadlineDate && estatusFinal === '✅ Legalizado') {
            if (state.fechaHoy > deadlineDate) {
                estatusFinal = '⚠️ Legalizado Fuera de Plazo';
            } else {
                estatusFinal = '✅ Legalizado (A Tiempo)';
            }
        } else if (deadlineDate && estatusFinal === '❌ Pendiente Legalización') {
            if (state.fechaHoy > deadlineDate) {
                estatusFinal = '🔴 Pendiente (Fuera de Plazo)';
            } else {
                estatusFinal = '🔵 Pendiente (En Plazo)';
            }
        }

        // Extraer todos los números que parecen pedidos (7-10 dígitos)
        const pedidoCandidates = rawReferencia.match(/\d{7,10}/g) || [];
        const pedidoClean = pedidoCandidates[0] || rawReferencia;

        let analistaFinal = null;
        let facturaEncontrada = 'N/A';
        let metodoCruce = 'Sin Asignar';
        let fechaFacturaStr = '-';
        let fechaFacturaDate = null;

        for (const cand of pedidoCandidates) {
            const cleanCand = cleanValue(cand);

            const facturaSAP = sapBridgeMap.get(cleanCand);

            if (facturaSAP) {
                facturaEncontrada = facturaSAP;
                const facturador = analistaFactMap.get(facturaSAP);
                if (facturador) {
                    analistaFinal = facturador;
                    metodoCruce = 'Facturador SAP';
                }
                const fObj = sapDateMap.get(cleanCand);
                if (fObj instanceof Date && !isNaN(fObj)) {
                    fechaFacturaStr = fObj.toLocaleDateString();
                    fechaFacturaDate = fObj;
                }
            }

            if (analistaFinal) break;
        }

        if (!analistaFinal) {
            analistaFinal = viaTransporte || 'PENDIENTE';
            metodoCruce = viaTransporte ? 'Vía Transporte' : 'Sin Asignar';
        }

        return {
            ESTATUS_FINAL:       estatusFinal,
            ESTADO_DUS_RAW:      estadoDUS,          // texto original de la planilla
            // P-06: usar pedidoClean como PEDIDO en vez de rawReferencia.
            // rawReferencia puede ser "REF-4500012345/4500012346" que cambia de formato
            // entre archivos DUS, generando llaves duplicadas en la memoria persistente.
            PEDIDO:              pedidoClean,
            PEDIDO_RAW:          rawReferencia,       // guardamos el original para display
            CONSIGNATARIO:       consignatarioStr,
            RESPONSABLE:         analistaFinal,
            FECHA_FACTURA:       fechaFacturaStr,
            FECHA_FACTURA_DATE:  fechaFacturaDate,
            FACTURA_SAP:         facturaEncontrada,
            METODO_CRUCE:        metodoCruce,
            DEMORA:              fechaFacturaDate ? Math.max(0, Math.floor((state.fechaHoy - fechaFacturaDate) / (1000 * 60 * 60 * 24))) : 0,
            DEADLINE:            fechaFacturaDate ? calcularDeadlineDUS(fechaFacturaDate) : null,
            DEMORA_SLA:          (function() {
                if (!fechaFacturaDate) return 0;
                const dl = calcularDeadlineDUS(fechaFacturaDate);
                const ref = isGestionado(estatusFinal) ? state.fechaHoy : state.fechaHoy; // siempre vs hoy para pendientes
                return Math.max(0, Math.floor((ref - dl) / (1000 * 60 * 60 * 24)));
            })()
        };
    });
}



function renderDUS(selected = []) {
    const tbody = document.querySelector('#dus-table tbody');
    if (!tbody) return;

    let data = state.dusResults;
    // Agregar registros de memoria que no están en los resultados actuales
    const memOnly = getMemoryOnlyRecords('dus', data);
    const allData = [...data, ...memOnly];

    const containerId = 'dus-multi-select';
    const totalOptions = document.querySelectorAll(`#${containerId} .ms-option`).length;

    let filtered = allData;
    if (selected.length > 0 && selected.length < totalOptions) {
        filtered = filtered.filter(d => selected.some(s => d.ESTATUS_FINAL === s || d.ESTATUS_FINAL.includes(s)));
    }

    // Aplicar búsqueda de texto
    filtered = applySearch(filtered, state.searchQuery, true);
    // Aplicar filtro de grupo — FIX: comparar tanto código SAP como nombre completo
    // porque en DUS el RESPONSABLE puede venir como nombre largo ("MRAMIREZ") o vía de transporte
    const gFilter = getGrupoAnalysts();
    if (gFilter) {
        const gFilterNames = getGrupoAnalystNames();
        filtered = filtered.filter(d => {
            const resp = String(d.RESPONSABLE || '').toUpperCase().trim();
            return gFilter.has(resp) || gFilterNames.has(resp);
        });
    }
    // Aplicar orden de columna
    const ss = state.sortStateDUS;
    filtered = sortResults(filtered.map(d => ({...d, _isDUS: true})), ss.col, ss.dir);

    // Columna de Observaciones (oculta su cabecera si no hay obs aplicables)
    const obs = loadObs();
    const hasAnyObs = filtered.some(d => {
        const eg = getDUSEstadoGeneral(d.ESTATUS_FINAL);
        return eg.label !== '\u2705 Legalizado' && !!obs[String(d.PEDIDO || '').trim()];
    });
    const obsHeader = document.getElementById('obs-col-header');
    if (obsHeader) obsHeader.style.display = hasAnyObs ? '' : 'none';

    tbody.innerHTML = filtered.map(function(d) {
        const eg = getDUSEstadoGeneral(d.ESTATUS_FINAL);
        const isLegalizado = eg.label === '\u2705 Legalizado';
        const nota = isLegalizado ? '' : (obs[String(d.PEDIDO || '').trim()] || '');
        // OBS-6: escHtml para fechas string (evita XSS si Excel envía texto en vez de fecha)
        const fFac = (d.FECHA_FACTURA instanceof Date) ? d.FECHA_FACTURA.toLocaleDateString() : escHtml(d.FECHA_FACTURA || '-');
        // C-05: escHtml() sanitiza datos de Excel (evita XSS si una celda contiene HTML)
        // P-06: PEDIDO_RAW contiene la referencia completa legible; PEDIDO es el número limpio para llaves
        const ePed  = escHtml(d.PEDIDO_RAW || d.PEDIDO);
        const eCons = escHtml(d.CONSIGNATARIO);
        const eResp = escHtml(d.RESPONSABLE);
        const eDet  = escHtml(d.ESTADO_DUS_RAW || d.ESTATUS_FINAL);
        const eNota = escHtml(nota);
        const obsCell = hasAnyObs
            ? '<td class="obs-cell" title="' + eNota + '">' + eNota + '</td>'
            : '';
        return '<tr class="' + (d._fromMemory ? 'memory-row' : '') + '">' +
            '<td><span class="badge ' + eg.badgeClass + '">' + eg.label + '</span></td>' +
            '<td title="' + eDet + '">' + eDet + '</td>' +
            '<td title="' + ePed + '">' + ePed + '</td>' +
            '<td title="' + eCons + '">' + eCons + '</td>' +
            '<td title="' + eResp + '">' + eResp + '</td>' +
            '<td style="white-space:nowrap">' + fFac + '</td>' +
            obsCell +
            '</tr>';
    }).join('');

    // Actualizar contador de búsqueda
    const countEl = document.getElementById('search-count');
    if (countEl) countEl.textContent = state.searchQuery ? `${filtered.length} resultado${filtered.length !== 1 ? 's' : ''}` : '';
}

/** Clasifica ESTADO GENERAL para el módulo DUS */
function getDUSEstadoGeneral(estatusFinal) {
    const s = String(estatusFinal || '');
    if (s.includes('Legalizado') && s.includes('A Tiempo'))       return { label: '✅ Legalizado (A Tiempo)',       badgeClass: 'badge-green' };
    if (s.includes('Legalizado') && s.includes('Fuera de Plazo')) return { label: '⚠️ Legalizado Fuera de Plazo',   badgeClass: 'badge-orange' };
    if (s.includes('✅') || s.includes('Legalizado'))              return { label: '✅ Legalizado',                   badgeClass: 'badge-green' };
    if (s.includes('Pendiente') && s.includes('Fuera de Plazo'))  return { label: '🔴 Pendiente (Fuera de Plazo)',   badgeClass: 'badge-critical' };
    if (s.includes('Pendiente') && s.includes('En Plazo'))        return { label: '🔵 Pendiente (En Plazo)',         badgeClass: 'badge-blue' };
    if (s.includes('Sin Zarpe'))                                  return { label: '⚠️ Sin Zarpe',                    badgeClass: 'badge-yellow' };
    if (s.includes('Anulado'))                                    return { label: '❌ Anulado',                      badgeClass: 'badge-critical' };
    return                                                               { label: '❌ Pendiente',                     badgeClass: 'badge-red' };
}

// updateChartForDUS movido arriba junto con las funciones Plotly (ver sección PLOTLY.JS)

function exportToExcel() {
    const isAuditoria = state.currentModule === 'auditoria';
    const selected = getSelectedValues(isAuditoria ? 'status-multi-select' : 'dus-multi-select');
    let dataToExport = isAuditoria ? state.results : state.dusResults;

    const containerId = isAuditoria ? 'status-multi-select' : 'dus-multi-select';
    const totalOptions = document.querySelectorAll(`#${containerId} .ms-option`).length;

    if (selected.length > 0 && selected.length < totalOptions) {
        if (isAuditoria) {
            // Para Auditoría: los filtros son etiquetas agrupadas (ej: '✅ Procesado')
            // La comparación debe hacerse via getEstadoGeneral, igual que renderResults.
            // Sin esto, '🚚 Procesado Terrestre' y '⚠️ Validado Manualmente' quedaban excluidos
            // al exportar con el filtro '✅ Procesado' activo.
            dataToExport = dataToExport.filter(r =>
                selected.includes(getEstadoGeneral(r.ESTATUS_FINAL).label)
            );
        } else {
            // Para DUS: los filtros son etiquetas de getDUSEstadoGeneral
            dataToExport = dataToExport.filter(r =>
                selected.some(s => r.ESTATUS_FINAL === s || r.ESTATUS_FINAL.includes(s))
            );
        }
    }

    // Aplicar filtro de grupo activo
    const gFilter = getGrupoAnalysts();
    if (gFilter) {
        dataToExport = dataToExport.filter(r => gFilter.has(String(r.RESPONSABLE || '').toUpperCase().trim()));
    }

    // Aplicar búsqueda activa (lo que ves = lo que descargas)
    if (state.searchQuery) {
        dataToExport = applySearch(dataToExport, state.searchQuery, !isAuditoria);
    }

    if (dataToExport.length === 0) {
        alert("No hay datos para exportar con el filtro actual.");
        return;
    }

    const wsData = isAuditoria
        ? dataToExport.map(r => {
            const eg = getEstadoGeneral(r.ESTATUS_FINAL);
            return {
                'ESTADO GENERAL':  eg.label,
                'DETALLE':         r.ESTATUS_FINAL,
                'DEMORA':          r.DEMORA,
                'T. GESTIÓN':     r.T_GESTIÓN !== null ? r.T_GESTIÓN : '-',
                'N° PEDIDO':      r.PEDIDO,
                'SOLICITANTE':     r.CLIENTE,
                'RESPONSABLE':     r.RESPONSABLE,
                'FECHA FAC.': (r.FECHA_FACTURA instanceof Date && !isNaN(r.FECHA_FACTURA)) ? r.FECHA_FACTURA : '',
                'FECHA BL':   (r.FECHA_BL instanceof Date && !isNaN(r.FECHA_BL)) ? r.FECHA_BL : '',
                'MOTIVO':          r.MOTIVO
              };
          })
        : dataToExport.map(r => {
            const eg = getDUSEstadoGeneral(r.ESTATUS_FINAL);
            const isLegalizado = eg.label === '✅ Legalizado';
            const obsData = loadObs();
            const nota = isLegalizado ? '' : (obsData[String(r.PEDIDO || '').trim()] || '');
            return {
                'ESTADO GENERAL': eg.label,
                'DETALLE':        r.ESTADO_DUS_RAW || r.ESTATUS_FINAL,
                'N° PEDIDO':      r.PEDIDO_RAW || r.PEDIDO,
                'CONSIGNATARIO':  r.CONSIGNATARIO,
                'RESPONSABLE':    r.RESPONSABLE,
                'FECHA FACTURA':  (r.FECHA_FACTURA_DATE instanceof Date && !isNaN(r.FECHA_FACTURA_DATE)) ? r.FECHA_FACTURA_DATE : '',
                'DEMORA (días)':  r.DEMORA || 0,
                'DEADLINE':       r.DEADLINE instanceof Date ? r.DEADLINE : '',
                'ATRASO SLA':     r.DEMORA_SLA || 0,
                'OBSERVACIONES':  nota
            };
        });

    const ws = XLSX.utils.json_to_sheet(wsData, { cellDates: true });

    // Aplicar formato de fecha DD-MM-YYYY a las columnas de fecha
    // para que Excel las ordene correctamente
    if (ws['!ref']) {
        const range = XLSX.utils.decode_range(ws['!ref']);
        const dateFormat = 'DD-MM-YYYY';
        // Detectar columnas de fecha por encabezado
        const dateCols = [];
        for (let C = range.s.c; C <= range.e.c; C++) {
            const hCell = ws[XLSX.utils.encode_cell({c: C, r: 0})];
            if (hCell && /fecha/i.test(String(hCell.v))) dateCols.push(C);
        }
        for (let R = range.s.r + 1; R <= range.e.r; R++) {
            dateCols.forEach(C => {
                const ref = XLSX.utils.encode_cell({c: C, r: R});
                if (ws[ref] && ws[ref].v instanceof Date) {
                    ws[ref].t = 'd';
                    ws[ref].z = dateFormat;
                }
            });
        }
    }

    
    // Aplicar estilos y colores con xlsx-js-style
    if (ws['!ref']) {
        const range = XLSX.utils.decode_range(ws['!ref']);
        
        // Colorear contenido (ESTATUS FINAL es la Columna A = 0)
        for(let R = range.s.r + 1; R <= range.e.r; ++R) {
            const cellRef = XLSX.utils.encode_cell({c: 0, r: R});
            const cell = ws[cellRef];
            if(!cell || !cell.v) continue;
            
            const val = String(cell.v);
            let bgColor = "FFFFFF";
            let fontColor = "000000";
            
            if (val.includes('✅') || val.includes('Validado')) {
                bgColor = "10B981"; // Verde
                fontColor = "FFFFFF";
            } else if (val.includes('IVV')) {
                bgColor = "F59E0B"; // Naranjo
                fontColor = "FFFFFF";
            } else if (val.includes('📦')) {
                bgColor = "F59E0B"; // Amarillo/Naranjo
                fontColor = "FFFFFF";
            } else if (val.includes('Error')) {
                bgColor = "9F1239"; // Crítico/Burdeo
                fontColor = "FFFFFF";
            } else if (val.includes('❌')) {
                bgColor = "EF4444"; // Rojo
                fontColor = "FFFFFF";
            } else if (val.includes('🚚')) {
                bgColor = "3B82F6"; // Azul
                fontColor = "FFFFFF";
            }
            
            cell.s = {
                fill: { fgColor: { rgb: bgColor } },
                font: { color: { rgb: fontColor }, bold: true }
            };
        }
        
        // Colorear la cabecera (Fila 0)
        for(let C = range.s.c; C <= range.e.c; ++C) {
            const headRef = XLSX.utils.encode_cell({c: C, r: 0});
            if(ws[headRef]) {
                ws[headRef].s = {
                    fill: { fgColor: { rgb: "1E293B" } }, // Dark header
                    font: { color: { rgb: "FFFFFF" }, bold: true }
                };
            }
        }
    }

    const wb = XLSX.utils.book_new();
    const sheetName = isAuditoria ? "Auditoría Comex" : "Legalización DUS";
    const grupoSuffix  = state.activeGrupo ? `_${state.activeGrupo.charAt(0).toUpperCase() + state.activeGrupo.slice(1)}` : '';
    const searchSuffix = state.searchQuery  ? `_busqueda_${state.searchQuery.trim().slice(0,15).replace(/\s+/g,'_')}` : '';
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    XLSX.writeFile(wb, `${sheetName}${grupoSuffix}${searchSuffix}_${new Date().toISOString().slice(0,10)}.xlsx`);
}

function updateExportButtonText(selected) {
    const isAuditoria = state.currentModule === 'auditoria';
    const btnId = isAuditoria ? 'export-excel-btn' : 'export-excel-btn-dus';
    const btn = document.getElementById(btnId);
    if (!btn) return;

    // Calcular cuántas filas realmente visibles hay (estado + grupo + búsqueda)
    let visible = isAuditoria ? (state.results || []) : (state.dusResults || []);
    const containerId = isAuditoria ? 'status-multi-select' : 'dus-multi-select';
    const total = document.querySelectorAll(`#${containerId} .ms-option`).length;
    if (selected.length > 0 && selected.length < total) {
        if (isAuditoria) {
            // Para Auditoría: misma lógica de getEstadoGeneral que exportToExcel y renderResults
            visible = visible.filter(r => selected.includes(getEstadoGeneral(r.ESTATUS_FINAL).label));
        } else {
            visible = visible.filter(r => selected.some(s => r.ESTATUS_FINAL === s || r.ESTATUS_FINAL.includes(s)));
        }
    }
    const gf = getGrupoAnalysts();
    if (gf) visible = visible.filter(r => gf.has(String(r.RESPONSABLE || '').toUpperCase().trim()));
    if (state.searchQuery) visible = applySearch(visible, state.searchQuery, !isAuditoria);

    const count = visible.length;
    const hasFilters = selected.length > 0 && selected.length < total || state.activeGrupo || state.searchQuery;
    btn.innerHTML = hasFilters
        ? `<i data-lucide="download"></i> Exportar (${count} resultado${count !== 1 ? 's' : ''})`
        : `<i data-lucide="download"></i> Exportar (Todos)`;
    if (window.lucide) lucide.createIcons();
}

function getSelectedValues(id) {
    return Array.from(document.querySelectorAll(`#${id} input:checked`)).map(i => i.value);
}

// =============================================================
// DIRECTORIO DE EQUIPO (localStorage)
// Asocia código SAP ↔ nombre completo del analista
// =============================================================

const TEAM_KEY = 'accomex_team_directory';

function loadTeamDirectory() {
    try { return JSON.parse(localStorage.getItem(TEAM_KEY) || '{}'); }
    catch { return {}; }
}

function saveTeamDirectory(dir) {
    localStorage.setItem(TEAM_KEY, JSON.stringify(dir));
}

/** Pre-carga el listado de analistas conocidos (solo si el código no existe aún) */
function seedTeamDirectory() {
    const defaultTeam = {
        // ❄️ Congelado
        'EMARTINEZG':  { name: 'Eugenia Martínez',   grupo: 'congelado' },
        'GCHACANO':    { name: 'Gina Chacano',        grupo: 'congelado' },
        'JOSANCHEZM':  { name: 'Jonathan Sánchez',    grupo: 'congelado' },
        'LCONEJEROSM': { name: 'Libni Conejeros',     grupo: 'congelado' },
        'MIFIGUEROA':  { name: 'Matías Figueroa',     grupo: 'congelado' },
        'MNSOTO':      { name: 'Matías Soto',         grupo: 'congelado' },
        'MRAMIREZ':    { name: 'Marcelo Ramírez',     grupo: 'congelado' },
        'SNAGUIL':     { name: 'Soraya Naguil',       grupo: 'congelado' },
        'YVALERO':     { name: 'Andreina Valero',     grupo: 'congelado' },
        // 🌿 Fresco
        'CZAMORANO':   { name: 'Cristina Zamorano',   grupo: 'fresco' },
        'DPINCOL':     { name: 'Daphne Pincol',       grupo: 'fresco' },
        'ESILVAJ':     { name: 'Evelyn Silva',        grupo: 'fresco' },
        'SMOHOR':      { name: 'Scandar Mohor',       grupo: 'fresco' },
        'VVENEGAS':    { name: 'Victor Venegas',      grupo: 'fresco' },
        // Sin grupo
        'XLEICHTLE':   { name: 'Ximena Leichtle',     grupo: null }
    };
    const dir = loadTeamDirectory();
    let changed = false;
    Object.entries(dir).forEach(([code, entry]) => {
        // Migrar formato antiguo: string -> { name, grupo }
        if (typeof entry === 'string') {
            dir[code] = { name: entry, grupo: null };
            changed = true;
        }
    });
    Object.entries(defaultTeam).forEach(([code, seed]) => {
        if (!dir[code]) {
            dir[code] = seed;
            changed = true;
        } else if (dir[code].grupo === null && seed.grupo !== null) {
            // Actualizar grupo si el seed tiene uno y el actual no
            dir[code].grupo = seed.grupo;
            changed = true;
        }
    });
    if (changed) saveTeamDirectory(dir);
}

/** Devuelve el nombre completo del analista dado su código SAP, o el propio código si no existe */
function resolveAnalystName(code) {
    const dir = loadTeamDirectory();
    const entry = dir[String(code).toUpperCase().trim()];
    if (!entry) return code;
    return typeof entry === 'string' ? entry : (entry.name || code);
}

/** Devuelve el grupo de un código SAP, o null si no tiene */
function resolveAnalystGrupo(code) {
    const dir = loadTeamDirectory();
    const entry = dir[String(code).toUpperCase().trim()];
    if (!entry || typeof entry === 'string') return null;
    return entry.grupo || null;
}

/** Devuelve un Set con los códigos del grupo activo, o null si es 'Todos' */
function getGrupoAnalysts() {
    if (!state.activeGrupo) return null;
    const dir = loadTeamDirectory();
    const codes = new Set();
    Object.entries(dir).forEach(([code, entry]) => {
        const g = typeof entry === 'string' ? null : (entry.grupo || null);
        if (g === state.activeGrupo) codes.add(code.toUpperCase().trim());
    });
    return codes;
}

/** Devuelve Set de nombres completos (uppercase) del grupo activo.
 *  Usado en DUS donde el RESPONSABLE puede ser el nombre largo del analista */
function getGrupoAnalystNames() {
    if (!state.activeGrupo) return new Set();
    const dir = loadTeamDirectory();
    const names = new Set();
    Object.entries(dir).forEach(([code, entry]) => {
        const g = typeof entry === 'string' ? null : (entry.grupo || null);
        const name = typeof entry === 'string' ? entry : (entry.name || '');
        if (g === state.activeGrupo) {
            names.add(name.toUpperCase().trim());
            // También agregar el código por si acaso
            names.add(code.toUpperCase().trim());
        }
    });
    return names;
}

/** Activa un grupo y re-renderiza todo */
function setActiveGrupo(grupo) {
    state.activeGrupo = grupo || null;
    document.querySelectorAll('.group-btn').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.grupo === (grupo || ''));
    });
    const infoEl = document.getElementById('group-filter-info');
    if (infoEl) {
        if (state.activeGrupo) {
            const counts = getGrupoAnalysts()?.size || 0;
            infoEl.textContent = `${counts} analista${counts !== 1 ? 's' : ''}`;
        } else {
            infoEl.textContent = '';
        }
    }
    // Re-renderizar todo
    renderResults(getSelectedValues('status-multi-select'));
    renderDUS(getSelectedValues('dus-multi-select'));
    updateStats();
    refreshAnalystFilters();
    refreshAnalystChart();
}

function openTeamModal() {
    document.getElementById('team-modal-overlay').style.display = 'flex';
    renderTeamList();
    if (window.lucide) lucide.createIcons();
}

function closeTeamModal() {
    document.getElementById('team-modal-overlay').style.display = 'none';
}

function renderTeamList() {
    const dir = loadTeamDirectory();
    const list = document.getElementById('team-list');
    const countEl = document.getElementById('team-count');
    const entries = Object.entries(dir);

    if (countEl) countEl.textContent = `${entries.length} miembro(s) registrado(s)`;

    if (entries.length === 0) {
        list.innerHTML = `<div style="color:rgba(255,255,255,0.3); text-align:center; padding:24px; font-size:0.85rem;">
            Sin miembros. Agrega el primero arriba.</div>`;
        return;
    }

    const getGrupoBadge = (entry) => {
        const g = typeof entry === 'string' ? null : (entry.grupo || null);
        if (!g) return '<span class="grupo-badge sin-grupo">Sin grupo</span>';
        return g === 'congelado'
            ? '<span class="grupo-badge congelado">❄️ Congelado</span>'
            : '<span class="grupo-badge fresco">🌿 Fresco</span>';
    };
    const getName = (entry) => typeof entry === 'string' ? entry : (entry.name || '');

    list.innerHTML = entries
        .sort((a, b) => {
            const ga = typeof a[1] === 'string' ? '' : (a[1].grupo || '');
            const gb = typeof b[1] === 'string' ? '' : (b[1].grupo || '');
            return ga.localeCompare(gb) || a[0].localeCompare(b[0]);
        })
        .map(([code, entry]) => {
            // C-05 extendido: sanitizar código SAP y nombre antes de inyectar en innerHTML
            // Previene XSS y el caso de apostrofes en nombres (ej: "O'Brien") que rompen onclick
            const safeCode = escHtml(code);
            const safeName = escHtml(getName(entry));
            // Para el onclick usamos el original escapado con apostrofe para JS string
            // OBS-7: escapar backslash antes que apóstrofe para no romper el JS string
            const jsCode   = code.replace(/\\/g, '\\\\').replace(/'/g, "\\'");
            return `
        <div class="team-member-row">
            <div class="team-member-info">
                <div style="display:flex; align-items:center; gap:8px;">
                    <span class="team-member-code">${safeCode}</span>
                    ${getGrupoBadge(entry)}
                </div>
                <span class="team-member-name">${safeName}</span>
            </div>
            <button class="team-member-delete" onclick="removeTeamMember('${jsCode}')" title="Eliminar">
                <i data-lucide="trash-2"></i>
            </button>
        </div>
    `;
        }).join('');

    if (window.lucide) lucide.createIcons();
}

function addTeamMember() {
    const codeEl = document.getElementById('team-input-code');
    const nameEl = document.getElementById('team-input-name');
    const grupoEl = document.getElementById('team-input-grupo');
    const code = String(codeEl.value || '').toUpperCase().trim();
    const name = String(nameEl.value || '').trim();
    const grupo = grupoEl ? (grupoEl.value || null) : null;

    if (!code || !name) { alert('Completa ambos campos: Código SAP y Nombre.'); return; }

    const dir = loadTeamDirectory();
    dir[code] = { name, grupo };
    saveTeamDirectory(dir);

    codeEl.value = '';
    nameEl.value = '';
    if (grupoEl) grupoEl.value = '';
    codeEl.focus();
    renderTeamList();
    // Refrescar filtros del gráfico
    refreshAnalystFilters();
    refreshAnalystChart();
}

function removeTeamMember(code) {
    if (!confirm(`¿Eliminar a ${resolveAnalystName(code)} (${code}) del directorio?`)) return;
    const dir = loadTeamDirectory();
    delete dir[code];
    saveTeamDirectory(dir);
    renderTeamList();
    refreshAnalystFilters();
    refreshAnalystChart();
}

// =============================================================
// GRÁFICO DE CARGA POR ANALISTA
// Horizontal grouped bars: Asignados vs Procesados/Legalizados
// Fuente: Memoria acumulada del módulo activo
// =============================================================

let analystChart = null;
let selectedAnalysts = new Set(); // vacío = todos
let noAnalystSelected = false;    // A-04: bandera explícita para "ninguno seleccionado"

function initAnalystChart() {
    const ctx = document.getElementById('analystChart')?.getContext('2d');
    if (!ctx) return;

    analystChart = new Chart(ctx, {
        type: 'bar',
        data: { labels: [], datasets: [] },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    labels: { color: 'rgba(255,255,255,0.7)', font: { size: 11 } }
                },
                tooltip: {
                    callbacks: {
                        label: ctx => ` ${ctx.dataset.label}: ${ctx.parsed.x} pedido(s)`
                    }
                }
            },
            scales: {
                x: {
                    ticks: { color: 'rgba(255,255,255,0.5)', font: { size: 10 } },
                    grid: { color: 'rgba(255,255,255,0.06)' },
                    beginAtZero: true,
                    precision: 0
                },
                y: {
                    ticks: { color: 'rgba(255,255,255,0.8)', font: { size: 11 } },
                    grid: { display: false }
                }
            }
        }
    });

    refreshAnalystFilters();
    refreshAnalystChart();
}

/** Construye los datos del gráfico combinando memoria (procesados históricos) + sesión actual (pendientes) */
function buildAnalystChartData(selectedMonth, analystFilter, statusFilter) {
    const module = state.currentModule;

    // --- FUENTE 1: Memoria (procesados/legalizados históricos) ---
    const mem = loadMemory(module);
    const memRecords = Object.values(mem);

    // --- FUENTE 2: Sesión actual (pendientes del día) ---
    const sessionData = (module === 'auditoria' ? state.results : state.dusResults) || [];

    // Agrupar por analista
    const grouped = {};

    const ensureAnalyst = (code) => {
        if (!grouped[code]) grouped[code] = { procesados: 0, pendientes: 0, sinZarpe: 0, pedidos: [] };
    };

    const matchesMonthFilter = (fechaStr) => {
        if (!selectedMonth || selectedMonth === 'all') return true;
        const fecha = fechaStr ? new Date(fechaStr) : null;
        if (!fecha || isNaN(fecha)) return false;
        const key = `${fecha.getFullYear()}-${String(fecha.getMonth()+1).padStart(2,'0')}`;
        return key === selectedMonth;
    };

    // Procesados desde memoria
    memRecords.forEach(r => {
        if (!matchesMonthFilter(r.FECHA_FACTURA)) return;
        const code = String(r.RESPONSABLE || 'N/D').toUpperCase().trim();
        // A-04: respetar bandera noAnalystSelected (ninguno = no mostrar ninguno)
        if (noAnalystSelected) return;
        if (analystFilter.size > 0 && !analystFilter.has(code)) return;
        ensureAnalyst(code);
        grouped[code].procesados++;
        grouped[code].pedidos.push({ ...r, _source: 'memoria', _category: 'procesado' });
    });

    // Pendientes/Sin Zarpe desde sesión actual
    sessionData.forEach(r => {
        if (!matchesMonthFilter(r.FECHA_FACTURA || (r.FECHA_FACTURA_DATE ? r.FECHA_FACTURA_DATE.toISOString() : null))) return;
        const s = String(r.ESTATUS_FINAL || '');
        const egLabel = module === 'dus'
            ? getDUSEstadoGeneral(s).label
            : getEstadoGeneral(s).label;
        const isProcessed = egLabel.includes('✅');
        const isSinZarpe  = egLabel.includes('⚠️') && module === 'dus';
        const isPending   = !isProcessed && !isSinZarpe;
        if (!isPending && !isSinZarpe) return;

        const code = String(r.RESPONSABLE || 'N/D').toUpperCase().trim();
        // A-04: respetar bandera noAnalystSelected
        if (noAnalystSelected) return;
        if (analystFilter.size > 0 && !analystFilter.has(code)) return;
        ensureAnalyst(code);
        if (isSinZarpe) {
            grouped[code].sinZarpe++;
            grouped[code].pedidos.push({ ...r, _source: 'sesión', _category: 'sin_zarpe' });
        } else {
            grouped[code].pendientes++;
            grouped[code].pedidos.push({ ...r, _source: 'sesión', _category: 'pendiente' });
        }
    });

    // Ordenar por total desc
    const sorted = Object.entries(grouped)
        .map(([code, v]) => ({ code, ...v, total: v.procesados + v.pendientes + v.sinZarpe }))
        .sort((a, b) => b.total - a.total);

    // Aplicar filtro de status
    const showProcessed = !statusFilter || statusFilter.has('procesados');
    const showPending   = !statusFilter || statusFilter.has('pendientes');
    const showSinZarpe  = !statusFilter || statusFilter.has('sin_zarpe');

    const labels       = sorted.map(e => resolveAnalystName(e.code));
    const procesados   = sorted.map(e => showProcessed ? e.procesados : 0);
    const pendientes   = sorted.map(e => showPending   ? e.pendientes : 0);
    const sinZarpe     = sorted.map(e => showSinZarpe  ? e.sinZarpe   : 0);
    const allPedidos   = sorted.flatMap(e => e.pedidos); // para export

    return { labels, procesados, pendientes, sinZarpe, allPedidos, sorted, hasData: sorted.length > 0 };
}

/** Rellena el select de mes con los meses disponibles en la memoria */
function refreshAnalystFilters() {
    const module = state.currentModule;
    const mem = loadMemory(module);
    const records = Object.values(mem);

    // Meses disponibles
    const monthSet = new Set();
    records.forEach(r => {
        const fecha = r.FECHA_FACTURA ? new Date(r.FECHA_FACTURA) : null;
        if (fecha && !isNaN(fecha)) {
            monthSet.add(`${fecha.getFullYear()}-${String(fecha.getMonth() + 1).padStart(2,'0')}`);
        }
    });

    const monthSel = document.getElementById('analyst-month-filter');
    if (monthSel) {
        const current = monthSel.value;
        const months = Array.from(monthSet).sort().reverse();
        monthSel.innerHTML = `<option value="all">Todos los meses</option>` +
            months.map(m => {
                const [y, mo] = m.split('-');
                const label = new Date(+y, +mo - 1).toLocaleString('es-CL', { month: 'long', year: 'numeric' });
                const cap = label.charAt(0).toUpperCase() + label.slice(1);
                return `<option value="${m}" ${m === current ? 'selected' : ''}>${cap}</option>`;
            }).join('');
    }

    // Multi-select analistas: combina from memory + directorio (filtrado por grupo activo)
    const analystSet = new Set(records.map(r => String(r.RESPONSABLE || '').toUpperCase().trim()).filter(Boolean));
    const dir = loadTeamDirectory();
    Object.keys(dir).forEach(k => analystSet.add(k.toUpperCase().trim()));

    // Pre-filtrar por grupo activo
    const gFilter = getGrupoAnalysts();
    const filteredAnalysts = gFilter
        ? Array.from(analystSet).filter(code => gFilter.has(code))
        : Array.from(analystSet);

    const optionsEl = document.getElementById('analyst-ms-options');
    if (optionsEl) {
        // Encabezado: Seleccionar todo / Desmarcar todo
        const selectAllHTML = `
            <div style="display:flex; justify-content:space-between; align-items:center; padding:6px 8px 10px; border-bottom:1px solid rgba(255,255,255,0.07); margin-bottom:4px; min-width:0;">
                <button onclick="analystSelectAll(true)" style="background:none;border:none;color:rgba(99,102,241,0.9);font-size:0.78rem;cursor:pointer;padding:2px 4px;white-space:nowrap;flex-shrink:0;">✔ Todos</button>
                <button onclick="analystSelectAll(false)" style="background:none;border:none;color:rgba(248,113,113,0.8);font-size:0.78rem;cursor:pointer;padding:2px 4px;white-space:nowrap;flex-shrink:0;">✕ Ninguno</button>
            </div>`;

        optionsEl.innerHTML = selectAllHTML + filteredAnalysts.sort().map(code => `
            <label class="ms-option">
                <input type="checkbox" value="${escHtml(code)}" ${selectedAnalysts.size === 0 || selectedAnalysts.has(code) ? 'checked' : ''}>
                <span>${escHtml(resolveAnalystName(code))} <small style="opacity:0.5;">(${escHtml(code)})</small></span>
            </label>
        `).join('');

        // Wire checkboxes
        optionsEl.querySelectorAll('input').forEach(cb => {
            cb.addEventListener('change', () => {
                const checked = Array.from(optionsEl.querySelectorAll('input:checked')).map(i => i.value);
                const all = Array.from(optionsEl.querySelectorAll('input')).length;
                // A-04: resetear bandera noAnalystSelected cuando el usuario selecciona individualmente
                noAnalystSelected = false;
                selectedAnalysts = checked.length === all ? new Set() : new Set(checked);
                updateAnalystTriggerLabel(checked, all);
                refreshAnalystChart();
            });
        });
    }
}

/** Actualiza el texto de la píldora del multi-select con texto inteligente + animación */
function updateAnalystTriggerLabel(checked, total) {
    const labelEl = document.getElementById('analyst-ms-label');
    if (!labelEl) return;

    let text;
    if (!checked || checked.length === 0 || checked.length === total) {
        text = '👤 Todos';
    } else if (checked.length === 1) {
        text = `👤 ${resolveAnalystName(checked[0])}`;
    } else if (checked.length === 2) {
        text = `👤 ${checked.map(c => resolveAnalystName(c)).join(' · ')}`;
    } else {
        text = `👤 ${checked.length} analistas`;
    }

    // Animación de pulso suave al cambiar
    labelEl.style.transition = 'opacity 0.15s ease, transform 0.15s ease';
    labelEl.style.opacity = '0';
    labelEl.style.transform = 'translateY(-4px)';
    requestAnimationFrame(() => {
        labelEl.textContent = text;
        labelEl.style.opacity = '1';
        labelEl.style.transform = 'translateY(0)';
    });
}

function setupAnalystMultiSelect() {
    const trigger = document.getElementById('analyst-ms-trigger');
    const container = document.getElementById('analyst-select-container');
    if (!trigger || !container) return;

    trigger.addEventListener('click', (e) => {
        e.stopPropagation();
        container.classList.toggle('open');
    });
    document.addEventListener('click', (e) => {
        if (!e.target.closest('#analyst-select-container'))
            container.classList.remove('open');
    });
}

function refreshAnalystChart() {
    if (!analystChart) return;
    const selectedMonth = document.getElementById('analyst-month-filter')?.value || 'all';

    // Filtro de status (qué datasets mostrar)
    let statusFilter = null; // null = mostrar todos
    const sfChecks = document.querySelectorAll('.analyst-status-filter:checked');
    if (sfChecks.length > 0 && sfChecks.length < document.querySelectorAll('.analyst-status-filter').length) {
        statusFilter = new Set(Array.from(sfChecks).map(c => c.value));
    }

    const { labels, procesados, pendientes, sinZarpe, hasData } =
        buildAnalystChartData(selectedMonth, selectedAnalysts, statusFilter);

    document.getElementById('analyst-empty').style.display = hasData ? 'none' : 'flex';

    const module = state.currentModule;
    const isDUS = module === 'dus';

    analystChart.data.labels = labels;
    analystChart.data.datasets = [
        {
            label: isDUS ? 'Legalizados' : 'Procesados',
            data: procesados,
            backgroundColor: 'rgba(52,211,153,0.6)',
            borderColor: 'rgba(52,211,153,1)',
            borderWidth: 1,
            borderRadius: 4
        },
        {
            label: isDUS ? 'Pendiente Legalización' : 'Pendientes a Facturar',
            data: pendientes,
            backgroundColor: 'rgba(248,113,113,0.6)',
            borderColor: 'rgba(248,113,113,1)',
            borderWidth: 1,
        },
        ...(isDUS ? [{
            label: 'Sin Zarpe',
            data: sinZarpe,
            backgroundColor: 'rgba(251,146,60,0.6)',
            borderColor: 'rgba(251,146,60,1)',
            borderWidth: 1,
            borderRadius: 4
        }] : [])
    ];

    const h = Math.max(200, labels.length * 65);
    const canvas = document.getElementById('analystChart');
    if (canvas) canvas.parentElement.style.maxHeight = h + 'px';

    analystChart.update();
}

/** Exporta el gráfico + tablas en UN SOLO Excel usando ExcelJS (imagen embebida) */
async function exportAnalystReport() {
    if (typeof ExcelJS === 'undefined') {
        alert('La librería ExcelJS aún se está cargando. Intenta en unos segundos.');
        return;
    }

    const selectedMonth = document.getElementById('analyst-month-filter')?.value || 'all';
    let statusFilter = null;
    const sfChecks = document.querySelectorAll('.analyst-status-filter:checked');
    if (sfChecks.length > 0 && sfChecks.length < document.querySelectorAll('.analyst-status-filter').length) {
        statusFilter = new Set(Array.from(sfChecks).map(c => c.value));
    }

    // Aplicar filtro de grupo al reporte de analistas
    const gFilter = getGrupoAnalysts();
    let analystFilterForExport = selectedAnalysts;
    if (gFilter) {
        // Si hay grupo activo, intersectar con la selección manual de analistas
        // Si selectedAnalysts está vacío (= todos), usamos directamente el grupo como filtro
        analystFilterForExport = selectedAnalysts.size === 0
            ? gFilter
            : new Set([...selectedAnalysts].filter(c => gFilter.has(c)));
    }

    const { sorted, allPedidos } = buildAnalystChartData(selectedMonth, analystFilterForExport, statusFilter);
    const module = state.currentModule;
    const isDUS = module === 'dus';
    const fecha = new Date().toISOString().slice(0, 10);
    const grupoSuffix = state.activeGrupo ? `_${state.activeGrupo.charAt(0).toUpperCase() + state.activeGrupo.slice(1)}` : '';
    const titulo = (isDUS ? 'Legalización DUS' : 'Auditoría Comex') + (state.activeGrupo ? ` — ${state.activeGrupo === 'congelado' ? '❄️ Congelado' : '🌿 Fresco'}` : '');

    // ── 1. Capturar imagen del gráfico ──────────────────────────
    const canvas = document.getElementById('analystChart');
    let imgBase64 = null;
    let chartW = 620, chartH = 300;
    if (canvas) {
        chartW = canvas.width;
        chartH = canvas.height;
        // Canvas temporal con fondo oscuro para que las etiquetas blancas sean visibles en Excel
        const exportCanvas = document.createElement('canvas');
        exportCanvas.width  = chartW;
        exportCanvas.height = chartH;
        const ectx = exportCanvas.getContext('2d');
        ectx.fillStyle = '#1e293b';
        ectx.fillRect(0, 0, chartW, chartH);
        ectx.drawImage(canvas, 0, 0);
        imgBase64 = exportCanvas.toDataURL('image/png')
            .replace(/^data:image\/png;base64,/, '');
    }

    // ── 2. Crear workbook ExcelJS ────────────────────────────────
    const wb = new ExcelJS.Workbook();
    wb.creator = 'AQUASHIELD · ExportDesk';
    wb.created = new Date();

    // Estilo de cabecera (oscuro con texto blanco)
    const headerFill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
    const headerFont  = { color: { argb: 'FFFFFFFF' }, bold: true, name: 'Calibri', size: 11 };
    const dataFont    = { name: 'Calibri', size: 10, color: { argb: 'FF1E293B' } };
    const borderStyle = { style: 'thin', color: { argb: 'FFCBD5E1' } };
    const allBorders  = { top: borderStyle, left: borderStyle, bottom: borderStyle, right: borderStyle };

    const applyHeaderRow = (row) => {
        row.eachCell(cell => {
            cell.fill = headerFill;
            cell.font = headerFont;
            cell.border = allBorders;
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
        });
        row.height = 22;
    };

    // Filas de datos: fondo blanco / gris muy claro alternado, texto oscuro
    const applyDataRow = (row, isEven) => {
        row.eachCell(cell => {
            cell.font = dataFont;
            cell.border = allBorders;
            cell.fill = isEven
                ? { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } }   // blanco
                : { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };  // gris muy claro
            cell.alignment = { vertical: 'middle' };
        });
        row.height = 18;
    };


    // ── HOJA 1: Gráfico + Resumen ────────────────────────────────
    const wsGrafico = wb.addWorksheet('Gráfico');
    wsGrafico.views = [{ showGridLines: false }];

    // Título
    wsGrafico.mergeCells('A1:G1');
    const titleCell = wsGrafico.getCell('A1');
    titleCell.value = `Carga por Analista — ${titulo} — ${fecha}`;
    titleCell.font = { bold: true, size: 13, color: { argb: 'FFFFFFFF' }, name: 'Calibri' };
    titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } };
    titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
    wsGrafico.getRow(1).height = 30;

    // Imagen del gráfico
    const chartRowStart = 2;
    // 1 Excel point (height=18) ≈ 24px at 96dpi (18pt * 1.333px/pt)
    const PX_PER_ROW = 24;
    const chartRows = Math.ceil(chartH / PX_PER_ROW) + 5; // +5 filas buffer
    if (imgBase64) {
        const imageId = wb.addImage({ base64: imgBase64, extension: 'png' });
        wsGrafico.addImage(imageId, {
            tl: { col: 0, row: chartRowStart - 1 },
            ext: { width: Math.min(chartW, 620), height: chartH }
        });
        // Fijar altura de cada fila reservada
        for (let i = chartRowStart; i <= chartRowStart + chartRows; i++) {
            wsGrafico.getRow(i).height = 18;
        }
    }

    // Tabla resumen debajo de la imagen
    const tableStartRow = chartRowStart + chartRows + 2;
    const sumHeaders = ['Analista', 'Código SAP', 'Procesados', 'Pendientes',
        ...(isDUS ? ['Sin Zarpe'] : []), 'Total'];
    const sumRow = wsGrafico.getRow(tableStartRow);
    sumHeaders.forEach((h, i) => { sumRow.getCell(i + 1).value = h; });
    applyHeaderRow(sumRow);

    sorted.forEach((e, idx) => {
        const row = wsGrafico.getRow(tableStartRow + 1 + idx);
        const vals = [resolveAnalystName(e.code), e.code, e.procesados, e.pendientes,
            ...(isDUS ? [e.sinZarpe] : []), e.total];
        vals.forEach((v, i) => { row.getCell(i + 1).value = v; });
        applyDataRow(row, idx % 2 === 0);
    });

    // Ancho columnas
    wsGrafico.columns = sumHeaders.map((_, i) => ({ width: i === 0 ? 28 : 14 }));

    // ── HOJA 2: Pedidos (detalle) ────────────────────────────────
    const wsPedidos = wb.addWorksheet('Pedidos');
    wsPedidos.views = [{ showGridLines: false }];

    const pedHeaders = isDUS
        ? ['Analista', 'Código SAP', 'Categoría', 'Estatus', 'N° Pedido', 'Consignatario', 'Factura SAP', 'Fecha Factura']
        : ['Analista', 'Código SAP', 'Categoría', 'Estado General', 'Detalle', 'N° Pedido', 'Cliente', 'Fecha Fac.'];

    const headRow = wsPedidos.getRow(1);
    pedHeaders.forEach((h, i) => { headRow.getCell(i + 1).value = h; });
    applyHeaderRow(headRow);

    allPedidos.forEach((p, idx) => {
        const code = String(p.RESPONSABLE || '').toUpperCase();
        const catLabel = p._category === 'procesado'
            ? '✅ Procesado'
            : p._category === 'sin_zarpe' ? '⚠️ Sin Zarpe' : '❌ Pendiente';

        const vals = isDUS
            ? [resolveAnalystName(code), code, catLabel, p.ESTATUS_FINAL, p.PEDIDO,
               p.CONSIGNATARIO || '-', p.FACTURA_SAP || '-',
               p.FECHA_FACTURA_DATE instanceof Date ? p.FECHA_FACTURA_DATE : '']
            : [resolveAnalystName(code), code, catLabel,
               getEstadoGeneral(p.ESTATUS_FINAL).label, p.ESTATUS_FINAL,
               p.PEDIDO, p.CLIENTE || '-',
               p.FECHA_FACTURA instanceof Date ? p.FECHA_FACTURA : ''];

        const row = wsPedidos.getRow(idx + 2);
        vals.forEach((v, i) => { row.getCell(i + 1).value = v; });
        applyDataRow(row, idx % 2 === 0);

        // Formato fecha para la última columna
        const dateCell = row.getCell(vals.length);
        if (dateCell.value instanceof Date) dateCell.numFmt = 'DD-MM-YYYY';
    });

    wsPedidos.columns = pedHeaders.map((_, i) => ({
        width: [28, 14, 16, 18, 22, 16, 26, 16][i] || 14
    }));

    // ── 3. Descargar ─────────────────────────────────────────────
    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `Reporte_Analistas_${isDUS ? 'DUS' : 'Auditoria'}_${new Date().toISOString().slice(0,10)}.xlsx`;
    link.click();
    URL.revokeObjectURL(url);
}

/** Selecciona o desmarca todos los analistas del multi-select */
function analystSelectAll(selectAll) {
    const optionsEl = document.getElementById('analyst-ms-options');
    if (!optionsEl) return;

    const checkboxes = Array.from(optionsEl.querySelectorAll('input[type="checkbox"]'));
    checkboxes.forEach(cb => { cb.checked = selectAll; });

    if (selectAll) {
        selectedAnalysts = new Set(); // vacío = todos
        noAnalystSelected = false;    // A-04: limpiar bandera
        updateAnalystTriggerLabel(checkboxes.map(c => c.value), checkboxes.length);
    } else {
        // A-04: usar bandera booleana explícita en vez del sentinel '__ninguno__'
        // que podría colisionar con un código SAP real en el futuro
        selectedAnalysts = new Set();
        noAnalystSelected = true;
        const labelEl = document.getElementById('analyst-ms-label');
        if (labelEl) {
            labelEl.style.transition = 'opacity 0.15s ease, transform 0.15s ease';
            labelEl.style.opacity = '0';
            labelEl.style.transform = 'translateY(-4px)';
            requestAnimationFrame(() => {
                labelEl.textContent = '👤 Ninguno';
                labelEl.style.opacity = '1';
                labelEl.style.transform = 'translateY(0)';
            });
        }
    }
    refreshAnalystChart();
}

// =============================================================
// MÓDULO: DUS OBSERVACIONES
// =============================================================

const OBS_KEY = 'dus_observaciones';

function loadObs() {
    try { return JSON.parse(localStorage.getItem(OBS_KEY) || '{}'); }
    catch { return {}; }
}
function saveObs(obs) {
    // A-06 aplicado también a observaciones: capturar QuotaExceededError
    try {
        localStorage.setItem(OBS_KEY, JSON.stringify(obs));
    } catch (e) {
        if (e.name === 'QuotaExceededError' || e.code === 22) {
            alert('⚠️ La memoria de observaciones está llena. Exporta y limpia el historial.');
        }
    }
}

function mergeObs(dest, pedido, nota) {
    const key = String(pedido || '').trim();
    const text = String(nota || '').trim();
    if (!key || !text) return;
    if (!dest[key]) { dest[key] = text; }
    else if (!dest[key].split(' | ').includes(text)) { dest[key] = dest[key] + ' | ' + text; }
}

async function parseObsExcel(file) {
    const data = await file.arrayBuffer();
    const wb = XLSX.read(data, { type: 'array' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    if (rows.length < 2) { alert('El archivo no contiene datos.'); return; }
    const headers = rows[0].map(h => String(h || '').toLowerCase().trim());
    const pedidoIdx = headers.findIndex(h => h.includes('pedido'));
    const obsIdx = headers.findIndex(h => h.includes('observ'));
    if (pedidoIdx === -1 || obsIdx === -1) {
        alert('No se encontraron columnas "N Pedido" y "Observaciones".\nRevisa que el Excel tenga esas columnas.');
        return;
    }
    const obs = loadObs();
    let added = 0, updated = 0, skipped = 0;
    for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        const pedido = String(row[pedidoIdx] || '').trim();
        const nota = String(row[obsIdx] || '').trim();
        if (!pedido || !nota) { skipped++; continue; }
        const before = obs[pedido];
        mergeObs(obs, pedido, nota);
        if (!before) added++;
        else if (obs[pedido] !== before) updated++;
        else skipped++;
    }
    saveObs(obs);
    renderObsList();
    renderDUS(getSelectedValues('dus-multi-select'));
    alert('Observaciones cargadas:\n' + added + ' nuevas  -  ' + updated + ' actualizadas  -  ' + skipped + ' omitidas');
}

function addObsManual() {
    const pedidoEl = document.getElementById('obs-input-pedido');
    const notaEl = document.getElementById('obs-input-nota');
    const pedido = String(pedidoEl?.value || '').trim();
    const nota = String(notaEl?.value || '').trim();
    if (!pedido || !nota) { alert('Completa el N de Pedido y la observacion.'); return; }
    const obs = loadObs();
    mergeObs(obs, pedido, nota);
    saveObs(obs);
    pedidoEl.value = ''; notaEl.value = ''; pedidoEl.focus();
    renderObsList();
    renderDUS(getSelectedValues('dus-multi-select'));
}

function deleteObs(pedido) {
    const obs = loadObs();
    delete obs[String(pedido).trim()];
    saveObs(obs);
    renderObsList();
    renderDUS(getSelectedValues('dus-multi-select'));
}

function renderObsList() {
    const obs = loadObs();
    const container = document.getElementById('obs-list-container');
    const listEl = document.getElementById('obs-list');
    const badge = document.getElementById('obs-count-badge');
    const entries = Object.entries(obs);
    if (badge) badge.textContent = entries.length > 0 ? String(entries.length) : '';
    if (!listEl) return;
    if (entries.length === 0) {
        if (container) container.style.display = 'none';
        listEl.innerHTML = '';
        return;
    }
    if (container) container.style.display = 'block';
    listEl.innerHTML = '<table class="obs-table"><thead><tr><th>Pedido</th><th>Observacion</th><th></th></tr></thead><tbody>' +
        entries.map(function(e) {
            var p = escHtml(e[0]),   // C-05: sanitizar pedido
                n = escHtml(e[1]);  // C-05: sanitizar nota de observación
            return '<tr><td class="obs-pedido-cell">' + p + '</td>' +
                   '<td class="obs-nota-cell" title="' + n + '">' + n + '</td>' +
                   // OBS-7: escapar backslash + apóstrofe para onclick seguro
                   '<td><button class="obs-delete-btn" onclick="deleteObs(\'' + e[0].replace(/\\/g,'\\\\').replace(/'/g,"\\'") + '\')" title="Eliminar"><i data-lucide="x"></i></button></td></tr>';
        }).join('') +
        '</tbody></table>';
    if (window.lucide) lucide.createIcons();
}

function setupObsWidget() {
    const dropzone = document.getElementById('obs-dropzone');
    const fileInput = document.getElementById('obs-file-input');
    if (!dropzone || !fileInput) return;
    dropzone.addEventListener('click', function() { fileInput.click(); });
    fileInput.addEventListener('change', function() {
        if (fileInput.files[0]) parseObsExcel(fileInput.files[0]);
        fileInput.value = '';
    });
    dropzone.addEventListener('dragover', function(e) { e.preventDefault(); dropzone.classList.add('drag-over'); });
    dropzone.addEventListener('dragleave', function() { dropzone.classList.remove('drag-over'); });
    dropzone.addEventListener('drop', function(e) {
        e.preventDefault();
        dropzone.classList.remove('drag-over');
        var file = e.dataTransfer.files[0];
        if (file && /\.(xlsx|xls)$/i.test(file.name)) parseObsExcel(file);
        else alert('Por favor arrastra un archivo Excel (.xlsx / .xls)');
    });
    var notaInput = document.getElementById('obs-input-nota');
    if (notaInput) notaInput.addEventListener('keydown', function(e) { if (e.key === 'Enter') addObsManual(); });
    renderObsList();
}

/** Toggle expand/collapse of DUS Observaciones widget body */
function toggleObsWidget() {
    const body = document.getElementById('obs-collapsible-body');
    const chevron = document.getElementById('obs-chevron');
    if (!body) return;
    const isOpen = body.classList.toggle('obs-expanded');
    if (chevron) {
        chevron.style.transform = isOpen ? 'rotate(180deg)' : 'rotate(0deg)';
    }
}

// =============================================================
// KPI ENGINE — Panel de Indicadores Clave
// =============================================================

/** Helper global: detecta si un ESTATUS_FINAL ya está "gestionado"
 *  (procesado, terrestre, legalizado, o validado manualmente) */
function isGestionado(est) {
    return est.includes('Legalizado') || est.includes('✅') || est.includes('🚚') || est.includes('Terrestre') || est.includes('Validado Manualmente');
}

/** Calcula el deadline DUS: día 8 del mes siguiente a la fecha de factura */
function calcularDeadlineDUS(fechaFactura) {
    if (!fechaFactura || !(fechaFactura instanceof Date) || isNaN(fechaFactura)) return null;
    const y = fechaFactura.getFullYear();
    const m = fechaFactura.getMonth(); // 0-based
    // Mes siguiente, día 8, medianoche
    return new Date(y, m + 1, 8, 0, 0, 0, 0);
}

/** Calcula todos los KPIs a partir de los resultados actuales */
function calculateKPIs(results) {
    if (!results || results.length === 0) return null;

    const isDUS = state.currentModule === 'dus';
    const total = results.length;
    const dayMs = 1000 * 60 * 60 * 24;

    // --- Demoras (solo pedidos PENDIENTES) ---
    const demorasPendientes = results
        .filter(r => typeof r.DEMORA === 'number' && r.DEMORA > 0 && !isGestionado(r.ESTATUS_FINAL || ''))
        .map(r => r.DEMORA);
    const demoraPromedio = demorasPendientes.length > 0
        ? (demorasPendientes.reduce((a, b) => a + b, 0) / demorasPendientes.length)
        : 0;
    const totalPendientes = demorasPendientes.length;

    // --- SLA Auditoría: dentro de 3 días | DUS: antes del día 8 del mes siguiente ---
    let dentroPlazo, slaPct;
    if (isDUS) {
        // SLA DUS: legalizados ANTES de su deadline (día 8 mes siguiente)
        dentroPlazo = results.filter(r => {
            if (!isGestionado(r.ESTATUS_FINAL || '')) return false;
            const dl = r.DEADLINE;
            if (!dl) return true; // sin fecha → no penalizar
            return state.fechaHoy <= dl; // se legalizó a tiempo
        }).length;
    } else {
        dentroPlazo = results.filter(r => {
            const d = parseInt(r.DEMORA) || 0;
            return isGestionado(r.ESTATUS_FINAL || '') && d <= 3;
        }).length;
    }
    const totalProcesados = results.filter(r => isGestionado(r.ESTATUS_FINAL || '')).length;
    slaPct = totalProcesados > 0 ? Math.round((dentroPlazo / totalProcesados) * 100) : 0;

    // --- DUS Fuera de SLA: pendientes cuyo deadline ya venció ---
    let fueraSLA = 0, demoraPromSLA = 0;
    if (isDUS) {
        const fueraSLAList = results.filter(r => {
            if (isGestionado(r.ESTATUS_FINAL || '')) return false;
            const dl = r.DEADLINE;
            return dl && state.fechaHoy > dl;
        });
        fueraSLA = fueraSLAList.length;
        if (fueraSLA > 0) {
            demoraPromSLA = fueraSLAList.reduce((sum, r) => {
                return sum + Math.max(0, Math.floor((state.fechaHoy - r.DEADLINE) / dayMs));
            }, 0) / fueraSLA;
        }
    }

    // --- En Riesgo (>7 días, no gestionados) ---
    const enRiesgo = results.filter(r => {
        const d = parseInt(r.DEMORA) || 0;
        return d > 7 && !isGestionado(r.ESTATUS_FINAL || '');
    }).length;

    // --- Críticos (>14 días, no gestionados) ---
    const criticos = results.filter(r => {
        const d = parseInt(r.DEMORA) || 0;
        return d > 14 && !isGestionado(r.ESTATUS_FINAL || '');
    }).length;

    // --- T. Gestión Promedio ---
    // Auditoría: Factura → BL | DUS: Factura → Hoy (total de días sin gestionar)
    let tGestionProm = null;
    if (!isDUS) {
        const tGestiones = results
            .filter(r => typeof r.T_GESTIÓN === 'number' && r.T_GESTIÓN >= 0)
            .map(r => r.T_GESTIÓN);
        if (tGestiones.length > 0) {
            tGestionProm = (tGestiones.reduce((a, b) => a + b, 0) / tGestiones.length);
        }
    } else {
        // DUS: promedio de DEMORA total (factura → hoy) de TODOS los registros con fecha
        const demorasTotales = results
            .filter(r => typeof r.DEMORA === 'number' && r.DEMORA > 0)
            .map(r => r.DEMORA);
        if (demorasTotales.length > 0) {
            tGestionProm = (demorasTotales.reduce((a, b) => a + b, 0) / demorasTotales.length);
        }
    }

    // --- Por Analista ---
    const porAnalista = {};
    results.forEach(r => {
        const resp = r.RESPONSABLE || 'Sin Asignar';
        if (!porAnalista[resp]) porAnalista[resp] = { total: 0, proc: 0, demoras: [], tgestiones: [], fueraSLA: 0 };
        porAnalista[resp].total++;
        const est = r.ESTATUS_FINAL || '';
        if (isGestionado(est)) porAnalista[resp].proc++;
        if (typeof r.DEMORA === 'number' && r.DEMORA > 0 && !isGestionado(est)) porAnalista[resp].demoras.push(r.DEMORA);
        if (typeof r.T_GESTIÓN === 'number' && r.T_GESTIÓN >= 0) porAnalista[resp].tgestiones.push(r.T_GESTIÓN);
        // DUS: contar fuera de SLA por analista
        if (isDUS && !isGestionado(est) && r.DEADLINE && state.fechaHoy > r.DEADLINE) {
            porAnalista[resp].fueraSLA++;
        }
    });

    // --- Por Mes ---
    const porMes = {};
    results.forEach(r => {
        let mes = 'Sin Fecha';
        if (r.FECHA_FACTURA instanceof Date && !isNaN(r.FECHA_FACTURA)) {
            mes = r.FECHA_FACTURA.toISOString().slice(0, 7);
        } else if (r.FECHA_FACTURA_DATE instanceof Date && !isNaN(r.FECHA_FACTURA_DATE)) {
            mes = r.FECHA_FACTURA_DATE.toISOString().slice(0, 7);
        }
        if (!porMes[mes]) porMes[mes] = { total: 0, proc: 0, demoras: [], fueraSLA: 0 };
        porMes[mes].total++;
        const est = r.ESTATUS_FINAL || '';
        if (isGestionado(est)) porMes[mes].proc++;
        if (typeof r.DEMORA === 'number' && r.DEMORA > 0 && !isGestionado(est)) porMes[mes].demoras.push(r.DEMORA);
        if (isDUS && !isGestionado(est) && r.DEADLINE && state.fechaHoy > r.DEADLINE) {
            porMes[mes].fueraSLA++;
        }
    });

    // --- Por Cliente ---
    const porCliente = {};
    results.forEach(r => {
        const cli = r.CLIENTE || r.CONSIGNATARIO || 'Sin Cliente';
        if (!porCliente[cli]) porCliente[cli] = { total: 0, proc: 0, demoras: [] };
        porCliente[cli].total++;
        const est = r.ESTATUS_FINAL || '';
        if (isGestionado(est)) porCliente[cli].proc++;
        if (typeof r.DEMORA === 'number' && r.DEMORA > 0 && !isGestionado(est)) porCliente[cli].demoras.push(r.DEMORA);
    });

    return {
        total, totalProcesados, totalPendientes, demoraPromedio, slaPct, dentroPlazo,
        enRiesgo, criticos, tGestionProm,
        fueraSLA, demoraPromSLA,
        porAnalista, porMes, porCliente
    };
}

/** Renderiza las tarjetas KPI con semáforo */
function renderKPIPanel(results) {
    const panel = document.getElementById('kpi-panel-widget');
    if (!panel) return;

    const kpis = calculateKPIs(results);
    if (!kpis) {
        panel.style.display = 'none';
        return;
    }
    panel.style.display = '';

    // Helper: set card value + color
    function setCard(id, value, color) {
        const valEl = document.getElementById(id + '-val');
        const cardEl = document.getElementById(id);
        if (valEl) valEl.textContent = value;
        if (cardEl) {
            cardEl.classList.remove('kpi-green', 'kpi-yellow', 'kpi-red', 'kpi-critical');
            if (color) cardEl.classList.add(color);
        }
    }

    const isDUS = state.currentModule === 'dus';

    // 1. Demora Promedio
    const dp = kpis.demoraPromedio.toFixed(1);
    setCard('kpi-demora', dp + 'd',
        kpis.demoraPromedio <= 3 ? 'kpi-green' : kpis.demoraPromedio <= 7 ? 'kpi-yellow' : 'kpi-red');
    const dpSub = document.getElementById('kpi-demora-sub');
    if (dpSub) dpSub.textContent = `${kpis.totalPendientes} pendientes`;
    const dpLabel = document.getElementById('kpi-demora')?.querySelector('.kpi-label');
    if (dpLabel) dpLabel.textContent = 'Demora Prom.';

    // 2. SLA %
    setCard('kpi-sla', kpis.slaPct + '%',
        kpis.slaPct >= 85 ? 'kpi-green' : kpis.slaPct >= 60 ? 'kpi-yellow' : 'kpi-red');
    const slaSub = document.getElementById('kpi-sla-sub');
    const slaLabel = document.getElementById('kpi-sla')?.querySelector('.kpi-label');
    if (isDUS) {
        if (slaLabel) slaLabel.textContent = 'SLA (≤8 día sig.)';
        if (slaSub) slaSub.textContent = `${kpis.dentroPlazo} de ${kpis.totalProcesados} a tiempo`;
    } else {
        if (slaLabel) slaLabel.textContent = 'SLA (≤3 días)';
        if (slaSub) slaSub.textContent = `${kpis.dentroPlazo} de ${kpis.totalProcesados} en plazo`;
    }

    // 3. En Riesgo / Fuera SLA
    if (isDUS) {
        setCard('kpi-riesgo', String(kpis.fueraSLA),
            kpis.fueraSLA === 0 ? 'kpi-green' : 'kpi-red');
        const rLabel = document.getElementById('kpi-riesgo')?.querySelector('.kpi-label');
        if (rLabel) rLabel.textContent = 'Fuera de SLA';
        const rSub = document.getElementById('kpi-riesgo-sub');
        if (rSub) rSub.textContent = kpis.fueraSLA > 0 ? `${kpis.demoraPromSLA.toFixed(1)}d prom. atraso` : 'Todo al día';
    } else {
        const riesgoPct = kpis.total > 0 ? Math.round((kpis.enRiesgo / kpis.total) * 100) : 0;
        setCard('kpi-riesgo', String(kpis.enRiesgo),
            kpis.enRiesgo === 0 ? 'kpi-green' : riesgoPct <= 5 ? 'kpi-yellow' : 'kpi-red');
        const rLabel = document.getElementById('kpi-riesgo')?.querySelector('.kpi-label');
        if (rLabel) rLabel.textContent = 'En Riesgo';
        const rSub = document.getElementById('kpi-riesgo-sub');
        if (rSub) rSub.textContent = `${riesgoPct}% del total`;
    }

    // 4. Críticos
    setCard('kpi-criticos', String(kpis.criticos),
        kpis.criticos === 0 ? 'kpi-green' : kpis.criticos <= 3 ? 'kpi-red' : 'kpi-critical');
    const cSub = document.getElementById('kpi-criticos-sub');
    if (cSub) cSub.textContent = kpis.criticos > 0 ? '¡Acción inmediata!' : 'Sin críticos';

    // 5. T. Gestión
    const tgLabel = document.getElementById('kpi-tgestion')?.querySelector('.kpi-label');
    if (kpis.tGestionProm !== null) {
        const tg = kpis.tGestionProm.toFixed(1);
        if (isDUS) {
            if (tgLabel) tgLabel.textContent = 'T. Total Prom.';
            setCard('kpi-tgestion', tg + 'd',
                kpis.tGestionProm <= 20 ? 'kpi-green' : kpis.tGestionProm <= 35 ? 'kpi-yellow' : 'kpi-red');
            const tgSub = document.getElementById('kpi-tgestion-sub');
            if (tgSub) tgSub.textContent = 'Factura → Hoy';
        } else {
            if (tgLabel) tgLabel.textContent = 'T. Gestión Prom.';
            setCard('kpi-tgestion', tg + 'd',
                kpis.tGestionProm <= 5 ? 'kpi-green' : kpis.tGestionProm <= 10 ? 'kpi-yellow' : 'kpi-red');
            const tgSub = document.getElementById('kpi-tgestion-sub');
            if (tgSub) tgSub.textContent = 'Factura → BL';
        }
    } else {
        setCard('kpi-tgestion', '—', null);
        if (tgLabel) tgLabel.textContent = isDUS ? 'T. Total Prom.' : 'T. Gestión Prom.';
    }

    if (window.lucide) lucide.createIcons();
}

/** Exporta KPIs a Excel multi-hoja CON COLORES (xlsx-js-style) */
function exportKPIExcel() {
    const results = state.currentModule === 'dus' ? state.dusResults : state.results;
    const kpis = calculateKPIs(results);
    if (!kpis) { alert('Sin datos para exportar. Ejecuta un análisis primero.'); return; }

    const wb = XLSX.utils.book_new();
    const hoy = new Date().toLocaleDateString('es-CL');
    const mod = state.currentModule === 'dus' ? 'DUS' : 'Auditoría';

    // ═══ Estilos reutilizables ═══
    const sHeader = { font: { bold: true, color: { rgb: 'FFFFFF' }, sz: 11 }, fill: { fgColor: { rgb: '1e293b' } }, alignment: { horizontal: 'center' } };
    const sTitle  = { font: { bold: true, sz: 14, color: { rgb: '0ea5e9' } } };
    const sGreen  = { fill: { fgColor: { rgb: 'dcfce7' } }, font: { color: { rgb: '166534' }, bold: true } };
    const sYellow = { fill: { fgColor: { rgb: 'fef9c3' } }, font: { color: { rgb: '854d0e' }, bold: true } };
    const sRed    = { fill: { fgColor: { rgb: 'fee2e2' } }, font: { color: { rgb: '991b1b' }, bold: true } };
    const sCrit   = { fill: { fgColor: { rgb: 'fecaca' } }, font: { color: { rgb: '7f1d1d' }, bold: true } };
    const sNeutral= { fill: { fgColor: { rgb: 'f1f5f9' } }, font: { color: { rgb: '475569' } } };

    function colorForEstado(text) {
        if (text.includes('Verde')) return sGreen;
        if (text.includes('Naranja')) return sYellow;
        if (text.includes('Rojo')) return sRed;
        if (text.includes('Crítico')) return sCrit;
        return sNeutral;
    }

    function styleHeaders(ws, cols) {
        for (let c = 0; c < cols; c++) {
            const addr = XLSX.utils.encode_cell({ r: 0, c });
            if (ws[addr]) ws[addr].s = sHeader;
        }
    }

    function styleCumplCol(ws, rows, colIdx) {
        for (let r = 1; r <= rows; r++) {
            const addr = XLSX.utils.encode_cell({ r, c: colIdx });
            if (!ws[addr]) continue;
            const val = parseInt(String(ws[addr].v)) || 0;
            ws[addr].s = val >= 85 ? sGreen : val >= 60 ? sYellow : sRed;
        }
    }

    function styleDemoraCol(ws, rows, colIdx) {
        for (let r = 1; r <= rows; r++) {
            const addr = XLSX.utils.encode_cell({ r, c: colIdx });
            if (!ws[addr] || ws[addr].v === '-') continue;
            const val = parseFloat(ws[addr].v) || 0;
            ws[addr].s = val <= 3 ? sGreen : val <= 7 ? sYellow : sRed;
        }
    }

    // ═══ Hoja 1: Resumen ═══
    const isDUSExport = state.currentModule === 'dus';
    const estadoDemora = kpis.demoraPromedio <= 3 ? '🟢 Verde' : kpis.demoraPromedio <= 7 ? '🟠 Naranja' : '🔴 Rojo';
    const estadoSLA    = kpis.slaPct >= 85 ? '🟢 Verde' : kpis.slaPct >= 60 ? '🟠 Naranja' : '🔴 Rojo';
    const estadoRiesgo = isDUSExport ? (kpis.fueraSLA === 0 ? '🟢 Verde' : '🔴 Rojo') : (kpis.enRiesgo === 0 ? '🟢 Verde' : '🔴 Rojo');
    const estadoCrit   = kpis.criticos === 0 ? '🟢 Verde' : '💀 Crítico';
    const estadoTG     = kpis.tGestionProm !== null ? (isDUSExport ? (kpis.tGestionProm <= 20 ? '🟢 Verde' : kpis.tGestionProm <= 35 ? '🟠 Naranja' : '🔴 Rojo') : (kpis.tGestionProm <= 5 ? '🟢 Verde' : '🟠 Naranja')) : '⚪ N/A';

    const resumen = [
        ['PANEL KPI — ExportDesk', '', '', hoy],
        ['Módulo:', mod],
        [],
        ['INDICADOR', 'VALOR', 'ESTADO', 'DESCRIPCIÓN'],
        ['Demora Promedio', kpis.demoraPromedio.toFixed(1) + ' días', estadoDemora, 'Promedio días demora pedidos pendientes'],
        [isDUSExport ? 'SLA (≤8 día sig.)' : 'SLA (≤3 días)', kpis.slaPct + '%', estadoSLA, kpis.dentroPlazo + ' de ' + kpis.totalProcesados + (isDUSExport ? ' legalizados a tiempo' : ' procesados dentro de plazo')],
        [isDUSExport ? 'Fuera de SLA' : 'En Riesgo (>7d)', isDUSExport ? kpis.fueraSLA : kpis.enRiesgo, estadoRiesgo, isDUSExport ? 'Pendientes cuyo deadline (día 8 mes sig.) ya venció' : 'Pedidos pendientes con más de 7 días de demora'],
        ['Críticos (>14d)', kpis.criticos, estadoCrit, 'Pedidos pendientes con más de 14 días — requieren acción inmediata'],
        [isDUSExport ? 'T. Total Prom.' : 'T. Gestión Prom.', kpis.tGestionProm !== null ? kpis.tGestionProm.toFixed(1) + ' días' : 'N/A', estadoTG, isDUSExport ? 'Promedio días desde factura hasta hoy (todos)' : 'Promedio de días entre factura y BL'],
        [],
        ['Total Pedidos', kpis.total],
        [isDUSExport ? 'Total Legalizados' : 'Total Procesados', kpis.totalProcesados],
        ['Total Pendientes', kpis.totalPendientes],
        ['% Cumplimiento', kpis.total > 0 ? Math.round((kpis.totalProcesados / kpis.total) * 100) + '%' : '0%'],
    ];
    const wsResumen = XLSX.utils.aoa_to_sheet(resumen);
    wsResumen['!cols'] = [{ wch: 22 }, { wch: 15 }, { wch: 14 }, { wch: 55 }];
    // Estilo título
    if (wsResumen['A1']) wsResumen['A1'].s = sTitle;
    // Estilo headers fila 4 (index 3)
    for (let c = 0; c < 4; c++) {
        const addr = XLSX.utils.encode_cell({ r: 3, c });
        if (wsResumen[addr]) wsResumen[addr].s = sHeader;
    }
    // Colorear columna ESTADO (C5:C9)
    const estadoMap = [estadoDemora, estadoSLA, estadoRiesgo, estadoCrit, estadoTG];
    estadoMap.forEach((txt, i) => {
        const addr = XLSX.utils.encode_cell({ r: 4 + i, c: 2 });
        if (wsResumen[addr]) wsResumen[addr].s = colorForEstado(txt);
    });
    XLSX.utils.book_append_sheet(wb, wsResumen, 'Resumen');

    // ═══ Hoja 2: Por Analista ═══
    const analHeaders = isDUSExport
        ? ['ANALISTA', 'CÓDIGO', 'GRUPO', 'TOTAL', 'LEGALIZADOS', 'A TIEMPO', 'FUERA PLAZO', 'PEND. FUERA SLA', '% SLA', 'DEMORA PROM.']
        : ['ANALISTA', 'CÓDIGO', 'GRUPO', 'TOTAL', 'PROCESADOS', '% CUMPL.', 'DEMORA PROM.', 'SLA ≤3d', 'T. GESTIÓN PROM.'];
    const analRows = [analHeaders];
    Object.entries(kpis.porAnalista)
        .sort((a, b) => b[1].total - a[1].total)
        .forEach(([code, data]) => {
            const name = resolveAnalystName(code);
            const grupo = resolveAnalystGrupo(code) || 'Sin grupo';
            if (isDUSExport) {
                // Contar sub-estatus por analista desde los results
                let aTiempo = 0, fueraPlazo = 0;
                results.filter(r => (r.RESPONSABLE || 'Sin Asignar') === code).forEach(r => {
                    const est = r.ESTATUS_FINAL || '';
                    if (est.includes('Legalizado') && est.includes('A Tiempo')) aTiempo++;
                    else if (est.includes('Legalizado') && est.includes('Fuera de Plazo')) fueraPlazo++;
                });
                const slaPct = (aTiempo + fueraPlazo) > 0 ? Math.round((aTiempo / (aTiempo + fueraPlazo)) * 100) : 0;
                const demProm = data.demoras.length > 0 ? (data.demoras.reduce((a, b) => a + b, 0) / data.demoras.length).toFixed(1) : '-';
                analRows.push([name, code, grupo, data.total, data.proc, aTiempo, fueraPlazo, data.fueraSLA || 0, slaPct + '%', demProm]);
            } else {
                const cumpl = data.total > 0 ? Math.round((data.proc / data.total) * 100) : 0;
                const demProm = data.demoras.length > 0 ? (data.demoras.reduce((a, b) => a + b, 0) / data.demoras.length).toFixed(1) : '-';
                const sla = data.demoras.length > 0 ? Math.round((data.demoras.filter(d => d <= 3).length / data.demoras.length) * 100) + '%' : '-';
                const tg = data.tgestiones.length > 0 ? (data.tgestiones.reduce((a, b) => a + b, 0) / data.tgestiones.length).toFixed(1) : '-';
                analRows.push([name, code, grupo, data.total, data.proc, cumpl + '%', demProm, sla, tg]);
            }
        });
    const wsAnal = XLSX.utils.aoa_to_sheet(analRows);
    wsAnal['!cols'] = isDUSExport
        ? [{ wch: 22 }, { wch: 14 }, { wch: 12 }, { wch: 8 }, { wch: 14 }, { wch: 12 }, { wch: 14 }, { wch: 16 }, { wch: 10 }, { wch: 14 }]
        : [{ wch: 22 }, { wch: 14 }, { wch: 12 }, { wch: 8 }, { wch: 12 }, { wch: 10 }, { wch: 14 }, { wch: 10 }, { wch: 16 }];
    styleHeaders(wsAnal, analHeaders.length);
    if (isDUSExport) {
        styleCumplCol(wsAnal, analRows.length - 1, 8); // % SLA
        styleDemoraCol(wsAnal, analRows.length - 1, 9); // DEMORA PROM.
    } else {
        styleCumplCol(wsAnal, analRows.length - 1, 5);
        styleDemoraCol(wsAnal, analRows.length - 1, 6);
    }
    XLSX.utils.book_append_sheet(wb, wsAnal, 'Por Analista');

    // ═══ Hoja 3: Por Mes / Período ═══
    const mesHeaders = isDUSExport
        ? ['PERÍODO', 'DEADLINE', 'TOTAL', 'LEGALIZADOS', 'A TIEMPO', 'FUERA PLAZO', 'PEND. FUERA SLA', '% SLA', 'DEMORA PROM.']
        : ['MES', 'TOTAL', 'PROCESADOS', '% CUMPL.', 'DEMORA PROM.', 'SLA ≤3d'];
    const mesRows = [mesHeaders];
    Object.entries(kpis.porMes)
        .sort((a, b) => a[0].localeCompare(b[0]))
        .forEach(([mes, data]) => {
            if (isDUSExport) {
                // Calcular deadline del período
                const [y, m] = mes !== 'Sin Fecha' ? mes.split('-').map(Number) : [0, 0];
                const dl = y > 0 ? new Date(y, m, 8).toLocaleDateString('es-CL') : '-';
                const periodoLabel = y > 0 ? `${mes}` : 'Sin Fecha';
                let aTiempo = 0, fueraPlazo = 0;
                results.filter(r => {
                    const fd = r.FECHA_FACTURA_DATE;
                    if (!fd || !(fd instanceof Date)) return mes === 'Sin Fecha';
                    return fd.toISOString().slice(0, 7) === mes;
                }).forEach(r => {
                    const est = r.ESTATUS_FINAL || '';
                    if (est.includes('Legalizado') && est.includes('A Tiempo')) aTiempo++;
                    else if (est.includes('Legalizado') && est.includes('Fuera de Plazo')) fueraPlazo++;
                });
                const slaPct = (aTiempo + fueraPlazo) > 0 ? Math.round((aTiempo / (aTiempo + fueraPlazo)) * 100) : 0;
                const demProm = data.demoras.length > 0 ? (data.demoras.reduce((a, b) => a + b, 0) / data.demoras.length).toFixed(1) : '-';
                mesRows.push([periodoLabel, dl, data.total, data.proc, aTiempo, fueraPlazo, data.fueraSLA || 0, slaPct + '%', demProm]);
            } else {
                const cumpl = data.total > 0 ? Math.round((data.proc / data.total) * 100) : 0;
                const demProm = data.demoras.length > 0 ? (data.demoras.reduce((a, b) => a + b, 0) / data.demoras.length).toFixed(1) : '-';
                const sla = data.demoras.length > 0 ? Math.round((data.demoras.filter(d => d <= 3).length / data.demoras.length) * 100) + '%' : '-';
                mesRows.push([mes, data.total, data.proc, cumpl + '%', demProm, sla]);
            }
        });
    const wsMes = XLSX.utils.aoa_to_sheet(mesRows);
    wsMes['!cols'] = isDUSExport
        ? [{ wch: 12 }, { wch: 14 }, { wch: 8 }, { wch: 14 }, { wch: 12 }, { wch: 14 }, { wch: 16 }, { wch: 10 }, { wch: 14 }]
        : [{ wch: 12 }, { wch: 8 }, { wch: 12 }, { wch: 10 }, { wch: 14 }, { wch: 10 }];
    styleHeaders(wsMes, mesHeaders.length);
    if (isDUSExport) {
        styleCumplCol(wsMes, mesRows.length - 1, 7); // % SLA
        styleDemoraCol(wsMes, mesRows.length - 1, 8); // DEMORA PROM.
    } else {
        styleCumplCol(wsMes, mesRows.length - 1, 3);
        styleDemoraCol(wsMes, mesRows.length - 1, 4);
    }
    XLSX.utils.book_append_sheet(wb, wsMes, isDUSExport ? 'Por Período' : 'Por Mes');

    // ═══ Hoja 4: Por Cliente ═══
    const cliRows = [['CLIENTE', 'TOTAL', 'PROCESADOS', '% CUMPL.', 'DEMORA PROM.']];
    Object.entries(kpis.porCliente)
        .sort((a, b) => b[1].total - a[1].total)
        .slice(0, 50)
        .forEach(([cli, data]) => {
            const cumpl = data.total > 0 ? Math.round((data.proc / data.total) * 100) : 0;
            const demProm = data.demoras.length > 0 ? (data.demoras.reduce((a, b) => a + b, 0) / data.demoras.length).toFixed(1) : '-';
            cliRows.push([cli, data.total, data.proc, cumpl + '%', demProm]);
        });
    const wsCli = XLSX.utils.aoa_to_sheet(cliRows);
    wsCli['!cols'] = [{ wch: 30 }, { wch: 8 }, { wch: 12 }, { wch: 10 }, { wch: 14 }];
    styleHeaders(wsCli, 5);
    styleCumplCol(wsCli, cliRows.length - 1, 3); // Col D = % CUMPL.
    styleDemoraCol(wsCli, cliRows.length - 1, 4); // Col E = DEMORA PROM.
    XLSX.utils.book_append_sheet(wb, wsCli, 'Por Cliente');

    // ═══ Hoja 5: Detalle Riesgo (>7 días, no gestionados) ═══
    const riesgoRows = [['PEDIDO', 'CLIENTE/CONSIG.', 'RESPONSABLE', 'DEMORA (DÍAS)', 'ESTADO', 'FECHA FACTURA']];
    const riesgoData = results
        .filter(r => {
            const d = parseInt(r.DEMORA) || 0;
            const est = r.ESTATUS_FINAL || '';
            return d > 7 && !isGestionado(est);
        })
        .sort((a, b) => (b.DEMORA || 0) - (a.DEMORA || 0));
    riesgoData.forEach(r => {
        const fFac = r.FECHA_FACTURA instanceof Date ? r.FECHA_FACTURA.toLocaleDateString() : (r.FECHA_FACTURA || '-');
        riesgoRows.push([
            r.PEDIDO_RAW || r.PEDIDO,
            r.CLIENTE || r.CONSIGNATARIO || '-',
            resolveAnalystName(r.RESPONSABLE || '-'),
            r.DEMORA || 0,
            r.ESTATUS_FINAL,
            fFac
        ]);
    });
    const wsRiesgo = XLSX.utils.aoa_to_sheet(riesgoRows);
    wsRiesgo['!cols'] = [{ wch: 16 }, { wch: 28 }, { wch: 22 }, { wch: 14 }, { wch: 24 }, { wch: 14 }];
    styleHeaders(wsRiesgo, 6);
    // Colorear demora: 1-3 verde, 4-7 amarillo, 8+ rojo
    styleDemoraCol(wsRiesgo, riesgoData.length, 3); // Col D = DEMORA
    XLSX.utils.book_append_sheet(wb, wsRiesgo, 'Detalle Riesgo');

    const fileName = `KPI_ExportDesk_${mod}_${new Date().toISOString().slice(0, 10)}.xlsx`;
    XLSX.writeFile(wb, fileName);
}

