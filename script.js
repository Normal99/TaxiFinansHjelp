// ─── CONFIG & STATE ──────────────────────────────────────────────
const DB_NAME = 'taxifinans-docs';
const STORE_NAME = 'documents';
let allData = []; // Combined data from all versions
let currentFilter = '';
let currentSort = 'none';
let currentLang = 'no';

// ─── INIT ────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
    initUpload();
    loadRecentDocuments();
});

// ─── INDEXEDDB HELPERS ───────────────────────────────────────────
function openDB() {
    return new Promise((resolve, reject) => {
        const request = indexedDB.open(DB_NAME, 1);
        
        request.onerror = () => reject(request.error);
        request.onsuccess = () => resolve(request.result);
        
        request.onupgradeneeded = (e) => {
            const db = e.target.result;
            
            if (!db.objectStoreNames.contains(STORE_NAME)) {
                const store = db.createObjectStore(STORE_NAME, { keyPath: 'id', autoIncrement: true });
                store.createIndex('name', 'name', { unique: false });
                store.createIndex('timestamp', 'timestamp', { unique: false });
            }
        };
    });
}

function saveDocument(name, data) {
    return openDB().then(db => {
        return new Promise((resolve, reject) => {
            const tx = db.transaction(STORE_NAME, 'readwrite');
            const store = tx.objectStore(STORE_NAME);
            
            // Sjekk om samme dokument finnes (basert på navn + data hash)
            const checkRequest = store.index('name').getAll(name);
            
            checkRequest.onsuccess = () => {
                const existingDocs = checkRequest.result;
                
                // Sjekk om data allerede eksisterer
                let exists = false;
                for (const doc of existingDocs) {
                    if (JSON.stringify(doc.data[0]) === JSON.stringify(data[0])) {
                        exists = true;
                        break;
                    }
                }
                
                if (!exists) {
                    const newDoc = {
                        name: name,
                        timestamp: Date.now(),
                        rowCount: data.length,
                        data: data
                    };
                    
                    store.add(newDoc);
                }
                
                resolve();
            };
            
            checkRequest.onerror = () => reject(checkRequest.error);
        });
    }).catch(err => {
        console.error('Lagringsfeil:', err);
    });
}

function loadRecentDocuments() {
    openDB().then(db => {
        const tx = db.transaction(STORE_NAME, 'readonly');
        const store = tx.objectStore(STORE_NAME);
        const request = store.getAll();
        
        request.onsuccess = () => {
            renderRecentList(request.result);
        };
        
        request.onerror = () => {
            console.error('Feil ved lasting:', request.error);
        };
    }).catch(err => {
        console.error('DB feil:', err);
    });
}

function deleteDocument(id) {
    return openDB().then(db => {
        return new Promise((resolve, reject) => {
            const tx = db.transaction(STORE_NAME, 'readwrite');
            const store = tx.objectStore(STORE_NAME);
            store.delete(id);
            
            tx.oncomplete = () => resolve();
            tx.onerror = () => reject(tx.error);
        });
    }).catch(err => {
        console.error('Slettingsfeil:', err);
    });
}

// ─── UPLOAD HANDLING ─────────────────────────────────────────────
function initUpload() {
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');

    // Drag & Drop
    ['dragenter', 'dragover'].forEach(evt => {
        dropZone.addEventListener(evt, e => {
            e.preventDefault();
            dropZone.style.borderColor = '#F7C520';
            dropZone.style.background = 'rgba(247,197,32,0.04)';
        });
    });

    ['dragleave', 'drop'].forEach(evt => {
        dropZone.addEventListener(evt, e => {
            e.preventDefault();
            dropZone.style.borderColor = '';
            dropZone.style.background = '';
        });
    });

    dropZone.addEventListener('drop', e => {
        const files = e.dataTransfer.files;
        if (files.length) handleFile(files[0]);
    });

    fileInput.addEventListener('change', () => {
        if (fileInput.files.length) handleFile(fileInput.files[0]);
    });
}

function handleFile(file) {
    const loading = document.getElementById('loading');
    const errorContainer = document.getElementById('errorContainer');
    
    // Show loading, hide previous errors
    loading.style.display = 'block';
    errorContainer.innerHTML = '';

    if (!file.name.match(/\.xlsx|\.xls$/i)) {
        showError('Ugyldig filformat. Bruk .xlsx eller .xls');
        return;
    }

    const reader = new FileReader();
    reader.onload = e => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            if (workbook.SheetNames.length === 0) throw new Error('Ingen ark i filen');
            
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(sheet, { defval: '' });

            if (!jsonData.length) throw new Error('Filen er tom');

            // Auto-detect sjåførkolonne
            const headers = Object.keys(jsonData[0]).map(h => h.toLowerCase());
            let driverCol = headers.find(h => 
                h.includes('sjåfør') || h.includes('fører') || h.includes('navn')
            ) || headers.find(h => 
                h.includes('driver') || h.includes('name') || h.includes('chauffeur')
            );

            if (!driverCol) {
                // Fallback: første kolonne med tekst
                driverCol = Object.keys(jsonData[0])[0];
            } else {
                // Finn originalt navn (case-insensitive match)
                const originalKey = Object.keys(jsonData[0]).find(k => 
                    k.toLowerCase() === driverCol
                );
                if (originalKey) driverCol = originalKey;
            }

            // Normaliser data
            allData = jsonData.map(row => {
                const normalized = {};
                for (const [key, value] of Object.entries(row)) {
                    const cleanKey = key.trim();
                    let cleanValue = value !== null && value !== undefined ? String(value).trim() : '';
                    
                    // Fjern spesialtegn fra tall
                    if (!isNaN(parseFloat(cleanValue)) && cleanValue !== '') {
                        cleanValue = cleanValue.replace(/[^0-9.,\-+\/]/g, '');
                        cleanValue = cleanValue.includes(',') 
                            ? parseFloat(cleanValue.replace(',', '.')) || 0 
                            : parseFloat(cleanValue) || 0;
                    }
                    
                    normalized[cleanKey] = cleanValue;
                }
                return normalized;
            });

            // Lagre dokumentet
            const docName = file.name.replace(/\.xlsx|\.xls$/i, '');
            
            saveDocument(docName, allData).then(() => {
                console.log('Dokument lagret i IndexedDB');
                
                // Prosess data og vis resultater
                processAllData();
            });

        } catch (err) {
            console.error('Feil ved lesing:', err);
            showError(err.message || 'Kunne ikke lese filen');
        } finally {
            loading.style.display = 'none';
        }
    };
    reader.readAsArrayBuffer(file);
}

// ─── OPEN DOCUMENT FROM RECENT LIST ──────────────────────────────
function openDocument(doc) {
    allData = doc.data;
    
    // Update UI with document name in page title
    document.title = `${doc.name} - Sjåfør Analyse`;
    
    processAllData();
}

// ─── PROCESS & DISPLAY DATA ──────────────────────────────────────
function processAllData() {
    if (!allData.length) return;

    // Skjul upload og recent docs, vis resultater
    const uploadSection = document.getElementById('uploadSection');
    const recentListContainer = document.getElementById('recentListContainer');
    
    if (uploadSection) uploadSection.style.display = 'none';
    if (recentListContainer) recentListContainer.style.display = 'none';

    // Skjul loading, vis results
    document.getElementById('loading').style.display = 'none';
    const resultsContainer = document.getElementById('resultsSection');
    resultsContainer.style.display = 'block';
    
    setTimeout(() => {
        resultsContainer.classList.add('visible');
    }, 10);

    // Stats
    updateStats();

    // Driver cards
    renderDriverCards();

    // Table
    renderTable();

    // Populate filter dropdown
    populateFilterDropdown();
}

function updateStats() {
    const statsGrid = document.getElementById('statsGrid');
    if (!statsGrid) return;

    const totalRows = allData.length;

    // Detect columns by exact key name
    let amountCol, kmsCol, kmsOccupiedCandidates = [], tripsCol, cashCol, hoursCol, billedCol;
    for (const key of Object.keys(allData[0])) {
        const k = key.toLowerCase().trim();
        if (!amountCol    && (k.includes('lønn') || k.includes('grunnlag')))           amountCol  = key;
        if (!kmsCol       && (k === 'km total' || k === 'km totalt'))                   kmsCol     = key;
        if (k === 'km opptatt' || k === 'km besatt')                                    kmsOccupiedCandidates.push(key);
        if (!tripsCol     && k === 'antall turer')                                      tripsCol   = key;
        if (!cashCol      && (k === 'faktisk kont.' || k === 'faktisk kont'))           cashCol    = key;
        if (!hoursCol     && k === 'effektiv timer')                                    hoursCol   = key;
        if (!billedCol    && k.includes('innkjørt total lav'))                          billedCol  = key;
    }
    // Pick the km-occupied column that actually has non-zero data
    const kmsOccupiedCol = kmsOccupiedCandidates.find(col =>
        allData.some(row => parseFloat(row[col]) > 0)
    ) || null;

    const sum = col => col ? allData.reduce((acc, row) => acc + (parseFloat(row[col]) || 0), 0) : 0;

    const totalAmount      = sum(amountCol);
    const totalKms         = sum(kmsCol);
    const totalKmsOccupied = sum(kmsOccupiedCol);
    const totalTrips       = sum(tripsCol);
    const totalCash        = sum(cashCol);
    const totalHours       = sum(hoursCol);
    const totalBilled      = sum(billedCol);

    const uniqueDrivers = getUniqueDriverCount();

    statsGrid.innerHTML = `
        <div class="stat-tile">
            <span class="stat-tile-label" data-i18n="totalRows">Totalt skift</span>
            <span class="stat-tile-value accent">${totalRows.toLocaleString()}</span>
        </div>
        <div class="stat-tile">
            <span class="stat-tile-label" data-i18n="uniqueDrivers">Sjåfører</span>
            <span class="stat-tile-value accent">${uniqueDrivers}</span>
        </div>
        ${amountCol ? `
        <div class="stat-tile">
            <span class="stat-tile-label" data-i18n="totalAmount">Lønnsgrunnlag</span>
            <span class="stat-tile-value accent">${formatCurrency(totalAmount)}</span>
        </div>` : ''}
        ${cashCol ? `
        <div class="stat-tile">
            <span class="stat-tile-label">Kontant</span>
            <span class="stat-tile-value accent">${formatCurrency(totalCash)}</span>
        </div>` : ''}
        ${billedCol ? `
        <div class="stat-tile">
            <span class="stat-tile-label">Innkjørt total</span>
            <span class="stat-tile-value accent">${formatCurrency(totalBilled)}</span>
        </div>` : ''}
        ${tripsCol ? `
        <div class="stat-tile">
            <span class="stat-tile-label">Antall turer</span>
            <span class="stat-tile-value accent">${formatNumber(totalTrips)}</span>
        </div>` : ''}
        ${kmsCol ? `
        <div class="stat-tile">
            <span class="stat-tile-label" data-i18n="totalKms">KM totalt</span>
            <span class="stat-tile-value accent">${formatNumber(totalKms)} km</span>
        </div>` : ''}
        ${kmsOccupiedCol ? `
        <div class="stat-tile">
            <span class="stat-tile-label">KM m/passasjer</span>
            <span class="stat-tile-value accent">${formatNumber(totalKmsOccupied)} km</span>
        </div>` : ''}
        ${hoursCol ? `
        <div class="stat-tile">
            <span class="stat-tile-label">Effektive timer</span>
            <span class="stat-tile-value accent">${formatNumber(totalHours)} t</span>
        </div>` : ''}
    `;
}

function formatCurrency(amount) {
    return new Intl.NumberFormat('nb-NO', { 
        style: 'currency', 
        currency: 'NOK' 
    }).format(amount);
}

function formatNumber(num) {
    return new Intl.NumberFormat('nb-NO').format(Math.round(num));
}

function getUniqueDriverCount() {
    const drivers = new Set();
    for (const row of allData) {
        // Finn sjåførkolonnen
        for (const key of Object.keys(row)) {
            if (key.toLowerCase().includes('sjåfør') || 
                key.toLowerCase().includes('fører') || 
                key.toLowerCase().includes('navn')) {
                const driver = String(row[key]).trim();
                if (driver) drivers.add(driver);
                break;
            }
        }
    }
    return drivers.size.toLocaleString();
}

// ─── DRIVER CARDS ────────────────────────────────────────────────
function renderDriverCards() {
    const grid = document.getElementById('summaryGrid');
    if (!grid) return;

    // Hent driver kolonne
    let driverCol = null;
    for (const key of Object.keys(allData[0])) {
        if (key.toLowerCase().includes('sjåfør') || 
            key.toLowerCase().includes('fører') || 
            key.toLowerCase().includes('navn')) {
            driverCol = key;
            break;
        }
    }

    if (!driverCol) return;

    // Aggreger per sjåfør
    const drivers = {};
    
    for (const row of allData) {
        const name = String(row[driverCol]).trim();
        if (!name) continue;

        if (!drivers[name]) {
            drivers[name] = { name, amount: 0, kms: 0, kmsOccupied: 0, trips: 0, cash: 0, hours: 0, billed: 0, rows: [] };
        }

        drivers[name].rows.push(row);

        for (const key of Object.keys(row)) {
            const keyLower = key.toLowerCase().trim();
            const val = parseFloat(row[key]);
            if (!isNaN(val) && val !== 0) {
                if (keyLower.includes('lønn') || keyLower.includes('lønns') || keyLower.includes('grunnlag')) {
                    drivers[name].amount += val;
                } else if (keyLower === 'km total' || keyLower === 'km totalt') {
                    drivers[name].kms += val;
                } else if (keyLower === 'km opptatt' || keyLower === 'km besatt') {
                    // Only add if this row actually has a value (avoids zero-filled Km besatt winning)
                    if (val > 0) drivers[name].kmsOccupied += val;
                } else if (keyLower === 'antall turer') {
                    drivers[name].trips += val;
                } else if (keyLower === 'faktisk kont.' || keyLower === 'faktisk kont') {
                    drivers[name].cash += val;
                } else if (keyLower === 'effektiv timer') {
                    drivers[name].hours += val;
                } else if (keyLower.includes('innkjørt total lav') || keyLower === 'innkjørt total lav sats') {
                    drivers[name].billed += val;
                }
            }
        }
    }

    // Sorter på lønngrunnlag
    const sorted = Object.values(drivers).sort((a, b) => b.amount - a.amount);

    grid.innerHTML = '';
    
    for (const driver of sorted.slice(0, 12)) {
        const card = document.createElement('div');
        card.className = 'driver-card';
        
        card.innerHTML = `
            <div class="driver-card-header">
                <div class="driver-avatar">
                    <svg viewBox="0 0 24 24" fill="currentColor" xmlns="http://www.w3.org/2000/svg">
                        <path d="M12 11.5a4 4 0 1 0 0-8 4 4 0 0 0 0 8zm0 2c-4 0-7 2-7 4v1h14v-1c0-2-3-4-7-4z"/>
                    </svg>
                </div>
                <div class="driver-name-block">
                    <div class="driver-name">${escapeHtml(driver.name)}</div>
                    <div class="driver-hint" data-i18n="clickToFilter">Klikk for å filtrere</div>
                </div>
            </div>
            <div class="card-metrics">
                <div class="metric">
                    <div class="metric-label" data-i18n="rows">Skift</div>
                    <div class="metric-val">${driver.rows.length}</div>
                </div>
                ${driver.amount > 0 ? `
                <div class="metric">
                    <div class="metric-label" data-i18n="amount">Lønnsgr.</div>
                    <div class="metric-val hi">${formatCurrency(driver.amount)}</div>
                </div>` : ''}
                ${driver.cash > 0 ? `
                <div class="metric">
                    <div class="metric-label">Kontant</div>
                    <div class="metric-val" style="color:var(--green)">${formatCurrency(driver.cash)}</div>
                </div>` : ''}
                ${driver.billed > 0 ? `
                <div class="metric">
                    <div class="metric-label">Innkjørt</div>
                    <div class="metric-val">${formatCurrency(driver.billed)}</div>
                </div>` : ''}
                ${driver.trips > 0 ? `
                <div class="metric">
                    <div class="metric-label">Turer</div>
                    <div class="metric-val">${formatNumber(driver.trips)}</div>
                </div>` : ''}
                ${driver.kms > 0 ? `
                <div class="metric">
                    <div class="metric-label" data-i18n="kms">KM</div>
                    <div class="metric-val" style="color:var(--green)">${formatNumber(driver.kms)}</div>
                </div>` : ''}
                ${driver.kmsOccupied > 0 ? `
                <div class="metric">
                    <div class="metric-label">KM m/pass.</div>
                    <div class="metric-val">${formatNumber(driver.kmsOccupied)}</div>
                </div>` : ''}
                ${driver.hours > 0 ? `
                <div class="metric">
                    <div class="metric-label">Timer</div>
                    <div class="metric-val">${formatNumber(driver.hours)} t</div>
                </div>` : ''}
            </div>
        `;

        card.addEventListener('click', () => {
            currentFilter = driver.name;
            document.getElementById('driverFilter').value = driver.name;
            applyFilters();
            
            // Scroll til tabell
            setTimeout(() => {
                document.querySelector('.table-card')?.scrollIntoView({ behavior: 'smooth' });
            }, 100);
        });

        grid.appendChild(card);
    }
}

// ─── TABLE ───────────────────────────────────────────────────────
function populateFilterDropdown() {
    const select = document.getElementById('driverFilter');
    if (!select) return;

    // Hent driver kolonne
    let driverCol = null;
    for (const key of Object.keys(allData[0])) {
        if (key.toLowerCase().includes('sjåfør') || 
            key.toLowerCase().includes('fører') || 
            key.toLowerCase().includes('navn')) {
            driverCol = key;
            break;
        }
    }

    if (!driverCol) return;

    // Hent unike sjåfører
    const drivers = [...new Set(allData.map(row => String(row[driverCol]).trim()).filter(Boolean))];
    
    select.innerHTML = '<option value="" data-i18n="allDrivers">Alle sjåfører</option>';
    
    for (const driver of drivers.sort()) {
        const opt = document.createElement('option');
        opt.value = driver;
        opt.textContent = driver;
        select.appendChild(opt);
    }

    // Event listener for filter
    select.addEventListener('change', () => {
        currentFilter = select.value;
        applyFilters();
    });
}

function renderTable() {
    const table = document.getElementById('dataTable');
    const thead = document.getElementById('tableHead');
    const tbody = document.getElementById('tableBody');
    
    if (!table || !thead || !tbody) return;

    // Skjul tabell mens vi bygger
    table.style.display = 'none';

    applyFilters();
}

function applyFilters() {
    let dataToDisplay = [...allData];
    const driverCol = findDriverColumn();

    if (currentFilter) {
        dataToDisplay = dataToDisplay.filter(row => 
            String(row[driverCol]).trim() === currentFilter
        );
    }

    // Sortering
    if (currentSort !== 'none') {
        const [col, dir] = currentSort.split('_');
        dataToDisplay.sort((a, b) => {
            let valA = parseFloat(a[col]);
            let valB = parseFloat(b[col]);
            
            if (isNaN(valA)) valA = 0;
            if (isNaN(valB)) valB = 0;

            return dir === 'desc' ? valB - valA : valA - valB;
        });
    }

    renderTableContent(dataToDisplay, driverCol);
}

function findDriverColumn() {
    for (const key of Object.keys(allData[0])) {
        if (key.toLowerCase().includes('sjåfør') || 
            key.toLowerCase().includes('fører') || 
            key.toLowerCase().includes('navn')) {
            return key;
        }
    }
    return Object.keys(allData[0])[0]; // Fallback
}

function renderTableContent(data, driverCol) {
    const thead = document.getElementById('tableHead');
    const tbody = document.getElementById('tableBody');
    
    if (!thead || !tbody) return;

    // Hoder
    const headers = Object.keys(data[0] || {});
    thead.innerHTML = '<tr>' + headers.map(h => `<th>${escapeHtml(h)}</th>`).join('') + '</tr>';

    // Rader
    tbody.innerHTML = '';

    for (const row of data) {
        const tr = document.createElement('tr');
        
        for (const key of headers) {
            const td = document.createElement('td');
            
            if (key === driverCol) {
                td.className = 'driver-highlight';
                td.textContent = String(row[key]);
            } else {
                let val = row[key];
                
                // Formater tall
                if (!isNaN(parseFloat(val)) && String(val).trim() !== '') {
                    const numVal = parseFloat(val);
                    td.textContent = new Intl.NumberFormat('nb-NO', { 
                        maximumFractionDigits: 2 
                    }).format(numVal);
                    
                    // Fargelegg store tall
                    if (Math.abs(numVal) > 10000) {
                        td.style.color = 'var(--accent)';
                    }
                } else {
                    td.textContent = val;
                }
            }
            
            tr.appendChild(td);
        }
        
        tbody.appendChild(tr);
    }

    // Totalrad
    const totalRow = document.createElement('tr');
    totalRow.className = 'total-row';
    
    let amountTotal = 0, kmsTotal = 0;
    for (const key of headers) {
        if (key.toLowerCase().includes('lønn') || key.toLowerCase().includes('lønns')) {
            const sum = data.reduce((acc, row) => acc + (parseFloat(row[key]) || 0), 0);
            amountTotal += sum;
        } else if (key.toLowerCase().includes('km') || key.toLowerCase().includes('kilometer')) {
            const sum = data.reduce((acc, row) => acc + (parseFloat(row[key]) || 0), 0);
            kmsTotal += sum;
        }
    }

    totalRow.innerHTML = headers.map((h, i) => 
        i === 0 ? `<td><strong>Totalt</strong></td>` : '<td></td>'
    ).join('');
    
    tbody.appendChild(totalRow);

    // Vis tabell
    document.getElementById('dataTable').style.display = 'table';
}

// ─── SORTING ─────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
    const sortSelect = document.getElementById('sortSelect');
    if (sortSelect) {
        sortSelect.addEventListener('change', () => {
            currentSort = sortSelect.value;
            applyFilters();
        });
    }
});

// ─── RECENT DOCUMENTS LIST RENDERER ──────────────────────────────
function renderRecentList(docs) {
    const container = document.getElementById('recentListContainer');
    
    if (!container) return;
    
    // Tøm eksisterende innhold (hold h2, fjern resten)
    while (container.firstChild && container.firstChild.tagName !== 'H2') {
        container.removeChild(container.firstChild);
    }

    const noRecentHint = document.getElementById('noRecentHint');
    
    if (!docs || docs.length === 0) {
        if (noRecentHint) {
            noRecentHint.style.display = 'block';
        } else {
            const hint = document.createElement('p');
            hint.id = 'noRecentHint';
            hint.textContent = 'Ingen lagrede dokumenter';
            hint.style.cssText = 'font-size:0.75rem;color:var(--text-dim);margin-top:10px;';
            container.appendChild(hint);
        }
        return;
    }

    if (noRecentHint) noRecentHint.style.display = 'none';

    // Grupper etter dato (mnd/år)
    const grouped = {};
    
    docs.forEach(doc => {
        const date = new Date(doc.timestamp);
        const monthYear = date.toLocaleDateString('nb-NO', { 
            year: 'numeric', 
            month: 'long' 
        });
        
        if (!grouped[monthYear]) {
            grouped[monthYear] = [];
        }
        
        grouped[monthYear].push(doc);
    });

    // Sorter måneder nedad
    const sortedMonths = Object.keys(grouped).sort((a, b) => {
        const dateA = new Date(a + ' 1');
        const dateB = new Date(b + ' 1');
        return dateB - dateA;
    });

    // Bygg UI som liste
    sortedMonths.forEach(monthYear => {
        const groupDiv = document.createElement('div');
        groupDiv.className = 'recent-group';
        
        const label = document.createElement('label');
        label.textContent = monthYear;
        groupDiv.appendChild(label);

        grouped[monthYear].forEach(doc => {
            const docItem = document.createElement('div');
            docItem.className = 'doc-item';
            
            // Formatert dato
            const dateObj = new Date(doc.timestamp);
            const formattedDate = dateObj.toLocaleDateString('nb-NO', {
                day: '2-digit',
                month: 'short',
                year: 'numeric'
            });

            docItem.innerHTML = `
                <div class="doc-item-header">
                    <span class="doc-name">${escapeHtml(doc.name)}</span>
                    <span class="doc-date">${formattedDate}</span>
                </div>
                <div class="doc-meta">
                    <span class="doc-rows">${doc.rowCount} rader</span>
                    ${doc.data[0] ? `<span class="doc-cols">${Object.keys(doc.data[0]).length} kolonner</span>` : ''}
                </div>
            `;

            // Klikk for å åpne dokument
            docItem.addEventListener('click', () => {
                openDocument(doc);
            });

            groupDiv.appendChild(docItem);
        });

        container.appendChild(groupDiv);
    });
}

// ─── UTILS ───────────────────────────────────────────────────────
function showError(msg) {
    const container = document.getElementById('errorContainer');
    if (!container) return;

    container.innerHTML = `
        <div class="error-box">
            <h3 data-i18n="errorTitle">Feil</h3>
            <p>${escapeHtml(msg)}</p>
        </div>
    `;
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// ─── SHOW UPLOAD SECTION (when starting fresh) ──────────────────
function showUploadSection() {
    // Hide results
    const resultsContainer = document.getElementById('resultsSection');
    if (resultsContainer) resultsContainer.style.display = 'none';
    
    // Show upload section
    const uploadSection = document.getElementById('uploadSection');
    if (uploadSection) uploadSection.style.display = 'flex';
    
    // Hide new file button
    const newFileBtn = document.getElementById('newFileBtn');
    if (newFileBtn) newFileBtn.style.display = 'none';
}

// ─── LANGUAGE SWITCHING ──────────────────────────────────────────
function setLang(lang) {
    currentLang = lang;
    
    // Update buttons
    document.getElementById('langNO').classList.toggle('active', lang === 'no');
    document.getElementById('langTR').classList.toggle('active', lang === 'tr');
    
    // Translate all elements with data-i18n attribute
    const translations = {
        no: {
            title: 'Sjåfør',
            recentDocsTitle: 'Nylige dokumenter',
            uploadTitle: 'Last opp<br>Excel-fil',
            uploadSub: 'Støtter .xlsx og .xls — gjenkjenner sjåførkolonne automatisk',
            chooseFile: 'Velg fil',
            dragDrop: 'Klikk eller dra og slipp Excel-filen her',
            totalRows: 'Totalt antall rader',
            totalAmount: 'Totalt lønngrunnlag',
            totalKms: 'Totalt KM',
            uniqueDrivers: 'Unike sjåfører',
            cardsTitle: 'Sjåføroversikt',
            cardsSub: 'Klikk på et kort for å filtrere tabellen',
            tableTitle: 'Skift',
            filterBy: 'FILTRER:',
            allDrivers: 'Alle sjåfører',
            sortBy: 'SORTER:',
            sortNone: 'Ingen',
            sortWages: 'Lønngrunnlag ↓',
            sortKms: 'KM ↓',
            amount: 'Lønngrunnlag',
            kms: 'KM',
            rows: 'Rader',
            clickToFilter: 'Klikk for å filtrere',
            loading: 'Laster data...',
            errorTitle: 'Feil'
        },
        tr: {
            title: 'Şoför',
            recentDocsTitle: 'Son Belgeler',
            uploadTitle: 'Excel<br>Fili Yükle',
            uploadSub: '.xlsx ve .xls desteklenir — şoför sütunu otomatik algılanır',
            chooseFile: 'Dosya Seç',
            dragDrop: 'Excel dosyasını buraya sürükleyin veya tıklayın',
            totalRows: 'Toplam Satır Sayısı',
            totalAmount: 'Toplam Maaş Tabanı',
            totalKms: 'Toplam KM',
            uniqueDrivers: 'Benzersiz Şoförler',
            cardsTitle: 'Şoför Özeti',
            cardsSub: 'Tabloyu filtrelemek için karta tıklayın',
            tableTitle: 'Vardiya',
            filterBy: 'FİLTRELE:',
            allDrivers: 'Tüm Şoförler',
            sortBy: 'SIRALA:',
            sortNone: 'Yok',
            sortWages: 'Maaş Tabanı ↓',
            sortKms: 'KM ↓',
            amount: 'Maaş Tabanı',
            kms: 'KM',
            rows: 'Satırlar',
            clickToFilter: 'Filtrelemek için tıklayın',
            loading: 'Veriler yükleniyor...',
            errorTitle: 'Hata'
        }
    };

    const t = translations[lang];
    document.querySelectorAll('[data-i18n]').forEach(el => {
        const key = el.getAttribute('data-i18n');
        if (t[key]) {
            // Preserve HTML in title tags
            if (key === 'title' || key === 'uploadTitle') {
                el.innerHTML = t[key];
            } else {
                el.textContent = t[key];
            }
        }
    });
}

// ─── DEBUG CONSOLE (optional, hidden by default) ──────────────────
const debugConsole = document.getElementById('debugConsole');
if (debugConsole) {
    window.debugLog = (...args) => {
        if (!debugConsole) return;
        const time = new Date().toLocaleTimeString();
        debugConsole.textContent += `[${time}] ` + args.join(' ') + '\n';
        debugConsole.scrollTop = debugConsole.scrollHeight;
    };
}
