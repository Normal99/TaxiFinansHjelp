// ═══════════════════════════════════════════════════════════════
// TAXI FINANS — SJÅFØR ANALYZER  v2.0
// ═══════════════════════════════════════════════════════════════

// ─── CONSTANTS ───────────────────────────────────────────────────
const DB_NAME    = 'taxifinans-docs';
const STORE_NAME = 'documents';

const MONTHS_NO    = ['Januar','Februar','Mars','April','Mai','Juni','Juli','August','September','Oktober','November','Desember'];
const MONTHS_TR    = ['Ocak','Şubat','Mart','Nisan','Mayıs','Haziran','Temmuz','Ağustos','Eylül','Ekim','Kasım','Aralık'];
const MONTHS_AB_NO = ['Jan','Feb','Mar','Apr','Mai','Jun','Jul','Aug','Sep','Okt','Nov','Des'];
const MONTHS_AB_TR = ['Oca','Şub','Mar','Nis','May','Haz','Tem','Ağu','Eyl','Eki','Kas','Ara'];

// ─── STATE ───────────────────────────────────────────────────────
let allData            = [];    // All rows, tagged with _month/_year/_fileId
let loadedFiles        = [];    // [{ id, name, month, year, rowCount, data }]
let currentFilter      = '';
let currentSort        = 'none';
let currentLang        = 'no';
let currentMonthFilter = null;  // { month: 0-11, year: YYYY } | null
let selectedYear       = null;
let fileIdCounter      = 0;

// ─── INIT ────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
    initUpload();
    initAddFileInput();
    loadRecentDocuments();

    const sortSelect = document.getElementById('sortSelect');
    if (sortSelect) sortSelect.addEventListener('change', () => {
        currentSort = sortSelect.value;
        applyFilters();
    });
});

// ─── i18n ────────────────────────────────────────────────────────
function T(key) {
    const t = {
        no: {
            noRecentDocs:'Ingen lagrede dokumenter', docRows:'rader', docCols:'kolonner',
            statCash:'Kontant', statBilled:'Innkjørt total', statTrips:'Antall turer',
            statKmsOcc:'KM m/passasjer', statHours:'Effektive timer',
            metricCash:'Kontant', metricBilled:'Innkjørt', metricTrips:'Turer',
            metricKmsOcc:'KM m/pass.', metricHours:'Timer',
            tableTotal:'Totalt', newFile:'+ Ny fil', chartEmpty:'Ingen data å vise',
            clickToFilter:'Klikk for å filtrere', rows:'Skift', amount:'Lønnsgr.', kms:'KM',
            docTitleSuffix:'Sjåfør Analyse', backBtn:'Tilbake',
            allMonths:'Alle måneder', yearView:'Årsvisning',
            loadedFilesTitle:'Lastede filer', unknownDate:'Ukjent dato',
            shifts:'skift', addFileBtn:'+ Legg til fil', newAnalysis:'Ny analyse',
        },
        tr: {
            noRecentDocs:'Henüz belge kaydedilmedi', docRows:'satır', docCols:'sütun',
            statCash:'Nakit Tahsilat', statBilled:'Toplam Hasılat', statTrips:'Sefer Sayısı',
            statKmsOcc:'Yolculu KM', statHours:'Aktif Çalışma Saati',
            metricCash:'Nakit', metricBilled:'Hasılat', metricTrips:'Sefer',
            metricKmsOcc:'Yolculu KM', metricHours:'Aktif Saat',
            tableTotal:'Toplam', newFile:'+ Yeni Dosya', chartEmpty:'Gösterilecek veri yok',
            clickToFilter:'Filtrelemek için tıklayın', rows:'Vardiya', amount:'Kazanç', kms:'KM',
            docTitleSuffix:'Şoför Analizi', backBtn:'Geri',
            allMonths:'Tüm Aylar', yearView:'Yıllık Görünüm',
            loadedFilesTitle:'Yüklenen Dosyalar', unknownDate:'Tarih Bilinmiyor',
            shifts:'vardiya', addFileBtn:'+ Dosya Ekle', newAnalysis:'Yeni Analiz',
        }
    };
    const lang = currentLang || 'no';
    return (t[lang]?.[key]) || (t.no[key]) || key;
}

// ─── INDEXEDDB ───────────────────────────────────────────────────
function openDB() {
    return new Promise((resolve, reject) => {
        const req = indexedDB.open(DB_NAME, 1);
        req.onerror   = () => reject(req.error);
        req.onsuccess = () => resolve(req.result);
        req.onupgradeneeded = e => {
            const db = e.target.result;
            if (!db.objectStoreNames.contains(STORE_NAME)) {
                const store = db.createObjectStore(STORE_NAME, { keyPath: 'id', autoIncrement: true });
                store.createIndex('name',      'name',      { unique: false });
                store.createIndex('timestamp', 'timestamp', { unique: false });
            }
        };
    });
}

function saveDocument(name, data) {
    return openDB().then(db => new Promise((resolve, reject) => {
        const tx    = db.transaction(STORE_NAME, 'readwrite');
        const store = tx.objectStore(STORE_NAME);
        const check = store.index('name').getAll(name);
        check.onsuccess = () => {
            const exists = check.result.some(d => JSON.stringify(d.data[0]) === JSON.stringify(data[0]));
            if (!exists) store.add({ name, timestamp: Date.now(), rowCount: data.length, data });
            resolve();
        };
        check.onerror = () => reject(check.error);
    })).catch(err => console.error('Lagringsfeil:', err));
}

function loadRecentDocuments() {
    openDB().then(db => {
        const req = db.transaction(STORE_NAME, 'readonly').objectStore(STORE_NAME).getAll();
        req.onsuccess = () => renderRecentList(req.result);
        req.onerror   = () => console.error('Feil ved lasting:', req.error);
    }).catch(err => console.error('DB feil:', err));
}

function deleteDocument(id) {
    return openDB().then(db => new Promise((resolve, reject) => {
        const tx = db.transaction(STORE_NAME, 'readwrite');
        tx.objectStore(STORE_NAME).delete(id);
        tx.oncomplete = () => resolve();
        tx.onerror    = () => reject(tx.error);
    })).catch(err => console.error('Slettingsfeil:', err));
}

// ─── MONTH DETECTION ─────────────────────────────────────────────
function detectMonthYear(rawRows, filename) {
    const keys = Object.keys(rawRows[0] || {});
    const dateKeys = keys.filter(k => {
        const kl = k.toLowerCase();
        return kl.includes('dato') || kl.includes('date') || kl.includes('start');
    });

    for (const key of dateKeys) {
        for (const row of rawRows.slice(0, 30)) {
            const val = row[key];
            let date = null;

            if (val instanceof Date && !isNaN(val.getTime())) {
                date = val;
            } else if (typeof val === 'string' && val.trim()) {
                const s = val.trim();
                // DD.MM.YYYY  or  DD/MM/YYYY  or  D.M.YYYY
                const dmy = s.match(/^(\d{1,2})[.\/\-](\d{1,2})[.\/\-](\d{4})$/);
                if (dmy) {
                    date = new Date(+dmy[3], +dmy[2] - 1, +dmy[1]);
                } else {
                    // Try ISO or locale string
                    const d = new Date(s);
                    if (!isNaN(d.getTime())) date = d;
                }
            } else if (typeof val === 'number' && val > 40000) {
                // Excel serial date
                date = new Date((val - 25569) * 86400 * 1000);
            }

            if (date && !isNaN(date.getTime()) && date.getFullYear() > 2000 && date.getFullYear() < 2100) {
                return { month: date.getMonth(), year: date.getFullYear() };
            }
        }
    }
    return detectMonthFromFilename(filename);
}

function detectMonthFromFilename(filename) {
    const f = filename.toLowerCase().replace(/\.xlsx?$/i, '');

    const NAMES_NO = ['januar','februar','mars','april','mai','juni','juli','august','september','oktober','november','desember'];
    const NAMES_EN = ['january','february','march','april','may','june','july','august','september','october','november','december'];
    const ABBR     = ['jan','feb','mar','apr','mai','jun','jul','aug','sep','okt','nov','des'];

    let month = null;
    for (let i = 0; i < 12; i++) {
        if (f.includes(NAMES_NO[i]) || f.includes(NAMES_EN[i])) { month = i; break; }
    }
    if (month === null) {
        for (let i = 0; i < 12; i++) {
            if (new RegExp('\\b' + ABBR[i] + '\\b').test(f)) { month = i; break; }
        }
    }

    // Year
    const yearMatch = f.match(/\b(20[2-3]\d)\b/);
    const year = yearMatch ? +yearMatch[1] : new Date().getFullYear();

    if (month === null) {
        // YYYY-MM or MM-YYYY
        const m1 = f.match(/(20[2-3]\d)[-_](\d{2})/);
        if (m1 && +m1[2] >= 1 && +m1[2] <= 12) return { month: +m1[2] - 1, year: +m1[1] };
        const m2 = f.match(/(\d{2})[-_](20[2-3]\d)/);
        if (m2 && +m2[1] >= 1 && +m2[1] <= 12) return { month: +m2[1] - 1, year: +m2[2] };
    }

    return month !== null ? { month, year } : null;
}

// ─── ROW NORMALIZATION ────────────────────────────────────────────
function normalizeRow(row) {
    const out = {};
    for (const [key, value] of Object.entries(row)) {
        const cleanKey = key.trim();
        let cleanValue;

        if (value instanceof Date) {
            cleanValue = isNaN(value.getTime()) ? '' : value.toISOString().split('T')[0];
        } else {
            cleanValue = value !== null && value !== undefined ? String(value).trim() : '';
            if (cleanValue !== '') {
                const num = parseFloat(cleanValue.replace(',', '.'));
                if (!isNaN(num) && /^[\d\s,.\-+]+$/.test(cleanValue)) {
                    cleanValue = num;
                }
            }
        }
        out[cleanKey] = cleanValue;
    }
    return out;
}

// ─── FILE READING (Promise wrapper) ──────────────────────────────
function readFileAsWorkbook(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = e => {
            try {
                const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array', cellDates: true });
                resolve(wb);
            } catch (err) { reject(err); }
        };
        reader.onerror = () => reject(new Error('Lesefeil'));
        reader.readAsArrayBuffer(file);
    });
}

// ─── UPLOAD HANDLING ─────────────────────────────────────────────
function initUpload() {
    const dropZone  = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');

    ['dragenter','dragover'].forEach(evt => dropZone.addEventListener(evt, e => {
        e.preventDefault();
        dropZone.classList.add('dragover');
    }));
    ['dragleave','drop'].forEach(evt => dropZone.addEventListener(evt, e => {
        e.preventDefault();
        dropZone.classList.remove('dragover');
    }));

    dropZone.addEventListener('drop', e => {
        if (e.dataTransfer.files.length) handleFiles(e.dataTransfer.files);
    });
    fileInput.addEventListener('change', () => {
        if (fileInput.files.length) handleFiles(fileInput.files);
        fileInput.value = '';
    });
}

function initAddFileInput() {
    const input = document.getElementById('addFileInput');
    if (!input) return;
    input.addEventListener('change', () => {
        if (input.files.length) handleFiles(input.files, false);
        input.value = '';
    });
}

async function handleFiles(files, clearExisting = false) {
    if (clearExisting) {
        loadedFiles = [];
        allData     = [];
        currentMonthFilter = null;
        currentFilter = '';
    }

    const loading = document.getElementById('loading');
    loading.style.display = 'block';
    document.getElementById('errorContainer').innerHTML = '';

    const valid = Array.from(files).filter(f => /\.xlsx?$/i.test(f.name));
    if (!valid.length) {
        showError('Ingen gyldige Excel-filer (.xlsx/.xls)');
        loading.style.display = 'none';
        return;
    }

    let added = 0;
    for (const file of valid) {
        try {
            const wb      = await readFileAsWorkbook(file);
            if (!wb.SheetNames.length) continue;

            const sheet   = wb.Sheets[wb.SheetNames[0]];
            const rawData = XLSX.utils.sheet_to_json(sheet, { defval: '' });
            if (!rawData.length) continue;

            const detected      = detectMonthYear(rawData, file.name);
            const normalizedData = rawData.map(normalizeRow);

            const entry = {
                id:       ++fileIdCounter,
                name:     file.name.replace(/\.xlsx?$/i, ''),
                month:    detected?.month ?? null,
                year:     detected?.year  ?? null,
                rowCount: normalizedData.length,
                data:     normalizedData,
            };

            // Avoid exact duplicates (same name + same first row)
            const isDupe = loadedFiles.some(f =>
                f.name === entry.name &&
                JSON.stringify(f.data[0]) === JSON.stringify(entry.data[0])
            );
            if (!isDupe) {
                loadedFiles.push(entry);
                added++;
                saveDocument(entry.name, normalizedData).catch(console.error);
            }

        } catch (err) {
            console.error(`Feil ved ${file.name}:`, err);
            showError(`Kunne ikke lese: ${escapeHtml(file.name)}`);
        }
    }

    loading.style.display = 'none';

    if (added > 0) {
        rebuildAllData();
        renderLoadedFilesUI();
        if (isResultsVisible()) {
            refreshResults();
        } else {
            processAllData();
        }
    }
}

function rebuildAllData() {
    allData = [];
    for (const f of loadedFiles) {
        for (const row of f.data) {
            allData.push({ ...row, _month: f.month, _year: f.year, _fileId: f.id, _fileName: f.name });
        }
    }
}

function isResultsVisible() {
    const rs = document.getElementById('resultsSection');
    return rs && rs.style.display !== 'none';
}

// ─── LOADED FILES UI ─────────────────────────────────────────────
function renderLoadedFilesUI() {
    const section = document.getElementById('loadedFilesSection');
    const items   = document.getElementById('loadedFilesItems');
    const count   = document.getElementById('loadedFilesCount');
    if (!section || !items) return;

    if (!loadedFiles.length) { section.style.display = 'none'; return; }

    section.style.display = 'block';
    if (count) count.textContent = `${loadedFiles.length} ${loadedFiles.length === 1 ? 'fil' : 'filer'}`;

    items.innerHTML = '';

    // Sort loaded files by year/month
    const sorted = [...loadedFiles].sort((a, b) => {
        if (a.year !== b.year) return (a.year || 0) - (b.year || 0);
        return (a.month || 0) - (b.month || 0);
    });

    for (const f of sorted) {
        const months = currentLang === 'tr' ? MONTHS_TR : MONTHS_NO;
        const monthLabel = (f.month !== null && f.year !== null)
            ? `${months[f.month]} ${f.year}`
            : T('unknownDate');

        const chip = document.createElement('div');
        chip.className = 'file-chip';
        chip.innerHTML = `
            <div style="flex:1;min-width:0;">
                <div class="file-chip-name">${escapeHtml(f.name)}</div>
                <div class="file-chip-meta">
                    <span class="file-chip-month">${monthLabel}</span>
                    <span style="margin-left:8px;color:var(--text-dim);">${f.rowCount} ${T('docRows')}</span>
                </div>
            </div>
            <button class="file-chip-remove" onclick="removeLoadedFile(${f.id})" title="Fjern fil">✕</button>
        `;
        items.appendChild(chip);
    }

    // Show "View results" button if results already processed
    const viewBtn = document.getElementById('viewResultsBtn');
    if (viewBtn) viewBtn.style.display = allData.length > 0 ? 'block' : 'none';
}

function removeLoadedFile(id) {
    loadedFiles = loadedFiles.filter(f => f.id !== id);
    rebuildAllData();
    renderLoadedFilesUI();

    if (!loadedFiles.length) {
        document.getElementById('resultsSection').style.display = 'none';
        currentMonthFilter = null;
        if (shiftChart) { shiftChart.destroy(); shiftChart = null; }
    } else if (isResultsVisible()) {
        // Validate month filter still has data
        if (currentMonthFilter) {
            const still = allData.some(r => r._month === currentMonthFilter.month && r._year === currentMonthFilter.year);
            if (!still) currentMonthFilter = null;
        }
        refreshResults();
    }
}

// ─── OPEN DOCUMENT FROM RECENT LIST ──────────────────────────────
function openDocument(doc) {
    // Wrap the saved doc as a loadedFile entry
    const detected = detectMonthYear(doc.data, doc.name);
    loadedFiles = [{
        id:       ++fileIdCounter,
        name:     doc.name,
        month:    detected?.month ?? null,
        year:     detected?.year  ?? null,
        rowCount: doc.data.length,
        data:     doc.data,
    }];
    rebuildAllData();
    renderLoadedFilesUI();
    document.title = `${doc.name} - ${T('docTitleSuffix')}`;
    processAllData();
}

// ─── PROCESS & DISPLAY ───────────────────────────────────────────
function processAllData() {
    if (!allData.length) return;

    document.getElementById('uploadSection').style.display = 'none';
    document.getElementById('backBtn').classList.add('visible');
    document.getElementById('loading').style.display = 'none';

    const resultsSection = document.getElementById('resultsSection');
    resultsSection.style.display = 'block';

    currentMonthFilter = null;  // Start with all data visible
    refreshResults();
}

function refreshResults() {
    renderMonthNav();
    updateStats();
    renderDriverCards();
    populateFilterDropdown();
    applyFilters();
    initChart();
}

// ─── ACTIVE DATA HELPER ───────────────────────────────────────────
function getActiveData() {
    if (!currentMonthFilter) return allData;
    return allData.filter(r =>
        r._month === currentMonthFilter.month &&
        r._year  === currentMonthFilter.year
    );
}

// ─── STATS ───────────────────────────────────────────────────────
function updateStats() {
    const statsGrid = document.getElementById('statsGrid');
    if (!statsGrid) return;

    const data = getActiveData();

    let amountCol, kmsCol, kmsOccCands = [], tripsCol, cashCol, hoursCol, billedCol;
    for (const key of Object.keys(data[0] || {})) {
        if (key.startsWith('_')) continue;
        const k = key.toLowerCase().trim();
        if (!amountCol && (k.includes('lønn') || k.includes('grunnlag')))          amountCol = key;
        if (!kmsCol    && (k === 'km total' || k === 'km totalt'))                  kmsCol    = key;
        if (k === 'km opptatt' || k === 'km besatt')                                kmsOccCands.push(key);
        if (!tripsCol  && k === 'antall turer')                                     tripsCol  = key;
        if (!cashCol   && (k === 'faktisk kont.' || k === 'faktisk kont'))          cashCol   = key;
        if (!hoursCol  && k === 'effektiv timer')                                   hoursCol  = key;
        if (!billedCol && k.includes('innkjørt total lav'))                         billedCol = key;
    }
    const kmsOccCol = kmsOccCands.find(col => data.some(r => parseFloat(r[col]) > 0)) || null;

    const sum = col => col ? data.reduce((acc, r) => acc + (parseFloat(r[col]) || 0), 0) : 0;

    const totalRows         = data.length;
    const totalAmount       = sum(amountCol);
    const totalKms          = sum(kmsCol);
    const totalKmsOcc       = sum(kmsOccCol);
    const totalTrips        = sum(tripsCol);
    const totalCash         = sum(cashCol);
    const totalHours        = sum(hoursCol);
    const totalBilled       = sum(billedCol);
    const uniqueDriverCount = getUniqueDriverCount(data);

    // Active month label for context
    const months    = currentLang === 'tr' ? MONTHS_NO : MONTHS_NO;
    const monthLabel = currentMonthFilter
        ? `— ${MONTHS_NO[currentMonthFilter.month]} ${currentMonthFilter.year}`
        : '';

    statsGrid.innerHTML = `
        <div class="stat-tile">
            <span class="stat-tile-label">Totalt skift ${monthLabel}</span>
            <span class="stat-tile-value accent">${totalRows.toLocaleString('nb-NO')}</span>
        </div>
        <div class="stat-tile">
            <span class="stat-tile-label">Sjåfører</span>
            <span class="stat-tile-value accent">${uniqueDriverCount}</span>
        </div>
        ${loadedFiles.length > 1 ? `
        <div class="stat-tile">
            <span class="stat-tile-label">Filer</span>
            <span class="stat-tile-value accent">${loadedFiles.length}</span>
        </div>` : ''}
        ${amountCol ? `
        <div class="stat-tile">
            <span class="stat-tile-label">Lønnsgrunnlag</span>
            <span class="stat-tile-value accent">${formatCurrency(totalAmount)}</span>
        </div>` : ''}
        ${cashCol ? `
        <div class="stat-tile">
            <span class="stat-tile-label">${T('statCash')}</span>
            <span class="stat-tile-value accent">${formatCurrency(totalCash)}</span>
        </div>` : ''}
        ${billedCol ? `
        <div class="stat-tile">
            <span class="stat-tile-label">${T('statBilled')}</span>
            <span class="stat-tile-value accent">${formatCurrency(totalBilled)}</span>
        </div>` : ''}
        ${tripsCol ? `
        <div class="stat-tile">
            <span class="stat-tile-label">${T('statTrips')}</span>
            <span class="stat-tile-value accent">${formatNumber(totalTrips)}</span>
        </div>` : ''}
        ${kmsCol ? `
        <div class="stat-tile">
            <span class="stat-tile-label">KM totalt</span>
            <span class="stat-tile-value accent">${formatNumber(totalKms)} km</span>
        </div>` : ''}
        ${kmsOccCol ? `
        <div class="stat-tile">
            <span class="stat-tile-label">${T('statKmsOcc')}</span>
            <span class="stat-tile-value accent">${formatNumber(totalKmsOcc)} km</span>
        </div>` : ''}
        ${hoursCol ? `
        <div class="stat-tile">
            <span class="stat-tile-label">${T('statHours')}</span>
            <span class="stat-tile-value accent">${formatNumber(totalHours)} t</span>
        </div>` : ''}
    `;
}

function getUniqueDriverCount(data) {
    const drivers = new Set();
    for (const row of data) {
        for (const key of Object.keys(row)) {
            if (key.startsWith('_')) continue;
            const kl = key.toLowerCase();
            if (kl.includes('sjåfør') || kl.includes('fører') || kl.includes('navn')) {
                const d = String(row[key]).trim();
                if (d) drivers.add(d);
                break;
            }
        }
    }
    return drivers.size.toLocaleString('nb-NO');
}

function formatCurrency(amount) {
    return new Intl.NumberFormat('nb-NO', { style: 'currency', currency: 'NOK' }).format(amount);
}
function formatNumber(num) {
    return new Intl.NumberFormat('nb-NO').format(Math.round(num));
}
function formatCurrencyCompact(amount) {
    if (amount >= 1000000) return (amount / 1000000).toFixed(1) + 'M';
    if (amount >= 1000)    return Math.round(amount / 1000) + 'k';
    return Math.round(amount).toString();
}

// ─── MONTH NAVIGATION ─────────────────────────────────────────────
function getAvailableYears() {
    const years = new Set();
    for (const r of allData) if (r._year !== null) years.add(r._year);
    return [...years].sort((a, b) => a - b);
}

function getAvailableMonthsForYear(year) {
    const months = new Set();
    for (const r of allData) if (r._year === year && r._month !== null) months.add(r._month);
    return [...months].sort((a, b) => a - b);
}

function getMonthStats(month, year) {
    const rows = allData.filter(r => r._month === month && r._year === year);
    if (!rows.length) return null;

    let totalAmount = 0;
    for (const row of rows) {
        for (const key of Object.keys(row)) {
            if (key.startsWith('_')) continue;
            const kl = key.toLowerCase().trim();
            if (kl.includes('lønn') || kl.includes('grunnlag')) {
                totalAmount += parseFloat(row[key]) || 0;
            }
        }
    }
    return { shifts: rows.length, totalAmount };
}

function renderMonthNav() {
    const section = document.getElementById('monthNavSection');
    if (!section) return;

    const years = getAvailableYears();

    // Only show if we have meaningful date data (more than one month worth)
    const allMonths = allData.filter(r => r._month !== null && r._year !== null);
    const multiMonth = (() => {
        const unique = new Set(allMonths.map(r => `${r._year}-${r._month}`));
        return unique.size > 1;
    })();

    if (years.length === 0 || !multiMonth) {
        section.style.display = 'none';
        return;
    }

    section.style.display = 'block';

    // Default year = most recent
    if (!selectedYear || !years.includes(selectedYear)) {
        selectedYear = Math.max(...years);
    }

    // Year pills
    const yearPillsEl = document.getElementById('yearPills');
    if (yearPillsEl) {
        yearPillsEl.innerHTML = years.map(y => `
            <button class="month-pill ${y === selectedYear ? 'active' : ''}" onclick="selectYear(${y})">
                ${y}
            </button>
        `).join('');
    }

    // "All" pill
    const allPill = document.getElementById('allMonthsPill');
    if (allPill) {
        allPill.className = 'month-pill' + (!currentMonthFilter ? ' active' : '');
        allPill.textContent = T('allMonths');
    }

    buildMonthGrid(selectedYear);
    buildMonthPills();
}

function buildMonthGrid(year) {
    const grid = document.getElementById('monthGrid');
    if (!grid) return;

    const abbr = currentLang === 'tr' ? MONTHS_AB_TR : MONTHS_AB_NO;
    grid.innerHTML = '';

    for (let m = 0; m < 12; m++) {
        const stats   = getMonthStats(m, year);
        const isActive = currentMonthFilter?.month === m && currentMonthFilter?.year === year;
        const hasData  = stats !== null;

        const cell = document.createElement('div');
        cell.className = `month-cell${hasData ? ' has-data' : ''}${isActive ? ' active' : ''}`;

        const amountStr = (hasData && stats.totalAmount > 0)
            ? `<span style="color:var(--accent);font-size:0.75rem;font-weight:700;">${formatCurrencyCompact(stats.totalAmount)}</span>`
            : '';

        cell.innerHTML = `
            <div class="month-cell-name">${abbr[m]}</div>
            ${hasData
                ? `<div class="month-cell-shifts">${stats.shifts}</div>
                   <div class="month-cell-amount">${T('shifts')}</div>
                   <div style="margin-top:4px;">${amountStr}</div>`
                : `<div class="month-cell-shifts" style="color:var(--text-dim);font-size:1rem;">—</div>`
            }
        `;

        if (hasData) cell.addEventListener('click', () => setMonthFilter(m, year));
        grid.appendChild(cell);
    }
}

function buildMonthPills() {
    const container = document.getElementById('monthPillsContainer');
    if (!container) return;

    const abbr  = currentLang === 'tr' ? MONTHS_AB_TR : MONTHS_AB_NO;
    container.innerHTML = '';

    const years = getAvailableYears();
    for (const year of years) {
        const months = getAvailableMonthsForYear(year);
        for (const m of months) {
            const isActive = currentMonthFilter?.month === m && currentMonthFilter?.year === year;
            const btn = document.createElement('button');
            btn.className = `month-pill${isActive ? ' active' : ''}`;
            btn.textContent = `${abbr[m]} '${String(year).slice(2)}`;
            btn.addEventListener('click', () => setMonthFilter(m, year));
            container.appendChild(btn);
        }
    }
}

function selectYear(year) {
    selectedYear = year;
    buildMonthGrid(year);
    // Update year pills highlight
    document.querySelectorAll('#yearPills .month-pill').forEach(btn => {
        btn.className = 'month-pill' + (btn.textContent.trim() == year ? ' active' : '');
    });
}

function setMonthFilter(month, year) {
    currentMonthFilter = { month, year };
    selectedYear = year;
    currentFilter = '';
    const sel = document.getElementById('driverFilter');
    if (sel) sel.value = '';

    renderMonthNav();
    updateStats();
    renderDriverCards();
    populateFilterDropdown();
    applyFilters();
    initChart();

    document.getElementById('statsGrid')?.scrollIntoView({ behavior: 'smooth' });
}

function clearMonthFilter() {
    currentMonthFilter = null;
    renderMonthNav();
    updateStats();
    renderDriverCards();
    populateFilterDropdown();
    applyFilters();
    initChart();
}

// ─── DRIVER CARDS ─────────────────────────────────────────────────
function renderDriverCards() {
    const grid = document.getElementById('summaryGrid');
    if (!grid) return;

    const data = getActiveData();

    let driverCol = findDriverColumnIn(data);
    if (!driverCol) return;

    const drivers = {};
    for (const row of data) {
        const name = String(row[driverCol]).trim();
        if (!name) continue;
        if (!drivers[name]) drivers[name] = { name, amount:0, kms:0, kmsOcc:0, trips:0, cash:0, hours:0, billed:0, rows:[] };
        drivers[name].rows.push(row);
        for (const key of Object.keys(row)) {
            if (key.startsWith('_')) continue;
            const kl  = key.toLowerCase().trim();
            const val = parseFloat(row[key]);
            if (!isNaN(val) && val !== 0) {
                if (kl.includes('lønn') || kl.includes('grunnlag'))          drivers[name].amount += val;
                else if (kl === 'km total' || kl === 'km totalt')            drivers[name].kms    += val;
                else if ((kl === 'km opptatt' || kl === 'km besatt') && val > 0) drivers[name].kmsOcc += val;
                else if (kl === 'antall turer')                              drivers[name].trips  += val;
                else if (kl === 'faktisk kont.' || kl === 'faktisk kont')    drivers[name].cash   += val;
                else if (kl === 'effektiv timer')                            drivers[name].hours  += val;
                else if (kl.includes('innkjørt total lav'))                  drivers[name].billed += val;
            }
        }
    }

    const sorted = Object.values(drivers).sort((a, b) => b.amount - a.amount);
    grid.innerHTML = '';

    for (const driver of sorted.slice(0, 12)) {
        const card = document.createElement('div');
        card.className = 'driver-card';
        card.innerHTML = `
            <div class="driver-card-header">
                <div class="driver-avatar">
                    <svg viewBox="0 0 24 24" fill="currentColor">
                        <path d="M12 11.5a4 4 0 1 0 0-8 4 4 0 0 0 0 8zm0 2c-4 0-7 2-7 4v1h14v-1c0-2-3-4-7-4z"/>
                    </svg>
                </div>
                <div class="driver-name-block">
                    <div class="driver-name">${escapeHtml(driver.name)}</div>
                    <div class="driver-hint">${T('clickToFilter')}</div>
                </div>
            </div>
            <div class="card-metrics">
                <div class="metric">
                    <div class="metric-label">${T('rows')}</div>
                    <div class="metric-val">${driver.rows.length}</div>
                </div>
                ${driver.amount > 0 ? `<div class="metric"><div class="metric-label">${T('amount')}</div><div class="metric-val hi">${formatCurrency(driver.amount)}</div></div>` : ''}
                ${driver.cash   > 0 ? `<div class="metric"><div class="metric-label">${T('metricCash')}</div><div class="metric-val green">${formatCurrency(driver.cash)}</div></div>` : ''}
                ${driver.billed > 0 ? `<div class="metric"><div class="metric-label">${T('metricBilled')}</div><div class="metric-val">${formatCurrency(driver.billed)}</div></div>` : ''}
                ${driver.trips  > 0 ? `<div class="metric"><div class="metric-label">${T('metricTrips')}</div><div class="metric-val">${formatNumber(driver.trips)}</div></div>` : ''}
                ${driver.kms    > 0 ? `<div class="metric"><div class="metric-label">${T('kms')}</div><div class="metric-val green">${formatNumber(driver.kms)}</div></div>` : ''}
                ${driver.kmsOcc > 0 ? `<div class="metric"><div class="metric-label">${T('metricKmsOcc')}</div><div class="metric-val">${formatNumber(driver.kmsOcc)}</div></div>` : ''}
                ${driver.hours  > 0 ? `<div class="metric"><div class="metric-label">${T('metricHours')}</div><div class="metric-val">${formatNumber(driver.hours)} t</div></div>` : ''}
            </div>
        `;

        card.addEventListener('click', () => {
            currentFilter = driver.name;
            const sel = document.getElementById('driverFilter');
            if (sel) sel.value = driver.name;
            applyFilters();
            // Sync chart
            chartDriverFilter = driver.name;
            const chartSel = document.getElementById('chartDriverFilter');
            if (chartSel) chartSel.value = driver.name;
            updateChipSelection(driver.name);
            renderChart();
            setTimeout(() => {
                document.querySelector('.table-card')?.scrollIntoView({ behavior: 'smooth' });
            }, 100);
        });

        grid.appendChild(card);
    }
}

// ─── TABLE ────────────────────────────────────────────────────────
function populateFilterDropdown() {
    const select = document.getElementById('driverFilter');
    if (!select) return;

    const data   = getActiveData();
    const dCol   = findDriverColumnIn(data);
    if (!dCol) return;

    const drivers = [...new Set(data.map(r => String(r[dCol]).trim()).filter(Boolean))].sort();
    select.innerHTML = `<option value="">Alle sjåfører</option>`;
    for (const d of drivers) {
        const o = document.createElement('option');
        o.value = d; o.textContent = d;
        select.appendChild(o);
    }
    select.value = currentFilter;
    select.onchange = () => { currentFilter = select.value; applyFilters(); };
}

function renderTable() {
    applyFilters();
}

function applyFilters() {
    let data = getActiveData();
    const driverCol = findDriverColumnIn(data);

    if (currentFilter && driverCol) {
        data = data.filter(r => String(r[driverCol]).trim() === currentFilter);
    }

    if (currentSort !== 'none') {
        const [colHint, dir] = currentSort.split('_');
        // Find matching column
        const sortCol = Object.keys(data[0] || {}).find(k => {
            const kl = k.toLowerCase();
            return colHint === 'amount' ? (kl.includes('lønn') || kl.includes('grunnlag')) : kl.includes('km');
        });
        if (sortCol) {
            data.sort((a, b) => {
                const vA = parseFloat(a[sortCol]) || 0;
                const vB = parseFloat(b[sortCol]) || 0;
                return dir === 'desc' ? vB - vA : vA - vB;
            });
        }
    }

    renderTableContent(data, driverCol);
}

function renderTableContent(data, driverCol) {
    const thead = document.getElementById('tableHead');
    const tbody = document.getElementById('tableBody');
    if (!thead || !tbody) return;

    // Exclude internal _meta columns
    const headers = Object.keys(data[0] || {}).filter(h => !h.startsWith('_'));

    thead.innerHTML = '<tr>' + headers.map(h => `<th>${escapeHtml(h)}</th>`).join('') + '</tr>';
    tbody.innerHTML = '';

    for (const row of data) {
        const tr = document.createElement('tr');
        for (const key of headers) {
            const td  = document.createElement('td');
            const val = row[key];
            if (key === driverCol) {
                td.className  = 'driver-highlight';
                td.textContent = String(val);
            } else if (typeof val === 'number') {
                td.textContent = new Intl.NumberFormat('nb-NO', { maximumFractionDigits: 2 }).format(val);
                if (Math.abs(val) > 10000) td.style.color = 'var(--accent)';
            } else {
                td.textContent = val;
            }
            tr.appendChild(td);
        }
        tbody.appendChild(tr);
    }

    // Total row
    const totalRow = document.createElement('tr');
    totalRow.className = 'total-row';
    totalRow.innerHTML = headers.map((h, i) => {
        if (i === 0) return `<td><strong>${T('tableTotal')}</strong></td>`;
        if (typeof (data[0]?.[h]) === 'number') {
            const s = data.reduce((acc, r) => acc + (parseFloat(r[h]) || 0), 0);
            return `<td style="color:var(--accent);font-weight:700;">${s !== 0 ? new Intl.NumberFormat('nb-NO',{maximumFractionDigits:2}).format(s) : ''}</td>`;
        }
        return '<td></td>';
    }).join('');
    tbody.appendChild(totalRow);

    document.getElementById('dataTable').style.display = 'table';
}

// ─── COLUMN DETECTION HELPERS ─────────────────────────────────────
function findDriverColumn() {
    return findDriverColumnIn(allData);
}
function findDriverColumnIn(data) {
    if (!data.length) return null;
    for (const key of Object.keys(data[0])) {
        if (key.startsWith('_')) continue;
        const kl = key.toLowerCase();
        if (kl.includes('sjåfør') || kl.includes('fører') || kl.includes('navn')) return key;
    }
    return Object.keys(data[0]).find(k => !k.startsWith('_')) || null;
}

// ─── RECENT DOCUMENTS ─────────────────────────────────────────────
function renderRecentList(docs) {
    const container = document.getElementById('recentListContainer');
    if (!container) return;

    // Clear dynamic content (keep the H3)
    const h3 = container.querySelector('h3');
    container.innerHTML = '';
    if (h3) container.appendChild(h3);

    const noHint = document.getElementById('noRecentHint');

    if (!docs || !docs.length) {
        if (noHint) noHint.style.display = 'block';
        return;
    }
    if (noHint) noHint.style.display = 'none';

    // Group by month/year
    const grouped = {};
    docs.forEach(doc => {
        const d = new Date(doc.timestamp);
        const key = d.toLocaleDateString('nb-NO', { year: 'numeric', month: 'long' });
        if (!grouped[key]) grouped[key] = [];
        grouped[key].push(doc);
    });

    const sortedKeys = Object.keys(grouped).sort((a, b) => new Date(b) - new Date(a));

    for (const monthYear of sortedKeys) {
        const groupDiv = document.createElement('div');
        groupDiv.className = 'recent-group';

        const label = document.createElement('label');
        label.textContent = monthYear;
        groupDiv.appendChild(label);

        for (const doc of grouped[monthYear]) {
            const item = document.createElement('div');
            item.className = 'doc-item';
            const d = new Date(doc.timestamp);
            item.innerHTML = `
                <div class="doc-item-header" style="display:flex;justify-content:space-between;align-items:center;">
                    <span class="doc-name">${escapeHtml(doc.name)}</span>
                    <span class="doc-date">${d.toLocaleDateString('nb-NO', { day:'2-digit', month:'short', year:'numeric' })}</span>
                </div>
                <div class="doc-meta">
                    <span class="doc-rows">${doc.rowCount} ${T('docRows')}</span>
                    ${doc.data[0] ? `<span class="doc-cols">${Object.keys(doc.data[0]).length} ${T('docCols')}</span>` : ''}
                </div>
            `;
            item.addEventListener('click', () => openDocument(doc));
            groupDiv.appendChild(item);
        }
        container.appendChild(groupDiv);
    }
}

// ─── UPLOAD SECTION ───────────────────────────────────────────────
function showUploadSection(clearFiles = false) {
    if (clearFiles) {
        loadedFiles = [];
        allData = [];
        currentMonthFilter = null;
        currentFilter = '';
        if (shiftChart) { shiftChart.destroy(); shiftChart = null; }
    }

    document.getElementById('resultsSection').style.display = 'none';
    document.getElementById('uploadSection').style.display = 'flex';
    document.getElementById('backBtn').classList.remove('visible');

    const newFileBtn = document.getElementById('newFileBtn');
    if (newFileBtn) newFileBtn.style.display = 'none';

    renderLoadedFilesUI();
    loadRecentDocuments();
    window.scrollTo({ top: 0, behavior: 'smooth' });
}

// ─── LANGUAGE SWITCHING ───────────────────────────────────────────
function setLang(lang) {
    currentLang = lang;

    document.getElementById('langNO').classList.toggle('active', lang === 'no');
    document.getElementById('langTR').classList.toggle('active', lang === 'tr');

    const translations = {
        no: {
            title:'Sjåfør', recentDocsTitle:'Nylige dokumenter',
            uploadTitle:'Last opp<br>Excel-fil',
            uploadSub:'Støtter .xlsx og .xls — gjenkjenner sjåførkolonne automatisk',
            chooseFile:'Velg fil', dragDrop:'Klikk eller dra og slipp Excel-filen her',
            loading:'Laster data...', errorTitle:'Feil',
            totalRows:'Totalt skift', uniqueDrivers:'Sjåfører', totalAmount:'Lønnsgrunnlag',
            totalKms:'KM totalt', cardsTitle:'Sjåføroversikt',
            cardsSub:'Klikk på et kort for å filtrere tabellen',
            tableTitle:'Skiftdata', filterBy:'FILTRER:', allDrivers:'Alle sjåfører',
            sortBy:'SORTER:', sortNone:'Ingen sortering', sortWages:'Lønnsgrunnlag ↓',
            sortKms:'KM ↓', backBtn:'Tilbake',
            chartDriver:'Sjåfør', chartMetric:'Metrikk', chartAll:'Alle sjåfører',
            chartReset:'✕ Alle', yearViewTitle:'Årsvisning',
            loadedFilesTitle:'Lastede filer', allMonthsLabel:'Alle måneder',
            addFileBtn:'+ Legg til fil', newAnalysis:'Ny analyse',
            mLønn:'Lønnsgrunnlag', mInnkjort:'Innkjørt total', mKontant:'Kontant',
            mKmTotal:'KM totalt', mKmOcc:'KM m/passasjer', mTurer:'Antall turer',
            mTimer:'Effektive timer',
        },
        tr: {
            title:'Şoför', recentDocsTitle:'Son Belgeler',
            uploadTitle:'Excel<br>Dosyası Yükle',
            uploadSub:'.xlsx ve .xls desteklenir — sürücü sütunu otomatik algılanır',
            chooseFile:'Dosya Seç', dragDrop:'Tıklayın veya Excel dosyasını buraya sürükleyip bırakın',
            loading:'Veriler yükleniyor...', errorTitle:'Hata',
            totalRows:'Toplam Vardiya', uniqueDrivers:'Şoförler', totalAmount:'Ücret Tabanı',
            totalKms:'Toplam KM', cardsTitle:'Şoför Özeti',
            cardsSub:'Tabloyu filtrelemek için bir karta tıklayın',
            tableTitle:'Vardiya Verileri', filterBy:'FİLTRELE:', allDrivers:'Tüm Şoförler',
            sortBy:'SIRALA:', sortNone:'Sıralama yapma', sortWages:'Ücret Tabanı ↓',
            sortKms:'KM ↓', backBtn:'Geri',
            chartDriver:'Şoför', chartMetric:'Metrik', chartAll:'Tüm Şoförler',
            chartReset:'✕ Tümü', yearViewTitle:'Yıllık Görünüm',
            loadedFilesTitle:'Yüklenen Dosyalar', allMonthsLabel:'Tüm Aylar',
            addFileBtn:'+ Dosya Ekle', newAnalysis:'Yeni Analiz',
            mLønn:'Ücret Tabanı', mInnkjort:'Toplam Hasılat', mKontant:'Nakit',
            mKmTotal:'Toplam KM', mKmOcc:'Yolculu KM', mTurer:'Sefer Sayısı',
            mTimer:'Aktif Saat',
        }
    };

    const t = translations[lang];
    document.querySelectorAll('[data-i18n]').forEach(el => {
        const key = el.getAttribute('data-i18n');
        if (t[key]) {
            if (['title','uploadTitle'].includes(key)) el.innerHTML = t[key];
            else el.textContent = t[key];
        }
    });

    // Update buttons not covered by data-i18n
    const setEl = (id, key) => { const el = document.getElementById(id); if (el && t[key]) el.textContent = t[key]; };
    setEl('chartTypebar',  'chartBar');
    setEl('chartTypeline', 'chartLine');
    setEl('chartAccumBtn', 'chartAccum');
    setEl('allMonthsPill', 'allMonthsLabel');

    // Chart metric options
    const metricMap = {
        'Lønnsgrunnlag': t.mLønn, 'Innkjørt total Lav sats': t.mInnkjort,
        'Faktisk kont.': t.mKontant, 'Km total': t.mKmTotal,
        'Km opptatt': t.mKmOcc, 'Antall turer': t.mTurer, 'Effektiv timer': t.mTimer,
    };
    document.querySelectorAll('#chartMetricFilter option').forEach(opt => {
        if (metricMap[opt.value]) opt.textContent = metricMap[opt.value];
    });

    // Update chart metric labels live
    if (window.CHART_METRICS) {
        CHART_METRICS['Lønnsgrunnlag'].label           = t.mLønn;
        CHART_METRICS['Innkjørt total Lav sats'].label = t.mInnkjort;
        CHART_METRICS['Faktisk kont.'].label           = t.mKontant;
        CHART_METRICS['Km total'].label                = t.mKmTotal;
        CHART_METRICS['Km opptatt'].label              = t.mKmOcc;
        CHART_METRICS['Antall turer'].label            = t.mTurer;
        CHART_METRICS['Effektiv timer'].label          = t.mTimer;
    }

    if (allData.length) refreshResults();
}

// ─── CHART ────────────────────────────────────────────────────────
let shiftChart        = null;
let chartDriverFilter = '';
let chartType         = 'bar';
let chartAccumulate   = false;

const CHART_METRICS = {
    'Lønnsgrunnlag':           { label:'Lønnsgrunnlag',   color:'#F7C520', fmt: v => formatCurrency(v) },
    'Innkjørt total Lav sats': { label:'Innkjørt total',  color:'#60A5FA', fmt: v => formatCurrency(v) },
    'Faktisk kont.':           { label:'Kontant',          color:'#34D399', fmt: v => formatCurrency(v) },
    'Km total':                { label:'KM totalt',        color:'#A78BFA', fmt: v => formatNumber(v)+' km' },
    'Km opptatt':              { label:'KM m/passasjer',   color:'#FB923C', fmt: v => formatNumber(v)+' km' },
    'Antall turer':            { label:'Antall turer',     color:'#F472B6', fmt: v => formatNumber(v) },
    'Effektiv timer':          { label:'Effektive timer',  color:'#38BDF8', fmt: v => formatNumber(v)+' t' },
};

function setChartType(type) {
    chartType = type;
    const barBtn  = document.getElementById('chartTypebar');
    const lineBtn = document.getElementById('chartTypeline');
    if (barBtn && lineBtn) {
        barBtn.style.background  = type === 'bar'  ? 'var(--accent)' : 'transparent';
        barBtn.style.color       = type === 'bar'  ? '#000'          : 'var(--text-muted)';
        lineBtn.style.background = type === 'line' ? 'var(--accent)' : 'transparent';
        lineBtn.style.color      = type === 'line' ? '#000'          : 'var(--text-muted)';
    }
    renderChart();
}

function toggleAccumulate() {
    chartAccumulate = !chartAccumulate;
    const btn = document.getElementById('chartAccumBtn');
    if (btn) {
        btn.style.background  = chartAccumulate ? 'var(--accent)' : 'transparent';
        btn.style.borderColor = chartAccumulate ? 'var(--accent)' : 'var(--border2)';
        btn.style.color       = chartAccumulate ? '#000'          : 'var(--text-muted)';
    }
    renderChart();
}

function initChart() {
    const sel = document.getElementById('chartDriverFilter');
    if (!sel) return;

    const data      = getActiveData();
    const driverCol = findDriverColumnIn(data);
    const drivers   = [...new Set(data.map(r => driverCol ? String(r[driverCol]).trim() : '').filter(Boolean))].sort();

    sel.innerHTML = `<option value="">Alle sjåfører</option>`;
    drivers.forEach(d => { const o = document.createElement('option'); o.value = d; o.textContent = d; sel.appendChild(o); });

    // Chips
    const chips = document.getElementById('chartDriverChips');
    if (chips) {
        chips.innerHTML = '';
        chips.appendChild(makeChip('Alle', '', true));
        drivers.forEach(d => chips.appendChild(makeChip(d, d, false)));
    }

    sel.onchange = () => {
        chartDriverFilter = sel.value;
        updateChipSelection(chartDriverFilter);
        renderChart();
    };
    document.getElementById('chartMetricFilter').onchange = renderChart;

    // Validate chartDriverFilter still exists in new data
    if (chartDriverFilter && !drivers.includes(chartDriverFilter)) {
        chartDriverFilter = '';
        sel.value = '';
    }

    renderChart();
}

function makeChip(label, value, active) {
    const btn = document.createElement('button');
    btn.textContent = label;
    btn.dataset.value = value;
    btn.style.cssText = `
        padding:5px 14px; border-radius:20px; font-family:inherit;
        font-size:0.8rem; font-weight:600; cursor:pointer;
        border:1px solid ${active ? 'var(--accent)' : 'var(--border2)'};
        background:${active ? 'var(--accent)' : 'var(--surface2)'};
        color:${active ? '#000' : 'var(--text-muted)'};
        transition:all 0.15s ease;
    `;
    btn.addEventListener('click', () => {
        chartDriverFilter = value;
        const s = document.getElementById('chartDriverFilter');
        if (s) s.value = value;
        updateChipSelection(value);
        renderChart();
    });
    return btn;
}

function updateChipSelection(value) {
    const chips = document.getElementById('chartDriverChips');
    if (!chips) return;
    chips.querySelectorAll('button').forEach(btn => {
        const active = btn.dataset.value === value;
        btn.style.background  = active ? 'var(--accent)' : 'var(--surface2)';
        btn.style.borderColor = active ? 'var(--accent)' : 'var(--border2)';
        btn.style.color       = active ? '#000'          : 'var(--text-muted)';
    });
}

function resetChartFilter() {
    chartDriverFilter = '';
    const s = document.getElementById('chartDriverFilter');
    if (s) s.value = '';
    updateChipSelection('');
    renderChart();
}

function renderChart() {
    const metricKey = document.getElementById('chartMetricFilter')?.value || 'Lønnsgrunnlag';
    const metric    = CHART_METRICS[metricKey] || CHART_METRICS['Lønnsgrunnlag'];
    const data      = getActiveData();
    const driverCol = findDriverColumnIn(data);

    let rows = [...data];
    if (chartDriverFilter && driverCol) {
        rows = rows.filter(r => String(r[driverCol]).trim() === chartDriverFilter);
    }

    // Sort by shift number or date
    const shiftKey = Object.keys(allData[0] || {}).find(k => k.toLowerCase().includes('skiftnr')) || null;
    rows.sort((a, b) => shiftKey ? (parseInt(a[shiftKey]) || 0) - (parseInt(b[shiftKey]) || 0) : 0);

    // If showing all months, group by month for readability
    const isMultiMonth = !currentMonthFilter && getAvailableYears().length > 0 &&
        (() => { const u = new Set(allData.filter(r=>r._month!==null).map(r=>`${r._year}-${r._month}`)); return u.size > 1; })();

    let labels, values;

    if (isMultiMonth && !chartDriverFilter) {
        // Aggregate by month
        const monthTotals = {};
        const abbr = currentLang === 'tr' ? MONTHS_AB_TR : MONTHS_AB_NO;
        for (const row of rows) {
            if (row._month === null || row._year === null) continue;
            const key = `${row._year}-${String(row._month).padStart(2,'0')}`;
            if (!monthTotals[key]) monthTotals[key] = { label: `${abbr[row._month]} ${row._year}`, total: 0 };
            const val = row[metricKey] ?? (() => {
                const found = Object.keys(row).find(k => k.toLowerCase().trim() === metricKey.toLowerCase().trim());
                return found ? row[found] : 0;
            })();
            monthTotals[key].total += parseFloat(val) || 0;
        }
        const sorted = Object.entries(monthTotals).sort(([a],[b]) => a.localeCompare(b));
        labels = sorted.map(([,v]) => v.label);
        values = sorted.map(([,v]) => v.total);
    } else {
        const dateKey = Object.keys(allData[0] || {}).find(k => k.toLowerCase().includes('start dato')) || null;
        labels = rows.map(r => {
            const shift  = shiftKey ? '#' + r[shiftKey] : '';
            const date   = dateKey  ? r[dateKey]        : '';
            const driver = chartDriverFilter ? '' : (driverCol ? ' · ' + String(r[driverCol]).trim() : '');
            return [shift + driver, date].filter(Boolean).join(' ');
        });
        values = rows.map(r => {
            let val = r[metricKey];
            if (val === undefined) {
                const found = Object.keys(r).find(k => k.toLowerCase().trim() === metricKey.toLowerCase().trim());
                val = found ? r[found] : 0;
            }
            return parseFloat(val) || 0;
        });
    }

    const displayValues = chartAccumulate
        ? values.reduce((acc, v, i) => { acc.push((acc[i-1] || 0) + v); return acc; }, [])
        : values;

    const empty  = document.getElementById('chartEmpty');
    const canvas = document.getElementById('shiftChart');
    if (!displayValues.some(v => v > 0)) {
        if (empty)  empty.style.display  = 'block';
        if (canvas) canvas.style.display = 'none';
        return;
    }
    if (empty)  empty.style.display  = 'none';
    if (canvas) canvas.style.display = 'block';

    if (shiftChart) { shiftChart.destroy(); shiftChart = null; }

    const isLine = chartType === 'line';
    const dataset = isLine ? {
        label: metric.label, data: displayValues,
        borderColor: metric.color, backgroundColor: metric.color + '22',
        borderWidth: 2.5, pointBackgroundColor: metric.color,
        pointBorderColor: '#12151C', pointBorderWidth: 2,
        pointRadius: 5, pointHoverRadius: 7, fill: true, tension: 0.35,
    } : {
        label: metric.label, data: displayValues,
        backgroundColor: metric.color + '99', borderColor: metric.color,
        borderWidth: 2, borderRadius: 6, borderSkipped: false,
    };

    shiftChart = new Chart(document.getElementById('shiftChart').getContext('2d'), {
        type: isLine ? 'line' : 'bar',
        data: { labels, datasets: [dataset] },
        options: {
            responsive: true, maintainAspectRatio: true,
            plugins: {
                legend: { display: false },
                tooltip: {
                    backgroundColor:'#191D27', borderColor:'#252C3E', borderWidth:1,
                    titleColor:'#E6E9F4', bodyColor:'#7B849A', padding:12,
                    callbacks: { label: ctx => '  ' + metric.fmt(ctx.parsed.y) }
                }
            },
            scales: {
                x: { ticks: { color:'#4A5268', font:{family:'Inter',size:11}, maxRotation:45 }, grid:{color:'#1E2433'}, border:{color:'#1E2433'} },
                y: { ticks: { color:'#7B849A', font:{family:'Inter',size:11}, callback: v => metric.fmt(v) }, grid:{color:'#1E2433'}, border:{color:'#1E2433'}, beginAtZero:true }
            }
        }
    });
}

// ─── UTILS ────────────────────────────────────────────────────────
function showError(msg) {
    const container = document.getElementById('errorContainer');
    if (!container) return;
    container.innerHTML = `
        <div class="error-box">
            <h3>${T('errorTitle')}</h3>
            <p>${escapeHtml(msg)}</p>
        </div>
    `;
    document.getElementById('loading').style.display = 'none';
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}
