(function () {
    'use strict';

    const STORAGE_KEYS = {
        config: 'importscore.config.v1',
        proxy: 'importscore.proxy.v1',
        students: 'importscore.students.v1',
        recent: 'importscore.recent-files.v1',
        fileName: 'importscore.file-name.v1',
        ui: 'importscore.ui.v1'
    };

    const DEFAULT_CONFIG = {
        soCauToiDa: 40,
        danhSachMaDE: ['701', '702', '703', '704'],
        tenCotHS: 'Họ và tên'
    };

    const state = {
        students: [],
        searchText: '',
        activeFilter: 'all',
        sortBy: 'name',
        sortDir: 'asc',
        showMaDE: true,
        selectedIndex: -1,
        selectedStudentId: null,
        chart: null,
        pendingWorkbook: null,
        pendingFileName: '',
        pendingSheets: [],
        importRows: [],
        searchFocusTimer: null,
        undoStack: [],
        redoStack: [],
        maxHistory: 50,
        cloudBusy: false,
        config: loadJson(STORAGE_KEYS.config, DEFAULT_CONFIG),
        proxyUrl: localStorage.getItem(STORAGE_KEYS.proxy) || 'https://proxyscore.mctran2005.workers.dev',
        currentFileName: localStorage.getItem(STORAGE_KEYS.fileName) || '',
        recentFiles: loadJson(STORAGE_KEYS.recent, []),
        ui: loadJson(STORAGE_KEYS.ui, { sidebarCollapsed: false })
    };

    const dom = {};

    document.addEventListener('DOMContentLoaded', init);

    function init() {
        cacheDom();
        bindEvents();
        hydrateState();
        renderAll();
    }

    function cacheDom() {
        dom.toastContainer = byId('toastContainer');
        dom.fileNameDisplay = byId('fileNameDisplay');
        dom.fileBadge = byId('fileBadge');

        dom.hStatTotal = byId('hStatTotal');
        dom.hStatEntered = byId('hStatEntered');
        dom.hStatExcellent = byId('hStatExcellent');
        dom.hStatGood = byId('hStatGood');
        dom.hStatAverage = byId('hStatAverage');
        dom.hStatWeak = byId('hStatWeak');

        dom.progressFill = byId('progressFill');
        dom.progressText = byId('progressText');
        dom.scoreExtremes = byId('scoreExtremes');

        dom.sidebar = byId('sidebar');
        dom.appBody = byId('appBody');

        dom.cfgSoCauToiDa = byId('cfgSoCauToiDa');
        dom.cfgDanhSachMaDE = byId('cfgDanhSachMaDE');
        dom.cfgTenCotHS = byId('cfgTenCotHS');
        dom.cfgProxyUrl = byId('cfgProxyUrl');

        dom.recentFilesList = byId('recentFilesList');

        dom.searchInput = byId('searchInput');
        dom.btnClearSearch = byId('btnClearSearch');
        dom.resultCount = byId('resultCount');

        dom.emptyState = byId('emptyState');
        dom.tableWrap = byId('tableWrap');
        dom.tableBody = byId('tableBody');
        dom.noResult = byId('noResult');
        dom.thMaDE = byId('thMaDE');

        dom.chartSection = byId('chartSection');
        dom.scoreChart = byId('scoreChart');
        dom.chartAvgBadge = byId('chartAvgBadge');

        dom.modalImport = byId('modalImport');
        dom.dropZone = byId('dropZone');
        dom.fileInput = byId('fileInput');
        dom.fileInfoBar = byId('fileInfoBar');
        dom.importFileName = byId('importFileName');
        dom.importFileSize = byId('importFileSize');
        dom.importPreview = byId('importPreview');
        dom.previewTableWrap = byId('previewTableWrap');
        dom.sheetSelector = byId('sheetSelector');
        dom.mapNameCol = byId('mapNameCol');
        dom.mapSkipRows = byId('mapSkipRows');
        dom.importSummary = byId('importSummary');
        dom.btnDoImport = byId('btnDoImport');
        dom.importCountSpan = byId('importCountSpan');

        dom.modalExport = byId('modalExport');
        dom.expStudentCount = byId('expStudentCount');
        dom.expEnteredCount = byId('expEnteredCount');
        dom.expAvgScore = byId('expAvgScore');
        dom.expRecordName = byId('expRecordName');
        dom.expClassName = byId('expClassName');
        dom.cbSaveToCloud = byId('cbSaveToCloud');
        dom.cloudSubOptions = byId('cloudSubOptions');
        dom.cbIsPublic = byId('cbIsPublic');

        dom.modalStudent = byId('modalStudent');
        dom.studentModalTitle = byId('studentModalTitle');
        dom.stuName = byId('stuName');
        dom.stuMaDE = byId('stuMaDE');
        dom.stuSoCau = byId('stuSoCau');
        dom.stuDiem = byId('stuDiem');

        dom.modalConfirm = byId('modalConfirm');
        dom.confirmTitle = byId('confirmTitle');
        dom.confirmMsg = byId('confirmMsg');
        dom.btnConfirmYes = byId('btnConfirmYes');
        dom.btnConfirmNo = byId('btnConfirmNo');

        dom.btnUndo = byId('btnUndo');
        dom.btnRedo = byId('btnRedo');
    }

    function bindEvents() {
        byId('btnImport').addEventListener('click', () => openModal(dom.modalImport));
        byId('btnExport').addEventListener('click', openExportModal);
        byId('btnToggleSidebar').addEventListener('click', toggleSidebar);
        byId('btnSaveConfig').addEventListener('click', saveExamConfig);
        byId('btnSaveCloudConfig').addEventListener('click', saveCloudConfig);
        byId('btnToggleChart').addEventListener('click', toggleChart);
        byId('btnCloseChart').addEventListener('click', () => setChartVisible(false));
        byId('btnToggleMaDE').addEventListener('click', toggleMaDE);
        byId('btnAddStudent').addEventListener('click', () => openStudentModal(-1));

        dom.searchInput.addEventListener('input', onSearchInput);
        dom.btnClearSearch.addEventListener('click', clearSearch);

        dom.btnUndo.addEventListener('click', undoChange);
        dom.btnRedo.addEventListener('click', redoChange);

        document.querySelectorAll('.fcl').forEach((button) => {
            button.addEventListener('click', () => {
                state.activeFilter = button.dataset.filter || 'all';
                document.querySelectorAll('.fcl').forEach((x) => x.classList.remove('active'));
                button.classList.add('active');
                renderTableAndStats();
            });
        });

        document.querySelectorAll('th.sortable').forEach((th) => {
            th.addEventListener('click', () => {
                const col = th.dataset.col;
                if (!col) return;
                if (state.sortBy === col) {
                    state.sortDir = state.sortDir === 'asc' ? 'desc' : 'asc';
                } else {
                    state.sortBy = col;
                    state.sortDir = 'asc';
                }
                renderTableAndStats();
            });
        });

        dom.tableBody.addEventListener('input', handleTableInput);
        dom.tableBody.addEventListener('keydown', handleTableKeydown);
        dom.tableBody.addEventListener('click', handleTableClick);

        bindImportModalEvents();
        bindExportModalEvents();
        bindStudentModalEvents();
        bindGenericModalEvents();
        bindKeyboardShortcuts();
    }

    function hydrateState() {
        const savedStudents = loadJson(STORAGE_KEYS.students, []);
        state.students = savedStudents.map(normalizeStudent);

        if (state.currentFileName) {
            dom.fileNameDisplay.textContent = state.currentFileName;
        }

        if (state.ui.sidebarCollapsed) {
            dom.sidebar.classList.add('collapsed');
            dom.appBody.classList.add('sidebar-collapsed');
        }

        dom.cfgSoCauToiDa.value = String(state.config.soCauToiDa);
        dom.cfgDanhSachMaDE.value = state.config.danhSachMaDE.join(',');
        dom.cfgTenCotHS.value = state.config.tenCotHS;
        dom.cfgProxyUrl.value = state.proxyUrl;

        rebuildStudentMaDEOptions();
        renderRecentFiles();
        updateUndoRedoState();
    }

    function renderAll() {
        renderTableAndStats();
        renderProgress();
    }

    function renderTableAndStats() {
        const rows = getFilteredSortedRows();

        if (state.students.length === 0) {
            dom.emptyState.classList.remove('hidden');
            dom.tableWrap.classList.add('hidden');
            dom.noResult.classList.add('hidden');
        } else {
            dom.emptyState.classList.add('hidden');
            dom.tableWrap.classList.remove('hidden');
            dom.noResult.classList.toggle('hidden', rows.length > 0);
        }

        dom.tableBody.innerHTML = rows.map((student, renderIndex) => {
            const score = getStudentScore(student);
            const scoreText = Number.isFinite(score) ? formatScore(score) : '';
            const soCau = Number.isFinite(student.socau) ? String(student.socau) : '';
            const xepLoai = getRank(score);

            return `
                <tr data-student-id="${escapeAttr(student.id)}" class="${state.selectedStudentId === student.id ? 'row-active' : ''}">
                    <td>${renderIndex + 1}</td>
                    <td>
                        <input class="cell-input" data-field="name" value="${escapeAttr(student.name)}" placeholder="Họ và tên">
                    </td>
                    <td class="made-col ${state.showMaDE ? '' : 'hidden'}">
                        <input class="cell-input" data-field="maDe" value="${escapeAttr(student.maDe || '')}" placeholder="Mã đề">
                    </td>
                    <td>
                        <input class="cell-input" data-field="socau" type="number" min="0" max="${state.config.soCauToiDa}" value="${escapeAttr(soCau)}" placeholder="Số câu">
                    </td>
                    <td>
                        <input class="cell-input" data-field="diem" type="number" min="0" max="10" step="0.01" value="${escapeAttr(scoreText)}" placeholder="Điểm">
                    </td>
                    <td>${xepLoai}</td>
                    <td>
                        <button class="row-btn" data-action="focus-score" title="Nhập nhanh">🎯</button>
                        <button class="row-btn" data-action="delete" title="Xóa">🗑️</button>
                    </td>
                </tr>
            `;
        }).join('');

        document.querySelectorAll('.made-col').forEach((col) => {
            col.classList.toggle('hidden', !state.showMaDE);
        });
        dom.thMaDE.classList.toggle('hidden', !state.showMaDE);

        dom.resultCount.textContent = `Hiển thị ${rows.length}/${state.students.length}`;

        renderHeaderStats();
        renderProgress();
        renderChartIfVisible();
    }

    function renderHeaderStats() {
        const scores = state.students.map(getStudentScore).filter(Number.isFinite);
        const entered = countEntered();

        dom.hStatTotal.textContent = String(state.students.length);
        dom.hStatEntered.textContent = String(entered);
        dom.hStatExcellent.textContent = String(scores.filter((s) => s >= 8).length);
        dom.hStatGood.textContent = String(scores.filter((s) => s >= 6.5 && s < 8).length);
        dom.hStatAverage.textContent = String(scores.filter((s) => s >= 5 && s < 6.5).length);
        dom.hStatWeak.textContent = String(scores.filter((s) => s < 5).length);
    }

    function renderProgress() {
        const total = state.students.length;
        const entered = countEntered();
        const pct = total === 0 ? 0 : Math.round((entered / total) * 100);

        dom.progressFill.style.width = `${pct}%`;
        dom.progressText.textContent = `${entered} / ${total} học sinh đã nhập điểm (${pct}%)`;

        const scores = state.students.map(getStudentScore).filter(Number.isFinite);
        if (scores.length > 0) {
            const min = Math.min(...scores);
            const max = Math.max(...scores);
            dom.scoreExtremes.textContent = `Thấp nhất: ${formatScore(min)} | Cao nhất: ${formatScore(max)}`;
        } else {
            dom.scoreExtremes.textContent = '';
        }
    }

    function onSearchInput() {
        state.searchText = normalizeText(dom.searchInput.value);
        dom.btnClearSearch.classList.toggle('hidden', !dom.searchInput.value.trim());
        renderTableAndStats();

        if (state.searchFocusTimer) {
            clearTimeout(state.searchFocusTimer);
            state.searchFocusTimer = null;
        }

        if (!state.searchText) {
            return;
        }

        state.searchFocusTimer = setTimeout(() => {
            const rows = getFilteredSortedRows();
            if (state.searchText && rows.length === 1) {
                state.selectedStudentId = rows[0].id;
                renderTableAndStats();
                focusDirectInputForStudent(rows[0].id);
            }
        }, 3000);
    }

    function clearSearch() {
        dom.searchInput.value = '';
        state.searchText = '';
        if (state.searchFocusTimer) {
            clearTimeout(state.searchFocusTimer);
            state.searchFocusTimer = null;
        }
        dom.btnClearSearch.classList.add('hidden');
        renderTableAndStats();
    }

    function getFilteredSortedRows() {
        const rows = state.students.filter((student) => {
            if (state.searchText) {
                const haystack = normalizeText(`${student.name} ${student.maDe || ''}`);
                if (!haystack.includes(state.searchText)) return false;
            }

            const score = getStudentScore(student);
            const hasScore = Number.isFinite(score);

            if (state.activeFilter === 'no-score') return !hasScore;
            if (state.activeFilter === 'excellent') return hasScore && score >= 8;
            if (state.activeFilter === 'good') return hasScore && score >= 6.5 && score < 8;
            if (state.activeFilter === 'average') return hasScore && score >= 5 && score < 6.5;
            if (state.activeFilter === 'weak') return hasScore && score < 5;
            return true;
        });

        rows.sort((a, b) => compareStudents(a, b, state.sortBy, state.sortDir));
        return rows;
    }

    function compareStudents(a, b, col, dir) {
        const order = dir === 'asc' ? 1 : -1;
        if (col === 'socau') {
            return order * ((a.socau || -1) - (b.socau || -1));
        }
        if (col === 'diem') {
            return order * ((getStudentScore(a) || -1) - (getStudentScore(b) || -1));
        }
        return order * String(a.name || '').localeCompare(String(b.name || ''), 'vi', { sensitivity: 'base' });
    }

    function handleTableInput(event) {
        const input = event.target;
        if (!(input instanceof HTMLInputElement)) return;

        const row = input.closest('tr');
        if (!row) return;

        const student = findStudentByRow(row);
        if (!student) return;

        const field = input.dataset.field;
        if (!field) return;

        snapshotBeforeMutation();

        if (field === 'name') {
            student.name = input.value.trimStart();
        }
        if (field === 'maDe') {
            student.maDe = input.value.trim();
        }
        if (field === 'socau') {
            student.socau = parseNullableNumber(input.value);
            if (Number.isFinite(student.socau)) {
                student.socau = clamp(student.socau, 0, state.config.soCauToiDa);
            }
        }
        if (field === 'diem') {
            student.diem = parseNullableNumber(input.value);
            if (Number.isFinite(student.diem)) {
                student.diem = clamp(student.diem, 0, 10);
            }
        }

        // Do not re-render table on each keystroke to avoid losing input focus while typing.
        persistStudents(false);
        renderHeaderStats();
        renderProgress();
        renderChartIfVisible();
        updateUndoRedoState();
    }

    function handleTableKeydown(event) {
        if (event.key !== 'Enter') return;

        const input = event.target;
        if (!(input instanceof HTMLInputElement)) return;

        const row = input.closest('tr');
        if (!row) return;

        const field = input.dataset.field;
        if (field === 'socau') {
            event.preventDefault();

            const student = findStudentByRow(row);
            if (student) {
                const socau = parseNullableNumber(input.value);
                student.socau = Number.isFinite(socau)
                    ? clamp(socau, 0, state.config.soCauToiDa)
                    : NaN;

                if (Number.isFinite(student.socau)) {
                    student.diem = clamp(student.socau * getAutoDiemMoiCau(), 0, 10);
                }

                persistStudents(false);
                renderHeaderStats();
                renderProgress();
                renderChartIfVisible();
                updateUndoRedoState();
            }

            const diemInput = row.querySelector('input[data-field="diem"]');
            if (diemInput instanceof HTMLInputElement) {
                if (student && Number.isFinite(student.diem)) {
                    diemInput.value = formatScore(student.diem);
                }
                diemInput.focus();
                diemInput.select();
            }
            return;
        }

        if (field === 'diem') {
            event.preventDefault();
            persistStudents();

            const currentRow = row;
            const nextRow = currentRow.nextElementSibling;
            if (nextRow instanceof HTMLTableRowElement) {
                const nextDiem = nextRow.querySelector('input[data-field="diem"]');
                if (nextDiem instanceof HTMLInputElement) {
                    nextDiem.focus();
                    nextDiem.select();
                }
            }
        }
    }

    function handleTableClick(event) {
        const rowForSelection = event.target.closest('tr[data-student-id]');
        if (rowForSelection) {
            state.selectedStudentId = rowForSelection.getAttribute('data-student-id');
            renderTableAndStats();
        }

        const btn = event.target.closest('button[data-action]');
        if (!btn) return;

        const row = btn.closest('tr');
        if (!row) return;
        const student = findStudentByRow(row);
        if (!student) return;

        const action = btn.dataset.action;
        if (action === 'focus-score') {
            const targetInput = row.querySelector('input[data-field="socau"]');
            if (targetInput instanceof HTMLElement) {
                targetInput.focus();
                targetInput.select();
            }
        }

        if (action === 'delete') {
            confirmDialog('Xóa học sinh', `Bạn có chắc muốn xóa "${student.name || 'Học sinh'}"?`)
                .then((ok) => {
                    if (!ok) return;
                    snapshotBeforeMutation();
                    state.students = state.students.filter((x) => x.id !== student.id);
                    persistStudents();
                    renderTableAndStats();
                    updateUndoRedoState();
                    showToast('Đã xóa học sinh', 'success');
                });
        }
    }

    function bindImportModalEvents() {
        byId('btnChooseFile').addEventListener('click', () => dom.fileInput.click());
        dom.fileInput.addEventListener('change', () => {
            const file = dom.fileInput.files && dom.fileInput.files[0];
            if (file) handleImportFile(file);
        });

        byId('btnClearFile').addEventListener('click', clearPendingImport);
        dom.sheetSelector.addEventListener('change', () => buildImportPreview());
        dom.mapNameCol.addEventListener('change', () => buildImportPreview());
        dom.mapSkipRows.addEventListener('input', () => buildImportPreview());
        dom.btnDoImport.addEventListener('click', doImportStudents);

        ['dragenter', 'dragover'].forEach((eventName) => {
            dom.dropZone.addEventListener(eventName, (event) => {
                event.preventDefault();
                event.stopPropagation();
                dom.dropZone.classList.add('drag-over');
            });
        });
        ['dragleave', 'drop'].forEach((eventName) => {
            dom.dropZone.addEventListener(eventName, (event) => {
                event.preventDefault();
                event.stopPropagation();
                dom.dropZone.classList.remove('drag-over');
            });
        });
        dom.dropZone.addEventListener('drop', (event) => {
            const file = event.dataTransfer && event.dataTransfer.files && event.dataTransfer.files[0];
            if (file) handleImportFile(file);
        });
    }

    async function handleImportFile(file) {
        if (!window.XLSX) {
            showToast('Thiếu thư viện đọc Excel (XLSX).', 'error');
            return;
        }

        try {
            const bytes = await file.arrayBuffer();
            const workbook = XLSX.read(bytes, { type: 'array' });

            state.pendingWorkbook = workbook;
            state.pendingFileName = file.name;
            state.pendingSheets = workbook.SheetNames.slice();

            dom.fileInfoBar.classList.remove('hidden');
            dom.importFileName.textContent = file.name;
            dom.importFileSize.textContent = `${(file.size / 1024).toFixed(1)} KB`;

            dom.sheetSelector.innerHTML = state.pendingSheets
                .map((name) => `<option value="${escapeAttr(name)}">${escapeHtml(name)}</option>`)
                .join('');

            buildImportPreview();
            dom.importPreview.classList.remove('hidden');
            dom.btnDoImport.classList.remove('hidden');
        } catch (error) {
            console.error(error);
            showToast('Không thể đọc file Excel. Vui lòng kiểm tra định dạng.', 'error');
        }
    }

    function clearPendingImport() {
        state.pendingWorkbook = null;
        state.pendingSheets = [];
        state.importRows = [];
        dom.fileInput.value = '';
        dom.fileInfoBar.classList.add('hidden');
        dom.importPreview.classList.add('hidden');
        dom.btnDoImport.classList.add('hidden');
        dom.previewTableWrap.innerHTML = '';
        dom.importSummary.textContent = '';
        dom.importCountSpan.textContent = '0';
    }

    function buildImportPreview() {
        if (!state.pendingWorkbook) return;

        const sheetName = dom.sheetSelector.value || state.pendingSheets[0];
        const sheet = state.pendingWorkbook.Sheets[sheetName];
        if (!sheet) return;

        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
        if (!Array.isArray(rows) || rows.length === 0) {
            dom.previewTableWrap.innerHTML = '<p class="empty-hint">Sheet trống.</p>';
            state.importRows = [];
            return;
        }

        const headerRow = rows[0].map((x) => String(x || '').trim());
        const headerOptions = headerRow
            .map((name, idx) => ({ name: name || `Cột ${idx + 1}`, idx }));

        const currentNameCol = parseInt(dom.mapNameCol.value, 10);
        const selectedNameIndex = Number.isInteger(currentNameCol)
            ? currentNameCol
            : resolveNameColumnIndex(headerOptions);
        dom.mapNameCol.innerHTML = headerOptions
            .map((item) => `
                <option value="${item.idx}" ${item.idx === selectedNameIndex ? 'selected' : ''}>${escapeHtml(item.name)}</option>
            `)
            .join('');

        const skipRows = clamp(parseInt(dom.mapSkipRows.value, 10) || 0, 0, 200);
        const nameCol = parseInt(dom.mapNameCol.value, 10);
        const importRows = [];

        for (let i = 1 + skipRows; i < rows.length; i += 1) {
            const row = rows[i] || [];
            const nameRaw = row[nameCol];
            const name = String(nameRaw || '').trim();
            if (!name) continue;

            const parsed = {
                id: makeId(),
                name,
                maDe: extractByHeader(headerRow, row, ['ma de', 'm de', 'm d', 'ma']) || '',
                socau: parseNullableNumber(extractByHeader(headerRow, row, ['so cau', 's cau', 'socau'])),
                diem: parseNullableNumber(extractByHeader(headerRow, row, ['diem', 'i m', 'score']))
            };

            importRows.push(normalizeStudent(parsed));
        }

        state.importRows = importRows;
        dom.importCountSpan.textContent = String(importRows.length);
        dom.importSummary.textContent = `Sẵn sàng import ${importRows.length} học sinh từ sheet ${sheetName}.`;

        const previewRows = rows.slice(0, Math.min(rows.length, 8));
        const previewHtml = `
            <table class="preview-tbl">
                <tbody>
                    ${previewRows.map((r, rowIdx) => `
                        <tr>
                            ${r.slice(0, 8).map((cell, colIdx) => {
                                const tag = rowIdx === 0 ? 'th' : 'td';
                                const mark = rowIdx === 0 && colIdx === nameCol ? ' style="background: rgba(16,185,129,0.2);"' : '';
                                return `<${tag}${mark}>${escapeHtml(String(cell || ''))}</${tag}>`;
                            }).join('')}
                        </tr>
                    `).join('')}
                </tbody>
            </table>
        `;

        dom.previewTableWrap.innerHTML = previewHtml;
    }

    function resolveNameColumnIndex(headerOptions) {
        const configured = normalizeText(state.config.tenCotHS || '');
        const exact = headerOptions.find((x) => normalizeText(x.name) === configured);
        if (exact) return exact.idx;

        const fallback = headerOptions.find((x) => {
            const value = normalizeText(x.name);
            return value.includes('ho va ten') || value.includes('ten hoc sinh') || value.includes('student');
        });

        return fallback ? fallback.idx : 0;
    }

    function doImportStudents() {
        if (!state.importRows.length) {
            showToast('Không có dữ liệu để import.', 'error');
            return;
        }

        snapshotBeforeMutation();
        const existingKey = new Set(state.students.map((x) => normalizeText(x.name)));
        let added = 0;

        state.importRows.forEach((row) => {
            const key = normalizeText(row.name);
            if (!key) return;
            if (existingKey.has(key)) return;
            existingKey.add(key);
            state.students.push(normalizeStudent(row));
            added += 1;
        });

        persistStudents();
        setCurrentFileName(state.pendingFileName || 'Excel import');
        addRecentFile(state.pendingFileName || `Import ${new Date().toLocaleString('vi-VN')}`);
        clearPendingImport();
        closeModal(dom.modalImport);
        renderTableAndStats();
        updateUndoRedoState();
        showToast(`Đã import ${added} học sinh`, 'success');
    }

    function bindExportModalEvents() {
        byId('btnExportExcel').addEventListener('click', async () => {
            await exportExcel();
        });
        byId('btnExportPdf').addEventListener('click', async () => {
            await exportPdf();
        });

        dom.cbSaveToCloud.addEventListener('change', () => {
            dom.cloudSubOptions.classList.toggle('hidden', !dom.cbSaveToCloud.checked);
        });
    }

    function openExportModal() {
        const total = state.students.length;
        const entered = countEntered();
        const scores = state.students.map(getStudentScore).filter(Number.isFinite);
        const avg = scores.length ? scores.reduce((a, b) => a + b, 0) / scores.length : NaN;

        dom.expStudentCount.textContent = String(total);
        dom.expEnteredCount.textContent = String(entered);
        dom.expAvgScore.textContent = Number.isFinite(avg) ? formatScore(avg) : '—';

        if (!dom.expRecordName.value) {
            dom.expRecordName.value = state.currentFileName ? stripExtension(state.currentFileName) : '';
        }

        openModal(dom.modalExport);
    }

    async function exportExcel() {
        if (!window.XLSX) {
            showToast('Thiếu thư viện xuất Excel.', 'error');
            return;
        }

        if (!state.students.length) {
            showToast('Chưa có dữ liệu để xuất.', 'error');
            return;
        }

        const fileBase = buildExportName();
        const rows = state.students.map((s, idx) => ({
            STT: idx + 1,
            'Họ và tên': s.name,
            'Mã đề': s.maDe || '',
            'Số câu đúng': Number.isFinite(s.socau) ? s.socau : '',
            'Điểm': Number.isFinite(getStudentScore(s)) ? getStudentScore(s) : '',
            'Xếp loại': getRank(getStudentScore(s))
        }));

        const ws = XLSX.utils.json_to_sheet(rows);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Bang diem');
        XLSX.writeFile(wb, `${fileBase}.xlsx`);

        const cloudOk = await maybeSaveToCloud(fileBase);
        if (cloudOk === false) return;

        addRecentFile(`${fileBase}.xlsx`);
        closeModal(dom.modalExport);
        showToast('Đã xuất Excel thành công', 'success');
    }

    async function exportPdf() {
        if (!window.jspdf || !window.jspdf.jsPDF) {
            showToast('Thiếu thư viện jsPDF.', 'error');
            return;
        }

        if (!state.students.length) {
            showToast('Chưa có dữ liệu để xuất.', 'error');
            return;
        }

        const fileBase = buildExportName();
        const doc = new window.jspdf.jsPDF({ orientation: 'p', unit: 'pt', format: 'a4' });
        doc.setFontSize(14);
        doc.text(`Bảng điểm - ${fileBase}`, 40, 40);
        doc.setFontSize(10);
        doc.text(`Ngày xuất: ${new Date().toLocaleString('vi-VN')}`, 40, 58);

        const body = state.students.map((s, idx) => [
            idx + 1,
            s.name,
            s.maDe || '',
            Number.isFinite(s.socau) ? s.socau : '',
            Number.isFinite(getStudentScore(s)) ? formatScore(getStudentScore(s)) : '',
            getRank(getStudentScore(s))
        ]);

        if (typeof doc.autoTable === 'function') {
            doc.autoTable({
                startY: 72,
                head: [['#', 'Họ và tên', 'Mã đề', 'Số câu', 'Điểm', 'Xếp loại']],
                body
            });
        }

        doc.save(`${fileBase}.pdf`);

        const cloudOk = await maybeSaveToCloud(fileBase);
        if (cloudOk === false) return;

        addRecentFile(`${fileBase}.pdf`);
        closeModal(dom.modalExport);
        showToast('Đã xuất PDF thành công', 'success');
    }

    async function maybeSaveToCloud(fileBase) {
        if (!dom.cbSaveToCloud.checked) return true;

        if (!state.proxyUrl.trim()) {
            showToast('Chưa cấu hình Proxy URL.', 'error');
            return false;
        }

        if (state.cloudBusy) return false;
        state.cloudBusy = true;

        try {
            const payload = {
                id: `record_${Date.now()}`,
                recordName: dom.expRecordName.value.trim() || fileBase,
                recordClass: dom.expClassName.value.trim(),
                lastModified: new Date().toISOString(),
                scoreColumns: ['Số câu đúng', 'Điểm'],
                students: state.students.map((s) => ({
                    name: s.name,
                    scores: {
                        'Số câu đúng': Number.isFinite(s.socau) ? String(s.socau) : '',
                        'Điểm': Number.isFinite(getStudentScore(s)) ? formatScore(getStudentScore(s)) : ''
                    }
                })),
                isPublic: !!dom.cbIsPublic.checked
            };

            const response = await fetch(`${state.proxyUrl.replace(/\/$/, '')}/records`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });

            if (!response.ok) {
                throw new Error(`HTTP ${response.status}`);
            }

            showToast('Đã lưu bản ghi lên cloud', 'success');
            return true;
        } catch (error) {
            console.error(error);
            showToast('Lưu cloud thất bại. Kiểm tra Proxy URL.', 'error');
            return false;
        } finally {
            state.cloudBusy = false;
        }
    }

    function bindStudentModalEvents() {
        byId('btnSaveStudent').addEventListener('click', saveStudentFromModal);

        dom.stuSoCau.addEventListener('input', () => {
            if (dom.stuDiem.value.trim()) return;
            const socau = parseNullableNumber(dom.stuSoCau.value);
            if (!Number.isFinite(socau)) return;
            const diem = clamp(socau * getAutoDiemMoiCau(), 0, 10);
            dom.stuDiem.value = formatScore(diem);
        });
    }

    function openStudentModal(index) {
        state.selectedIndex = index;

        rebuildStudentMaDEOptions();

        if (index >= 0) {
            const s = state.students[index];
            dom.studentModalTitle.textContent = '✏️ Sửa học sinh';
            dom.stuName.value = s.name || '';
            dom.stuMaDE.value = s.maDe || '';
            dom.stuSoCau.value = Number.isFinite(s.socau) ? String(s.socau) : '';
            dom.stuDiem.value = Number.isFinite(s.diem) ? formatScore(s.diem) : '';
        } else {
            dom.studentModalTitle.textContent = '➕ Thêm học sinh';
            dom.stuName.value = '';
            dom.stuMaDE.value = '';
            dom.stuSoCau.value = '';
            dom.stuDiem.value = '';
        }

        openModal(dom.modalStudent);
        dom.stuName.focus();
    }

    function saveStudentFromModal() {
        const name = dom.stuName.value.trim();
        if (!name) {
            showToast('Vui lòng nhập họ và tên.', 'error');
            dom.stuName.focus();
            return;
        }

        const payload = normalizeStudent({
            id: state.selectedIndex >= 0 ? state.students[state.selectedIndex].id : makeId(),
            name,
            maDe: dom.stuMaDE.value.trim(),
            socau: parseNullableNumber(dom.stuSoCau.value),
            diem: parseNullableNumber(dom.stuDiem.value)
        });

        snapshotBeforeMutation();
        if (state.selectedIndex >= 0) {
            state.students[state.selectedIndex] = payload;
        } else {
            state.students.push(payload);
        }

        persistStudents();
        closeModal(dom.modalStudent);
        renderTableAndStats();
        updateUndoRedoState();
        showToast('Đã lưu học sinh', 'success');
    }

    function bindGenericModalEvents() {
        document.querySelectorAll('[data-close]').forEach((element) => {
            element.addEventListener('click', () => {
                const modalId = element.getAttribute('data-close');
                const modal = byId(modalId);
                if (modal) closeModal(modal);
            });
        });

        document.querySelectorAll('.modal-overlay').forEach((overlay) => {
            overlay.addEventListener('click', (event) => {
                if (event.target === overlay) closeModal(overlay);
            });
        });
    }

    function bindKeyboardShortcuts() {
        document.addEventListener('keydown', (event) => {
            if (event.ctrlKey && event.key.toLowerCase() === 'f') {
                event.preventDefault();
                dom.searchInput.focus();
                dom.searchInput.select();
            }

            if (event.ctrlKey && !event.altKey && !event.shiftKey && !event.metaKey && event.key.toLowerCase() === 'g') {
                event.preventDefault();
                triggerDirectScoreInput('socau');
            }

            if (event.ctrlKey && !event.altKey && !event.shiftKey && !event.metaKey && event.key.toLowerCase() === 'b') {
                event.preventDefault();
                triggerDirectScoreInput('diem');
            }

            if (event.ctrlKey && event.key.toLowerCase() === 'z') {
                event.preventDefault();
                undoChange();
            }

            if (event.ctrlKey && event.key.toLowerCase() === 'y') {
                event.preventDefault();
                redoChange();
            }
        });
    }

    function saveExamConfig() {
        const soCau = clamp(parseInt(dom.cfgSoCauToiDa.value, 10) || DEFAULT_CONFIG.soCauToiDa, 1, 500);
        const danhSach = dom.cfgDanhSachMaDE.value
            .split(',')
            .map((x) => x.trim())
            .filter(Boolean);

        state.config = {
            soCauToiDa: soCau,
            danhSachMaDE: danhSach,
            tenCotHS: dom.cfgTenCotHS.value.trim() || DEFAULT_CONFIG.tenCotHS
        };

        localStorage.setItem(STORAGE_KEYS.config, JSON.stringify(state.config));
        rebuildStudentMaDEOptions();
        renderTableAndStats();
        showToast('Đã lưu cấu hình bài thi', 'success');
    }

    function saveCloudConfig() {
        state.proxyUrl = dom.cfgProxyUrl.value.trim();
        localStorage.setItem(STORAGE_KEYS.proxy, state.proxyUrl);
        showToast('Đã lưu cấu hình cloud', 'success');
    }

    function rebuildStudentMaDEOptions() {
        const options = ['<option value="">— Không có —</option>']
            .concat(state.config.danhSachMaDE.map((ma) => `<option value="${escapeAttr(ma)}">${escapeHtml(ma)}</option>`));
        dom.stuMaDE.innerHTML = options.join('');
    }

    function toggleSidebar() {
        state.ui.sidebarCollapsed = !state.ui.sidebarCollapsed;
        dom.sidebar.classList.toggle('collapsed', state.ui.sidebarCollapsed);
        dom.appBody.classList.toggle('sidebar-collapsed', state.ui.sidebarCollapsed);
        localStorage.setItem(STORAGE_KEYS.ui, JSON.stringify(state.ui));
    }

    function toggleChart() {
        setChartVisible(dom.chartSection.classList.contains('hidden'));
    }

    function setChartVisible(visible) {
        dom.chartSection.classList.toggle('hidden', !visible);
        if (visible) renderChart();
    }

    function renderChartIfVisible() {
        if (!dom.chartSection.classList.contains('hidden')) {
            renderChart();
        }
    }

    function renderChart() {
        if (!window.Chart) return;
        const scores = state.students.map(getStudentScore).filter(Number.isFinite);

        const buckets = {
            '0-4.9': 0,
            '5-6.4': 0,
            '6.5-7.9': 0,
            '8-10': 0
        };

        scores.forEach((s) => {
            if (s < 5) buckets['0-4.9'] += 1;
            else if (s < 6.5) buckets['5-6.4'] += 1;
            else if (s < 8) buckets['6.5-7.9'] += 1;
            else buckets['8-10'] += 1;
        });

        const avg = scores.length ? scores.reduce((a, b) => a + b, 0) / scores.length : NaN;
        dom.chartAvgBadge.textContent = Number.isFinite(avg) ? `Điểm TB: ${formatScore(avg)}` : 'Điểm TB: —';

        if (state.chart) {
            state.chart.destroy();
        }

        state.chart = new Chart(dom.scoreChart, {
            type: 'bar',
            data: {
                labels: Object.keys(buckets),
                datasets: [{
                    label: 'Số học sinh',
                    data: Object.values(buckets),
                    backgroundColor: ['#ef4444', '#3b82f6', '#f59e0b', '#10b981']
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                scales: {
                    y: { beginAtZero: true, ticks: { precision: 0 } }
                }
            }
        });
    }

    function toggleMaDE() {
        state.showMaDE = !state.showMaDE;
        renderTableAndStats();
    }

    function snapshotBeforeMutation() {
        state.undoStack.push(JSON.stringify(state.students));
        if (state.undoStack.length > state.maxHistory) {
            state.undoStack.shift();
        }
        state.redoStack = [];
    }

    function undoChange() {
        if (!state.undoStack.length) return;
        state.redoStack.push(JSON.stringify(state.students));
        const prev = state.undoStack.pop();
        state.students = JSON.parse(prev).map(normalizeStudent);
        persistStudents(false);
        renderTableAndStats();
        updateUndoRedoState();
        showToast('Đã hoàn tác', 'success');
    }

    function redoChange() {
        if (!state.redoStack.length) return;
        state.undoStack.push(JSON.stringify(state.students));
        const next = state.redoStack.pop();
        state.students = JSON.parse(next).map(normalizeStudent);
        persistStudents(false);
        renderTableAndStats();
        updateUndoRedoState();
        showToast('Đã làm lại', 'success');
    }

    function updateUndoRedoState() {
        dom.btnUndo.disabled = state.undoStack.length === 0;
        dom.btnRedo.disabled = state.redoStack.length === 0;
    }

    function persistStudents(render = true) {
        localStorage.setItem(STORAGE_KEYS.students, JSON.stringify(state.students));
        if (render) renderTableAndStats();
    }

    function countEntered() {
        return state.students.filter((s) => Number.isFinite(getStudentScore(s))).length;
    }

    function getAutoDiemMoiCau() {
        const totalQuestions = clamp(parseInt(state.config.soCauToiDa, 10) || DEFAULT_CONFIG.soCauToiDa, 1, 500);
        return 10 / totalQuestions;
    }

    function getStudentScore(student) {
        if (Number.isFinite(student.diem)) {
            return clamp(student.diem, 0, 10);
        }
        if (Number.isFinite(student.socau)) {
            return clamp(student.socau * getAutoDiemMoiCau(), 0, 10);
        }
        return NaN;
    }

    function triggerDirectScoreInput(targetField = 'socau') {
        const rows = getFilteredSortedRows();
        if (!rows.length) {
            showToast('Không có kết quả để nhập nhanh.', 'warning');
            return;
        }

        if (rows.length === 1) {
            state.selectedStudentId = rows[0].id;
            renderTableAndStats();
            focusDirectInputForStudent(rows[0].id, targetField);
            return;
        }

        if (!state.selectedStudentId || !rows.some((x) => x.id === state.selectedStudentId)) {
            showToast('Có nhiều kết quả. Hãy bấm chọn 1 học sinh rồi dùng Ctrl+G hoặc Ctrl+B.', 'info');
            return;
        }

        focusDirectInputForStudent(state.selectedStudentId, targetField);
    }

    function focusDirectInputForStudent(studentId, targetField = 'socau') {
        const row = dom.tableBody.querySelector(`tr[data-student-id="${CSS.escape(studentId)}"]`);
        if (!row) return;

        const soCauInput = row.querySelector('input[data-field="socau"]');
        const diemInput = row.querySelector('input[data-field="diem"]');
        const preferred = targetField === 'diem' ? diemInput : soCauInput;
        const fallback = targetField === 'diem' ? soCauInput : diemInput;
        const target = (preferred instanceof HTMLInputElement && !preferred.disabled)
            ? preferred
            : fallback;

        if (target instanceof HTMLInputElement) {
            target.focus();
            target.select();
        }
    }

    function getRank(score) {
        if (!Number.isFinite(score)) return '—';
        if (score >= 8) return 'Giỏi';
        if (score >= 6.5) return 'Khá';
        if (score >= 5) return 'Trung bình';
        return 'Yếu';
    }

    function normalizeStudent(input) {
        return {
            id: input.id || makeId(),
            name: String(input.name || '').trim(),
            maDe: String(input.maDe || '').trim(),
            socau: parseNullableNumber(input.socau),
            diem: parseNullableNumber(input.diem)
        };
    }

    function setCurrentFileName(fileName) {
        state.currentFileName = fileName;
        localStorage.setItem(STORAGE_KEYS.fileName, fileName);
        dom.fileNameDisplay.textContent = fileName || 'Chưa có file';
    }

    function addRecentFile(fileName) {
        if (!fileName) return;
        const timestamp = Date.now();
        const item = { fileName, timestamp };
        state.recentFiles = [item]
            .concat(state.recentFiles.filter((x) => x.fileName !== fileName))
            .slice(0, 8);
        localStorage.setItem(STORAGE_KEYS.recent, JSON.stringify(state.recentFiles));
        renderRecentFiles();
    }

    function renderRecentFiles() {
        if (!state.recentFiles.length) {
            dom.recentFilesList.innerHTML = '<p class="empty-hint">Chưa có file nào</p>';
            return;
        }

        dom.recentFilesList.innerHTML = state.recentFiles
            .map((item) => {
                const date = new Date(item.timestamp);
                return `
                    <button class="recent-file-item" data-file-name="${escapeAttr(item.fileName)}">
                        <span class="rf-name">${escapeHtml(item.fileName)}</span>
                        <small class="rf-time">${date.toLocaleString('vi-VN')}</small>
                    </button>
                `;
            })
            .join('');

        dom.recentFilesList.querySelectorAll('.recent-file-item').forEach((button) => {
            button.addEventListener('click', () => {
                const fileName = button.getAttribute('data-file-name') || '';
                setCurrentFileName(fileName);
                showToast(`Đang làm việc với: ${fileName}`, 'success');
            });
        });
    }

    function buildExportName() {
        const customName = dom.expRecordName.value.trim();
        if (customName) return customName;
        if (state.currentFileName) return stripExtension(state.currentFileName);
        return `Bang-diem-${formatDateForFile(new Date())}`;
    }

    function stripExtension(fileName) {
        return fileName.replace(/\.[^/.]+$/, '');
    }

    function formatDateForFile(date) {
        const yyyy = date.getFullYear();
        const mm = String(date.getMonth() + 1).padStart(2, '0');
        const dd = String(date.getDate()).padStart(2, '0');
        const hh = String(date.getHours()).padStart(2, '0');
        const min = String(date.getMinutes()).padStart(2, '0');
        return `${yyyy}${mm}${dd}-${hh}${min}`;
    }

    function confirmDialog(title, message) {
        return new Promise((resolve) => {
            dom.confirmTitle.textContent = title;
            dom.confirmMsg.textContent = message;
            openModal(dom.modalConfirm);

            const onYes = () => {
                cleanup();
                resolve(true);
            };
            const onNo = () => {
                cleanup();
                resolve(false);
            };

            const cleanup = () => {
                dom.btnConfirmYes.removeEventListener('click', onYes);
                dom.btnConfirmNo.removeEventListener('click', onNo);
                closeModal(dom.modalConfirm);
            };

            dom.btnConfirmYes.addEventListener('click', onYes);
            dom.btnConfirmNo.addEventListener('click', onNo);
        });
    }

    function openModal(modal) {
        if (!modal) return;
        modal.classList.remove('hidden');
    }

    function closeModal(modal) {
        if (!modal) return;
        modal.classList.add('hidden');
    }

    function showToast(message, type) {
        if (!dom.toastContainer) return;
        const toast = document.createElement('div');
        toast.className = `toast toast-${type || 'info'}`;
        toast.textContent = message;
        dom.toastContainer.appendChild(toast);
        setTimeout(() => {
            toast.classList.add('toast-out');
            setTimeout(() => toast.remove(), 180);
        }, 2200);
    }

    function findStudentByRow(row) {
        const id = row.getAttribute('data-student-id');
        if (!id) return null;
        return state.students.find((x) => x.id === id) || null;
    }

    function parseNullableNumber(value) {
        if (value === null || value === undefined) return NaN;
        const text = String(value).trim().replace(',', '.');
        if (!text) return NaN;
        const n = Number(text);
        return Number.isFinite(n) ? n : NaN;
    }

    function extractByHeader(headerRow, row, aliases) {
        const normalizedAliases = aliases.map(normalizeText);
        for (let i = 0; i < headerRow.length; i += 1) {
            const header = normalizeText(headerRow[i]);
            if (!header) continue;
            if (normalizedAliases.some((alias) => header.includes(alias))) {
                return row[i];
            }
        }
        return '';
    }

    function normalizeText(value) {
        return String(value || '')
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .toLowerCase()
            .trim();
    }

    function loadJson(key, fallback) {
        try {
            const raw = localStorage.getItem(key);
            if (!raw) return fallback;
            return JSON.parse(raw);
        } catch (error) {
            return fallback;
        }
    }

    function byId(id) {
        return document.getElementById(id);
    }

    function clamp(n, min, max) {
        return Math.min(max, Math.max(min, n));
    }

    function formatScore(value) {
        return Number(value).toFixed(2).replace(/\.00$/, '').replace(/(\.\d)0$/, '$1');
    }

    function makeId() {
        return `st_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 8)}`;
    }

    function escapeHtml(value) {
        return String(value)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
    }

    function escapeAttr(value) {
        return escapeHtml(value);
    }
})();
