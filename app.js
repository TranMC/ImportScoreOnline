(function () {
    'use strict';

    const RUNTIME_PROXY_URL = (() => {
        if (window.APP_CONFIG && typeof window.APP_CONFIG.PROXY_URL === 'string') {
            return window.APP_CONFIG.PROXY_URL.trim();
        }

        // Backward compatibility for older deployments that expose `const CONFIG`.
        if (typeof CONFIG !== 'undefined' && CONFIG && typeof CONFIG.PROXY_URL === 'string') {
            return CONFIG.PROXY_URL.trim();
        }

        return '';
    })();

    const STORAGE_KEYS = {
        config: 'importscore.config.v1',
        proxy: 'importscore.proxy.v1',
        students: 'importscore.students.v1',
        recent: 'importscore.recent-files.v1',
        fileName: 'importscore.file-name.v1',
        currentMaDE: 'importscore.current-made.v1',
        ui: 'importscore.ui.v1'
    };

    const ALLOWED_CHART_TYPES = new Set(['bar', 'line', 'doughnut', 'pie', 'polarArea', 'radar', 'barHorizontal']);
    const ALLOWED_SINGLE_SEARCH_FOCUS_TARGETS = new Set(['none', 'socau', 'made']);

    const DEFAULT_CONFIG = {
        soCauToiDa: 40,
        diemToiDa: 10,
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
        chartRenderTimer: null,
        chartDrag: {
            active: false,
            pointerId: null,
            offsetX: 0,
            offsetY: 0
        },
        currentMaDE: (localStorage.getItem(STORAGE_KEYS.currentMaDE) || '').trim(),
        config: loadJson(STORAGE_KEYS.config, DEFAULT_CONFIG),
        proxyUrl: RUNTIME_PROXY_URL || localStorage.getItem(STORAGE_KEYS.proxy) || '',
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
        dom.cfgDiemToiDa = byId('cfgDiemToiDa');
        dom.cfgDanhSachMaDE = byId('cfgDanhSachMaDE');
        dom.cfgTenCotHS = byId('cfgTenCotHS');
        dom.cfgSearchSingleFocusTarget = byId('cfgSearchSingleFocusTarget');
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
        dom.chartDragHandle = byId('chartDragHandle');
        dom.scoreChart = byId('scoreChart');
        dom.chartAvgBadge = byId('chartAvgBadge');
        dom.chartTypeSelect = byId('chartTypeSelect');

        dom.modalReport = byId('modalReport');
        dom.reportSection = byId('reportSection');
        dom.reportSubtitle = byId('reportSubtitle');
        dom.reportPrintable = byId('reportPrintable');
        dom.repTotal = byId('repTotal');
        dom.repEntered = byId('repEntered');
        dom.repAvg = byId('repAvg');
        dom.repMedian = byId('repMedian');
        dom.repStd = byId('repStd');
        dom.repPassRate = byId('repPassRate');
        dom.repDistExcellent = byId('repDistExcellent');
        dom.repDistGood = byId('repDistGood');
        dom.repDistAverage = byId('repDistAverage');
        dom.repDistWeak = byId('repDistWeak');
        dom.repDistExcellentText = byId('repDistExcellentText');
        dom.repDistGoodText = byId('repDistGoodText');
        dom.repDistAverageText = byId('repDistAverageText');
        dom.repDistWeakText = byId('repDistWeakText');
        dom.repTopList = byId('repTopList');
        dom.repBottomList = byId('repBottomList');
        dom.repByMaDEBody = byId('repByMaDEBody');
        dom.repScoreBody = byId('repScoreBody');

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
        dom.mapMaDeCol = byId('mapMaDeCol');
        dom.mapDiemCol = byId('mapDiemCol');
        dom.mapSkipRows = byId('mapSkipRows');
        dom.cbClearBeforeImport = byId('cbClearBeforeImport');
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
        dom.currentMaDESelect = byId('currentMaDESelect');
    }

    function bindEvents() {
        byId('btnImport').addEventListener('click', () => openModal(dom.modalImport));
        byId('btnExport').addEventListener('click', openExportModal);
        byId('btnToggleSidebar').addEventListener('click', toggleSidebar);
        byId('btnSaveConfig').addEventListener('click', saveExamConfig);
        byId('btnSaveCloudConfig').addEventListener('click', saveCloudConfig);
        byId('btnToggleChart').addEventListener('click', toggleChart);
        byId('btnCloseChart').addEventListener('click', () => setChartVisible(false));
        byId('btnOpenAdvancedReport').addEventListener('click', () => setAdvancedReportVisible(true));
        byId('btnCloseAdvancedReport').addEventListener('click', () => setAdvancedReportVisible(false));
        byId('btnExportReportPdf').addEventListener('click', async () => {
            await exportPdf();
        });
        byId('btnToggleMaDE').addEventListener('click', toggleMaDE);
        byId('btnAddStudent').addEventListener('click', () => openStudentModal(-1));

        dom.searchInput.addEventListener('input', onSearchInput);
        dom.btnClearSearch.addEventListener('click', clearSearch);

        dom.btnUndo.addEventListener('click', undoChange);
        dom.btnRedo.addEventListener('click', redoChange);
        dom.chartTypeSelect.addEventListener('change', () => {
            const chartType = ALLOWED_CHART_TYPES.has(dom.chartTypeSelect.value)
                ? dom.chartTypeSelect.value
                : 'bar';
            state.ui.chartType = chartType;
            localStorage.setItem(STORAGE_KEYS.ui, JSON.stringify(state.ui));
            renderChartIfVisible({ immediate: true });
        });
        dom.currentMaDESelect.addEventListener('change', () => {
            state.currentMaDE = dom.currentMaDESelect.value.trim();
            localStorage.setItem(STORAGE_KEYS.currentMaDE, state.currentMaDE);
            const valueLabel = state.currentMaDE || 'Không có';
            showToast(`Đã chọn mã đề hiện hành: ${valueLabel}`, 'info');
        });
        dom.cfgSearchSingleFocusTarget.addEventListener('change', () => {
            const nextTarget = ALLOWED_SINGLE_SEARCH_FOCUS_TARGETS.has(dom.cfgSearchSingleFocusTarget.value)
                ? dom.cfgSearchSingleFocusTarget.value
                : 'none';
            state.ui.searchSingleFocusTarget = nextTarget;
            localStorage.setItem(STORAGE_KEYS.ui, JSON.stringify(state.ui));
        });

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
        dom.tableBody.addEventListener('change', handleTableChange);
        dom.tableBody.addEventListener('keydown', handleTableKeydown);
        dom.tableBody.addEventListener('click', handleTableClick);

        bindImportModalEvents();
        bindExportModalEvents();
        bindStudentModalEvents();
        bindGenericModalEvents();
        bindChartDragEvents();
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

        if (!ALLOWED_CHART_TYPES.has(state.ui.chartType)) {
            state.ui.chartType = 'bar';
            localStorage.setItem(STORAGE_KEYS.ui, JSON.stringify(state.ui));
        }

        // Migrate legacy boolean setting to new explicit focus target mode.
        if (typeof state.ui.autoFocusSingleResult === 'boolean') {
            state.ui.searchSingleFocusTarget = state.ui.autoFocusSingleResult ? 'socau' : 'none';
            delete state.ui.autoFocusSingleResult;
            localStorage.setItem(STORAGE_KEYS.ui, JSON.stringify(state.ui));
        }

        if (!ALLOWED_SINGLE_SEARCH_FOCUS_TARGETS.has(state.ui.searchSingleFocusTarget)) {
            state.ui.searchSingleFocusTarget = 'none';
            localStorage.setItem(STORAGE_KEYS.ui, JSON.stringify(state.ui));
        }

        if (!state.ui.chartPosition
            || !Number.isFinite(state.ui.chartPosition.left)
            || !Number.isFinite(state.ui.chartPosition.top)) {
            state.ui.chartPosition = null;
            localStorage.setItem(STORAGE_KEYS.ui, JSON.stringify(state.ui));
        }

        dom.cfgSoCauToiDa.value = String(state.config.soCauToiDa);
        dom.cfgDiemToiDa.value = String(state.config.diemToiDa);
        dom.cfgDanhSachMaDE.value = state.config.danhSachMaDE.join(',');
        dom.cfgTenCotHS.value = state.config.tenCotHS;
        dom.cfgSearchSingleFocusTarget.value = state.ui.searchSingleFocusTarget;
        dom.cfgProxyUrl.value = state.proxyUrl;
        dom.chartTypeSelect.value = state.ui.chartType;

        applySavedChartPosition();
        rebuildMaDEOptions();
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
            const rowRankClass = getRowRankClass(score);
            const badgeClass = getBadgeClass(score);
            const maDeText = student.maDe ? escapeHtml(student.maDe) : '—';

            return `
                <tr data-student-id="${escapeAttr(student.id)}" class="${rowRankClass} ${state.selectedStudentId === student.id ? 'row-active' : ''}">
                    <td class="td-stt">${renderIndex + 1}</td>
                    <td class="td-name">
                        <input class="cell-input" data-field="name" value="${escapeAttr(student.name)}" placeholder="Họ và tên">
                    </td>
                    <td class="td-made made-col ${state.showMaDE ? '' : 'hidden'}">
                        <span class="made-pill">${maDeText}</span>
                    </td>
                    <td class="td-socau">
                        <input class="cell-input" data-field="socau" type="number" min="0" max="${state.config.soCauToiDa}" value="${escapeAttr(soCau)}" placeholder="Số câu">
                    </td>
                    <td class="td-diem">
                        <input class="cell-input" data-field="diem" type="number" min="0" max="10" step="0.01" value="${escapeAttr(scoreText)}" placeholder="Điểm">
                    </td>
                    <td class="td-xl"><span class="badge ${badgeClass}">${xepLoai}</span></td>
                    <td class="td-act">
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
        renderAdvancedReport();
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
                const selectedRow = dom.tableBody.querySelector(`tr[data-student-id="${CSS.escape(rows[0].id)}"]`);
                if (selectedRow instanceof HTMLElement) {
                    selectedRow.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
                }

                if (state.ui.searchSingleFocusTarget === 'socau') {
                    focusDirectInputForStudent(rows[0].id, 'socau');
                }

                if (state.ui.searchSingleFocusTarget === 'made') {
                    focusDirectInputForStudent(rows[0].id, 'diem');
                }
            }
        }, 500);
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
        if (!(input instanceof HTMLInputElement) && !(input instanceof HTMLSelectElement)) return;

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
        if (field === 'socau') {
            student.socau = parseNullableNumber(input.value);
            if (Number.isFinite(student.socau)) {
                student.socau = clamp(student.socau, 0, state.config.soCauToiDa);
                student.maDe = state.currentMaDE;
            }
        }
        if (field === 'diem') {
            student.diem = parseNullableNumber(input.value);
            if (Number.isFinite(student.diem)) {
                student.diem = clamp(student.diem, 0, 10);
                student.maDe = state.currentMaDE;
            }
        }

        if (field === 'socau' || field === 'diem') {
            const maDeCell = row.querySelector('.made-pill');
            if (maDeCell) {
                maDeCell.textContent = student.maDe || '—';
            }
        }

        // Do not re-render table on each keystroke to avoid losing input focus while typing.
        persistStudents(false);
        renderHeaderStats();
        renderProgress();
        updateUndoRedoState();
    }

    function handleTableChange(event) {
        const input = event.target;
        if (!(input instanceof HTMLInputElement) && !(input instanceof HTMLSelectElement)) return;

        if (!input.dataset.field) return;

        renderChartIfVisible({ immediate: true });
    }

    function handleTableKeydown(event) {
        if (event.key !== 'Enter') return;

        const input = event.target;
        if (!(input instanceof HTMLInputElement)) return;

        const row = input.closest('tr');
        if (!row) return;

        const field = input.dataset.field;
        if (field !== 'socau' && field !== 'diem') return;

        event.preventDefault();

        const student = findStudentByRow(row);
        if (!student) return;

        if (field === 'socau') {
            const socau = parseNullableNumber(input.value);
            student.socau = Number.isFinite(socau)
                ? clamp(socau, 0, state.config.soCauToiDa)
                : NaN;

            if (Number.isFinite(student.socau)) {
                student.diem = clamp(student.socau * getAutoDiemMoiCau(), 0, 10);
                student.maDe = state.currentMaDE;
            }
        }

        if (field === 'diem') {
            const diem = parseNullableNumber(input.value);
            student.diem = Number.isFinite(diem)
                ? clamp(diem, 0, 10)
                : NaN;

            if (Number.isFinite(student.diem)) {
                student.maDe = state.currentMaDE;
            }
        }

        input.blur();
        persistStudents();
        renderChartIfVisible({ immediate: true });
    }

    function handleTableClick(event) {
        const target = event.target;
        const clickedInteractive = target.closest('input, select, button, option, textarea');

        const rowForSelection = target.closest('tr[data-student-id]');
        if (rowForSelection && !clickedInteractive) {
            const nextSelectedId = rowForSelection.getAttribute('data-student-id');
            if (nextSelectedId !== state.selectedStudentId) {
                state.selectedStudentId = nextSelectedId;
                renderTableAndStats();
            }
        }

        const btn = target.closest('button[data-action]');
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
        dom.mapMaDeCol.addEventListener('change', () => buildImportPreview());
        dom.mapDiemCol.addEventListener('change', () => buildImportPreview());
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
        if (dom.cbClearBeforeImport) {
            dom.cbClearBeforeImport.checked = true;
        }
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

        const currentMaDeCol = parseInt(dom.mapMaDeCol.value, 10);
        const selectedMaDeIndex = Number.isInteger(currentMaDeCol)
            ? currentMaDeCol
            : resolveMaDeColumnIndex(headerOptions);

        const currentDiemCol = parseInt(dom.mapDiemCol.value, 10);
        const selectedDiemIndex = Number.isInteger(currentDiemCol)
            ? currentDiemCol
            : resolveScoreColumnIndex(headerOptions);

        dom.mapNameCol.innerHTML = headerOptions
            .map((item) => `
                <option value="${item.idx}" ${item.idx === selectedNameIndex ? 'selected' : ''}>${escapeHtml(item.name)}</option>
            `)
            .join('');

        dom.mapMaDeCol.innerHTML = ['<option value="-1">— Không có mã đề —</option>']
            .concat(headerOptions.map((item) => `
                <option value="${item.idx}" ${item.idx === selectedMaDeIndex ? 'selected' : ''}>${escapeHtml(item.name)}</option>
            `))
            .join('');

        dom.mapDiemCol.innerHTML = ['<option value="-1">— Không import điểm —</option>']
            .concat(headerOptions.map((item) => `
                <option value="${item.idx}" ${item.idx === selectedDiemIndex ? 'selected' : ''}>${escapeHtml(item.name)}</option>
            `))
            .join('');

        const skipRows = clamp(parseInt(dom.mapSkipRows.value, 10) || 0, 0, 200);
        const nameCol = parseInt(dom.mapNameCol.value, 10);
        const maDeCol = parseInt(dom.mapMaDeCol.value, 10);
        const diemCol = parseInt(dom.mapDiemCol.value, 10);
        const importRows = [];

        for (let i = 1 + skipRows; i < rows.length; i += 1) {
            const row = rows[i] || [];
            const nameRaw = row[nameCol];
            const name = String(nameRaw || '').trim();
            if (!name) continue;

            const parsed = {
                id: makeId(),
                name,
                maDe: Number.isInteger(maDeCol) && maDeCol >= 0 ? String(row[maDeCol] || '').trim() : '',
                socau: parseNullableNumber(extractByHeader(headerRow, row, ['so cau', 's cau', 'socau'])),
                diem: Number.isInteger(diemCol) && diemCol >= 0
                    ? parseNullableNumber(row[diemCol])
                    : parseNullableNumber(extractByHeader(headerRow, row, ['diem', 'diem so', 'score', 'mark', 'grade']))
            };

            importRows.push(normalizeStudent(parsed));
        }

        state.importRows = importRows;
        dom.importCountSpan.textContent = String(importRows.length);
        const maDeText = Number.isInteger(maDeCol) && maDeCol >= 0
            ? `Mã đề: ${headerOptions[maDeCol]?.name || 'Cột đã chọn'}`
            : 'Mã đề: Không import';
        const diemText = Number.isInteger(diemCol) && diemCol >= 0
            ? `Điểm: ${headerOptions[diemCol]?.name || 'Cột đã chọn'}`
            : 'Điểm: Tự suy đoán nếu có';
        dom.importSummary.textContent = `Sẵn sàng import ${importRows.length} học sinh từ sheet ${sheetName}. ${maDeText}. ${diemText}.`;

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

    function resolveMaDeColumnIndex(headerOptions) {
        const validDirect = new Set(['ma de', 'made', 'ma_de', 'ma de thi', 'made thi']);
        const rejectTokens = ['id', 'dinhdanh', 'mahs', 'mahocsinh', 'studentid'];

        for (const option of headerOptions) {
            const normalized = normalizeText(option.name);
            const compact = normalized.replace(/[^a-z0-9]/g, '');
            if (!normalized) continue;
            if (rejectTokens.some((token) => compact.includes(token))) continue;

            if (validDirect.has(normalized)) return option.idx;
            if (compact === 'made' || compact === 'madethi') return option.idx;
            if (compact.includes('made') && !compact.includes('dinhdanh')) return option.idx;
        }

        return -1;
    }

    function resolveScoreColumnIndex(headerOptions) {
        const directSet = new Set(['diem', 'diem so', 'score', 'mark', 'grade']);

        for (const option of headerOptions) {
            const normalized = normalizeText(option.name);
            const compact = normalized.replace(/[^a-z0-9]/g, '');
            if (!normalized) continue;
            if (directSet.has(normalized)) return option.idx;
            if (compact === 'diem' || compact === 'diemso') return option.idx;
            if (compact.includes('diem') || compact.includes('score')) return option.idx;
        }

        return -1;
    }

    function doImportStudents() {
        if (!state.importRows.length) {
            showToast('Không có dữ liệu để import.', 'error');
            return;
        }

        snapshotBeforeMutation();
        const clearBeforeImport = !!(dom.cbClearBeforeImport && dom.cbClearBeforeImport.checked);
        const cleanedImportRows = dedupeImportRows(state.importRows.map(normalizeStudent));
        let added = 0;

        if (clearBeforeImport) {
            state.students = cleanedImportRows;
            added = cleanedImportRows.length;
        } else {
            const existingKey = new Set(state.students.map((x) => `${normalizeText(x.name)}|${normalizeText(x.maDe || '')}`));

            cleanedImportRows.forEach((row) => {
                const key = `${normalizeText(row.name)}|${normalizeText(row.maDe || '')}`;
                if (!key) return;
                if (existingKey.has(key)) return;
                existingKey.add(key);
                state.students.push(row);
                added += 1;
            });
        }

        persistStudents();
        setCurrentFileName(state.pendingFileName || 'Excel import');
        addRecentFile(state.pendingFileName || `Import ${new Date().toLocaleString('vi-VN')}`);
        clearPendingImport();
        closeModal(dom.modalImport);
        renderTableAndStats();
        updateUndoRedoState();
        const modeText = clearBeforeImport ? ' (đã xóa dữ liệu cũ)' : '';
        showToast(`Đã import ${added} học sinh${modeText}`, 'success');
    }

    function dedupeImportRows(rows) {
        const uniqueRows = [];
        const seen = new Set();

        rows.forEach((row) => {
            const nameKey = normalizeText(row.name);
            if (!nameKey) return;
            const key = `${nameKey}|${normalizeText(row.maDe || '')}`;
            if (seen.has(key)) return;
            seen.add(key);
            uniqueRows.push(row);
        });

        return uniqueRows;
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

        renderAdvancedReport();

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

        if (!window.html2canvas) {
            showToast('Thiếu thư viện html2canvas.', 'error');
            return;
        }

        if (!state.students.length) {
            showToast('Chưa có dữ liệu để xuất.', 'error');
            return;
        }

        const fileBase = buildExportName();
        const wasHidden = dom.modalReport.classList.contains('hidden');
        let exportModeApplied = false;
        let exportedWithJsPdf = false;

        try {
            if (wasHidden) {
                setAdvancedReportVisible(true);
            }

            dom.reportPrintable.classList.add('report-export-mode');
            exportModeApplied = true;
            renderAdvancedReport();
            await waitForNextPaint();

            const canvas = await window.html2canvas(dom.reportPrintable, {
                scale: 2,
                backgroundColor: '#0b122d',
                useCORS: true,
                logging: false
            });

            const doc = new window.jspdf.jsPDF({ orientation: 'l', unit: 'pt', format: 'a4' });
            const pageWidth = doc.internal.pageSize.getWidth();
            const pageHeight = doc.internal.pageSize.getHeight();
            const margin = 16;
            const targetWidth = pageWidth - margin * 2;
            const scaleRatio = targetWidth / canvas.width;
            const sliceHeightPx = Math.max(1, Math.floor((pageHeight - margin * 2) / scaleRatio));

            let renderedY = 0;
            let pageIndex = 0;

            while (renderedY < canvas.height) {
                const currentSliceHeight = Math.min(sliceHeightPx, canvas.height - renderedY);
                const pageCanvas = document.createElement('canvas');
                pageCanvas.width = canvas.width;
                pageCanvas.height = currentSliceHeight;

                const pageContext = pageCanvas.getContext('2d');
                if (!pageContext) break;

                pageContext.drawImage(
                    canvas,
                    0,
                    renderedY,
                    canvas.width,
                    currentSliceHeight,
                    0,
                    0,
                    canvas.width,
                    currentSliceHeight
                );

                if (pageIndex > 0) {
                    doc.addPage();
                }

                const imageHeight = currentSliceHeight * scaleRatio;
                doc.addImage(
                    pageCanvas.toDataURL('image/png'),
                    'PNG',
                    margin,
                    margin,
                    targetWidth,
                    imageHeight,
                    undefined,
                    'FAST'
                );

                renderedY += currentSliceHeight;
                pageIndex += 1;
            }

            if (pageIndex > 0) {
                doc.save(`${fileBase}.pdf`);
                exportedWithJsPdf = true;
            }
        } catch (error) {
            console.error('PDF canvas export failed, fallback to print window:', error);
        } finally {
            if (exportModeApplied) {
                dom.reportPrintable.classList.remove('report-export-mode');
            }
            if (wasHidden) {
                setAdvancedReportVisible(false);
            }
        }

        if (!exportedWithJsPdf) {
            const opened = exportPdfViaPrintWindow(fileBase);
            if (!opened) {
                showToast('Không thể mở cửa sổ in PDF. Kiểm tra trình chặn popup.', 'error');
                return;
            }

            showToast('Đã mở cửa sổ in. Chọn Save as PDF để lưu đúng font tiếng Việt.', 'warning');
        }

        const cloudOk = await maybeSaveToCloud(fileBase);
        if (cloudOk === false) return;

        addRecentFile(`${fileBase}.pdf`);
        closeModal(dom.modalExport);
        showToast('Đã xuất PDF thành công', 'success');
    }

    function exportPdfViaPrintWindow(fileBase) {
        const printWindow = window.open('', '_blank', 'noopener,noreferrer,width=1024,height=720');
        if (!printWindow) return false;

        const rows = state.students.map((student, index) => {
            const score = getStudentScore(student);
            const soCauText = Number.isFinite(student.socau) ? String(student.socau) : '—';
            const scoreText = Number.isFinite(score) ? formatScore(score) : '—';
            return `
                <tr>
                    <td>${index + 1}</td>
                    <td class="cell-left">${escapeHtml(student.name)}</td>
                    <td>${escapeHtml(student.maDe || '—')}</td>
                    <td>${soCauText}</td>
                    <td>${scoreText}</td>
                    <td>${escapeHtml(getRank(score))}</td>
                </tr>
            `;
        }).join('');

        const html = `
            <!DOCTYPE html>
            <html lang="vi">
            <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <title>${escapeHtml(fileBase)} - PDF</title>
                <style>
                    @page { size: A4 portrait; margin: 14mm; }
                    body {
                        font-family: 'Segoe UI', 'Noto Sans', Tahoma, Arial, sans-serif;
                        color: #0f172a;
                        margin: 0;
                        font-size: 12px;
                        line-height: 1.4;
                    }
                    .header {
                        margin-bottom: 12px;
                    }
                    .title {
                        font-size: 20px;
                        font-weight: 800;
                        margin-bottom: 2px;
                    }
                    .meta {
                        color: #475569;
                        font-size: 12px;
                    }
                    table {
                        width: 100%;
                        border-collapse: collapse;
                        margin-top: 10px;
                    }
                    th, td {
                        border: 1px solid #cbd5e1;
                        padding: 6px 8px;
                        text-align: center;
                    }
                    th {
                        background: #0ea5e9;
                        color: #ffffff;
                        font-size: 11px;
                        text-transform: uppercase;
                        letter-spacing: 0.2px;
                    }
                    .cell-left { text-align: left; }
                    tbody tr:nth-child(even) { background: #f8fafc; }
                </style>
            </head>
            <body>
                <div class="header">
                    <div class="title">Bảng điểm học sinh</div>
                    <div class="meta">Bản ghi: ${escapeHtml(fileBase)}</div>
                    <div class="meta">Ngày xuất: ${escapeHtml(new Date().toLocaleString('vi-VN'))}</div>
                </div>

                <table>
                    <thead>
                        <tr>
                            <th>#</th>
                            <th>Họ và tên</th>
                            <th>Mã đề</th>
                            <th>Số câu</th>
                            <th>Điểm</th>
                            <th>Xếp loại</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${rows}
                    </tbody>
                </table>

                <script>
                    window.onload = function () {
                        setTimeout(function () { window.print(); }, 350);
                    };
                </script>
            </body>
            </html>
        `;

        printWindow.document.open();
        printWindow.document.write(html);
        printWindow.document.close();
        return true;
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

        rebuildMaDEOptions();

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

            if (!event.ctrlKey && !event.altKey && !event.shiftKey && !event.metaKey && event.key.toLowerCase() === 'x') {
                const activeElement = document.activeElement;
                const isTypingContext = activeElement instanceof HTMLInputElement
                    || activeElement instanceof HTMLTextAreaElement
                    || activeElement instanceof HTMLSelectElement
                    || (activeElement instanceof HTMLElement && activeElement.isContentEditable);

                if (isTypingContext) return;
                event.preventDefault();
                state.currentMaDE = '';
                dom.currentMaDESelect.value = '';
                localStorage.setItem(STORAGE_KEYS.currentMaDE, state.currentMaDE);
                showToast('Đã xóa mã đề hiện hành', 'info');
            }
        });
    }

    function saveExamConfig() {
        const soCau = clamp(parseInt(dom.cfgSoCauToiDa.value, 10) || DEFAULT_CONFIG.soCauToiDa, 1, 500);
        const diem = clamp(parseFloat(dom.cfgDiemToiDa.value) || DEFAULT_CONFIG.diemToiDa, 0.5, 1000);
        const danhSach = dom.cfgDanhSachMaDE.value
            .split(',')
            .map((x) => x.trim())
            .filter(Boolean);

        state.config = {
            soCauToiDa: soCau,
            diemToiDa: diem,
            danhSachMaDE: danhSach,
            tenCotHS: dom.cfgTenCotHS.value.trim() || DEFAULT_CONFIG.tenCotHS
        };

        localStorage.setItem(STORAGE_KEYS.config, JSON.stringify(state.config));
        rebuildMaDEOptions();
        renderTableAndStats();
        
        const pointPerQuestion = (diem / soCau).toFixed(2);
        showToast(`Đã lưu cấu hình bài thi - Mỗi câu ${pointPerQuestion} điểm`, 'success');
    }

    function saveCloudConfig() {
        state.proxyUrl = dom.cfgProxyUrl.value.trim();
        localStorage.setItem(STORAGE_KEYS.proxy, state.proxyUrl);
        showToast('Đã lưu cấu hình cloud', 'success');
    }

    function rebuildMaDEOptions() {
        const values = state.config.danhSachMaDE.slice();
        if (state.currentMaDE && !values.includes(state.currentMaDE)) {
            values.push(state.currentMaDE);
        }

        const options = ['<option value="">— Không có —</option>']
            .concat(values.map((ma) => `<option value="${escapeAttr(ma)}">${escapeHtml(ma)}</option>`));

        dom.stuMaDE.innerHTML = options.join('');
        dom.currentMaDESelect.innerHTML = options.join('');
        dom.currentMaDESelect.value = state.currentMaDE;
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
        if (visible) {
            applySavedChartPosition();
            renderChartIfVisible({ immediate: true });
            return;
        }

        if (state.chartRenderTimer) {
            clearTimeout(state.chartRenderTimer);
            state.chartRenderTimer = null;
        }
    }

    function bindChartDragEvents() {
        if (!dom.chartDragHandle || !dom.chartSection) return;

        dom.chartDragHandle.addEventListener('pointerdown', onChartDragStart);
        window.addEventListener('pointermove', onChartDragMove);
        window.addEventListener('pointerup', onChartDragEnd);
        window.addEventListener('pointercancel', onChartDragEnd);
        window.addEventListener('resize', () => {
            if (!state.ui.chartPosition) return;
            applySavedChartPosition();
        });
    }

    function onChartDragStart(event) {
        if (!(event instanceof PointerEvent)) return;
        if (event.button !== 0) return;
        if (event.target instanceof HTMLElement && event.target.closest('button, select, input, option')) return;
        if (dom.chartSection.classList.contains('hidden')) return;

        const rect = dom.chartSection.getBoundingClientRect();
        state.chartDrag.active = true;
        state.chartDrag.pointerId = event.pointerId;
        state.chartDrag.offsetX = event.clientX - rect.left;
        state.chartDrag.offsetY = event.clientY - rect.top;

        applyChartPosition(rect.left, rect.top, false);
        dom.chartSection.classList.add('chart-dragging');

        if (dom.chartDragHandle.setPointerCapture) {
            try {
                dom.chartDragHandle.setPointerCapture(event.pointerId);
            } catch (error) {
                // Ignore pointer capture errors on unsupported browsers.
            }
        }

        event.preventDefault();
    }

    function onChartDragMove(event) {
        if (!(event instanceof PointerEvent)) return;
        if (!state.chartDrag.active) return;
        if (state.chartDrag.pointerId !== event.pointerId) return;

        const rawLeft = event.clientX - state.chartDrag.offsetX;
        const rawTop = event.clientY - state.chartDrag.offsetY;
        const clamped = clampChartCoordinates(rawLeft, rawTop);
        applyChartPosition(clamped.left, clamped.top, false);
    }

    function onChartDragEnd(event) {
        if (!(event instanceof PointerEvent)) return;
        if (!state.chartDrag.active) return;
        if (state.chartDrag.pointerId !== event.pointerId) return;

        state.chartDrag.active = false;
        state.chartDrag.pointerId = null;
        dom.chartSection.classList.remove('chart-dragging');

        if (dom.chartDragHandle.releasePointerCapture) {
            try {
                dom.chartDragHandle.releasePointerCapture(event.pointerId);
            } catch (error) {
                // Ignore pointer capture errors on unsupported browsers.
            }
        }

        const rect = dom.chartSection.getBoundingClientRect();
        const clamped = clampChartCoordinates(rect.left, rect.top);
        applyChartPosition(clamped.left, clamped.top, true);
    }

    function applySavedChartPosition() {
        if (!state.ui.chartPosition) {
            dom.chartSection.style.left = '';
            dom.chartSection.style.top = '';
            dom.chartSection.style.right = '';
            dom.chartSection.style.bottom = '';
            return;
        }

        const clamped = clampChartCoordinates(state.ui.chartPosition.left, state.ui.chartPosition.top);
        applyChartPosition(clamped.left, clamped.top, true);
    }

    function applyChartPosition(left, top, persist) {
        dom.chartSection.style.left = `${left}px`;
        dom.chartSection.style.top = `${top}px`;
        dom.chartSection.style.right = 'auto';
        dom.chartSection.style.bottom = 'auto';

        if (!persist) return;
        state.ui.chartPosition = {
            left: Math.round(left),
            top: Math.round(top)
        };
        localStorage.setItem(STORAGE_KEYS.ui, JSON.stringify(state.ui));
    }

    function clampChartCoordinates(left, top) {
        const margin = 8;
        const topBoundary = getChartTopBoundary();
        const maxLeft = Math.max(margin, window.innerWidth - dom.chartSection.offsetWidth - margin);
        const maxTop = Math.max(topBoundary, window.innerHeight - dom.chartSection.offsetHeight - margin);

        return {
            left: clamp(left, margin, maxLeft),
            top: clamp(top, topBoundary, maxTop)
        };
    }

    function getChartTopBoundary() {
        if (window.innerWidth <= 480) return 8;

        const rootStyles = getComputedStyle(document.documentElement);
        const headerHeight = parseInt(rootStyles.getPropertyValue('--header-h'), 10) || 60;
        const progressHeight = parseInt(rootStyles.getPropertyValue('--prog-h'), 10) || 32;
        return headerHeight + progressHeight + 6;
    }

    function renderChartIfVisible(options = {}) {
        if (dom.chartSection.classList.contains('hidden')) return;

        const delayMs = options.immediate ? 0 : 160;
        if (state.chartRenderTimer) {
            clearTimeout(state.chartRenderTimer);
        }

        state.chartRenderTimer = setTimeout(() => {
            state.chartRenderTimer = null;
            renderChart();
        }, delayMs);
    }

    function renderChart() {
        if (!window.Chart) return;
        const scores = state.students.map(getStudentScore).filter(Number.isFinite);
        const selectedChartType = ALLOWED_CHART_TYPES.has(state.ui.chartType) ? state.ui.chartType : 'bar';
        const chartType = selectedChartType === 'barHorizontal' ? 'bar' : selectedChartType;

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

        const labels = Object.keys(buckets);
        const values = Object.values(buckets);
        const palette = ['#ef4444', '#3b82f6', '#f59e0b', '#10b981'];
        const isCircularChart = selectedChartType === 'doughnut'
            || selectedChartType === 'pie'
            || selectedChartType === 'polarArea';
        const isRadarChart = selectedChartType === 'radar';
        const isHorizontalBar = selectedChartType === 'barHorizontal';

        const dataset = {
            label: 'Số HS theo mức điểm',
            data: values
        };

        if (selectedChartType === 'line') {
            dataset.borderColor = '#60a5fa';
            dataset.backgroundColor = 'rgba(96, 165, 250, 0.25)';
            dataset.borderWidth = 2;
            dataset.fill = true;
            dataset.tension = 0.35;
            dataset.pointRadius = 4;
            dataset.pointHoverRadius = 5;
            dataset.pointBackgroundColor = '#c7d2fe';
        } else if (isRadarChart) {
            dataset.backgroundColor = 'rgba(96, 165, 250, 0.25)';
            dataset.borderColor = '#60a5fa';
            dataset.borderWidth = 2;
            dataset.pointRadius = 3;
            dataset.pointBackgroundColor = '#c7d2fe';
        } else {
            dataset.backgroundColor = palette;
            dataset.borderColor = chartType === 'bar' ? 'rgba(255,255,255,0.15)' : '#12122b';
            dataset.borderWidth = chartType === 'bar' ? 1 : 2;
            dataset.borderRadius = chartType === 'bar' ? 6 : 0;
        }

        const nextOptions = {
            responsive: true,
            maintainAspectRatio: false,
            animation: {
                duration: 1120,
                easing: 'easeOutQuart'
            },
            indexAxis: isHorizontalBar ? 'y' : 'x',
            plugins: {
                legend: {
                    display: isCircularChart || isRadarChart,
                    position: 'bottom',
                    labels: {
                        color: '#94a3b8',
                        boxWidth: 12,
                        usePointStyle: true,
                        pointStyle: 'circle'
                    }
                },
                tooltip: {
                    callbacks: {
                        label(context) {
                            const count = Number(context.raw) || 0;
                            return ` ${count} học sinh`;
                        }
                    }
                }
            },
            scales: isCircularChart
                ? {}
                : isRadarChart
                    ? {
                        r: {
                            beginAtZero: true,
                            ticks: {
                                precision: 0,
                                color: '#94a3b8',
                                backdropColor: 'transparent'
                            },
                            grid: { color: 'rgba(148, 163, 184, 0.2)' },
                            angleLines: { color: 'rgba(148, 163, 184, 0.14)' },
                            pointLabels: { color: '#94a3b8' }
                        }
                    }
                    : {
                        x: {
                            ticks: { color: '#94a3b8' },
                            grid: { color: 'rgba(148, 163, 184, 0.08)' }
                        },
                        y: {
                            beginAtZero: true,
                            ticks: {
                                precision: 0,
                                color: '#94a3b8'
                            },
                            grid: { color: 'rgba(148, 163, 184, 0.1)' }
                        }
                    }
        };

        const canReuseChart = state.chart
            && state.chart.config
            && state.chart.config.type === chartType;

        if (canReuseChart) {
            state.chart.data.labels = labels;
            state.chart.data.datasets = [dataset];
            state.chart.options = nextOptions;
            state.chart.update();
            return;
        }

        if (state.chart) {
            state.chart.destroy();
        }

        state.chart = new Chart(dom.scoreChart, {
            type: chartType,
            data: {
                labels,
                datasets: [dataset]
            },
            options: nextOptions
        });
    }

    function setAdvancedReportVisible(visible) {
        if (!dom.modalReport) return;
        dom.modalReport.classList.toggle('hidden', !visible);
        if (visible) {
            renderAdvancedReport();
        }
    }

    function renderAdvancedReport() {
        if (!dom.reportSection || !dom.reportPrintable) return;

        const total = state.students.length;
        const scoredStudents = state.students
            .map((student) => ({
                student,
                score: getStudentScore(student)
            }))
            .filter((item) => Number.isFinite(item.score));

        const entered = scoredStudents.length;
        const scores = scoredStudents.map((item) => item.score);
        const avg = entered ? scores.reduce((sum, value) => sum + value, 0) / entered : NaN;
        const sortedScoresAsc = scores.slice().sort((a, b) => a - b);
        const median = entered
            ? (entered % 2 === 0
                ? (sortedScoresAsc[entered / 2 - 1] + sortedScoresAsc[entered / 2]) / 2
                : sortedScoresAsc[Math.floor(entered / 2)])
            : NaN;
        const variance = entered
            ? scores.reduce((sum, value) => sum + ((value - avg) ** 2), 0) / entered
            : NaN;
        const std = Number.isFinite(variance) ? Math.sqrt(variance) : NaN;
        const passCount = scores.filter((score) => score >= 5).length;
        const passRate = entered ? (passCount / entered) * 100 : NaN;

        const dist = {
            excellent: scores.filter((score) => score >= 8).length,
            good: scores.filter((score) => score >= 6.5 && score < 8).length,
            average: scores.filter((score) => score >= 5 && score < 6.5).length,
            weak: scores.filter((score) => score < 5).length
        };

        const distWidth = {
            excellent: entered ? (dist.excellent / entered) * 100 : 0,
            good: entered ? (dist.good / entered) * 100 : 0,
            average: entered ? (dist.average / entered) * 100 : 0,
            weak: entered ? (dist.weak / entered) * 100 : 0
        };

        dom.repTotal.textContent = String(total);
        dom.repEntered.textContent = String(entered);
        dom.repAvg.textContent = Number.isFinite(avg) ? formatScore(avg) : '—';
        dom.repMedian.textContent = Number.isFinite(median) ? formatScore(median) : '—';
        dom.repStd.textContent = Number.isFinite(std) ? formatScore(std) : '—';
        dom.repPassRate.textContent = Number.isFinite(passRate) ? `${passRate.toFixed(1)}%` : '—';

        dom.repDistExcellent.style.width = `${distWidth.excellent}%`;
        dom.repDistGood.style.width = `${distWidth.good}%`;
        dom.repDistAverage.style.width = `${distWidth.average}%`;
        dom.repDistWeak.style.width = `${distWidth.weak}%`;

        dom.repDistExcellentText.textContent = `${dist.excellent}`;
        dom.repDistGoodText.textContent = `${dist.good}`;
        dom.repDistAverageText.textContent = `${dist.average}`;
        dom.repDistWeakText.textContent = `${dist.weak}`;

        const topFive = scoredStudents
            .slice()
            .sort((a, b) => b.score - a.score)
            .slice(0, 5);
        const bottomFive = scoredStudents
            .slice()
            .sort((a, b) => a.score - b.score)
            .slice(0, 5);

        dom.repTopList.innerHTML = topFive.length
            ? topFive.map((item) => `
                <li>${escapeHtml(item.student.name)} - <span class="rank-score">${formatScore(item.score)}</span></li>
            `).join('')
            : '<li>Chưa có dữ liệu điểm</li>';

        dom.repBottomList.innerHTML = bottomFive.length
            ? bottomFive.map((item) => `
                <li>${escapeHtml(item.student.name)} - <span class="rank-score">${formatScore(item.score)}</span></li>
            `).join('')
            : '<li>Chưa có dữ liệu điểm</li>';

        const byMaDe = new Map();
        state.students.forEach((student) => {
            const key = student.maDe || '—';
            const score = getStudentScore(student);
            if (!byMaDe.has(key)) {
                byMaDe.set(key, {
                    total: 0,
                    entered: 0,
                    excellent: 0,
                    weak: 0,
                    scoreSum: 0
                });
            }

            const bucket = byMaDe.get(key);
            bucket.total += 1;

            if (Number.isFinite(score)) {
                bucket.entered += 1;
                bucket.scoreSum += score;
                if (score >= 8) bucket.excellent += 1;
                if (score < 5) bucket.weak += 1;
            }
        });

        const rows = Array.from(byMaDe.entries())
            .sort(([a], [b]) => String(a).localeCompare(String(b), 'vi'))
            .map(([maDe, bucket]) => {
                const avgByCode = bucket.entered ? bucket.scoreSum / bucket.entered : NaN;
                return `
                    <tr>
                        <td>${escapeHtml(String(maDe))}</td>
                        <td>${bucket.total}</td>
                        <td>${bucket.entered}</td>
                        <td>${Number.isFinite(avgByCode) ? formatScore(avgByCode) : '—'}</td>
                        <td>${bucket.excellent}</td>
                        <td>${bucket.weak}</td>
                    </tr>
                `;
            });

        dom.repByMaDEBody.innerHTML = rows.length
            ? rows.join('')
            : '<tr><td colspan="6">Chưa có dữ liệu</td></tr>';

        const scoreRows = state.students
            .map((student, index) => {
                const score = getStudentScore(student);
                const soCauText = Number.isFinite(student.socau) ? String(student.socau) : '—';
                const scoreText = Number.isFinite(score) ? formatScore(score) : '—';
                return `
                    <tr>
                        <td>${index + 1}</td>
                        <td class="rep-cell-left">${escapeHtml(student.name)}</td>
                        <td>${escapeHtml(student.maDe || '—')}</td>
                        <td>${soCauText}</td>
                        <td>${scoreText}</td>
                        <td>${escapeHtml(getRank(score))}</td>
                    </tr>
                `;
            });

        dom.repScoreBody.innerHTML = scoreRows.length
            ? scoreRows.join('')
            : '<tr><td colspan="6">Chưa có dữ liệu học sinh</td></tr>';

        dom.reportSubtitle.textContent = `Cập nhật lúc ${new Date().toLocaleString('vi-VN')} - ${entered}/${total} học sinh đã có điểm`;
    }

    function toggleMaDE() {
        state.showMaDE = !state.showMaDE;
        renderTableAndStats();
    }

    function waitForNextPaint() {
        return new Promise((resolve) => {
            requestAnimationFrame(() => {
                requestAnimationFrame(resolve);
            });
        });
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

    function getRowRankClass(score) {
        if (!Number.isFinite(score)) return '';
        if (score >= 8) return 'row-excellent';
        if (score >= 6.5) return 'row-good';
        if (score >= 5) return 'row-average';
        return 'row-weak';
    }

    function getBadgeClass(score) {
        if (!Number.isFinite(score)) return 'badge-none';
        if (score >= 8) return 'badge-excellent';
        if (score >= 6.5) return 'badge-good';
        if (score >= 5) return 'badge-average';
        return 'badge-weak';
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
        const toastType = ['success', 'error', 'warning', 'info'].includes(type) ? type : 'info';
        const iconMap = {
            success: '✓',
            error: '✕',
            warning: '!',
            info: 'i'
        };

        toast.className = `toast toast-${toastType}`;
        toast.innerHTML = `
            <span class="toast-icon" aria-hidden="true">${iconMap[toastType]}</span>
            <span class="toast-msg">${escapeHtml(message)}</span>
        `;
        dom.toastContainer.appendChild(toast);
        setTimeout(() => {
            toast.classList.add('toast-out');
            setTimeout(() => toast.remove(), 220);
        }, 2800);
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
            .replace(/[đĐ]/g, 'd')
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
