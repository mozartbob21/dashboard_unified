const themeToggleBtn = document.getElementById("themeToggleBtn");
const runCheckBtn = document.getElementById("runCheckBtn");
const aboutProjectBtn = document.getElementById("aboutProjectBtn");
const toggleProblemSectionBtn = document.getElementById("toggleProblemSectionBtn");

const loaderBlock = document.getElementById("loaderBlock");
const statusMessage = document.getElementById("statusMessage");
const systemStateBadge = document.getElementById("systemStateBadge");
const currentStage = document.getElementById("currentStage");
const currentStageMessage = document.getElementById("currentStageMessage");

const metricTotal = document.getElementById("metricTotal");
const metricWorking = document.getElementById("metricWorking");
const metricNotWorking = document.getElementById("metricNotWorking");
const metricNotConnected = document.getElementById("metricNotConnected");

const checkTimesBlock = document.getElementById("checkTimesBlock");
const startedAtValue = document.getElementById("startedAtValue");
const finishedAtValue = document.getElementById("finishedAtValue");

const emptyStateCard = document.getElementById("emptyStateCard");
const problemSection = document.getElementById("problemSection");
const problemResultsContainer = document.getElementById("problemResultsContainer");
const detailsSection = document.getElementById("detailsSection");
const detailsContent = document.getElementById("detailsContent");

const aboutProjectModal = document.getElementById("aboutProjectModal");
const closeAboutProjectBtn = document.getElementById("closeAboutProjectBtn");
const closeAboutProjectFooterBtn = document.getElementById("closeAboutProjectFooterBtn");

const openAddressUploadBtn = document.getElementById("openAddressUploadBtn");
const openAddressUploadBtnInline = document.getElementById("openAddressUploadBtnInline");
const addressUploadModal = document.getElementById("addressUploadModal");
const closeAddressUploadBtn = document.getElementById("closeAddressUploadBtn");
const closeAddressUploadFooterBtn = document.getElementById("closeAddressUploadFooterBtn");
const addressUploadForm = document.getElementById("addressUploadForm");
const addressFileInput = document.getElementById("addressFileInput");
const submitAddressUploadBtn = document.getElementById("submitAddressUploadBtn");
const uploadStatusBox = document.getElementById("uploadStatusBox");

const addressTablePath = document.getElementById("addressTablePath");
const addressTableRows = document.getElementById("addressTableRows");
const addressTableUpdatedAt = document.getElementById("addressTableUpdatedAt");
const addressTableSize = document.getElementById("addressTableSize");
const addressTablePreview = document.getElementById("addressTablePreview");

const statusFilter = document.getElementById("statusFilter");
const groupByFilter = document.getElementById("groupByFilter");

let statusPollTimer = null;
let lastShownFinalStatus = null;
let isProblemSectionCollapsed = false;

function getCurrentFilters() {
    return {
        status: statusFilter?.value || window.__INITIAL_FILTERS__?.status || "all",
        groupBy: groupByFilter?.value || window.__INITIAL_FILTERS__?.groupBy || "none"
    };
}

function getCurrentTheme() {
    return document.documentElement.dataset.theme || "dark";
}

function applyTheme(theme) {
    document.documentElement.dataset.theme = theme;
    localStorage.setItem("camera-dashboard-theme", theme);

    if (!themeToggleBtn) {
        return;
    }

    const icon = themeToggleBtn.querySelector(".theme-toggle-icon");
    const text = themeToggleBtn.querySelector(".theme-toggle-text");

    if (theme === "light") {
        if (icon) {
            icon.textContent = "☀️";
        }
        if (text) {
            text.textContent = "Светлая";
        }
    } else {
        if (icon) {
            icon.textContent = "🌙";
        }
        if (text) {
            text.textContent = "Тёмная";
        }
    }
}

function toggleTheme() {
    const currentTheme = getCurrentTheme();
    const nextTheme = currentTheme === "dark" ? "light" : "dark";
    applyTheme(nextTheme);
}

function showLoader() {
    if (loaderBlock) {
        loaderBlock.classList.remove("hidden");
    }
}

function hideLoader() {
    if (loaderBlock) {
        loaderBlock.classList.add("hidden");
    }
}

function setSystemReady() {
    if (!systemStateBadge) {
        return;
    }
    systemStateBadge.textContent = "Готова";
    systemStateBadge.className = "system-badge ready";
}

function setSystemRunning() {
    if (!systemStateBadge) {
        return;
    }
    systemStateBadge.textContent = "Выполняется проверка";
    systemStateBadge.className = "system-badge running";
}

function disableControlsForRun() {
    if (runCheckBtn) {
        runCheckBtn.disabled = true;
        runCheckBtn.textContent = "Проверка выполняется...";
    }
    if (aboutProjectBtn) {
        aboutProjectBtn.disabled = true;
    }
}

function enableControlsAfterRun() {
    if (runCheckBtn) {
        runCheckBtn.disabled = false;
        runCheckBtn.textContent = "Запустить проверку";
    }
    if (aboutProjectBtn) {
        aboutProjectBtn.disabled = false;
    }
}

function scrollToTopSmooth() {
    window.scrollTo({
        top: 0,
        behavior: "smooth"
    });
}

function showToast(message, kind = "success", scrollTop = false) {
    if (!statusMessage) {
        return;
    }

    statusMessage.textContent = message;
    statusMessage.className = `status-message ${kind}`;

    if (scrollTop) {
        scrollToTopSmooth();
    }

    clearTimeout(window.__cameraToastTimer);

    window.__cameraToastTimer = setTimeout(() => {
        statusMessage.className = "status-message hidden";
        statusMessage.textContent = "";
    }, 3500);
}

function clearToast() {
    if (!statusMessage) {
        return;
    }

    statusMessage.className = "status-message hidden";
    statusMessage.textContent = "";
}

function openModal(modal) {
    if (!modal) {
        return;
    }

    modal.classList.remove("hidden");
    document.body.classList.add("modal-open");
}

function closeModal(modal) {
    if (!modal) {
        return;
    }

    modal.classList.add("hidden");
    document.body.classList.remove("modal-open");
}

function openAboutProjectModal() {
    openModal(aboutProjectModal);
}

function closeAboutProjectModal() {
    closeModal(aboutProjectModal);
}

function openAddressUploadModal() {
    openModal(addressUploadModal);
}

function closeAddressUploadModal() {
    closeModal(addressUploadModal);

    if (uploadStatusBox) {
        uploadStatusBox.className = "upload-status-box hidden";
        uploadStatusBox.textContent = "";
    }

    if (addressUploadForm) {
        addressUploadForm.reset();
    }
}

function stopStatusPolling() {
    if (statusPollTimer) {
        clearTimeout(statusPollTimer);
        statusPollTimer = null;
    }
}

function scheduleNextPoll() {
    stopStatusPolling();
    statusPollTimer = setTimeout(fetchStatus, 1000);
}

function getStatusLabel(status) {
    if (status === "working") {
        return "Работает";
    }
    if (status === "not_working") {
        return "Не работает";
    }
    if (status === "not_connected") {
        return "Не подключена";
    }
    return "Неизвестно";
}

function getRowClass(status) {
    if (status === "working") {
        return "status-ok";
    }
    if (status === "not_working") {
        return "status-critical";
    }
    return "status-risk";
}

function getPillClass(status) {
    if (status === "working") {
        return "pill pill-ok";
    }
    if (status === "not_working") {
        return "pill pill-critical";
    }
    return "pill pill-risk";
}

function getProblemBadgeText(status) {
    if (status === "not_working") {
        return "Камера не работает";
    }
    if (status === "not_connected") {
        return "Не подключена";
    }
    if (status === "unknown") {
        return "Статус не определён";
    }
    return "Статус не определён";
}


function getProblemDescription(item) {
    const address = item?.address || "";
    const owner = item?.owner || "Не указан";
    const status = item?.camera_status;

    if (status === "not_working") {
        return `По адресу "${address}" камера не работает. Ответственный: ${owner}.`;
    }

    if (status === "not_connected") {
        return `По адресу "${address}" камера не подключена или адрес отсутствует в таблице камер. Ответственный: ${owner}.`;
    }

    if (status === "unknown") {
        return `По адресу "${address}" статус камеры не удалось определить автоматически. Ответственный: ${owner}.`;
    }

    return `По адресу "${address}" статус камеры не удалось определить автоматически. Ответственный: ${owner}.`;
}

function escapeHtml(value) {
    return String(value ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
}

function buildStreamLink(url) {
    if (!url) {
        return "—";
    }

    return `<a href="${escapeHtml(url)}" target="_blank" rel="noopener noreferrer" class="table-link">Открыть</a>`;
}

function updateMetrics(metrics) {
    if (metricTotal) {
        metricTotal.textContent = metrics.total ?? 0;
    }
    if (metricWorking) {
        metricWorking.textContent = metrics.working ?? 0;
    }
    if (metricNotWorking) {
        metricNotWorking.textContent = metrics.not_working ?? 0;
    }
    if (metricNotConnected) {
        metricNotConnected.textContent = metrics.not_connected ?? 0;
    }
}

function updateRunState(checkState) {
    if (currentStage) {
        currentStage.textContent = checkState.stage || "Ожидание запуска";
    }

    if (currentStageMessage) {
        currentStageMessage.textContent = checkState.message || "";
    }

    if (startedAtValue) {
        startedAtValue.textContent = checkState.started_at || "—";
    }

    if (finishedAtValue) {
        finishedAtValue.textContent = checkState.finished_at || "—";
    }

    if (checkState.is_running) {
        setSystemRunning();
        showLoader();
        disableControlsForRun();
    } else {
        setSystemReady();
        hideLoader();
        enableControlsAfterRun();
    }
}

function updateProblemSectionCollapsedState() {
    if (!problemResultsContainer || !toggleProblemSectionBtn) {
        return;
    }

    if (isProblemSectionCollapsed) {
        problemResultsContainer.classList.add("hidden");
        toggleProblemSectionBtn.textContent = "Развернуть";
    } else {
        problemResultsContainer.classList.remove("hidden");
        toggleProblemSectionBtn.textContent = "Свернуть";
    }
}

function toggleProblemSection() {
    isProblemSectionCollapsed = !isProblemSectionCollapsed;
    updateProblemSectionCollapsedState();
}

function renderProblemResults(results) {
    if (!problemResultsContainer || !problemSection) {
        return;
    }

    const problemItems = results.filter((item) => item.camera_status !== "working");
    problemResultsContainer.innerHTML = "";

    if (!problemItems.length) {
        problemSection.classList.add("hidden");
        return;
    }

    problemSection.classList.remove("hidden");

    for (const item of problemItems) {
        const block = document.createElement("div");
        block.className = "message-block";
        block.innerHTML = `
            <div class="message-head">
                <div class="message-head-badges">
                    <strong>${escapeHtml(item.owner || "Не указан")}</strong>
                    <span class="${getPillClass(item.camera_status)}">${escapeHtml(getProblemBadgeText(item.camera_status))}</span>
                </div>
            </div>

            <div class="message-meta">
                ${escapeHtml(item.address || "")} / Проверено: ${escapeHtml(item.checked_at || "—")}
            </div>

            <pre>${escapeHtml(getProblemDescription(item))}</pre>

            <div class="message-actions">
                ${item.stream_url ? `<a class="action-btn action-btn-success" href="${escapeHtml(item.stream_url)}" target="_blank" rel="noopener noreferrer">Открыть ссылку</a>` : ""}
            </div>
        `;
        problemResultsContainer.appendChild(block);
    }

    updateProblemSectionCollapsedState();
}

function renderFlatTable(results) {
    return `
        <div class="table-wrap">
            <table>
                <thead>
                    <tr>
                        <th>ID</th>
                        <th>Адрес</th>
                        <th>Город</th>
                        <th>Ответственный</th>
                        <th>Статус камеры</th>
                        <th>Ссылка</th>
                        <th>Время проверки</th>
                    </tr>
                </thead>
                <tbody>
                    ${results.map((row) => `
                        <tr class="${getRowClass(row.camera_status)}">
                            <td>${escapeHtml(row.id ?? "")}</td>
                            <td>${escapeHtml(row.address ?? "")}</td>
                            <td>${escapeHtml(row.city || "Не указан")}</td>
                            <td>${escapeHtml(row.owner || "Не указан")}</td>
                            <td><span class="${getPillClass(row.camera_status)}">${escapeHtml(getStatusLabel(row.camera_status))}</span></td>
                            <td>${buildStreamLink(row.link_url || row.stream_url)}</td>
                            <td>${escapeHtml(row.checked_at ?? "—")}</td>
                        </tr>
                    `).join("")}
                </tbody>
            </table>
        </div>
    `;
}

function renderGroupedTable(groups) {
    return `
        <div id="groupedResultsContainer">
            ${groups.map((group) => `
                <div class="group-block">
                    <div class="group-header">
                        <h3>${escapeHtml(group.name)}</h3>
                        <span class="group-count">${escapeHtml(group.items.length)}</span>
                    </div>

                    <div class="table-wrap">
                        <table>
                            <thead>
                                <tr>
                                    <th>ID</th>
                                    <th>Адрес</th>
                                    <th>Город</th>
                                    <th>Ответственный</th>
                                    <th>Статус камеры</th>
                                    <th>Ссылка</th>
                                    <th>Время проверки</th>
                                </tr>
                            </thead>
                            <tbody>
                                ${group.items.map((row) => `
                                    <tr class="${getRowClass(row.camera_status)}">
                                        <td>${escapeHtml(row.id ?? "")}</td>
                                        <td>${escapeHtml(row.address ?? "")}</td>
                                        <td>${escapeHtml(row.city || "Не указан")}</td>
                                        <td>${escapeHtml(row.owner || "Не указан")}</td>
                                        <td><span class="${getPillClass(row.camera_status)}">${escapeHtml(getStatusLabel(row.camera_status))}</span></td>
                                        <td>${buildStreamLink(row.link_url || row.stream_url)}</td>
                                        <td>${escapeHtml(row.checked_at ?? "—")}</td>
                                    </tr>
                                `).join("")}
                            </tbody>
                        </table>
                    </div>
                </div>
            `).join("")}
        </div>
    `;
}

function renderDetailsContent(filteredResults, groupedResults, groupBy) {
    if (!detailsContent || !detailsSection) {
        return;
    }

    if (!filteredResults.length && (!groupedResults || !groupedResults.length)) {
        detailsSection.classList.add("hidden");
        return;
    }

    detailsSection.classList.remove("hidden");

    if (groupBy === "none") {
        detailsContent.innerHTML = renderFlatTable(filteredResults);
    } else {
        detailsContent.innerHTML = renderGroupedTable(groupedResults || []);
    }
}

function updateVisibility(results, checkState) {
    const hasResults = Array.isArray(results) && results.length > 0;
    const hasAnyTime = Boolean(checkState.started_at || checkState.finished_at || checkState.is_running);

    if (checkTimesBlock) {
        checkTimesBlock.classList.toggle("hidden", !hasAnyTime);
    }

    if (emptyStateCard) {
        emptyStateCard.classList.toggle("hidden", hasResults || checkState.is_running);
    }
}

function updateAddressTableInfo(table) {
    if (!table) {
        return;
    }

    if (addressTablePath) {
        addressTablePath.textContent = table.path || "—";
    }

    if (addressTableRows) {
        addressTableRows.textContent = table.rows_count ?? 0;
    }

    if (addressTableUpdatedAt) {
        addressTableUpdatedAt.textContent = table.updated_at || "—";
    }

    if (addressTableSize) {
        addressTableSize.textContent = `${table.size_bytes ?? 0} байт`;
    }

    if (addressTablePreview) {
        addressTablePreview.textContent = Array.isArray(table.preview) && table.preview.length
            ? table.preview.join("\n")
            : "Файл пока не загружен";
    }
}

async function fetchAddressTableInfo() {
    try {
        const response = await fetch("/cameras/address-table-info", {
            cache: "no-store"
        });
        const data = await response.json();

        if (data.ok) {
            updateAddressTableInfo(data.table);
        }
    } catch (error) {
        console.error("Ошибка получения информации о таблице адресов", error);
    }
}

async function fetchStatus() {
    try {
        const filters = getCurrentFilters();
        const params = new URLSearchParams({
            status: filters.status,
            group_by: filters.groupBy
        });

        const response = await fetch(`/cameras/status?${params.toString()}`, {
            cache: "no-store"
        });
        const data = await response.json();

        const checkState = data.check_state || {};
        const results = data.results || [];
        const filteredResults = data.filtered_results || [];
        const groupedResults = data.grouped_results || [];
        const metrics = data.metrics || {};
        const groupBy = data.group_by || "none";

        updateRunState(checkState);
        updateMetrics(metrics);
        renderProblemResults(results);
        renderDetailsContent(filteredResults, groupedResults, groupBy);
        updateVisibility(results, checkState);

        if (checkState.is_running) {
            lastShownFinalStatus = null;
            scheduleNextPoll();
            return;
        }

        stopStatusPolling();

        if (checkState.status === "done" && lastShownFinalStatus !== "done") {
            lastShownFinalStatus = "done";
            showToast("Проверка успешно завершена", "success", true);
        }

        if (checkState.status === "error" && lastShownFinalStatus !== "error") {
            lastShownFinalStatus = "error";
            showToast(checkState.last_error || "Произошла ошибка", "error", true);
        }
    } catch (error) {
        stopStatusPolling();
        hideLoader();
        enableControlsAfterRun();
        setSystemReady();
        showToast("Ошибка при получении статуса", "error", true);
    }
}

async function startCheck() {
    disableControlsForRun();
    showLoader();
    clearToast();
    setSystemRunning();
    scrollToTopSmooth();

    if (emptyStateCard) {
        emptyStateCard.classList.add("hidden");
    }

    try {
        const response = await fetch("/cameras/run-check", {
            method: "POST"
        });

        const data = await response.json();

        if (!response.ok || !data.ok) {
            throw new Error(data.message || "Не удалось запустить проверку");
        }

        lastShownFinalStatus = null;
        await fetchStatus();
    } catch (error) {
        hideLoader();
        enableControlsAfterRun();
        setSystemReady();
        showToast(error.message || "Не удалось выполнить проверку", "error", true);
    }
}

async function uploadAddressTable(event) {
    event.preventDefault();

    if (!addressFileInput || !addressFileInput.files || !addressFileInput.files.length) {
        showUploadStatus("Выберите файл для загрузки", "error");
        return;
    }

    const file = addressFileInput.files[0];
    const formData = new FormData();
    formData.append("file", file);

    if (submitAddressUploadBtn) {
        submitAddressUploadBtn.disabled = true;
        submitAddressUploadBtn.textContent = "Загрузка...";
    }

    showUploadStatus("Файл загружается и преобразуется в TSV...", "loading");

    try {
        const response = await fetch("/api/address-table/upload", {
            method: "POST",
            body: formData
        });

        const data = await response.json();

        if (!response.ok || !data.ok) {
            throw new Error(data.detail || data.message || "Ошибка загрузки файла");
        }

        updateAddressTableInfo(data.table);
        showUploadStatus(data.message || "Таблица успешно обновлена", "success");
        showToast("Таблица адресов успешно обновлена", "success");
    } catch (error) {
        showUploadStatus(error.message || "Ошибка загрузки файла", "error");
        showToast(error.message || "Ошибка загрузки файла", "error");
    } finally {
        if (submitAddressUploadBtn) {
            submitAddressUploadBtn.disabled = false;
            submitAddressUploadBtn.textContent = "Загрузить и обновить";
        }
    }
}

function showUploadStatus(message, kind) {
    if (!uploadStatusBox) {
        return;
    }

    uploadStatusBox.textContent = message;
    uploadStatusBox.className = `upload-status-box ${kind}`;
}

if (themeToggleBtn) {
    themeToggleBtn.addEventListener("click", toggleTheme);
}

if (runCheckBtn) {
    runCheckBtn.addEventListener("click", startCheck);
}

if (aboutProjectBtn) {
    aboutProjectBtn.addEventListener("click", openAboutProjectModal);
}

if (toggleProblemSectionBtn) {
    toggleProblemSectionBtn.addEventListener("click", toggleProblemSection);
}

if (closeAboutProjectBtn) {
    closeAboutProjectBtn.addEventListener("click", closeAboutProjectModal);
}

if (closeAboutProjectFooterBtn) {
    closeAboutProjectFooterBtn.addEventListener("click", closeAboutProjectModal);
}

if (openAddressUploadBtn) {
    openAddressUploadBtn.addEventListener("click", openAddressUploadModal);
}

if (openAddressUploadBtnInline) {
    openAddressUploadBtnInline.addEventListener("click", openAddressUploadModal);
}

if (closeAddressUploadBtn) {
    closeAddressUploadBtn.addEventListener("click", closeAddressUploadModal);
}

if (closeAddressUploadFooterBtn) {
    closeAddressUploadFooterBtn.addEventListener("click", closeAddressUploadModal);
}

if (addressUploadForm) {
    addressUploadForm.addEventListener("submit", uploadAddressTable);
}

if (aboutProjectModal) {
    aboutProjectModal.addEventListener("click", (event) => {
        if (event.target === aboutProjectModal) {
            closeAboutProjectModal();
        }
    });
}

if (addressUploadModal) {
    addressUploadModal.addEventListener("click", (event) => {
        if (event.target === addressUploadModal) {
            closeAddressUploadModal();
        }
    });
}

document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
        closeAboutProjectModal();
        closeAddressUploadModal();
    }
});

applyTheme(localStorage.getItem("camera-dashboard-theme") || "dark");
updateProblemSectionCollapsedState();
fetchAddressTableInfo();
fetchStatus();