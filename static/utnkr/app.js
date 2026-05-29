(function () {
    const SERVICE_BASE = "/utnkr";

    const root = document.documentElement;

    const syncButton = document.getElementById("syncButton");

    const loaderBlock = document.getElementById("loaderBlock");
    const statusError = document.getElementById("statusError");

    const systemStateBadge = document.getElementById("systemStateBadge");
    const statusText = document.getElementById("statusText");
    const currentStageMessage = document.getElementById("currentStageMessage");
    const stateMetricValue = document.getElementById("stateMetricValue");

    const aboutProjectBtn = document.getElementById("aboutProjectBtn");
    const aboutProjectModal = document.getElementById("aboutProjectModal");
    const closeAboutProjectBtn = document.getElementById("closeAboutProjectBtn");
    const closeAboutProjectFooterBtn = document.getElementById("closeAboutProjectFooterBtn");

    const violatorsCollapseBtn = document.getElementById("violatorsCollapseBtn");
    const violatorsSectionBody = document.getElementById("violatorsSectionBody");

    let statusTimer = null;

    function showElement(element) {
        if (element) {
            element.classList.remove("hidden");
        }
    }

    function hideElement(element) {
        if (element) {
            element.classList.add("hidden");
        }
    }

    function setError(message) {
        if (!statusError) {
            alert(message);
            return;
        }

        statusError.textContent = message;
        statusError.classList.remove("hidden");
    }

    function clearError() {
        if (!statusError) return;

        statusError.textContent = "";
        statusError.classList.add("hidden");
    }

    function updateStatusUi(status) {
        const running = Boolean(status.running);

        if (statusText) {
            statusText.textContent = status.stage || "Ожидание запуска";
        }

        if (currentStageMessage) {
            currentStageMessage.textContent = status.message || "Система готова к запуску проверки.";
        }

        if (stateMetricValue) {
            stateMetricValue.textContent = running ? "Проверка" : "Готово";
        }

        if (systemStateBadge) {
            systemStateBadge.classList.remove("running", "ready");

            if (running) {
                systemStateBadge.classList.add("running");
                systemStateBadge.textContent = "Выполняется проверка";
            } else {
                systemStateBadge.classList.add("ready");
                systemStateBadge.textContent = "Готова";
            }
        }

        if (syncButton) {
            syncButton.disabled = running;
            syncButton.textContent = running ? "Проверка выполняется..." : "Запустить проверку";
        }

        if (running) {
            showElement(loaderBlock);
        } else {
            hideElement(loaderBlock);
        }

        if (status.last_error) {
            setError(status.last_error);
        }
    }

    async function fetchRunStatus() {
        const response = await fetch(`${SERVICE_BASE}/run-status`, {
            method: "GET",
            cache: "no-store"
        });

        if (!response.ok) {
            const text = await response.text();
            throw new Error(`HTTP ${response.status}: ${text}`);
        }

        return await response.json();
    }

    async function pollStatusOnce() {
        try {
            const status = await fetchRunStatus();
            updateStatusUi(status);

            if (!status.running && statusTimer) {
                clearInterval(statusTimer);
                statusTimer = null;

                if (!status.last_error && status.stage === "Готово") {
                    window.location.reload();
                }
            }
        } catch (error) {
            console.error(error);
            setError(`Не удалось получить статус проверки: ${error.message}`);

            if (statusTimer) {
                clearInterval(statusTimer);
                statusTimer = null;
            }
        }
    }

    function startStatusPolling() {
        if (statusTimer) {
            clearInterval(statusTimer);
        }

        pollStatusOnce();
        statusTimer = setInterval(pollStatusOnce, 1500);
    }

    async function runCheck() {
        if (!syncButton) return;

        try {
            clearError();

            syncButton.disabled = true;
            syncButton.textContent = "Запуск...";

            showElement(loaderBlock);

            const response = await fetch(`${SERVICE_BASE}/run-check`, {
                method: "POST"
            });

            const text = await response.text();

            let payload = null;

            try {
                payload = text ? JSON.parse(text) : null;
            } catch (error) {
                payload = null;
            }

            if (!response.ok) {
                const message =
                    payload && payload.message
                        ? payload.message
                        : text || "Backend вернул ошибку.";

                throw new Error(`HTTP ${response.status}: ${message}`);
            }

            if (payload && payload.ok === false) {
                throw new Error(payload.message || "Проверка не была запущена.");
            }

            startStatusPolling();
        } catch (error) {
            console.error(error);
            setError(`Не удалось запустить проверку. ${error.message}`);

            syncButton.disabled = false;
            syncButton.textContent = "Запустить проверку";

            hideElement(loaderBlock);
        }
    }

    function initRunButton() {
        if (!syncButton) return;

        syncButton.addEventListener("click", function () {
            runCheck();
        });
    }

    function initAboutModal() {
        if (aboutProjectBtn && aboutProjectModal) {
            aboutProjectBtn.addEventListener("click", function () {
                aboutProjectModal.classList.remove("hidden");
            });
        }

        if (closeAboutProjectBtn && aboutProjectModal) {
            closeAboutProjectBtn.addEventListener("click", function () {
                aboutProjectModal.classList.add("hidden");
            });
        }

        if (closeAboutProjectFooterBtn && aboutProjectModal) {
            closeAboutProjectFooterBtn.addEventListener("click", function () {
                aboutProjectModal.classList.add("hidden");
            });
        }

        if (aboutProjectModal) {
            aboutProjectModal.addEventListener("click", function (event) {
                if (event.target === aboutProjectModal) {
                    aboutProjectModal.classList.add("hidden");
                }
            });
        }

        document.addEventListener("keydown", function (event) {
            if (event.key === "Escape" && aboutProjectModal) {
                aboutProjectModal.classList.add("hidden");
            }
        });
    }

    function initViolatorsCollapse() {
        if (!violatorsCollapseBtn || !violatorsSectionBody) return;

        violatorsCollapseBtn.addEventListener("click", function () {
            const expanded = violatorsCollapseBtn.getAttribute("aria-expanded") === "true";
            const nextExpanded = !expanded;

            violatorsCollapseBtn.setAttribute("aria-expanded", String(nextExpanded));

            const text = violatorsCollapseBtn.querySelector(".violators-collapse-btn-text");
            const icon = violatorsCollapseBtn.querySelector(".violators-collapse-btn-icon");

            if (nextExpanded) {
                violatorsSectionBody.classList.remove("hidden");
                if (text) text.textContent = "Свернуть";
                if (icon) icon.textContent = "−";
            } else {
                violatorsSectionBody.classList.add("hidden");
                if (text) text.textContent = "Развернуть";
                if (icon) icon.textContent = "+";
            }
        });
    }

    async function initInitialStatus() {
        try {
            const status = await fetchRunStatus();
            updateStatusUi(status);

            if (status.running) {
                startStatusPolling();
            }
        } catch (error) {
            console.warn("Не удалось получить начальный статус:", error);
        }
    }

    document.addEventListener("DOMContentLoaded", function () {
        initRunButton();
        initAboutModal();
        initViolatorsCollapse();
        initInitialStatus();
    });
})();
