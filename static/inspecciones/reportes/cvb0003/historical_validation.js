(function () {
    "use strict";

    const letters = new Set(["a", "b", "c", "d", "e", "f", "g"]);

    function fieldLetter(input) {
        const part = (input.name || "").split("-").pop().toLowerCase();
        return letters.has(part) ? part : null;
    }

    function validateInput(input) {
        const maximum = Number(input.dataset.historicalValue);
        const current = Number(input.value);
        const letter = (fieldLetter(input) || "").toUpperCase();
        let message = "";

        if (input.value !== "" && Number.isFinite(maximum) && current > maximum) {
            message = "Valor inválido. La nueva medición no puede ser mayor que la medición histórica equivalente: " + input.dataset.historicalValue + " mm.";
        }

        input.setCustomValidity(message);
        input.classList.toggle("cvb3-invalid", Boolean(message));
        input.setAttribute("aria-invalid", message ? "true" : "false");

        let hint = input.parentElement.querySelector(".cvb3-history-hint");
        if (!hint) {
            hint = document.createElement("small");
            hint.className = "cvb3-history-hint";
            hint.style.cssText = "display:block;color:#334155;margin-top:3px;font-weight:600";
            input.insertAdjacentElement("afterend", hint);
        }
        hint.textContent = "Anterior " + letter + ": " + input.dataset.historicalValue + " mm";

        let error = input.parentElement.querySelector(".cvb3-field-error");
        if (message && !error) {
            error = document.createElement("small");
            error.className = "cvb3-field-error";
            error.style.cssText = "display:block;color:#b91c1c;font-weight:700;margin-top:3px";
            input.insertAdjacentElement("afterend", error);
        }
        if (error) {
            error.textContent = message;
            if (!message) error.remove();
        }
        return !message;
    }

    function measurementInputs(container) {
        return Array.from(container.querySelectorAll("input[type='number']"))
            .filter(function (input) { return fieldLetter(input); });
    }

    function numericValues(container) {
        return measurementInputs(container)
            .filter(function (input) { return input.value !== ""; })
            .map(function (input) { return Number(input.value); })
            .filter(Number.isFinite);
    }

    function updateTable(table) {
        const headers = Array.from(table.querySelectorAll("thead th"));
        const minIndex = headers.findIndex(function (th) {
            return th.textContent.trim().toLowerCase().startsWith("mín");
        });
        const avgIndex = headers.findIndex(function (th) {
            return th.textContent.trim().toLowerCase().startsWith("prom");
        });

        table.querySelectorAll("tbody tr").forEach(function (row) {
            const values = numericValues(row);
            const cells = row.querySelectorAll("td");
            if (minIndex >= 0 && cells[minIndex]) {
                cells[minIndex].textContent = values.length ? Math.min.apply(null, values).toFixed(2) : "-";
            }
            if (avgIndex >= 0 && cells[avgIndex]) {
                cells[avgIndex].textContent = values.length
                    ? (values.reduce(function (a, b) { return a + b; }, 0) / values.length).toFixed(2)
                    : "-";
            }
        });
    }

    function phaseSummary(section, previousMin, compareWithHistory) {
        let summary = section.querySelector(":scope > .cvb3-phase-summary");
        if (!summary) {
            summary = document.createElement("div");
            summary.className = "cvb3-phase-summary";
            summary.style.cssText = "margin:8px 0;padding:8px;background:#e5eef7;font-weight:700;color:#18324f";
            section.appendChild(summary);
        }

        const values = numericValues(section);
        if (!values.length) {
            summary.textContent = "Mínimo actual: - · Promedio actual: -";
            return null;
        }

        const currentMin = Math.min.apply(null, values);
        const currentAvg = values.reduce(function (a, b) { return a + b; }, 0) / values.length;
        let text = "Mínimo actual: " + currentMin.toFixed(2) + " mm · Promedio actual: " + currentAvg.toFixed(2) + " mm";
        if (compareWithHistory && Number.isFinite(previousMin)) {
            text += " · Mínimo anterior: " + previousMin.toFixed(2) + " mm · Variación: " + (currentMin - previousMin).toFixed(2) + " mm";
        }
        summary.textContent = text;
        return currentMin;
    }

    function updateBlock(block) {
        block.querySelectorAll("table").forEach(updateTable);
        const history = block.querySelector(".cvb3-history");
        const previousMin = history ? Number(history.dataset.historicalMinimum) : NaN;
        const selector = block.querySelector(".tipo-medicion-selector");
        const campaign = selector && selector.value === "CAMPANA";
        const sections = campaign
            ? Array.from(block.querySelectorAll(".campaign-phase"))
            : Array.from(block.querySelectorAll(".normal-measurements"));

        let comparableCurrentMin = null;
        sections.forEach(function (section) {
            const isStart = section.textContent.toUpperCase().includes("INICIO DE CAMPAÑA");
            const currentMin = phaseSummary(section, previousMin, !isStart);
            if (!isStart && currentMin !== null && comparableCurrentMin === null) {
                comparableCurrentMin = currentMin;
            }
        });

        const historySummary = block.querySelector(".cvb3-live-summary");
        if (historySummary) {
            const previous = Number.isFinite(previousMin) ? previousMin.toFixed(2) + " mm" : "-";
            const current = comparableCurrentMin === null ? "-" : comparableCurrentMin.toFixed(2) + " mm";
            const variation = comparableCurrentMin === null || !Number.isFinite(previousMin)
                ? "-"
                : (comparableCurrentMin - previousMin).toFixed(2) + " mm";
            historySummary.textContent = "Mínimo anterior: " + previous + " · Mínimo actual: " + current + " · Variación: " + variation;
        }
    }

    document.addEventListener("DOMContentLoaded", function () {
        const form = document.querySelector("form");

        document.querySelectorAll("[data-history-validation='1']").forEach(function (input) {
            const block = input.closest(".polea-block, .shaft-block") || form;
            const run = function () {
                validateInput(input);
                if (block) updateBlock(block);
            };
            input.addEventListener("input", run);
            input.addEventListener("change", run);
            run();
        });

        document.querySelectorAll(".polea-block, .shaft-block").forEach(function (block) {
            measurementInputs(block).forEach(function (input) {
                input.addEventListener("input", function () { updateBlock(block); });
            });
            const selector = block.querySelector(".tipo-medicion-selector");
            if (selector) selector.addEventListener("change", function () { updateBlock(block); });
            updateBlock(block);
        });

        if (form) {
            form.addEventListener("submit", function (event) {
                let valid = true;
                document.querySelectorAll("[data-history-validation='1']").forEach(function (input) {
                    valid = validateInput(input) && valid;
                });
                if (!valid) {
                    event.preventDefault();
                    const firstInvalid = form.querySelector(":invalid");
                    if (firstInvalid) {
                        firstInvalid.scrollIntoView({ behavior: "smooth", block: "center" });
                        firstInvalid.reportValidity();
                    }
                }
            });
        }
    });
}());
