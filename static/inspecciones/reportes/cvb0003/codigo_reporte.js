(function () {
    "use strict";

    const SUFIJOS = {
        FAJA: "FAJA",
        POLEAS: "POLEAS",
        LIFE_SHAFT: "LIFE-SHAFT",
    };

    const fechaInput = document.querySelector(
        'input[name="fecha_inspeccion"], input[name="inspeccion-fecha_inspeccion"]'
    );
    const codigoVisible = document.querySelector(
        "[data-codigo-reporte-cvb0003]"
    );

    if (!fechaInput || !codigoVisible) {
        return;
    }

    function actualizarCodigoVisible() {
        const partes = fechaInput.value.split("-");
        const sufijo = SUFIJOS[codigoVisible.dataset.tipoReporte];

        if (partes.length !== 3 || !sufijo) {
            return;
        }

        const fecha = `${partes[0]}${partes[1]}${partes[2]}`;
        codigoVisible.textContent = `${fecha}-VTUT-CVB0003-${sufijo}`;
    }

    fechaInput.addEventListener("input", actualizarCodigoVisible);
    fechaInput.addEventListener("change", actualizarCodigoVisible);
    actualizarCodigoVisible();
}());
