
function revisarErrores() {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            const texto = result.value;
            const corregido = texto.replace(/\b(errores?|descuidos?)\b/gi, "correcciones automáticas");
            document.getElementById("resultado").innerText = "Texto corregido: " + corregido;
        }
    });
}

function mejorarRedaccion() {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            const texto = result.value;
            const mejorado = "Desde un enfoque técnico, " + texto.toLowerCase();
            document.getElementById("resultado").innerText = "Texto mejorado: " + mejorado;
        }
    });
}

function revisarPlanteamiento() {
    document.getElementById("resultado").innerText = "Análisis de cobertura simulado: informe conforme a condiciones estándar.";
}

function finalizar() {
    document.getElementById("resultado").innerText = "Guardando documento como Word y PDF... (simulado)";
}
