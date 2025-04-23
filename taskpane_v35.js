
Office.onReady(() => {
    function revisarErrores() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                fetch("http://192.168.0.34:8000/revisar-errores", {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ texto: result.value })
                })
                .then(response => response.json())
                .then(data => {
                    document.getElementById("resultado").innerText = data.respuesta;
                });
            }
        });
    }

    function mejorarRedaccion() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                fetch("http://192.168.0.34:8000/mejorar-redaccion", {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ texto: result.value })
                })
                .then(response => response.json())
                .then(data => {
                    document.getElementById("resultado").innerText = data.respuesta;
                });
            }
        });
    }

    function revisarPlanteamiento() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                fetch("http://192.168.0.34:8000/revisar-planteamiento", {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ texto: result.value })
                })
                .then(response => response.json())
                .then(data => {
                    document.getElementById("resultado").innerText = data.respuesta;
                });
            }
        });
    }

    function finalizar() {
        document.getElementById("resultado").innerText = "Finalización simulada: se guardarían Word y PDF.";
    }

    // Asignar handlers a botones una vez cargado Office
    document.querySelector("button[onclick*='revisarErrores']").onclick = revisarErrores;
    document.querySelector("button[onclick*='mejorarRedaccion']").onclick = mejorarRedaccion;
    document.querySelector("button[onclick*='revisarPlanteamiento']").onclick = revisarPlanteamiento;
    document.querySelector("button[onclick*='finalizar']").onclick = finalizar;

    console.log("✅ Joseph listo dentro de Office.onReady - Versión v35");
});
