
Office.onReady(() => {
    const log = msg => document.getElementById("consola").innerText = msg;

    const revisarErrores = () => {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, result => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                fetch("http://192.168.0.34:8000/revisar-errores", {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ texto: result.value })
                })
                .then(response => response.json())
                .then(data => document.getElementById("resultado").innerText = data.respuesta)
                .catch(error => log("❌ Error: " + error.message));
            }
        });
    };

    const mejorarRedaccion = () => {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, result => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                fetch("http://192.168.0.34:8000/mejorar-redaccion", {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ texto: result.value })
                })
                .then(response => response.json())
                .then(data => document.getElementById("resultado").innerText = data.respuesta)
                .catch(error => log("❌ Error: " + error.message));
            }
        });
    };

    const revisarPlanteamiento = () => {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, result => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                fetch("http://192.168.0.34:8000/revisar-planteamiento", {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ texto: result.value })
                })
                .then(response => response.json())
                .then(data => document.getElementById("resultado").innerText = data.respuesta)
                .catch(error => log("❌ Error: " + error.message));
            }
        });
    };

    const finalizar = () => {
        document.getElementById("resultado").innerText = "Finalización simulada: se guardarían Word y PDF.";
    };

    document.getElementById("btnErrores").onclick = revisarErrores;
    document.getElementById("btnRedaccion").onclick = mejorarRedaccion;
    document.getElementById("btnPlanteamiento").onclick = revisarPlanteamiento;
    document.getElementById("btnFinalizar").onclick = finalizar;

    log("✅ Joseph listo con Office.onReady - v36");
});
