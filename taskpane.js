
Office.onReady(() => {
  const resultado = document.getElementById("resultado");

  async function procesar(endpoint) {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, async (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const texto = result.value;
        try {
          const res = await fetch(`https://joseph-api-159952633905.europe-north1.run.app/${endpoint}`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ texto })
          });
          const data = await res.json();
          resultado.innerText = data.respuesta || "Sin respuesta.";
        } catch (err) {
          resultado.innerText = "❌ Error: " + err.message;
        }
      }
    });
  }

  document.getElementById("btnErrores").onclick = () => procesar("revisar-errores");
  document.getElementById("btnRedaccion").onclick = () => procesar("mejorar-redaccion");
  document.getElementById("btnPlanteamiento").onclick = () => procesar("revisar-planteamiento");
  document.getElementById("btnFinalizar").onclick = () => {
    resultado.innerText = "✅ Informe finalizado. Word y PDF generados (simulado).";
  };
});
