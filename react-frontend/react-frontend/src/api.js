export const API_BASE = "http://localhost:8000";

export async function generarWord() {
  const url = `${API_BASE}/generate-report`;
  // POST sin body (demo). Luego podrás enviar parámetros.
  const resp = await fetch(url, {
    method: "POST",
  });

  if (!resp.ok) {
    const data = await resp.json().catch(() => ({}));
    const msg = data?.error || `Error HTTP ${resp.status}`;
    throw new Error(msg);
  }

  // Recibir como blob y forzar descarga
  const blob = await resp.blob();
  const contentDisposition = resp.headers.get("content-disposition");
  let filename = "reporte.docx";

  // Intentar extraer nombre del archivo que envía el backend
  if (contentDisposition) {
    const match = contentDisposition.match(/filename="?([^"]+)"?/);
    if (match && match[1]) filename = match[1];
  }

  // Crear URL temporal y descargar
  const urlBlob = window.URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = urlBlob;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  window.URL.revokeObjectURL(urlBlob);
}
