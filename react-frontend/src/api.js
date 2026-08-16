// Es una constante que guarda el URL del backend, el que usa FastAPI,
// este sirve para identificar el puerto que se utiliza
export const API_BASE = "http://localhost:8000";

// Se exporta el front-end
export const API_URL = import.meta.env.VITE_API_URL ?? "http://localhost:8000";

// Es una función asincrona. E sdecir, es una función que puede tardar 
// algún tiempo en completarse, ya que es una petición al backend.
// Este pide al backend que genere el .docx y lo descargue.
export async function generarWord() {

  // El URL final que usa el URL del backend y accede a la función que definimos en el main.py
  const url = `${API_BASE}/generate-report`;

  // Se realiza una petición HTTP con fetch.
  const resp = await fetch(url, {
    method: "POST",
  });

  // Se comprueba que el servidor funcione.
  if (!resp.ok) {
    const data = await resp.json().catch(() => ({}));
    const msg = data?.error || `Error HTTP ${resp.status}`;
    throw new Error(msg);
  }

  // Se lee el archivo como blob (trozo de datos binarios)
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

// Se crea una funcion para obtener los funcinoarios de la base de datos

export async function getFuncionarios() {
  const resp = await fetch(`${API_URL}/funcionarios`, {
    headers: {"Accept":"application/json" },
  });

  if (!resp.ok) throw new Error(`Error ${resp.status} al cargar funcionarios`);
  return await resp.json();
}

