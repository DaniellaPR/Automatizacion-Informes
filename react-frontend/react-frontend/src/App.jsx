import { useState } from "react";
import { generarWord } from "./api";

export default function App() {
  const [loading, setLoading] = useState(false);
  const [mensaje, setMensaje] = useState("");

  const handleClick = async () => {
    setMensaje("");
    setLoading(true);
    try {
      await generarWord();
      setMensaje("✅ Documento generado y descargado.");
    } catch (err) {
      console.error(err);
      setMensaje("❌ Hubo un error: " + err.message);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={{
      minHeight: "100vh",
      display: "grid",
      placeItems: "center",
      background: "#0f172a",
      color: "white",
      fontFamily: "system-ui, -apple-system, Segoe UI, Roboto, Arial"
    }}>
      <div style={{
        background: "#111827",
        padding: "2rem",
        borderRadius: "1rem",
        boxShadow: "0 10px 30px rgba(0,0,0,.5)",
        width: "min(520px, 92vw)",
        textAlign: "center"
      }}>
        <h1 style={{ margin: 0, fontSize: "1.5rem" }}>
          Generar informe (.docx)
        </h1>
        <p style={{ opacity: .8 }}>
          Frontend en React → llama a FastAPI → devuelve archivo.
        </p>

        <button
          onClick={handleClick}
          disabled={loading}
          style={{
            cursor: loading ? "not-allowed" : "pointer",
            background: loading ? "#374151" : "#2563eb",
            color: "white",
            border: "none",
            borderRadius: ".75rem",
            padding: ".9rem 1.25rem",
            fontSize: "1rem",
            fontWeight: 600,
            transition: "transform .05s ease",
          }}
        >
          {loading ? "Generando..." : "Generar Word"}
        </button>

        {mensaje && (
          <p style={{ marginTop: "1rem" }}>{mensaje}</p>
        )}

        <div style={{
          marginTop: "1.25rem",
          textAlign: "left",
          fontSize: ".9rem",
          opacity: .85,
          lineHeight: 1.5
        }}>
          <strong>¿Qué pasa al hacer clic?</strong>
          <ol>
            <li>React hace <code>POST /generate-report</code> al backend.</li>
            <li>FastAPI crea un <code>.docx</code> (por ahora demo).</li>
            <li>El navegador descarga el archivo automáticamente.</li>
          </ol>
        </div>
      </div>
    </div>
  );
}
