// src/pages/FuncionariosTable.jsx
import { useEffect, useState } from "react";
import { getFuncionarios } from "../api";

export default function FuncionariosTable() {
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  useEffect(() => {
    let alive = true;
    (async () => {
      try {
        setLoading(true);
        const data = await getFuncionarios();
        if (alive) setRows(data); // el backend ya envía: cedula, nombres, apellidos, direccion, cargo
      } catch (e) {
        if (alive) setError(e.message || "Error al cargar funcionarios");
      } finally {
        if (alive) setLoading(false);
      }
    })();
    return () => { alive = false; };
  }, []);

  if (loading) return <p>Cargando funcionarios…</p>;
  if (error) return <p style={{ color: "#c00" }}> {error}</p>;
  if (!rows.length) return <p>No hay funcionarios para mostrar.</p>;

  return (
    <div style={{ display: "grid", placeItems: "center", minHeight: "60vh", padding: "1rem" }}>
      <div style={{ width: "min(1100px, 95%)" }}>
        <h2 style={{ marginBottom: "1rem" }}>Selección de Personal</h2>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <colgroup>
              <col style={{ width: "14ch" }} />
              <col style={{ width: "22ch" }} />
              <col style={{ width: "24ch" }} />
              <col />
              <col style={{ width: "18ch" }} />
            </colgroup>
            <thead>
              <tr>
                <th style={th}>Cédula</th>
                <th style={th}>Nombres</th>
                <th style={th}>Apellidos</th>
                <th style={th}>Dirección</th>
                <th style={th}>Cargo</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((r, i) => (
                <tr 
                  key={i}
                  style = {{cursor: "pointer"}}
                  onClick = {async() =>{
                    localStorage.setItem("cedulaSeleccionada",r.cedula);

                    await fetch("http://localhost:8000/api/seleccion/funcionario", {
                      method: "POST",
                      headers:{"Content-Type": "application/json"},
                      body:JSON.stringify({cedula:r.cedula}),
                    });
                  }}
                >
                  <td style={td}>{r.cedula}</td>
                  <td style={td}>{r.nombres}</td>
                  <td style={td}>{r.apellidos}</td>
                  <td style={td}>{r.direccion}</td>
                  <td style={td}>{r.cargo}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

const th = { textAlign: "left", borderBottom: "2px solid #ddd", padding: ".6rem .75rem", fontWeight: 700 };
const td = { borderBottom: "1px solid #eee", padding: ".55rem .75rem", verticalAlign: "top" };
