// src/pages/Menu.jsx
import { Link } from "react-router-dom";

export default function Menu(){
  return (
    <div style={{minHeight:"100vh", padding:"32px"}}>
      <header style={{display:"flex", alignItems:"center", gap:12, marginBottom:24}}>
        <div style={{width:10,height:10,borderRadius:"50%",background:"var(--inec-azul)"}} />
        <div style={{width:10,height:10,borderRadius:"50%",background:"var(--inec-amarillo)"}} />
        <div style={{width:10,height:10,borderRadius:"50%",background:"var(--inec-rojo)"}} />
        <h2 style={{margin:0}}>Panel principal</h2>
      </header>

      <nav style={{display:"grid", gap:12, maxWidth:520}}>
        
        <button style={pill("var(--accent)")}>Infrome de Productos</button>
        <button style={pill("var(--primary)")}>Informe de Actividades y Productos Realizados</button>
        <button style={pill("var(--danger)")}>Informe de Aceptación de Productos Entregados</button>
        <br />
        <br />
        <br />
        <br />
        <Link to="/" style={pill("var(--black)")}>Volver al Login</Link>
      </nav>
    </div>
  );
}

function pill(bg){
  return {
    padding:"14px 16px", borderRadius:12, background:bg, color:"#fff",
    border:"1px solid rgba(0,0,0,0.05)", textAlign:"left", fontWeight:600
  };
}
