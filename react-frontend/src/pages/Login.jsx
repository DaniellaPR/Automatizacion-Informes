// src/pages/Login.jsx
import { useNavigate } from "react-router-dom";

export default function Login() {
  const navigate = useNavigate();

  const handleLogin = (e) => {
    e.preventDefault();
    navigate("/funcionarios");
  };

  return (
    <div style ={{
      minHeight: "100vh",
      display: "flex",
      alignItems: "center",
      justifyContent:"center",
      backgroundImage:"url('/INEC_V2.jpg')",
      backgroundSize:"cover",
      backgroundPosition: "center"
    }}>
      <div style={styles.wrapper}>
        <div style={styles.card}>
          <div style={styles.brandStrip}>
            <div style={styles.rings}>
              <span style={{...styles.ring, background:"var(--inec-azul)"}} />
              <span style={{...styles.ring, background:"var(--inec-amarillo)"}} />
              <span style={{...styles.ring, background:"var(--inec-rojo)"}} />
            </div>
            <h1 style={styles.title}>INEC · Acceso</h1>
            <p style={styles.subtitle}>Instituto Nacional de Estadística y Censos</p>
          </div>

          <form onSubmit={handleLogin} style={styles.form}>
            <label style={styles.label}>Usuario</label>
            <input style={styles.input} type="text" placeholder="tu_correo@inec.gob.ec" required/>
            <label style={styles.label}>Contraseña</label>
            <input style={styles.input} type="password" placeholder="••••••••" required/>

            <button type="submit" style={styles.buttonPrimary}>Entrar</button>
            <p style={styles.help}>
              * Prototipo visual (no funcional).
            </p>
          </form>
        </div>
      </div>
    </div>
  );
}

const styles = {
  wrapper:{
    minHeight:"100vh", display:"grid", placeItems:"center", padding:"clamp(16px, 3vw, 40px)"
  },
  card:{
    width:"100%", maxWidth:480, background:"var(--card)", borderRadius:18,
    boxShadow:"var(--shadow)", overflow:"hidden", border:"1px solid #e5e7eb"
  },
  brandStrip:{
    padding:"28px 24px",
    background: "linear-gradient(135deg, var(--inec-azul) 0%, #0a3ea6 50%, var(--inec-amarillo) 120%)",
    color:"#fff", position:"relative"
  },
  rings:{
    position:"absolute", right:16, top:16, display:"flex", gap:8
  },
  ring:{
    width:14, height:14, borderRadius:"50%", opacity:0.95, boxShadow:"0 0 0 3px rgba(255,255,255,0.25)"
  },
  title:{ margin:"0 0 4px 0", fontSize:24, letterSpacing:0.4 },
  subtitle:{ margin:0, opacity:0.9 },

  form:{ padding:24, display:"grid", gap:12 },
  label:{ fontSize:14, color:"var(--muted)" },
  input:{
    padding:"12px 14px", borderRadius:12, border:"1px solid #dbe3f0",
    outline:"none", background:"#fff"
  },
  buttonPrimary:{
    marginTop:8, padding:"12px 14px", borderRadius:12, border:"1px solid var(--primary)",
    background:"var(--primary)", color:"#fff", fontWeight:600,
    boxShadow:"0 0 0 3px var(--ring)"
  },
  help:{ margin:"8px 0 0", fontSize:12, color:"var(--muted)" }
};
