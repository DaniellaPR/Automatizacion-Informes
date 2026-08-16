import React, { useState } from "react";


function Menu_Principal() {

  const [step, setStep] = useState("seleccion");

  const [tipo, setTipo] = useState(null);

 

  const tiposDeInforme = ["Info_Productos", "Info_Act_Prod_Realizados", "Info_Aceptación"];

 

  const informesPorTipo = {

    Info_Productos: ["Producto 1", "Producto 2", "Todos"],

    Info_Act_Prod_Realizados: ["Informe de Actividades y productos realizados"],

    Info_Aceptación: ["Producto 1", "Producto 2", "Todos", "Personalizados"]

  };

 

  return (

    <div className="app-container">

      {step === "seleccion" && (

        <div className="card">

          <img

            src="logo_inec.jpg"

            alt="INEC"

            title="INEC"

            className="logo"

            style = {{width:"300px", height:"120px"}}


          />
          <h1 className="titulo">Selecciona un tipo de informe</h1>
          <div className="opciones">
            {tiposDeInforme.map((nombre, idx) => (
              <button
                key={idx}
                onClick={() => {
                  setTipo(nombre);
                  setStep("informes");
                }}
                className="btn btn-primary"
              >
                {nombre}
              </button>
            ))}
          </div>
        </div>
      )}
      {/* Pantalla de informes por tipo */}
      {step === "informes" && (
        <div className="card">
          <h1 className="titulo">{`Informes de ${tipo}`}</h1>
          <div className="opciones">
            {informesPorTipo[tipo].map((informe, idx) => (
              <button
                key={idx}
                onClick={() => alert(`Aquí se descargará: ${informe}`)}
                className="btn btn-accent"
              >
                {informe}
              </button>
            ))}
          </div>
          <button
            onClick={() => setStep("seleccion")}
            className="btn btn-secondary"
          >
            Volver
          </button>
        </div>
      )}
    </div>
  );
}

 

export default Menu_Principal;