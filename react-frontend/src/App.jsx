import { BrowserRouter as Router, Routes, Route } from "react-router-dom";

import Login from "./pages/Login";
import FuncionariosTable from "./pages/FuncionariosTable";
import Menu_Principal from "./pages/Menu_Principal";
export default function App(){
  return (
    <Router>
      <Routes>
        <Route path="/" element={<Login />}/>
        <Route path="/funcionarios" element={<FuncionariosTable />}/>
        <Route path="/Menu_PrincipalB" element={<Menu_Principal />}/>

      </Routes>
    </Router>
  );
}