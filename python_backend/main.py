from fastapi import HTTPException
from fastapi.responses import FileResponse
from .services.report_service import generar_word_cedula
from pydantic import BaseModel
from .db.postgres import PostgresPool
from .db.postgres import get_cursor
from .services.funcionarios_service import listar_funcionarios
from .services.funcionarios_service import obtener_funcionario_por_cedula
from .config import PG_CONFIG
from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from .db.postgres import get_cursor

app = FastAPI()
app.state.CEDULA_SELECCIONADA = None

@app.get("/")
def root():
    return {"status": "running", "service": "python_backend"}


app.state.CEDULAS = []
app.state.NOMBRES = []
app.state.APELLIDOS = []
app.state.DIRECCIONES = []
app.state.CARGOS = []
app.state.CORREOS = []
app.state.HONORARIOS = []

origins = [
    "http://localhost:5173",
    "http://127.0.0.1:5173",
    "http://localhost:3000",
    "http://127.0.0.1:3000",
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,      
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

class CedulaRequest(BaseModel):
    cedula:str

# ---------- ENDPOINTS ----------
@app.get("/dbtest")
def db_test():
    try:
        with get_cursor() as cur:
            cur.execute("SELECT version();")
            version = cur.fetchone()[0]
        return {"status": "ok", "version": version}
    except Exception as e:
        return {"status": "error", "detail": str(e)}


@app.get("/funcionarios")
def get_funcionarios():
    """Datos directo desde PostgreSQL para tu tabla React."""
    return listar_funcionarios()

@app.post("/api/informes/seleccion")
async def recibir_informes(request: Request):
    data = await request.json()
    funcionarios = data.get("funcionarios", [])

    app.state.CEDULAS     = [f.get("cedula")     for f in funcionarios]
    app.state.NOMBRES     = [f.get("nombres")    for f in funcionarios]
    app.state.APELLIDOS   = [f.get("apellidos")  for f in funcionarios]
    app.state.DIRECCIONES = [f.get("direccion")  for f in funcionarios]
    app.state.CARGOS      = [f.get("cargo")      for f in funcionarios]
    app.state.CORREOS     = [f.get("correo")     for f in funcionarios]
    app.state.HONORARIOS  = [f.get("honorario")  for f in funcionarios]

    print(f"Datos Recibidos. Cédulas: {app.state.CEDULAS}")
    return {"status": "ok", "total": len(app.state.CEDULAS)}

@app.get("/api/informes/cedulas")
async def obtener_cedulas():
    return {"cedulas": app.state.CEDULAS}

@app.on_event("shutdown")
def shutdown_event():
    # Solo liberar el pool si existe. Nada de usar app.state aquí.
    try:
        PostgresPool().closeall()
    except Exception as e:
        print("Shutdown cleanup error:", e)

@app.post("/api/seleccion/funcionario")
async def seleccionar_funcionario(payload:CedulaRequest, request: Request):
    request.app.state.CEDULA_SELECCIONADA = payload.cedula
    return {"ok": True, "cedula":payload.cedula}

@app.get("/api/seleccion/funcionario")
async def obtener_funcionario_seleccionado(request: Request):
    return {"cedula": request.app.state.CEDULA_SELECCIONADA}

@app.get("/report/cedula")
def generar_reporte_desde_seleccion(request: Request):
    ced = request.app.state.CEDULA_SELECCIONADA

    if not ced:
        raise HTTPException(status_code=400, detail="No hay cedula seleccionada. Has click en una fila primero.")
    info = obtener_funcionario_por_cedula(ced)
    path = generar_word_cedula(ced, info)

    return FileResponse(
        path,
        medi_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=path.name,
    )
@app.get("/report/cedula/{cedula}")
def generar_reporte_parametro(cedula:str):
    info = obtener_funcionario_por_cedula(cedula)
    path = generar_word_cedula(cedula, info)
    return FileResponse (
        path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=path.name,
    )