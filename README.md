# Sistema de Gestión de Funcionarios e Informes — INEC
 
Prototipo full-stack para la consulta de funcionarios civiles y la generación automatizada de informes en formato Word. Desarrollado como proyecto de práctica para el **Instituto Nacional de Estadística y Censos (INEC)**.
 
> **Estado del proyecto:** prototipo funcional en desarrollo. Login visual. Ver [Limitaciones](#limitaciones-y-pendientes).
 
## Descripción
 
El sistema permite:
- Consultar el listado de funcionarios civiles almacenado en una base de datos PostgreSQL.
- Seleccionar un funcionario desde una tabla interactiva.
- Generar y descargar automáticamente un informe en formato `.docx` con los datos del funcionario seleccionado.
## Arquitectura
 
```
├── python_backend/   # API REST (FastAPI)
└── react-frontend/   # Interfaz web (React + Vite)
```
 
**Backend → PostgreSQL**: consultas vía pool de conexiones.
**Frontend → Backend**: peticiones HTTP (fetch) al API REST, CORS habilitado para desarrollo local.
 
## Tech Stack
 
**Backend**
- Python 3.11 + FastAPI
- PostgreSQL (`psycopg2`, pool de conexiones singleton)
- `python-docx` para generación de documentos Word
- Pydantic para validación de datos
- Uvicorn como servidor ASGI
**Frontend**
- React 19 + Vite 7
- React Router DOM 7
- CSS personalizado (tema institucional INEC)
## Funcionalidades implementadas
 
- `GET /funcionarios` — Lista todos los funcionarios civiles (cédula, nombres, apellidos, dirección, cargo) desde la BD.
- `POST /api/seleccion/funcionario` — Registra el funcionario seleccionado en el estado de la sesión.
- `GET /api/seleccion/funcionario` — Devuelve la cédula actualmente seleccionada.
- `GET /report/cedula` — Genera y descarga el informe `.docx` del funcionario seleccionado.
- `GET /report/cedula/{cedula}` — Genera el informe directamente a partir de una cédula (sin selección previa).
- `GET /dbtest` — Verifica la conexión con PostgreSQL.
- Tabla interactiva en el frontend que consume `/funcionarios` y permite seleccionar filas.
## Instalación y ejecución
 
### Backend
 
```bash
cd python_backend
python -m venv .venv
source .venv/bin/activate        # En Windows: .venv\Scripts\activate
pip install fastapi uvicorn psycopg2-binary python-docx python-dotenv
 
# Crear archivo .env con las variables de conexión (ver sección Variables de entorno)
uvicorn main:app --reload
```
 
API disponible en `http://localhost:8000`.
 
### Frontend
 
```bash
cd react-frontend
npm install
npm run dev
```
 
App disponible en `http://localhost:5173`.
 
## Variables de entorno (backend)
 
Crear un archivo `.env` en `python_backend/` con:
 
```
PG_HOST=localhost
PG_PORT=5432
PG_DB=nombre_de_tu_base
PG_USER=usuario
PG_PASSWORD=contraseña
PG_MINCONN=1
PG_MAXCONN=5
```
 
## Esquema de base de datos esperado
 
El backend consulta las tablas `funcionario_civil` y `tdr` (términos de referencia), relacionadas por `id_tdr`, con al menos los siguientes campos:
 
- `funcionario_civil`: `cedula_ruc_civil`, `nombre_funcionario_civil`, `apellido_funcionario_civil`, `direccion_funcionario_civil`, `id_tdr`
- `tdr`: `id_tdr`, `cargo_tdr`
## Limitaciones y pendientes
 
- [ ] El login (`Login.jsx`) es solo visual; no valida credenciales ni implementa autenticación real.
- [ ] El documento `.docx` generado solo incluye 4 campos; falta plantilla institucional.
- [ ] Falta archivo `requirements.txt` con las dependencias fijadas del backend.
- [ ] Las credenciales de base de datos no deben mantenerse como valores por defecto en el código fuente.
## Autor

Daniela Pozo
Martin Herrera
