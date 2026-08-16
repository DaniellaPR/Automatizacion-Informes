# services/funcionarios_service.py
from typing import List, Dict, Optional
from ..db.postgres import get_cursor

def listar_funcionarios() -> List[Dict]:
    sql = """
        SELECT
            fc.cedula_ruc_civil                 AS cedula,
            fc.nombre_funcionario_civil         AS nombres,
            fc.apellido_funcionario_civil       AS apellidos,
            fc.direccion_funcionario_civil      AS direccion,
            COALESCE(t.cargo_tdr, '—')          AS cargo
        FROM funcionario_civil fc
        LEFT JOIN tdr t
            ON t.id_tdr = fc.id_tdr   -- <-- ajusta si tu relación es por otra columna
        ORDER BY fc.nombre_funcionario_civil, fc.apellido_funcionario_civil;
    """
    with get_cursor() as cur:
        cur.execute(sql)
        rows = cur.fetchall()
        cols = [d[0] for d in cur.description]
    return [dict(zip(cols, r)) for r in rows]


def obtener_funcionario_por_cedula(cedula: str) -> Optional[Dict]:
    sql = """
        SELECT
            fc.cedula_ruc_civil AS cedula,
            fc.nombre_funcionario_civil AS nombres,
            fc.apellido_funcionario_civil AS apellidos,
            fc.direccion_funcionario_civil AS direccion,
            COALESCE(t.cargo_tdr, '—') AS cargo

        FROM funcionario_civil fc
        LEFT JOIN tdr t ON t.id_tdr = fc.id_tdr
        WHERE fc.cedula_ruc_civil = %s
        LIMIT 1;
    """

    with get_cursor() as cur:
        cur.execute(sql,(cedula,))
        row = cur.fetchone()
        if not row:
            return None
        cols = [c[0] for c in cur.description]
        return dict(zip(cols,row))