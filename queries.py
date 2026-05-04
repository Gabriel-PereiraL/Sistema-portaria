# =============================================================================
# queries.py — Funções de acesso ao banco (CRUD)
# =============================================================================
# Responsabilidade: executar queries SQL e retornar os dados brutos.
# Não contém regras de negócio nem validações — isso fica em logic.py.
# Cada função abre e fecha sua própria conexão.
# =============================================================================

from datetime import datetime
from db import get_connection


def rg_ja_cadastrado(rg: str) -> bool:
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT id FROM visitante WHERE rg = %s", (rg,))
        resultado = cursor.fetchone()
        cursor.close()
        return resultado is not None
    finally:
        conn.close()


def cadastrar_visitante(nome: str, rg: str):
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("INSERT INTO visitante (nome, rg) VALUES (%s, %s)", (nome, rg))
        conn.commit()
        cursor.close()
    finally:
        conn.close()


def buscar_visitantes():
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT id, nome, rg FROM visitante ORDER BY nome")
        r = cursor.fetchall()
        cursor.close()
        return r
    finally:
        conn.close()


def buscar_funcionarios():
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT id, nome, setor FROM funcionario ORDER BY nome")
        r = cursor.fetchall()
        cursor.close()
        return r
    finally:
        conn.close()


def visitante_ja_esta_dentro(visitante_id: int) -> bool:
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute(
            "SELECT id FROM entrada WHERE visitante_id = %s AND data_hora_saida IS NULL",
            (visitante_id,),
        )
        r = cursor.fetchone()
        cursor.close()
        return r is not None
    finally:
        conn.close()





def registrar_entrada(visitante_id: int, funcionario_id: int, descricao: str = "", porteiro: str = ""):
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute(
            "INSERT INTO entrada "
            "(data_hora_entrada, data_hora_saida, visitante_id, funcionario_id, descricao, porteiro) "
            "VALUES (%s, NULL, %s, %s, %s, %s)",
            (datetime.now(), visitante_id, funcionario_id, descricao, porteiro),
        )
        conn.commit()
        cursor.close()
    finally:
        conn.close()





def buscar_visitantes_dentro():
    """
    Retorna visitantes sem saída registrada:
    [(id_entrada, nome_visitante, nome_funcionario, setor, data_hora_entrada), ...]
    """
    conn   = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT e.id, v.nome, f.nome, f.setor, e.data_hora_entrada
        FROM   entrada e
        JOIN   visitante   v ON e.visitante_id  = v.id
        JOIN   funcionario f ON e.funcionario_id = f.id
        WHERE  e.data_hora_saida IS NULL
        ORDER  BY e.data_hora_entrada DESC
    """)
    r = cursor.fetchall()
    cursor.close()
    conn.close()
    return r


def registrar_saida(entrada_id: int):
    """Atualiza o registro de entrada com o horário de saída atual."""
    conn   = get_connection()
    cursor = conn.cursor()
    agora  = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cursor.execute(
        "UPDATE entrada SET data_hora_saida = %s WHERE id = %s",
        (agora, entrada_id),
    )
    conn.commit()
    cursor.close()
    conn.close()


def buscar_historico_visitas():
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT e.id,
                   COALESCE(v.nome, '[Visitante removido]'),
                   COALESCE(f.nome, '[Funcionário removido]'),
                   COALESCE(f.setor, '—'),
                   e.descricao,
                   e.data_hora_entrada,
                   e.data_hora_saida,
                   COALESCE(e.porteiro, '') AS porteiro
            FROM   entrada e
            LEFT JOIN visitante   v ON e.visitante_id  = v.id
            LEFT JOIN funcionario f ON e.funcionario_id = f.id
            ORDER  BY e.data_hora_entrada DESC
        """)
        r = cursor.fetchall()
        cursor.close()
        return r
    finally:
        conn.close()


# ── Dashboard ─────────────────────────────────────────────────────────────────

def buscar_stats():
    conn = get_connection()
    try:
        cursor = conn.cursor()

        cursor.execute("SELECT COUNT(*) FROM visitante")
        visitantes = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM entrada WHERE data_hora_saida IS NULL")
        dentro = cursor.fetchone()[0]

        cursor.execute("""
            SELECT COUNT(*) FROM entrada 
            WHERE data_hora_entrada >= CURDATE()
        """)
        hoje = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM funcionario")
        funcionarios = cursor.fetchone()[0]

        return visitantes, dentro, hoje, funcionarios

    except Exception as e:
        print(f"[ERRO] buscar_stats: {e}")
        return 0, 0, 0, 0
    finally:
        conn.close()