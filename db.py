# =============================================================================
# db.py — Conexão com o banco de dados
# =============================================================================

import mysql.connector



DB_CONFIG = {
    "host":     "localhost",
    "user":     "ti",
    "password": "235689",
    "database": "portaria",
    "port":     3306,
}


def get_connection():
    """
    Cria e retorna uma nova conexão com o banco MySQL usando DB_CONFIG.
    Deve ser chamada no início de cada função de banco e fechada ao final.
    """
    return mysql.connector.connect(**DB_CONFIG)