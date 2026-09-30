import hashlib
import hmac
import os

from queries import (
    rg_ja_cadastrado,
    cadastrar_visitante,
    visitante_ja_esta_dentro,
    registrar_entrada,
)

LOGIN_USUARIO = os.environ.get("PORTARIA_LOGIN_USER", "TI")

def validar_login(usuario, senha):
    encoded = os.environ.get("PORTARIA_LOGIN_PASSWORD_HASH")
    if not encoded:
        return False
    try:
        algorithm, iterations, salt_hex, digest_hex = encoded.split("$", 3)
        if algorithm != "pbkdf2_sha256":
            return False
        actual = hashlib.pbkdf2_hmac("sha256", senha.encode("utf-8"), bytes.fromhex(salt_hex), int(iterations)).hex()
        return hmac.compare_digest(usuario, LOGIN_USUARIO) and hmac.compare_digest(actual, digest_hex)
    except (ValueError, TypeError):
        return False

def tentar_cadastrar_visitante(nome, rg):
    if len(nome) < 3:
        return False, "Nome inválido"
    if not rg:
        return False, "RG obrigatório"
    if rg_ja_cadastrado(rg):
        return False, "RG já cadastrado"
    cadastrar_visitante(nome, rg)
    return True, "Visitante cadastrado com sucesso"

def tentar_registrar_entrada(v_id, f_id, descricao, porteiro):
    if visitante_ja_esta_dentro(v_id):
        return False, "Visitante já está dentro"
    registrar_entrada(v_id, f_id, descricao, porteiro)
    return True, "Entrada registrada com sucesso"
