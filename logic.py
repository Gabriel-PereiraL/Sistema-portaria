from queries import (
    rg_ja_cadastrado,
    cadastrar_visitante,
    visitante_ja_esta_dentro,
    registrar_entrada,
)

LOGIN_USUARIO = "TI"
LOGIN_SENHA = "TI123"


def validar_login(usuario, senha):
    return usuario == LOGIN_USUARIO and senha == LOGIN_SENHA


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