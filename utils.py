# =============================================================================
# utils.py — Funções auxiliares puras
# =============================================================================
# Responsabilidade: funções utilitárias sem dependência de banco ou UI.
# =============================================================================

from datetime import datetime


def formatar_tempo_permanencia(entrada, saida=None) -> str:
    """
    Calcula e formata o tempo de permanência entre entrada e saída.
    - Se 'saida' não for fornecida, usa o horário atual (visita em andamento).
    - Retorna no formato "01h 30min".
    - Retorna "—" se 'entrada' não for um objeto datetime válido.
    """
    if not isinstance(entrada, datetime):
        return "—"
    fim           = saida if isinstance(saida, datetime) else datetime.now()
    delta         = fim - entrada
    total_minutos = int(delta.total_seconds() // 60)
    h = total_minutos // 60
    m = total_minutos % 60
    return f"{h:02d}h {m:02d}min"