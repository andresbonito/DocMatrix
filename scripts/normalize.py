"""Normalização conservadora: o valor original nunca é alterado."""
from __future__ import annotations
import re
from datetime import datetime
from typing import Any

_RE_BR = re.compile(r'^-?\d{1,3}(?:\.\d{3})+(?:,\d+)?$|^-?\d+,\d+$')
_RE_US = re.compile(r'^-?\d{1,3}(?:,\d{3})+(?:\.\d+)?$|^-?\d+\.\d+$')
_RE_INT = re.compile(r'^-?\d+$')
_RE_DATE_BR = re.compile(r'^(\d{2})/(\d{2})/(\d{4})$')
_RE_ISO = re.compile(r'^\d{4}-\d{2}-\d{2}$')


def normalize_value(value: Any) -> tuple[Any, str]:
    if value is None or isinstance(value, (int, float, bool)):
        return value, type(value).__name__ if value is not None else 'nulo'
    raw = str(value).strip()
    if not raw:
        return None, 'vazio'
    if _RE_ISO.fullmatch(raw):
        try:
            return datetime.strptime(raw, '%Y-%m-%d').strftime('%Y-%m-%d'), 'data'
        except ValueError:
            pass
    match = _RE_DATE_BR.fullmatch(raw)
    if match:
        try:
            return datetime.strptime(raw, '%d/%m/%Y').strftime('%Y-%m-%d'), 'data'
        except ValueError:
            pass
    clean = re.sub(r'^(?:R\$|US\$|€|£|\$)\s*', '', raw).strip()
    if _RE_INT.fullmatch(clean):
        # Evite coerção de identificadores com zero à esquerda (CNPJ, códigos etc.).
        if len(clean.lstrip('-')) > 1 and clean.lstrip('-').startswith('0'):
            return raw, 'texto'
        if len(clean.lstrip('-')) > 15:  # Precisão máxima exata do Excel.
            return raw, 'identificador'
        return int(clean), 'numero'
    if _RE_BR.fullmatch(clean):
        num = float(clean.replace('.', '').replace(',', '.'))
        return num, 'numero'
    if _RE_US.fullmatch(clean):
        num = float(clean.replace(',', ''))
        return num, 'numero'
    return raw, 'texto'


def excel_safe(value: Any) -> Any:
    """Evita formula injection em dados textuais não confiáveis."""
    if not isinstance(value, str):
        return value
    if value and (value[0] in ('=', '+', '-', '@', '\t', '\r', '\n')):
        return "'" + value
    return value
