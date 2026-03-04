"""
product_match.py — Поиск похожих продуктов для предложения связки вариант → канонический.
"""
from __future__ import annotations

import re


def _normalize(name: str) -> str:
    """Приводит название к виду для сравнения: нижний регистр, один пробел между словами."""
    if not name:
        return ""
    s = str(name).strip().lower()
    s = re.sub(r"\s+", " ", s)
    return s


def find_similar_canonicals(variant_name: str, canonical_names: list[str]) -> list[str]:
    """
    Ищет среди канонических названий те, с которыми вариант может совпадать.
    Возвращает список канонических названий (от более точного совпадения к менее).
    - Точное совпадение после нормализации
    - Один содержит другой (вариант в каноническом или наоборот)
    - Общее начало слов (первое слово совпадает)
    """
    if not variant_name or not canonical_names:
        return []
    vn = _normalize(variant_name)
    if not vn:
        return []
    exact = []
    contains = []
    first_word = []
    vn_words = set(vn.split())
    for c in canonical_names:
        cn = _normalize(c)
        if not cn:
            continue
        if vn == cn:
            exact.append(c)
            continue
        if vn in cn or cn in vn:
            contains.append(c)
            continue
        c_words = set(cn.split())
        if vn_words & c_words:
            first_word.append(c)
    return exact + contains + first_word
