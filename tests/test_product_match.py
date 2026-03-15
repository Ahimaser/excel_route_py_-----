"""Тесты для product_match.find_similar_canonicals."""
import pytest
from core.product_match import find_similar_canonicals


def test_empty_input():
    assert find_similar_canonicals("", ["Молоко"]) == []
    assert find_similar_canonicals("Молоко", []) == []


def test_exact_match():
    assert find_similar_canonicals("Молоко", ["Молоко", "Кефир"]) == ["Молоко"]
    assert find_similar_canonicals("  молоко  ", ["Молоко"]) == ["Молоко"]


def test_contains():
    assert "Молоко 3.2%" in find_similar_canonicals("Молоко", ["Молоко 3.2%", "Кефир"])
    assert "Молоко" in find_similar_canonicals("Молоко 3.2%", ["Молоко", "Кефир"])


def test_first_word():
    result = find_similar_canonicals("Молоко пастеризованное", ["Молоко топлёное", "Кефир"])
    assert "Молоко топлёное" in result
