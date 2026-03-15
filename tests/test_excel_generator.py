"""Тесты для excel_generator: calc_pcs, apply_replacements, merge_replacement_pairs_for_display."""
import pytest
from core.excel_generator import calc_pcs, apply_replacements, merge_replacement_pairs_for_display


def test_calc_pcs_round_up():
    assert calc_pcs(2.7, 1.5, round_up=True) == 2  # 2.7/1.5=1.8 -> ceil=2
    assert calc_pcs(3.0, 1.5, round_up=True) == 2
    assert calc_pcs(3.1, 1.5, round_up=True) == 3


def test_calc_pcs_round_down():
    assert calc_pcs(2.7, 1.5, round_up=False) == 1
    assert calc_pcs(3.0, 1.5, round_up=False) == 2


def test_calc_pcs_zeros():
    assert calc_pcs(0, 1.5) == 0
    assert calc_pcs(3.0, 0) == 0


def test_apply_replacements_empty():
    routes = [{"routeNum": "1", "products": [{"name": "Молоко", "quantity": 10}]}]
    assert apply_replacements(routes, []) == routes


def test_apply_replacements_full():
    routes = [
        {"routeNum": "1", "products": [{"name": "Молоко", "quantity": 10, "unit": "л"}]},
        {"routeNum": "2", "products": [{"name": "Молоко", "quantity": 5, "unit": "л"}]},
    ]
    repl = [{"fromProduct": "Молоко", "toProduct": "Кефир"}]
    result = apply_replacements(routes, repl)
    assert len(result) == 2
    assert result[0]["products"][0]["name"] == "Кефир"
    assert result[0]["products"][0]["quantity"] == 10


def test_merge_replacement_pairs_empty():
    products = [{"name": "Молоко", "quantity": 10, "unit": "л"}]
    assert merge_replacement_pairs_for_display(products, []) == products


def test_merge_replacement_pairs_simple():
    products = [
        {"name": "Молоко", "quantity": 7, "unit": "л", "pcs": 7},
        {"name": "Кефир", "quantity": 3, "unit": "л", "pcs": 3},
    ]
    repl = [{"fromProduct": "Молоко", "toProduct": "Кефир"}]
    result = merge_replacement_pairs_for_display(products, repl)
    assert len(result) == 1
    assert "замена" in result[0]["name"]
    assert "Молоко" in result[0]["name"] and "Кефир" in result[0]["name"]
