from project import suggest_item, add_to_cart, remove_from_cart
from _pytest.monkeypatch import MonkeyPatch


def test_suggest_item():
    assert suggest_item('apple', ['apple', 'mango']) == 1


def test_add_to_cart():
    # If item not present
    assert add_to_cart(item=['apple', 2.39, 7], cart=[]) == [['apple', 2.39, 7]]
    # If item already present
    assert add_to_cart(item=['apple', 2.39, 7], cart=[['apple', 2.39, 7]]) == [['apple', 2.39, 14]]


def test_remove_from_cart():
    # If item found in cart
    monkeypatch = MonkeyPatch()
    monkeypatch.setattr('builtins.input', lambda _: "3")
    assert remove_from_cart(item_name='apple', cart=[['apple', 2.39, 7]]) == [['apple', 2.39, 4]]
    # If item not in cart
    assert remove_from_cart(item_name='banana', cart=[['apple', 2.39, 7]]) == 0
