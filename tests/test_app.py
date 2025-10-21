from src.app import greet


def test_greet():
    out = greet("אלעד")
    assert "אלעד" in out
