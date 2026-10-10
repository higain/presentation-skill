"""``#RRGGBB`` strings and the ints UNO stores colours as, both directions."""


def to_hex(value) -> str:
    """``0x2A7DE1`` -> ``"#2A7DE1"``."""
    return f"#{int(value) & 0xFFFFFF:06X}"


def parse_hex(value):
    """``"#2A7DE1"``, ``"2A7DE1"`` or an int -> int; ``None`` if unreadable."""
    if isinstance(value, int):
        return value
    try:
        return int(str(value).strip().lstrip("#"), 16)
    except (TypeError, ValueError):
        return None
