from exstruct.io import dict_without_empty_values


def test_dict_without_empty_values_nested() -> None:
    data = {
        "a": 1,
        "b": "",
        "c": [],
        "d": {},
        "e": None,
        "f": {"x": "ok", "y": "", "z": {"k": None, "m": 2}},
        "g": [1, "", {}, []],
    }
    filtered = dict_without_empty_values(data)
    assert filtered == {"a": 1, "f": {"x": "ok", "z": {"m": 2}}, "g": [1]}
