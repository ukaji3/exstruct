import pytest
from tests.utils import parametrize

from exstruct.core.shapes import angle_to_compass, compute_line_angle_deg


@parametrize(
    "angle,expected",
    [
        (0, "E"),
        (22.4, "E"),
        (22.6, "NE"),
        (44.9, "NE"),
        (45, "NE"),
        (67.4, "NE"),
        (67.6, "N"),
        (89.9, "N"),
        (90, "N"),
        (112.4, "N"),
        (112.6, "NW"),
        (134.9, "NW"),
        (135, "NW"),
        (157.4, "NW"),
        (157.6, "W"),
        (180, "W"),
        (202.4, "W"),
        (202.6, "SW"),
        (225, "SW"),
        (247.4, "SW"),
        (247.6, "S"),
        (270, "S"),
        (292.4, "S"),
        (292.6, "SE"),
        (315, "SE"),
        (337.4, "SE"),
        (337.6, "E"),
    ],
)
def test_angle_to_compass_8方位(angle: float, expected: str) -> None:
    assert angle_to_compass(angle) == expected


@parametrize(
    "w,h,expected",
    [
        (10.0, 0.0, 0.0),
        (0.0, 10.0, 90.0),
        (-10.0, 0.0, 180.0),
        (0.0, -10.0, 270.0),
        (10.0, 10.0, 45.0),
        (10.0, -10.0, 315.0),
    ],
)
def test_compute_line_angle_deg_座標系確認(w: float, h: float, expected: float) -> None:
    assert compute_line_angle_deg(w, h) == pytest.approx(expected, abs=1e-6)
