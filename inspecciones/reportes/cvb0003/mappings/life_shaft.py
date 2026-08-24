from pathlib import Path

from django.conf import settings


MASTER_PATH = (
    Path(settings.BASE_DIR)
    / "inspecciones"
    / "reportes"
    / "cvb0003"
    / "assets"
    / "master_life_shaft_cvb0003_fotos_grandes.xlsx"
)
MASTER_SHA256 = "76929AC204D23A4C6E34C8C6055CE4DDCDCEDA0B79445935972B32C6D166A945"
SHEET_NAME = "Hoja1"
STATIC_IMAGE_MAX_ROW = 67
MEASUREMENT_FIELDS = tuple("abcdefg")


HEADER_CELLS = {
    "title": "L1",
    "condition": "I7",
    "plant": "I10",
    "process": "V10",
    "equipment": "I12",
    "tag": "V12",
    "stage": "I14",
    "equipment_condition": "V14",
    "inspection_date": "I16",
    "report_date": "V16",
    "inspector": "I18",
    "supervisor": "V18",
    "author": "I20",
    "validator": "V20",
    "circumstances": "I23",
    "background": "I26",
    "observations": "I28",
    "recommendation_rows": tuple(f"I{row}" for row in range(45, 53)),
    "diagram_title": "C54",
}


LIFE_SHAFT_BLOCKS = {
    1: {
        "slot_a": (68, 82), "slot_b": (83, 97), "visual": 98,
        "photo_rows": (99, 117), "caption": "C118", "photo_layout": "two_across",
    },
    2: {
        "slot_a": (120, 134), "slot_b": (135, 149), "visual": 150,
        "photo_rows": (151, 169), "caption": "C170", "photo_layout": "three_across",
    },
    3: {
        "slot_a": (172, 186), "slot_b": (187, 201), "visual": 202,
        "photo_rows": (203, 221), "caption": "C222", "photo_layout": "one",
    },
    4: {
        "slot_a": (224, 238), "slot_b": (239, 253), "visual": 254,
        "photo_rows": (255, 274), "caption": "C275", "photo_layout": "three_across",
    },
    5: {
        "slot_a": (277, 291), "slot_b": (292, 306), "visual": 307,
        "photo_rows": (308, 324), "caption": "C325", "photo_layout": "grid",
    },
}


TECHNICAL_LAYOUT = {
    "title_column": "C",
    "inner_title_column": "V",
    "calibration_column": "M",
    "point_column": "Y",
    "measurement_columns": ("AB", "AD", "AF", "AH", "AJ", "AL", "AN"),
    "average_column": "AP",
    "minimum_column": "AS",
    "note_column": "V",
    "inner_title_offset": 2,
    "calibration_offset": 3,
    "data_offset": 5,
    "note_offset": 12,
    "measurement_count": 4,
}
