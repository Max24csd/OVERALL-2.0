from pathlib import Path

from django.conf import settings


MASTER_PATH = (
    Path(settings.BASE_DIR)
    / "inspecciones"
    / "reportes"
    / "cvb0003"
    / "assets"
    / "master_poleas_cvb0003_fotos_grandes.xlsx"
)
MASTER_SHA256 = "3925A75DAC3DF5EAF8AE7F96BF2BEE9AA9D3F52261A9C3595C72D2D27E11D376"
SHEET_NAME = "Hoja1"
STATIC_IMAGE_MAX_ROW = 101
MEASUREMENT_FIELDS = tuple("abcdefg")


HEADER_CELLS = {
    "title": "M3",
    "condition": "K9",
    "plant": "K12",
    "process": "Y12",
    "equipment": "K14",
    "tag": "Y14",
    "stage": "K16",
    "equipment_condition": "Y16",
    "inspection_date": "K18",
    "report_date": "Y18",
    "inspector": "K20",
    "supervisor": "Y20",
    "author": "K22",
    "validator": "Y22",
    "circumstances": "K25",
    "background": "K27",
    "observations": "K28",
    "recommendation_rows": tuple(f"K{row}" for row in range(48, 58)),
    "diagram_title": "D59",
}


POLEA_BLOCKS = {
    1: {
        "slot_a": (73, 86), "slot_b": (87, 100), "visual": 101,
        "photo_rows": (102, 114), "caption": "E115",
        "data_offset": 5, "calibration_offset": 3, "inner_title_offset": 2,
        "title_column": "D", "inner_title_column": "W",
        "calibration_column": "N", "point_column": "Z",
        "measurement_columns": ("AC", "AE", "AG", "AI", "AK", "AM", "AO"),
        "average_column": "AQ", "minimum_column": "AT", "note_column": "Z",
        "photo_layout": "four_across",
    },
    2: {
        "slot_a": (117, 129), "slot_b": (130, 142), "visual": 143,
        "photo_rows": (144, 159), "caption": "D160",
        "data_offset": 4, "calibration_offset": 2, "inner_title_offset": 1,
        "title_column": "D", "inner_title_column": "W",
        "calibration_column": "N", "point_column": "Z",
        "measurement_columns": ("AC", "AE", "AG", "AI", "AK", "AM", "AO"),
        "average_column": "AQ", "minimum_column": "AT", "note_column": "AC",
        "photo_layout": "two_across",
    },
    3: {
        "slot_a": (164, 178), "slot_b": (179, 193), "visual": 194,
        "photo_rows": (195, 234), "caption": "D235",
        "data_offset": 5, "calibration_offset": 3, "inner_title_offset": 2,
        "title_column": "D", "inner_title_column": "W",
        "calibration_column": "N", "point_column": "Z",
        "measurement_columns": ("AC", "AE", "AG", "AI", "AK", "AM", "AO"),
        "average_column": "AQ", "minimum_column": "AT", "note_column": "Z",
        "photo_layout": "two_across",
    },
    4: {
        "slot_a": (236, 248), "slot_b": (249, 261), "visual": 262,
        "photo_rows": (263, 288), "caption": "D289",
        "data_offset": 4, "calibration_offset": 2, "inner_title_offset": 1,
        "title_column": "D", "inner_title_column": "W",
        "calibration_column": "N", "point_column": "Z",
        "measurement_columns": ("AC", "AE", "AG", "AI", "AK", "AM", "AO"),
        "average_column": "AQ", "minimum_column": "AT", "note_column": "AC",
        "photo_layout": "two_plus_one",
    },
    5: {
        "slot_a": (291, 303), "slot_b": (304, 316), "visual": 317,
        "photo_rows": (318, 343), "caption": "D344",
        "data_offset": 4, "calibration_offset": 2, "inner_title_offset": 1,
        "title_column": "D", "inner_title_column": "W",
        "calibration_column": "N", "point_column": "Z",
        "measurement_columns": ("AC", "AE", "AG", "AI", "AK", "AM", "AO"),
        "average_column": "AQ", "minimum_column": "AT", "note_column": "AC",
        "photo_layout": "grid_2x2",
    },
    6: {
        "slot_a": (346, 359), "slot_b": (360, 373), "visual": 374,
        "photo_rows": (375, 384), "caption": "D385",
        "data_offset": 5, "calibration_offset": 3, "inner_title_offset": 2,
        "title_column": "D", "inner_title_column": "W",
        "calibration_column": "N", "point_column": "Z",
        "measurement_columns": ("AC", "AE", "AG", "AI", "AK", "AM", "AO"),
        "average_column": "AQ", "minimum_column": "AT", "note_column": "Z",
        "photo_layout": "three_across",
    },
    7: {
        "slot_a": (386, 401), "slot_b": (402, 417), "visual": 418,
        "photo_rows": (419, 435), "caption": "D436",
        "data_offset": 5, "calibration_offset": 3, "inner_title_offset": 2,
        "title_column": "C", "inner_title_column": "V",
        "calibration_column": "M", "point_column": "Y",
        "measurement_columns": ("AB", "AD", "AF", "AH", "AJ", "AL", "AN"),
        "average_column": "AP", "minimum_column": "AS", "note_column": "Y",
        "photo_layout": "two_across",
    },
    8: {
        "slot_a": (437, 451), "slot_b": (452, 466), "visual": 467,
        "photo_rows": (468, 503), "caption": "D504",
        "data_offset": 5, "calibration_offset": 3, "inner_title_offset": 2,
        "title_column": "D", "inner_title_column": "W",
        "calibration_column": "N", "point_column": "Z",
        "measurement_columns": ("AC", "AE", "AG", "AI", "AK", "AM", "AO"),
        "average_column": "AQ", "minimum_column": "AT", "note_column": "Z",
        "photo_layout": "two_plus_one",
    },
    9: {
        "slot_a": (505, 518), "slot_b": (519, 532), "visual": 533,
        "photo_rows": (534, 555), "caption": "D556",
        "data_offset": 5, "calibration_offset": 3, "inner_title_offset": 2,
        "title_column": "D", "inner_title_column": "W",
        "calibration_column": "N", "point_column": "Z",
        "measurement_columns": ("AC", "AE", "AG", "AI", "AK", "AM", "AO"),
        "average_column": "AQ", "minimum_column": "AT", "note_column": "Z",
        "photo_layout": "two_across",
    },
}


HIDDEN_SHEET_POLEA_8 = {
    "sheet": "Hoja2",
    "title": "D8",
    "inner_title": "W10",
    "calibration_start": 11,
    "data_start": 13,
    "average": 18,
    "minimum": 19,
    "note": "Z20",
    "point_column": "Z",
    "measurement_columns": ("AC", "AE", "AG", "AI", "AK", "AM", "AO"),
    "average_column": "AQ",
    "minimum_column": "AT",
    "calibration_column": "N",
}
