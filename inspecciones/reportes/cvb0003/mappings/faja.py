from pathlib import Path

from django.conf import settings


LEGACY_MASTER_PATH = (
    Path(settings.BASE_DIR)
    / "inspecciones"
    / "reportes"
    / "cvb0003"
    / "assets"
    / "master_faja_cvb0003.xlsx"
)
MASTER_PATH = LEGACY_MASTER_PATH.with_name("master_faja_cvb0003_paginado.xlsx")
MASTER_SHA256 = "64B39D0EE0399C913E4C0A98FB1B5127E6BD2B9EB041DA06D70E7E7AA95ACA2A"
SHEET_NAME = "REPORTE DE INSPECCION CV0003"
CONTINUATION_SHEET_NAME = "FOTOS CONTINUACION"
CONTINUATION_PAGE_COUNT = 80
CONTINUATION_ROWS_PER_PAGE = 40
CONTINUATION_PHOTO_ROWS = (2, 37)
CONTINUATION_SUMMARY_OFFSET = 37
STATIC_IMAGE_MAX_ROW = 216
MEASUREMENT_FIELDS = tuple("abcdefg")

DIAGRAM_IMAGE_PATH = (
    Path(settings.BASE_DIR)
    / "static"
    / "inspecciones"
    / "faja"
    / "cvb003"
    / "diagrama_empalmes_cvb0003.png"
)
DIAGRAM_IMAGE_RANGE = "C192:U213"


HEADER_CELLS = {
    "title": "F2",
    "condition": "F7",
    "plant": "F10",
    "process": "L10",
    "equipment": "F12",
    "tag": "L12",
    "stage": "F14",
    "equipment_condition": "L14",
    "inspection_date": "F16",
    "report_date": "L16",
    "inspector": "F18",
    "supervisor": "L18",
    "author": "F20",
    "validator": "L20",
    "circumstances": "F23",
    "background": "F28",
    "observations": "F30",
    "recommendations": "F105",
}


UT_CALIBRATION_BLOCKS = tuple(
    {
        "number": number,
        "marca_equipo": f"I{start}",
        "modelo_equipo": f"I{start + 1}",
        "frecuencia_mhz": f"I{start + 2}",
        "rango_mm": f"I{start + 3}",
        "metodo_empleado": f"I{start + 4}",
        "acoplante": f"P{start}",
        "rectificacion": f"P{start + 1}",
        "velocidad_ms": f"P{start + 2}",
        "retardo_us": f"P{start + 3}",
        "tipo_scan": f"P{start + 4}",
    }
    for number, start in enumerate(range(141, 184, 6), start=1)
)


EMPALME_DATA_ROWS = tuple(range(222, 312))
CARGA_DATA_ROWS = tuple(range(893, 1013))
RETORNO_DATA_ROWS = tuple(range(1022, 1124))
VALUE_COLUMNS = ("H", "I", "J", "K", "L", "M", "N")


# Rango de imagen y rango de descripción obtenidos de los anchors aprobados.
EMPALME_PHOTO_SLOTS = (
    ("H322:O340", "C358"), ("O322:U340", "C358"),
    ("C331:G354", "C358"), ("H342:N357", "C358"),
    ("O342:U357", "C358"), ("C360:H378", "C399"),
    ("I360:O378", "C399"), ("O360:V378", "C399"),
    ("I379:O398", "C399"), ("C402:H420", "C441"),
    ("I402:O420", "C441"), ("O402:U420", "C441"),
    ("O420:U439", "C441"), ("C421:H440", "C441"),
    ("I421:O439", "C441"), ("C444:H459", "C477"),
    ("I444:O459", "C477"), ("P444:U459", "C477"),
    ("C461:H476", "C477"), ("I461:N476", "C477"),
    ("P461:U476", "C477"), ("C480:H499", "C528"),
    ("I480:O499", "C528"), ("P480:U499", "C528"),
    ("C500:H518", "C528"), ("I500:O520", "C528"),
    ("C531:H550", "C580"), ("I531:O551", "C580"),
    ("P531:V550", "C580"), ("C556:H575", "C580"),
    ("I556:O575", "C580"), ("P556:V574", "C580"),
    ("C583:K594", "C625"), ("M583:U594", "C625"),
    ("C594:K607", "C625"), ("M595:U607", "C625"),
    ("C607:K624", "C625"), ("M607:U624", "C625"),
    ("I629:O646", "C657"), ("C630:H645", "C657"),
    ("C646:H656", "C657"), ("I648:O657", "C657"),
    ("P648:U656", "C657"), ("C660:H678", "C690"),
    ("I660:O678", "C690"), ("P660:V678", "C690"),
    ("C679:H689", "C690"), ("I679:O689", "C690"),
    ("P679:U689", "C690"), ("C693:H711", "C723"),
    ("I693:O711", "C723"), ("P693:V711", "C723"),
    ("C712:H722", "C723"), ("I713:O722", "C723"),
    ("C726:K746", "C748"), ("L726:U746", "C748"),
    ("C746:K747", "C748"), ("L746:U747", "C748"),
    ("C751:H770", "C793"), ("I751:O769", "C793"),
    ("P751:U768", "C793"), ("I770:O789", "C793"),
    ("P770:V788", "C793"), ("C771:H789", "C793"),
    ("D796:K819", "C870"), ("L796:U819", "C870"),
    ("D819:K842", "C870"), ("L819:U842", "C870"),
    ("D843:K868", "C870"), ("C873:H882", "C885"),
    ("I873:O882", "C885"), ("O873:U882", "C885"),
    ("C882:H884", "C885"), ("I882:O884", "C885"),
    ("O882:U884", "C885"),
)


EMPALME_PHOTO_GROUPS = (
    {"identification": "EMPALME E-16 (RETORNO)", "title": "C317", "rows": (322, 357), "caption": "C358"},
    {"identification": "EMPALME E-01 (RETORNO)", "title": "C359", "rows": (360, 398), "caption": "C399"},
    {"identification": "EMPALME E-03 (RETORNO)", "title": "C401", "rows": (402, 440), "caption": "C441"},
    {"identification": "EMPALME E-04 (RETORNO)", "title": "C443", "rows": (444, 476), "caption": "C477"},
    {"identification": "EMPALME E-05 (RETORNO)", "title": "C479", "rows": (480, 527), "caption": "C528"},
    {"identification": "EMPALME E-06 (RETORNO)", "title": "C530", "rows": (531, 579), "caption": "C580"},
    {"identification": "EMPALME E-07 (CARGA)", "title": "C582", "rows": (583, 624), "caption": "C625"},
    {"identification": "EMPALME E-09 (CARGA)", "title": "C627", "rows": (629, 656), "caption": "C657"},
    {"identification": "EMPALME E-10 (CARGA)", "title": "C659", "rows": (660, 689), "caption": "C690"},
    {"identification": "EMPALME E-11 (CARGA)", "title": "C692", "rows": (693, 722), "caption": "C723"},
    {"identification": "EMPALME E-12 (CARGA)", "title": "C725", "rows": (726, 747), "caption": "C748"},
    {"identification": "EMPALME E-13 (CARGA)", "title": "C750", "rows": (751, 792), "caption": "C793"},
    {"identification": "EMPALME E-14 (CARGA)", "title": "C795", "rows": (796, 869), "caption": "C870"},
    {"identification": "EMPALME E-15 (CARGA)", "title": "C872", "rows": (873, 884), "caption": "C885"},
)


GENERAL_PHOTO_ROWS = (
    (1131, 1157), (1159, 1185), (1187, 1214), (1216, 1242),
    (1244, 1270), (1272, 1295), (1298, 1321), (1323, 1346),
    (1348, 1371), (1373, 1395), (1397, 1420), (1423, 1446),
    (1449, 1472), (1474, 1497), (1499, 1522),
)

GENERAL_PHOTO_SLOTS = tuple(
    item
    for start, description in GENERAL_PHOTO_ROWS
    for item in (
        (f"C{start}:K{description - 1}", f"C{description}"),
        (f"L{start}:T{description - 1}", f"L{description}"),
    )
) + (
    ("C1524:J1545", "C1546"), ("M1524:U1545", "L1546"),
    ("D1549:K1571", "C1572"), ("M1549:T1571", "L1572"),
)
