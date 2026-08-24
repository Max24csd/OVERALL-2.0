from io import BytesIO
from math import floor

from PIL import Image as PILImage
from PIL import ImageOps
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.cell import range_boundaries
from openpyxl.utils.units import pixels_to_EMU


DEFAULT_COLUMN_WIDTH = 8.43
DEFAULT_ROW_HEIGHT_POINTS = 15


def _column_width_to_pixels(width):
    width = DEFAULT_COLUMN_WIDTH if width is None else width
    return max(floor(((256 * width + floor(128 / 7)) / 256) * 7), 1)


def _row_height_to_pixels(height):
    height = DEFAULT_ROW_HEIGHT_POINTS if height is None else height
    return max(int(round(height * 96 / 72)), 1)


def medir_rango_celdas_px(ws, rango):
    min_col, min_row, max_col, max_row = range_boundaries(rango)
    default_col = ws.sheet_format.defaultColWidth or DEFAULT_COLUMN_WIDTH
    default_row = ws.sheet_format.defaultRowHeight or DEFAULT_ROW_HEIGHT_POINTS
    ancho = 0
    for numero in range(min_col, max_col + 1):
        dimension = next(
            (d for d in ws.column_dimensions.values() if d.min <= numero <= d.max),
            None,
        )
        ancho += _column_width_to_pixels(
            dimension.width if dimension and dimension.width is not None else default_col
        )
    alto = 0
    for numero in range(min_row, max_row + 1):
        dimension = ws.row_dimensions.get(numero)
        alto += _row_height_to_pixels(
            dimension.height if dimension and dimension.height is not None else default_row
        )
    return ancho, alto


def tamano_comun_cuadros_px(ws, rangos):
    tamanos = [medir_rango_celdas_px(ws, rango) for rango in rangos]
    if not tamanos:
        return 0, 0
    return min(item[0] for item in tamanos), min(item[1] for item in tamanos)


def _imagen_excel(ruta):
    with PILImage.open(ruta) as origen:
        orientacion = origen.getexif().get(274, 1)
        ajustada = ImageOps.exif_transpose(origen)
        ancho, alto = ajustada.size
        if orientacion == 1:
            return ExcelImage(str(ruta)), ancho, alto
        normalizada = ajustada.copy()
        if normalizada.mode not in {"RGB", "RGBA"}:
            normalizada = normalizada.convert("RGB")
        buffer = BytesIO()
        normalizada.save(buffer, format="PNG")
        buffer.seek(0)
        imagen = ExcelImage(buffer)
        imagen._cvb0004_buffer = buffer
        return imagen, ancho, alto


def insertar_imagen_ajustada(
    ws, ruta_imagen, rango_celdas, ancho_max_px=None, alto_max_px=None, margen_px=4
):
    if not ruta_imagen:
        return False
    try:
        imagen, ancho_original, alto_original = _imagen_excel(ruta_imagen)
    except (FileNotFoundError, OSError, ValueError):
        return False
    ancho_rango, alto_rango = medir_rango_celdas_px(ws, rango_celdas)
    ancho_caja = min(ancho_max_px or ancho_rango, ancho_rango)
    alto_caja = min(alto_max_px or alto_rango, alto_rango)
    escala = min(
        max(ancho_caja - margen_px * 2, 1) / ancho_original,
        max(alto_caja - margen_px * 2, 1) / alto_original,
    )
    ancho = max(int(round(ancho_original * escala)), 1)
    alto = max(int(round(alto_original * escala)), 1)
    min_col, min_row, _max_col, _max_row = range_boundaries(rango_celdas)
    marcador = AnchorMarker(
        col=min_col - 1,
        row=min_row - 1,
        colOff=pixels_to_EMU(max((ancho_rango - ancho) // 2, 0)),
        rowOff=pixels_to_EMU(max((alto_rango - alto) // 2, 0)),
    )
    imagen.width = ancho
    imagen.height = alto
    imagen.anchor = OneCellAnchor(
        _from=marcador,
        ext=XDRPositiveSize2D(cx=pixels_to_EMU(ancho), cy=pixels_to_EMU(alto)),
    )
    ws.add_image(imagen)
    return True
