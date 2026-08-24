from datetime import date

from django.contrib.auth import get_user_model
from django.test import TestCase

from inspecciones.models import Faja, Inspeccion
from inspecciones.reportes.cvb0003.code_utils import generar_codigo_cvb0003


class CodigoReporteCVB0003Tests(TestCase):
    def setUp(self):
        self.usuario = get_user_model().objects.create_user(
            "codigo-cvb0003",
            "codigo@example.com",
            "test",
        )
        self.faja = Faja.objects.create(nombre="CVB003", tag="CVB0003")

    def crear_inspeccion(self, tipo, fecha_programada, fecha_inspeccion):
        return Inspeccion.objects.create(
            faja=self.faja,
            tipo=tipo,
            codigo_reporte=f"TEMP-{tipo}-{fecha_programada}",
            fecha_programada=fecha_programada,
            fecha_inspeccion=fecha_inspeccion,
            inspector=self.usuario,
            supervisor=self.usuario,
            analista=self.usuario,
            creado_por=self.usuario,
        )

    def test_funcion_central_formatea_fechas_y_tipos(self):
        fechas = (
            (date(2026, 8, 12), "20260812"),
            (date(2026, 8, 15), "20260815"),
            (date(2026, 9, 1), "20260901"),
        )
        tipos = (
            ("FAJA", "FAJA"),
            ("POLEAS", "POLEAS"),
            ("LIFE_SHAFT", "LIFE-SHAFT"),
        )
        for fecha, fecha_codigo in fechas:
            for tipo, sufijo in tipos:
                with self.subTest(fecha=fecha, tipo=tipo):
                    self.assertEqual(
                        generar_codigo_cvb0003(fecha, tipo),
                        f"{fecha_codigo}-VTUT-CVB0003-{sufijo}",
                    )

    def test_inspeccion_nueva_regenera_codigo_al_cambiar_fecha(self):
        inspeccion = self.crear_inspeccion(
            Inspeccion.Tipo.POLEAS,
            date(2026, 8, 10),
            date(2026, 8, 12),
        )
        self.assertEqual(
            inspeccion.codigo_reporte,
            "20260812-VTUT-CVB0003-POLEAS",
        )

        inspeccion.fecha_inspeccion = date(2026, 8, 15)
        inspeccion.fecha_reporte = date(2026, 9, 1)
        inspeccion.save()
        inspeccion.refresh_from_db()
        self.assertEqual(
            inspeccion.codigo_reporte,
            "20260815-VTUT-CVB0003-POLEAS",
        )

    def test_codigo_historico_emitido_permanece_intacto(self):
        historica = self.crear_inspeccion(
            Inspeccion.Tipo.FAJA,
            date(2026, 8, 3),
            date(2026, 8, 3),
        )
        codigo_emitido = "20260731-VTUT-CVB0003-FAJA"
        Inspeccion.objects.filter(pk=historica.pk).update(
            codigo_reporte=codigo_emitido
        )
        self.crear_inspeccion(
            Inspeccion.Tipo.FAJA,
            date(2026, 8, 10),
            date(2026, 8, 12),
        )

        historica.refresh_from_db()
        historica.fecha_reporte = date(2026, 8, 8)
        historica.save()
        historica.refresh_from_db()
        self.assertEqual(historica.codigo_reporte, codigo_emitido)
