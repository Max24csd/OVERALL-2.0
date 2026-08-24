from datetime import date
from decimal import Decimal
from pathlib import Path

from django.contrib.auth import get_user_model
from django.test import TestCase
from django.urls import reverse

from inspecciones.models import (
    Faja,
    FotoPolea,
    Inspeccion,
    LifeShaftInspeccion,
    MedicionPolea,
    MedicionPoleaCampana,
    PoleaInspeccion,
)


class HistorialVisibleCVB0003Tests(TestCase):
    def setUp(self):
        self.usuario = get_user_model().objects.create_superuser(
            "ui-auditor", "ui@example.com", "test"
        )
        self.faja = Faja.objects.create(nombre="CVB003", tag="CVB0003")
        self.client.force_login(self.usuario)

    def inspeccion(self, tipo, codigo, fecha=date(2026, 8, 3)):
        return Inspeccion.objects.create(
            faja=self.faja,
            tipo=tipo,
            codigo_reporte=codigo,
            fecha_programada=fecha,
            fecha_inspeccion=fecha,
            inspector=self.usuario,
            supervisor=self.usuario,
            analista=self.usuario,
            creado_por=self.usuario,
        )

    def test_poleas_muestra_historial_guardado_en_el_inicio_y_en_la_tabla(self):
        historica = self.inspeccion(Inspeccion.Tipo.POLEAS, "UI-POLEAS-HIST")
        polea = PoleaInspeccion.objects.create(
            inspeccion=historica, numero=1, nombre="POLEA #01"
        )
        MedicionPolea.objects.create(
            polea=polea,
            punto=1,
            orden=1,
            a=Decimal("15.11"),
            b=Decimal("14.93"),
        )
        actual = self.inspeccion(
            Inspeccion.Tipo.POLEAS, "UI-POLEAS-ACTUAL", date(2026, 8, 10)
        )
        PoleaInspeccion.objects.create(
            inspeccion=actual, numero=1, nombre="POLEA #01"
        )

        respuesta = self.client.get(reverse("formulario_poleas", args=[actual.id]))

        self.assertEqual(respuesta.status_code, 200)
        self.assertContains(respuesta, "Última inspección: 03/08/2026")
        self.assertContains(respuesta, "Mínimo anterior")
        self.assertContains(respuesta, "Promedio anterior")
        self.assertContains(respuesta, "Ver historial")
        self.assertContains(respuesta, 'data-historical-value="15.11"')
        self.assertContains(respuesta, "historical_validation.js")

    def test_life_shaft_muestra_historial_guardado(self):
        historica = self.inspeccion(Inspeccion.Tipo.LIFE_SHAFT, "UI-LIFE-HIST")
        shaft = LifeShaftInspeccion.objects.create(
            inspeccion=historica, numero=1, nombre="LIFE SHAFT #01"
        )
        from inspecciones.models import MedicionLifeShaft

        MedicionLifeShaft.objects.create(
            life_shaft=shaft, punto=1, orden=1, a=Decimal("19.30")
        )
        actual = self.inspeccion(
            Inspeccion.Tipo.LIFE_SHAFT, "UI-LIFE-ACTUAL", date(2026, 8, 10)
        )
        LifeShaftInspeccion.objects.create(
            inspeccion=actual, numero=1, nombre="LIFE SHAFT #01"
        )

        respuesta = self.client.get(
            reverse("formulario_life_shaft", args=[actual.id])
        )

        self.assertEqual(respuesta.status_code, 200)
        self.assertContains(respuesta, "Última inspección: 03/08/2026")
        self.assertContains(respuesta, "LIFE SHAFT #01")
        self.assertContains(respuesta, 'data-historical-value="19.30"')

    def test_dashboard_cuenta_solo_mediciones_activas_y_fotos_cvb0003(self):
        inspeccion = self.inspeccion(Inspeccion.Tipo.POLEAS, "UI-DASHBOARD")
        normal = PoleaInspeccion.objects.create(
            inspeccion=inspeccion, numero=1, nombre="POLEA #01", tipo_medicion="NORMAL"
        )
        campana = PoleaInspeccion.objects.create(
            inspeccion=inspeccion, numero=2, nombre="POLEA #02", tipo_medicion="CAMPANA"
        )
        MedicionPolea.objects.create(polea=normal, punto=1, orden=1, a=Decimal("10"))
        MedicionPolea.objects.create(polea=campana, punto=1, orden=1, a=Decimal("9"))
        MedicionPoleaCampana.objects.create(
            polea=campana, fase="INICIO", punto=1, orden=1, a=Decimal("20")
        )
        MedicionPoleaCampana.objects.create(
            polea=campana, fase="FIN", punto=1, orden=1, a=Decimal("19")
        )
        FotoPolea.objects.create(
            polea=normal, imagen="inspecciones/poleas/historica.jpg", subida_por=self.usuario
        )

        respuesta = self.client.get(reverse("dashboard"))
        fila = next(item for item in respuesta.context["inspecciones"] if item.pk == inspeccion.pk)

        self.assertEqual(fila.total_mediciones, 3)
        self.assertEqual(fila.total_fotos, 1)

    def test_javascript_contiene_referencia_y_error_inmediato(self):
        ruta = Path(__file__).resolve().parent.parent / "static" / "inspecciones" / "reportes" / "cvb0003" / "historical_validation.js"
        contenido = ruta.read_text(encoding="utf-8")
        self.assertIn("Anterior ", contenido)
        self.assertIn("Valor inválido", contenido)
        self.assertIn("cvb3-invalid", contenido)
        self.assertIn("Variación", contenido)
