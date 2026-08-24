from datetime import date
from decimal import Decimal

from django.contrib.auth import get_user_model
from django.test import TestCase
from django.urls import reverse

from inspecciones.forms import MedicionPoleaFormSet, PoleaInspeccionForm
from inspecciones.models import (
    Faja,
    FaseCampana,
    Inspeccion,
    LifeShaftInspeccion,
    MedicionEmpalmeCVB0003,
    MedicionLifeShaft,
    MedicionPolea,
    MedicionPoleaCampana,
    PoleaInspeccion,
    TipoMedicionComponente,
)
from inspecciones.reportes.cvb0003.history import (
    historial_componente,
    historial_faja,
    modo_campana_seleccionado,
    validar_formset_historico,
)


class HistorialCVB0003Tests(TestCase):
    def setUp(self):
        usuario = get_user_model().objects.create_user("auditor", password="test")
        self.faja = Faja.objects.create(nombre="CVB003", tag="CVB0003")
        self.usuario = usuario

    def inspeccion(self, fecha, tipo=Inspeccion.Tipo.POLEAS):
        return Inspeccion.objects.create(
            faja=self.faja,
            tipo=tipo,
            codigo_reporte=f"TEMP-{tipo}-{fecha.isoformat()}",
            fecha_programada=fecha,
            fecha_inspeccion=fecha,
            inspector=self.usuario,
            supervisor=self.usuario,
            analista=self.usuario,
            creado_por=self.usuario,
        )

    def polea(self, inspeccion, modo=TipoMedicionComponente.NORMAL):
        return PoleaInspeccion.objects.create(
            inspeccion=inspeccion,
            numero=1,
            nombre="POLEA #01",
            tipo_medicion=modo,
        )

    def formset_normal(self, polea, valor):
        medicion = polea.mediciones.first()
        prefijo = f"mediciones-{polea.id}"
        data = {
            f"{prefijo}-TOTAL_FORMS": "1",
            f"{prefijo}-INITIAL_FORMS": "1",
            f"{prefijo}-MIN_NUM_FORMS": "0",
            f"{prefijo}-MAX_NUM_FORMS": "1000",
            f"{prefijo}-0-id": str(medicion.id),
            f"{prefijo}-0-punto": "1",
            f"{prefijo}-0-orden": "1",
            f"{prefijo}-0-posicion": "P1",
            f"{prefijo}-0-a": str(valor),
            f"{prefijo}-0-observacion": "",
        }
        return MedicionPoleaFormSet(data, instance=polea, prefix=prefijo)

    def test_caso_a_normal_permite_menor_y_bloquea_mayor(self):
        actual = self.polea(self.inspeccion(date(2026, 8, 4)))
        MedicionPolea.objects.create(polea=actual, punto=1, orden=1, a=Decimal("24.11"))
        mapa = {1: {"a": Decimal("24.11")}}

        permitido = self.formset_normal(actual, "23.80")
        self.assertTrue(validar_formset_historico(permitido, mapa))

        bloqueado = self.formset_normal(actual, "25.00")
        self.assertFalse(validar_formset_historico(bloqueado, mapa))
        self.assertIn("24.11 mm", bloqueado.forms[0].errors["a"][0])

    def test_caso_b_fin_permite_menor_y_bloquea_mayor(self):
        actual = self.polea(self.inspeccion(date(2026, 8, 5)))
        MedicionPolea.objects.create(polea=actual, punto=1, orden=1, a=Decimal("20.00"))
        mapa = {1: {"a": Decimal("20.00")}}
        self.assertTrue(validar_formset_historico(self.formset_normal(actual, "19.50"), mapa))
        self.assertFalse(validar_formset_historico(self.formset_normal(actual, "21.00"), mapa))

    def test_caso_c_inicio_no_se_restringe_y_define_referencia(self):
        anterior = self.polea(
            self.inspeccion(date(2026, 8, 3)), TipoMedicionComponente.CAMPANA
        )
        MedicionPoleaCampana.objects.create(
            polea=anterior, fase=FaseCampana.FIN, punto=1, orden=1, a=Decimal("15.00")
        )
        MedicionPoleaCampana.objects.create(
            polea=anterior, fase=FaseCampana.INICIO, punto=1, orden=1, a=Decimal("30.00")
        )
        actual = self.polea(self.inspeccion(date(2026, 8, 4)))
        MedicionPolea.objects.create(polea=actual, punto=1, orden=1, a=Decimal("30.00"))

        sin_restriccion = self.formset_normal(actual, "30.00")
        self.assertTrue(
            validar_formset_historico(
                sin_restriccion, {1: {"a": Decimal("15.00")}}, restringir=False
            )
        )
        historial = historial_componente(
            actual.inspeccion, actual, actual.inspeccion.fecha_inspeccion, "poleas_inspeccionadas"
        )
        self.assertEqual(historial["tipo"], "INICIO DE CAMPAÑA")
        self.assertEqual(historial["valores"][1]["a"], Decimal("30.00"))

    def test_caso_d_sin_historico_equivalente_permite(self):
        actual = self.polea(self.inspeccion(date(2026, 8, 6)))
        MedicionPolea.objects.create(polea=actual, punto=1, orden=1)
        self.assertTrue(validar_formset_historico(self.formset_normal(actual, "99.00"), {}))

    def test_caso_e_post_actual_controla_modo(self):
        actual = self.polea(self.inspeccion(date(2026, 8, 7)))
        prefijo = f"polea-{actual.id}"
        formulario = PoleaInspeccionForm(
            {f"{prefijo}-tipo_medicion": "CAMPANA"},
            instance=actual,
            prefix=prefijo,
        )
        self.assertTrue(modo_campana_seleccionado(formulario, actual))

        actual.tipo_medicion = TipoMedicionComponente.CAMPANA
        formulario_normal = PoleaInspeccionForm(
            {f"{prefijo}-tipo_medicion": "NORMAL"},
            instance=actual,
            prefix=prefijo,
        )
        self.assertFalse(modo_campana_seleccionado(formulario_normal, actual))

    def test_formularios_cvb0003_renderizan_historial_y_script(self):
        self.usuario.is_superuser = True
        self.usuario.save(update_fields=["is_superuser"])
        self.client.force_login(self.usuario)

        anterior = self.polea(self.inspeccion(date(2026, 8, 3)))
        MedicionPolea.objects.create(
            polea=anterior, punto=1, orden=1, a=Decimal("24.11")
        )
        actual = self.inspeccion(date(2026, 8, 4))
        respuesta = self.client.get(reverse("formulario_poleas", args=[actual.id]))
        self.assertEqual(respuesta.status_code, 200)
        self.assertContains(respuesta, "Historial técnico anterior")
        self.assertContains(respuesta, "historical_validation.js")
        self.assertContains(respuesta, 'data-historical-value="24.11"')

    def test_historial_funciona_para_faja_y_life_shaft(self):
        faja_anterior = self.inspeccion(date(2026, 8, 3), Inspeccion.Tipo.FAJA)
        fila_anterior = MedicionEmpalmeCVB0003.objects.create(
            inspeccion=faja_anterior,
            zona="Z1",
            empalme="E01",
            bastidor_lado="B1",
            posicion="P1",
            a=Decimal("18.20"),
        )
        faja_actual = self.inspeccion(date(2026, 8, 4), Inspeccion.Tipo.FAJA)
        fila_actual = MedicionEmpalmeCVB0003.objects.create(
            inspeccion=faja_actual,
            zona=fila_anterior.zona,
            empalme=fila_anterior.empalme,
            bastidor_lado=fila_anterior.bastidor_lado,
            posicion=fila_anterior.posicion,
        )
        mapa_faja = historial_faja(
            faja_actual, faja_actual.fecha_inspeccion, [fila_actual], "empalme"
        )
        self.assertEqual(mapa_faja[fila_actual.pk]["valores"]["a"], Decimal("18.20"))

        inspeccion_anterior = self.inspeccion(
            date(2026, 8, 3), Inspeccion.Tipo.LIFE_SHAFT
        )
        shaft_anterior = LifeShaftInspeccion.objects.create(
            inspeccion=inspeccion_anterior, numero=1, nombre="LIFE SHAFT #01"
        )
        MedicionLifeShaft.objects.create(
            life_shaft=shaft_anterior, punto=1, orden=1, a=Decimal("19.30")
        )
        inspeccion_actual = self.inspeccion(
            date(2026, 8, 4), Inspeccion.Tipo.LIFE_SHAFT
        )
        shaft_actual = LifeShaftInspeccion.objects.create(
            inspeccion=inspeccion_actual, numero=1, nombre="LIFE SHAFT #01"
        )
        mapa_shaft = historial_componente(
            inspeccion_actual,
            shaft_actual,
            inspeccion_actual.fecha_inspeccion,
            "life_shafts",
        )
        self.assertEqual(mapa_shaft["valores"][1]["a"], Decimal("19.30"))
