from datetime import date
from unittest.mock import patch

from django.contrib.auth.models import Group, User
from django.test import RequestFactory, TestCase
from django.urls import reverse

from inspecciones.models import Faja, Inspeccion
from inspecciones.reportes.cvb0003.permissions import (
    puede_acceder_inspeccion_cvb0003,
)
from inspecciones.views import _aplicar_accion_flujo, puede_editar_inspeccion


class SeguridadAsignacionCVB0003Tests(TestCase):
    @classmethod
    def setUpTestData(cls):
        cls.grupos = {
            nombre: Group.objects.create(name=nombre)
            for nombre in (
                "Administrador",
                "Inspector",
                "Supervisor",
                "Analista",
                "Cliente",
            )
        }
        cls.usuarios = {}
        for nombre in cls.grupos:
            usuario = User.objects.create_user(
                username=nombre.lower(), password="segura-123"
            )
            usuario.groups.add(cls.grupos[nombre])
            cls.usuarios[nombre] = usuario

        for nombre in ("Inspector", "Supervisor", "Analista", "Cliente"):
            usuario = User.objects.create_user(
                username=f"{nombre.lower()}-otro", password="segura-123"
            )
            usuario.groups.add(cls.grupos[nombre])
            cls.usuarios[f"{nombre} otro"] = usuario

        cls.admin = User.objects.create_superuser(
            username="root-security", email="root@example.com", password="segura-123"
        )
        cls.faja = Faja.objects.create(nombre="Faja 3", tag="CVB0003")
        cls.inspecciones = {}
        for indice, tipo in enumerate(Inspeccion.Tipo.values, start=1):
            cls.inspecciones[tipo] = Inspeccion.objects.create(
                faja=cls.faja,
                tipo=tipo,
                codigo_reporte=f"2026090{indice}-VTUT-CVB0003-{tipo}",
                fecha_programada=date(2026, 9, indice),
                fecha_inspeccion=date(2026, 9, indice),
                inspector=cls.usuarios["Inspector"],
                supervisor=cls.usuarios["Supervisor"],
                analista=cls.usuarios["Analista"],
                cliente=cls.usuarios["Cliente"],
                creado_por=cls.admin,
            )

    def setUp(self):
        self.factory = RequestFactory()

    def _set_estado(self, inspeccion, estado):
        Inspeccion.objects.filter(pk=inspeccion.pk).update(estado=estado)
        inspeccion.refresh_from_db()

    def test_matriz_helper_por_asignacion_y_estado(self):
        inspeccion = self.inspecciones[Inspeccion.Tipo.POLEAS]
        self.assertTrue(puede_acceder_inspeccion_cvb0003(self.admin, inspeccion))
        self.assertTrue(
            puede_acceder_inspeccion_cvb0003(self.usuarios["Inspector"], inspeccion)
        )
        self.assertFalse(
            puede_acceder_inspeccion_cvb0003(self.usuarios["Inspector otro"], inspeccion)
        )
        self.assertTrue(puede_editar_inspeccion(self.usuarios["Inspector"], inspeccion))

        self._set_estado(inspeccion, Inspeccion.Estado.DEVUELTO)
        self.assertTrue(puede_editar_inspeccion(self.usuarios["Inspector"], inspeccion))
        self._set_estado(inspeccion, Inspeccion.Estado.EN_REVISION)
        self.assertFalse(puede_editar_inspeccion(self.usuarios["Inspector"], inspeccion))
        self.assertTrue(puede_editar_inspeccion(self.usuarios["Supervisor"], inspeccion))
        self.assertFalse(
            puede_acceder_inspeccion_cvb0003(self.usuarios["Supervisor otro"], inspeccion)
        )
        self._set_estado(inspeccion, Inspeccion.Estado.REVISADO)
        self.assertTrue(puede_editar_inspeccion(self.usuarios["Analista"], inspeccion))
        self.assertFalse(
            puede_acceder_inspeccion_cvb0003(self.usuarios["Analista otro"], inspeccion)
        )
        self._set_estado(inspeccion, Inspeccion.Estado.PUBLICADO)
        self.assertTrue(
            puede_acceder_inspeccion_cvb0003(self.usuarios["Cliente"], inspeccion)
        )
        self.assertFalse(
            puede_acceder_inspeccion_cvb0003(self.usuarios["Cliente otro"], inspeccion)
        )

    def test_cliente_no_puede_ver_borrador_propio(self):
        for inspeccion in self.inspecciones.values():
            self.assertFalse(
                puede_acceder_inspeccion_cvb0003(
                    self.usuarios["Cliente"], inspeccion, "ver"
                )
            )

    def test_inspector_no_puede_editar_en_revision(self):
        rutas = {
            Inspeccion.Tipo.FAJA: "formulario_faja",
            Inspeccion.Tipo.POLEAS: "formulario_poleas",
            Inspeccion.Tipo.LIFE_SHAFT: "formulario_life_shaft",
        }
        self.client.force_login(self.usuarios["Inspector"])
        for tipo, ruta in rutas.items():
            inspeccion = self.inspecciones[tipo]
            self._set_estado(inspeccion, Inspeccion.Estado.EN_REVISION)
            respuesta = self.client.post(
                reverse(ruta, args=[inspeccion.pk]),
                {"workflow_action": "guardar"},
            )
            self.assertEqual(respuesta.status_code, 403)

    def test_formularios_y_reportes_rechazan_personal_no_asignado(self):
        casos = {
            Inspeccion.Tipo.FAJA: ("formulario_faja", "reporte_faja"),
            Inspeccion.Tipo.POLEAS: ("formulario_poleas", "reporte_poleas"),
            Inspeccion.Tipo.LIFE_SHAFT: (
                "formulario_life_shaft",
                "reporte_life_shaft",
            ),
        }
        for tipo, rutas in casos.items():
            inspeccion = self.inspecciones[tipo]
            for rol in ("Inspector otro", "Supervisor otro", "Analista otro"):
                self.client.force_login(self.usuarios[rol])
                for ruta in rutas:
                    with self.subTest(tipo=tipo, rol=rol, ruta=ruta):
                        respuesta = self.client.get(reverse(ruta, args=[inspeccion.pk]))
                        self.assertEqual(respuesta.status_code, 403)

    def test_accesos_asignados_por_estado_en_los_tres_tipos(self):
        rutas = {
            Inspeccion.Tipo.FAJA: ("formulario_faja", "reporte_faja"),
            Inspeccion.Tipo.POLEAS: ("formulario_poleas", "reporte_poleas"),
            Inspeccion.Tipo.LIFE_SHAFT: (
                "formulario_life_shaft",
                "reporte_life_shaft",
            ),
        }
        for tipo, (ruta_formulario, ruta_reporte) in rutas.items():
            inspeccion = self.inspecciones[tipo]

            self.client.force_login(self.admin)
            self.assertEqual(
                self.client.get(reverse(ruta_formulario, args=[inspeccion.pk])).status_code,
                200,
            )

            self.client.force_login(self.usuarios["Inspector"])
            self.assertEqual(
                self.client.get(reverse(ruta_formulario, args=[inspeccion.pk])).status_code,
                200,
            )

            self._set_estado(inspeccion, Inspeccion.Estado.EN_REVISION)
            self.client.force_login(self.usuarios["Supervisor"])
            self.assertEqual(
                self.client.get(reverse(ruta_formulario, args=[inspeccion.pk])).status_code,
                200,
            )

            self._set_estado(inspeccion, Inspeccion.Estado.REVISADO)
            self.client.force_login(self.usuarios["Analista"])
            self.assertEqual(
                self.client.get(reverse(ruta_formulario, args=[inspeccion.pk])).status_code,
                200,
            )

            self._set_estado(inspeccion, Inspeccion.Estado.PUBLICADO)
            self.client.force_login(self.usuarios["Cliente"])
            self.assertEqual(
                self.client.get(reverse(ruta_reporte, args=[inspeccion.pk])).status_code,
                200,
            )

    def test_exportaciones_rechazan_no_asignados_y_otro_cliente(self):
        rutas = {
            Inspeccion.Tipo.FAJA: "exportar_excel_faja_cvb0003",
            Inspeccion.Tipo.POLEAS: "exportar_excel_poleas_cvb0003",
            Inspeccion.Tipo.LIFE_SHAFT: "exportar_excel_life_shaft_cvb0003",
        }
        for tipo, ruta in rutas.items():
            inspeccion = self.inspecciones[tipo]
            self.client.force_login(self.usuarios["Inspector otro"])
            self.assertEqual(
                self.client.get(reverse(ruta, args=[inspeccion.pk])).status_code,
                403,
            )
            self._set_estado(inspeccion, Inspeccion.Estado.PUBLICADO)
            self.client.force_login(self.usuarios["Cliente otro"])
            self.assertEqual(
                self.client.get(reverse(ruta, args=[inspeccion.pk])).status_code,
                403,
            )

    @patch("inspecciones.reportes.cvb0003.faja_views.generar_excel_faja_cvb0003_master")
    @patch("inspecciones.reportes.cvb0003.poleas_views.generar_excel_poleas_cvb0003_master")
    @patch("inspecciones.reportes.cvb0003.views.generar_excel_life_shaft_cvb0003_master")
    def test_cliente_asignado_descarga_publicados_de_los_tres_tipos(
        self, generar_life, generar_poleas, generar_faja
    ):
        from io import BytesIO

        for generador in (generar_faja, generar_poleas, generar_life):
            generador.return_value = BytesIO(b"xlsx")
        rutas = {
            Inspeccion.Tipo.FAJA: "exportar_excel_faja_cvb0003",
            Inspeccion.Tipo.POLEAS: "exportar_excel_poleas_cvb0003",
            Inspeccion.Tipo.LIFE_SHAFT: "exportar_excel_life_shaft_cvb0003",
        }
        self.client.force_login(self.usuarios["Cliente"])
        for tipo, ruta in rutas.items():
            inspeccion = self.inspecciones[tipo]
            self._set_estado(inspeccion, Inspeccion.Estado.PUBLICADO)
            respuesta = self.client.get(reverse(ruta, args=[inspeccion.pk]))
            self.assertEqual(respuesta.status_code, 200)

    def test_endpoint_estado_rechaza_usuario_no_asignado(self):
        inspeccion = self.inspecciones[Inspeccion.Tipo.POLEAS]
        self.client.force_login(self.usuarios["Inspector otro"])
        respuesta = self.client.post(
            reverse("cambiar_estado_inspeccion", args=[inspeccion.pk, "enviar_revision"])
        )
        self.assertEqual(respuesta.status_code, 403)
        inspeccion.refresh_from_db()
        self.assertEqual(inspeccion.estado, Inspeccion.Estado.BORRADOR)

    def test_endpoint_estado_rechaza_accion_fuera_del_rol(self):
        inspeccion = self.inspecciones[Inspeccion.Tipo.FAJA]
        self._set_estado(inspeccion, Inspeccion.Estado.EN_REVISION)
        self.client.force_login(self.usuarios["Inspector"])
        respuesta = self.client.post(
            reverse(
                "cambiar_estado_inspeccion",
                args=[inspeccion.pk, "aprobar_supervisor"],
            )
        )
        self.assertEqual(respuesta.status_code, 403)
        inspeccion.refresh_from_db()
        self.assertEqual(inspeccion.estado, Inspeccion.Estado.EN_REVISION)

    def test_dashboard_cvb0003_no_expone_inspecciones_ajenas(self):
        propia = self.inspecciones[Inspeccion.Tipo.FAJA]
        ajena = self.inspecciones[Inspeccion.Tipo.POLEAS]
        Inspeccion.objects.filter(pk=ajena.pk).update(
            inspector=self.usuarios["Inspector otro"]
        )
        self.client.force_login(self.usuarios["Inspector"])
        respuesta = self.client.get(reverse("dashboard"))
        self.assertEqual(respuesta.status_code, 200)
        visibles = {
            inspeccion.pk
            for inspeccion in respuesta.context["inspecciones_nuevas_cvb0003"]
        } | {
            inspeccion.pk
            for inspeccion in respuesta.context["historial_completo_cvb0003"]
        }
        self.assertIn(propia.pk, visibles)
        self.assertNotIn(ajena.pk, visibles)

    def test_dashboard_cliente_solo_incluye_publicados_propios(self):
        propia = self.inspecciones[Inspeccion.Tipo.FAJA]
        ajena = self.inspecciones[Inspeccion.Tipo.POLEAS]
        Inspeccion.objects.filter(pk=propia.pk).update(
            estado=Inspeccion.Estado.PUBLICADO
        )
        Inspeccion.objects.filter(pk=ajena.pk).update(
            estado=Inspeccion.Estado.PUBLICADO,
            cliente=self.usuarios["Cliente otro"],
        )
        self.client.force_login(self.usuarios["Cliente"])
        respuesta = self.client.get(reverse("dashboard"))
        self.assertEqual(respuesta.status_code, 200)
        visibles = {
            inspeccion.pk
            for inspeccion in respuesta.context["inspecciones_nuevas_cvb0003"]
        } | {
            inspeccion.pk
            for inspeccion in respuesta.context["historial_completo_cvb0003"]
        }
        self.assertIn(propia.pk, visibles)
        self.assertNotIn(ajena.pk, visibles)

    def test_transiciones_respetan_asignacion(self):
        inspeccion = self.inspecciones[Inspeccion.Tipo.LIFE_SHAFT]
        request = self.factory.post("/", {"workflow_action": "enviar_supervisor"})
        request.user = self.usuarios["Inspector otro"]
        correcto, _mensaje = _aplicar_accion_flujo(
            request, inspeccion, "enviar_supervisor"
        )
        self.assertFalse(correcto)
        inspeccion.refresh_from_db()
        self.assertEqual(inspeccion.estado, Inspeccion.Estado.BORRADOR)

        request.user = self.usuarios["Inspector"]
        correcto, _mensaje = _aplicar_accion_flujo(
            request, inspeccion, "enviar_supervisor"
        )
        self.assertTrue(correcto)
        inspeccion.refresh_from_db()
        self.assertEqual(inspeccion.estado, Inspeccion.Estado.EN_REVISION)
