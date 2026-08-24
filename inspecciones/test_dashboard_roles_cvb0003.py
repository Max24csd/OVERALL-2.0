from datetime import date, timedelta

from django.contrib.auth.models import Group, User
from django.test import TestCase
from django.urls import reverse

from inspecciones.models import Faja, Inspeccion


class DashboardRolesCVB0003Tests(TestCase):
    @classmethod
    def setUpTestData(cls):
        cls.groups = {
            role: Group.objects.create(name=role)
            for role in ("Inspector", "Supervisor", "Analista", "Cliente")
        }
        cls.users = {}
        for role in cls.groups:
            for suffix in ("a", "b"):
                user = User.objects.create_user(f"{role.lower()}-{suffix}", password="test-123")
                user.groups.add(cls.groups[role])
                cls.users[(role, suffix)] = user
        cls.admin = User.objects.create_superuser("admin-dashboard", "admin@example.com", "test-123")
        cls.faja = Faja.objects.create(nombre="Faja CVB003", tag="CVB0003")
        cls.items = {}
        base = date(2026, 8, 12)
        states = (
            Inspeccion.Estado.BORRADOR,
            Inspeccion.Estado.EN_REVISION,
            Inspeccion.Estado.REVISADO,
            Inspeccion.Estado.APROBADO,
            Inspeccion.Estado.PUBLICADO,
        )
        for index, state in enumerate(states):
            suffix = "a" if index != 1 else "b"
            item = Inspeccion.objects.create(
                faja=cls.faja,
                tipo=Inspeccion.Tipo.values[index % 3],
                codigo_reporte=f"VISIBLE-{state}-{suffix}",
                fecha_programada=base + timedelta(days=index),
                fecha_inspeccion=base + timedelta(days=index),
                inspector=cls.users[("Inspector", suffix)],
                supervisor=cls.users[("Supervisor", suffix)],
                analista=cls.users[("Analista", suffix)],
                cliente=cls.users[("Cliente", suffix)],
                estado=state,
                creado_por=cls.admin,
            )
            item.refresh_from_db()
            cls.items[(state, suffix)] = item
        cls.published_b = Inspeccion.objects.create(
            faja=cls.faja,
            tipo=Inspeccion.Tipo.LIFE_SHAFT,
            codigo_reporte="HIDDEN-PUBLISHED-B",
            fecha_programada=base,
            fecha_inspeccion=base,
            inspector=cls.users[("Inspector", "b")],
            supervisor=cls.users[("Supervisor", "b")],
            analista=cls.users[("Analista", "b")],
            cliente=cls.users[("Cliente", "b")],
            estado=Inspeccion.Estado.PUBLICADO,
            creado_por=cls.admin,
        )
        cls.published_b.refresh_from_db()

    def _get(self, role, suffix="a"):
        user = self.admin if role == "Administrador" else self.users[(role, suffix)]
        self.client.force_login(user)
        return self.client.get(reverse("dashboard"))

    def test_administrador_ve_resumen_global_y_accesos_administrativos(self):
        response = self._get("Administrador")
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Panel de control global")
        self.assertContains(response, "Total inspecciones")
        self.assertContains(response, "Usuarios")
        self.assertContains(response, "Nueva parada · Próximamente")
        self.assertContains(response, self.published_b.codigo_reporte)

    def test_inspector_solo_recibe_sus_asignaciones(self):
        response = self._get("Inspector", "a")
        self.assertContains(response, "Mis inspecciones")
        self.assertContains(response, self.items[(Inspeccion.Estado.BORRADOR, "a")].codigo_reporte)
        self.assertNotContains(response, self.items[(Inspeccion.Estado.EN_REVISION, "b")].codigo_reporte)
        self.assertNotContains(response, "Usuarios")
        ids = {item.pk for item in response.context["inspecciones"]}
        self.assertTrue(ids)
        self.assertTrue(all(item.inspector_id == self.users[("Inspector", "a")].id for item in response.context["inspecciones"]))

    def test_supervisor_solo_recibe_sus_asignaciones(self):
        response_a = self._get("Supervisor", "a")
        self.assertContains(response_a, "Bandeja de revisión")
        self.assertNotContains(response_a, self.items[(Inspeccion.Estado.EN_REVISION, "b")].codigo_reporte)
        response_b = self._get("Supervisor", "b")
        self.assertContains(response_b, self.items[(Inspeccion.Estado.EN_REVISION, "b")].codigo_reporte)
        self.assertContains(response_b, "Pendientes de revisión")
        self.assertNotContains(response_b, "Usuarios")

    def test_analista_solo_recibe_revisados_aprobados_asignados(self):
        response = self._get("Analista", "a")
        self.assertContains(response, "Bandeja de aprobación")
        self.assertContains(response, self.items[(Inspeccion.Estado.REVISADO, "a")].codigo_reporte)
        self.assertContains(response, self.items[(Inspeccion.Estado.APROBADO, "a")].codigo_reporte)
        self.assertNotContains(response, self.items[(Inspeccion.Estado.EN_REVISION, "b")].codigo_reporte)
        self.assertContains(response, "Publicar al cliente")

    def test_cliente_solo_recibe_publicados_propios(self):
        response_a = self._get("Cliente", "a")
        self.assertContains(response_a, "Reportes finales")
        self.assertContains(response_a, self.items[(Inspeccion.Estado.PUBLICADO, "a")].codigo_reporte)
        self.assertNotContains(response_a, self.published_b.codigo_reporte)
        self.assertNotContains(response_a, self.items[(Inspeccion.Estado.BORRADOR, "a")].codigo_reporte)
        self.assertNotContains(response_a, "<dt>Inspector</dt>", html=True)
        self.assertNotContains(response_a, "Usuarios")

    def test_agrupacion_usa_fecha_inspeccion_como_fecha_tecnica(self):
        response = self._get("Administrador")
        groups = response.context["grupos_dashboard"]["historicas"] + response.context["grupos_dashboard"]["actuales"]
        inspection_dates = {
            item.pk: group["fecha"]
            for group in groups
            for item in group["inspecciones"]
        }
        for item in response.context["inspecciones"]:
            if item.pk in inspection_dates:
                self.assertEqual(inspection_dates[item.pk], item.fecha_inspeccion)

    def test_parada_chancado_agrupa_cvb001_cvb003_cvb004_sin_inventar_registros(self):
        faja_1 = Faja.objects.create(nombre="Faja CVB001", tag="CVB0001")
        faja_4 = Faja.objects.create(nombre="Faja CVB004", tag="CVB0004")
        fecha = date(2026, 8, 20)
        creadas = []
        for faja, tipo, codigo in (
            (faja_1, Inspeccion.Tipo.FAJA, "CHANCADO-CVB001-FAJA"),
            (self.faja, Inspeccion.Tipo.POLEAS, "CHANCADO-CVB003-POLEAS"),
            (faja_4, Inspeccion.Tipo.LIFE_SHAFT, "CHANCADO-CVB004-LIFE"),
        ):
            item = Inspeccion.objects.create(
                faja=faja,
                tipo=tipo,
                codigo_reporte=codigo,
                fecha_programada=fecha,
                fecha_inspeccion=fecha,
                inspector=self.users[("Inspector", "a")],
                supervisor=self.users[("Supervisor", "a")],
                analista=self.users[("Analista", "a")],
                cliente=self.users[("Cliente", "a")],
                creado_por=self.admin,
            )
            item.refresh_from_db()
            creadas.append(item)

        response = self._get("Administrador")
        actual = response.context["grupos_dashboard"]["actuales"]
        self.assertEqual(len(actual), 1)
        self.assertEqual(actual[0]["fecha"], fecha)
        self.assertEqual(
            [equipo["codigo"] for equipo in actual[0]["equipos"]],
            ["CVB001", "CVB003", "CVB004"],
        )
        self.assertEqual(
            sum(len(equipo["inspecciones"]) for equipo in actual[0]["equipos"]),
            len(creadas),
        )
        for item in creadas:
            self.assertContains(response, item.codigo_reporte)
