from datetime import date

from django.contrib.auth.models import Group, User
from django.test import RequestFactory, TestCase

from inspecciones.models import Faja, HistorialEstado, Inspeccion
from inspecciones.views import _aplicar_accion_flujo


class TrazabilidadFlujoCVB0003Tests(TestCase):
    @classmethod
    def setUpTestData(cls):
        cls.users = {}
        for role in ("Inspector", "Supervisor", "Analista"):
            group = Group.objects.create(name=role)
            user = User.objects.create_user(role.lower(), password="segura-123")
            user.groups.add(group)
            cls.users[role] = user
        cls.admin = User.objects.create_superuser(
            "admin-flow", "admin@example.com", "segura-123"
        )
        cls.faja = Faja.objects.create(nombre="Faja 3", tag="CVB0003")

    def setUp(self):
        self.inspection = Inspeccion.objects.create(
            faja=self.faja,
            tipo=Inspeccion.Tipo.POLEAS,
            codigo_reporte="20260920-VTUT-CVB0003-POLEAS",
            fecha_programada=date(2026, 9, 20),
            fecha_inspeccion=date(2026, 9, 20),
            inspector=self.users["Inspector"],
            supervisor=self.users["Supervisor"],
            analista=self.users["Analista"],
            creado_por=self.admin,
        )
        self.factory = RequestFactory()

    def transition(self, user, action, comment=""):
        request = self.factory.post("/", {"workflow_action": action})
        request.user = user
        return _aplicar_accion_flujo(request, self.inspection, action, comment)

    def assert_event(self, previous, new, action, user, role, comment=""):
        event = HistorialEstado.objects.get(inspeccion=self.inspection)
        self.assertEqual(event.estado_anterior, previous)
        self.assertEqual(event.estado_nuevo, new)
        self.assertEqual(event.accion, action)
        self.assertEqual(event.usuario, user)
        self.assertEqual(event.rol, role)
        self.assertEqual(event.comentario, comment)
        self.assertIsNotNone(event.fecha)

    def test_borrador_a_en_revision(self):
        ok, _ = self.transition(self.users["Inspector"], "enviar_supervisor")
        self.assertTrue(ok)
        self.inspection.refresh_from_db()
        self.assertEqual(self.inspection.estado, Inspeccion.Estado.EN_REVISION)
        self.assert_event(
            "BORRADOR", "EN_REVISION", "ENVIAR_A_REVISION",
            self.users["Inspector"], "Inspector",
        )

    def test_devolucion_supervisor_exige_y_conserva_comentario(self):
        Inspeccion.objects.filter(pk=self.inspection.pk).update(estado="EN_REVISION")
        self.inspection.refresh_from_db()
        ok, _ = self.transition(self.users["Supervisor"], "devolver_supervisor")
        self.assertFalse(ok)
        self.assertEqual(HistorialEstado.objects.count(), 0)
        comment = "Verificar nuevamente Polea 03."
        ok, _ = self.transition(
            self.users["Supervisor"], "devolver_supervisor", comment
        )
        self.assertTrue(ok)
        self.assert_event(
            "EN_REVISION", "DEVUELTO", "DEVOLVER_SUPERVISOR",
            self.users["Supervisor"], "Supervisor", comment,
        )

    def test_aprobaciones_y_publicacion_registran_responsable(self):
        cases = (
            ("EN_REVISION", self.users["Supervisor"], "aprobar_supervisor", "REVISADO", "APROBAR_SUPERVISOR", "Supervisor"),
            ("REVISADO", self.users["Analista"], "aprobar_analista", "APROBADO", "APROBAR_ANALISTA", "Analista"),
            ("APROBADO", self.users["Analista"], "publicar", "PUBLICADO", "PUBLICAR", "Analista"),
        )
        for previous, user, action, new, event_action, role in cases:
            HistorialEstado.objects.all().delete()
            Inspeccion.objects.filter(pk=self.inspection.pk).update(estado=previous)
            self.inspection.refresh_from_db()
            ok, _ = self.transition(user, action)
            self.assertTrue(ok)
            self.assert_event(previous, new, event_action, user, role)

    def test_no_autorizada_no_cambia_ni_registra(self):
        ok, _ = self.transition(self.users["Inspector"], "aprobar_supervisor")
        self.assertFalse(ok)
        self.inspection.refresh_from_db()
        self.assertEqual(self.inspection.estado, "BORRADOR")
        self.assertEqual(HistorialEstado.objects.count(), 0)

    def test_transicion_invalida_no_cambia_ni_registra(self):
        ok, _ = self.transition(self.users["Inspector"], "enviar_supervisor")
        self.assertTrue(ok)
        HistorialEstado.objects.all().delete()
        ok, _ = self.transition(self.users["Inspector"], "enviar_supervisor")
        self.assertFalse(ok)
        self.inspection.refresh_from_db()
        self.assertEqual(self.inspection.estado, "EN_REVISION")
        self.assertEqual(HistorialEstado.objects.count(), 0)

    def test_flujo_completo_crea_un_evento_por_transicion(self):
        flow = (
            (self.users["Inspector"], "enviar_supervisor", ""),
            (self.users["Supervisor"], "devolver_supervisor", "Corregir fotografía."),
            (self.users["Inspector"], "enviar_supervisor", "Fotografía corregida."),
            (self.users["Supervisor"], "aprobar_supervisor", ""),
            (self.users["Analista"], "aprobar_analista", ""),
            (self.users["Analista"], "publicar", ""),
        )
        for user, action, comment in flow:
            ok, message = self.transition(user, action, comment)
            self.assertTrue(ok, message)
        self.inspection.refresh_from_db()
        self.assertEqual(self.inspection.estado, "PUBLICADO")
        self.assertEqual(self.inspection.historial.count(), len(flow))
        self.example_history = list(
            self.inspection.historial.order_by("fecha").values_list(
                "usuario__username", "rol", "estado_anterior", "estado_nuevo",
                "accion", "comentario",
            )
        )

    def test_administrador_registra_rol_administrador(self):
        ok, _ = self.transition(self.admin, "enviar_supervisor")
        self.assertTrue(ok)
        self.assertEqual(HistorialEstado.objects.get().rol, "Administrador")
