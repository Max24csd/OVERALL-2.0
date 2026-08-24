from datetime import date

from django.contrib.auth.models import Group, User
from django.template.loader import render_to_string
from django.test import TestCase
from django.urls import reverse

from inspecciones.models import Faja, HistorialEstado, Inspeccion
from inspecciones.views import _contexto_workflow_ui_cvb0003, obtener_permisos_flujo


class WorkflowUICVB0003Tests(TestCase):
    @classmethod
    def setUpTestData(cls):
        cls.groups = {
            role: Group.objects.create(name=role)
            for role in ("Inspector", "Supervisor", "Analista", "Cliente")
        }
        cls.users = {}
        for role, group in cls.groups.items():
            user = User.objects.create_user(role.lower(), password="test-123")
            user.groups.add(group)
            cls.users[role] = user
        cls.admin = User.objects.create_superuser("admin-ui", "admin@example.com", "test-123")
        cls.faja = Faja.objects.create(nombre="Faja CVB003", tag="CVB0003")
        cls.items = {}
        for index, tipo in enumerate(Inspeccion.Tipo.values, 1):
            cls.items[tipo] = Inspeccion.objects.create(
                faja=cls.faja,
                tipo=tipo,
                codigo_reporte=f"2026081{index}-VTUT-CVB0003-{tipo}",
                fecha_programada=date(2026, 8, 10 + index),
                fecha_inspeccion=date(2026, 8, 10 + index),
                inspector=cls.users["Inspector"],
                supervisor=cls.users["Supervisor"],
                analista=cls.users["Analista"],
                cliente=cls.users["Cliente"],
                creado_por=cls.admin,
            )

    def _html(self, tipo, estado, role, section="actions"):
        item = self.items[tipo]
        Inspeccion.objects.filter(pk=item.pk).update(estado=estado)
        item.refresh_from_db()
        user = self.admin if role == "Administrador" else self.users[role]
        context = {"inspeccion": item, "workflow_section": section}
        context.update(obtener_permisos_flujo(user, item))
        context.update(_contexto_workflow_ui_cvb0003(item))
        return render_to_string("inspecciones/_workflow_controls_cvb0003.html", context)

    def test_matriz_visual_se_repite_en_los_tres_reportes(self):
        cases = (
            ("Inspector", Inspeccion.Estado.BORRADOR, ("Guardar borrador", "Enviar a revisión"), ("Dar visto bueno", "Publicar al cliente")),
            ("Inspector", Inspeccion.Estado.EN_REVISION, (), ("Guardar cambios", "Enviar a revisión", "Dar visto bueno")),
            ("Supervisor", Inspeccion.Estado.EN_REVISION, ("Guardar cambios", "Devolver al inspector", "Dar visto bueno"), ("Aprobar reporte", "Publicar al cliente")),
            ("Supervisor", Inspeccion.Estado.REVISADO, (), ("Guardar cambios", "Devolver al inspector", "Dar visto bueno")),
            ("Analista", Inspeccion.Estado.REVISADO, ("Guardar cambios", "Devolver al inspector", "Aprobar reporte"), ("Dar visto bueno", "Publicar al cliente")),
            ("Analista", Inspeccion.Estado.APROBADO, ("Publicar al cliente",), ("Guardar cambios", "Aprobar reporte")),
            ("Administrador", Inspeccion.Estado.BORRADOR, ("Guardar cambios", "Enviar a revisión"), ("Dar visto bueno", "Aprobar reporte", "Publicar al cliente")),
        )
        for tipo in Inspeccion.Tipo.values:
            for role, state, visible, hidden in cases:
                with self.subTest(tipo=tipo, role=role, state=state):
                    html = self._html(tipo, state, role)
                    for label in visible:
                        self.assertIn(label, html)
                    for label in hidden:
                        self.assertNotIn(label, html)

    def test_devolucion_del_inspector_muestra_ultima_observacion(self):
        for tipo, item in self.items.items():
            HistorialEstado.objects.create(
                inspeccion=item,
                estado_anterior=Inspeccion.Estado.EN_REVISION,
                estado_nuevo=Inspeccion.Estado.DEVUELTO,
                usuario=self.users["Supervisor"],
                rol="Supervisor",
                accion=HistorialEstado.Accion.DEVOLVER_SUPERVISOR,
                comentario=f"Corregir {tipo}",
            )
            html = self._html(tipo, Inspeccion.Estado.DEVUELTO, "Inspector", "header")
            self.assertIn("REPORTE DEVUELTO", html)
            self.assertIn(f"Corregir {tipo}", html)
            self.assertNotIn("Historial de revisión", html)

    def test_historial_interno_no_se_renderiza_para_cliente(self):
        for tipo in Inspeccion.Tipo.values:
            html = self._html(tipo, Inspeccion.Estado.PUBLICADO, "Cliente", "header")
            self.assertNotIn("Historial de revisión", html)
            actions = self._html(tipo, Inspeccion.Estado.PUBLICADO, "Cliente")
            self.assertNotIn("Vista previa", actions)
            self.assertNotIn('name="workflow_action"', actions)

    def test_partial_incluido_dos_veces_en_cada_formulario(self):
        paths = (
            "templates/inspecciones/formulario_faja_cvb0003.html",
            "templates/inspecciones/formulario_poleas_cvb0003.html",
            "templates/inspecciones/formulario_life_shaft_cvb0003.html",
        )
        from pathlib import Path
        from django.conf import settings

        for relative in paths:
            source = (Path(settings.BASE_DIR) / relative).read_text(encoding="utf-8")
            self.assertEqual(source.count("_workflow_controls_cvb0003.html"), 2)

    def test_confirmacion_modal_y_prevencion_doble_envio(self):
        html = self._html(Inspeccion.Tipo.POLEAS, Inspeccion.Estado.EN_REVISION, "Supervisor")
        self.assertIn("data-workflow-confirm", html)
        self.assertIn("Motivo de devolución *", html)
        self.assertIn("data-return-comment", html)
        self.assertIn("Procesando...", html)
        self.assertIn("let submitted=false", html)

    def test_cliente_publicado_va_directamente_al_reporte_final(self):
        route_names = {
            Inspeccion.Tipo.FAJA: ("formulario_faja", "reporte_faja"),
            Inspeccion.Tipo.POLEAS: ("formulario_poleas", "reporte_poleas"),
            Inspeccion.Tipo.LIFE_SHAFT: ("formulario_life_shaft", "reporte_life_shaft"),
        }
        self.client.force_login(self.users["Cliente"])
        for tipo, (form_name, report_name) in route_names.items():
            item = self.items[tipo]
            Inspeccion.objects.filter(pk=item.pk).update(estado=Inspeccion.Estado.PUBLICADO)
            response = self.client.get(reverse(form_name, args=[item.pk]))
            self.assertRedirects(
                response,
                reverse(report_name, args=[item.pk]),
                fetch_redirect_response=False,
            )
