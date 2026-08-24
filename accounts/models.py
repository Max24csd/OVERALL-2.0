from django.conf import settings
from django.db import models


class PerfilUsuario(models.Model):
    class TipoVinculo(models.TextChoices):
        PERMANENTE = "PERMANENTE", "Permanente"
        CONTRATO = "CONTRATO", "Por contrato"
        INTERMITENTE = "INTERMITENTE", "Intermitente"

    usuario = models.OneToOneField(
        settings.AUTH_USER_MODEL,
        on_delete=models.CASCADE,
        related_name="perfil_sistema",
    )

    tipo_vinculo = models.CharField(
        max_length=20,
        choices=TipoVinculo.choices,
        default=TipoVinculo.CONTRATO,
    )

    fecha_inicio_contrato = models.DateField(
        null=True,
        blank=True,
    )

    fecha_fin_contrato = models.DateField(
        null=True,
        blank=True,
    )

    creado_en = models.DateTimeField(
        auto_now_add=True,
    )

    actualizado_en = models.DateTimeField(
        auto_now=True,
    )

    class Meta:
        verbose_name = "Perfil de usuario"
        verbose_name_plural = "Perfiles de usuario"

    def __str__(self):
        return f"{self.usuario.username} - {self.get_tipo_vinculo_display()}"

    def contrato_vigente(self, fecha=None):
        from django.utils import timezone

        fecha = fecha or timezone.localdate()

        if self.tipo_vinculo == self.TipoVinculo.PERMANENTE:
            return True

        if self.tipo_vinculo == self.TipoVinculo.INTERMITENTE:
            return True

        if self.fecha_inicio_contrato and fecha < self.fecha_inicio_contrato:
            return False

        if self.fecha_fin_contrato and fecha > self.fecha_fin_contrato:
            return False

        return True

class AccesoParada(models.Model):
    class Rol(models.TextChoices):
        INSPECTOR = "Inspector", "Inspector"
        SUPERVISOR = "Supervisor", "Supervisor"
        ANALISTA = "Analista", "Analista"
        CLIENTE = "Cliente", "Cliente"

    parada = models.ForeignKey(
        "inspecciones.Parada",
        on_delete=models.CASCADE,
        related_name="accesos_temporales",
    )

    usuario = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.CASCADE,
        related_name="accesos_paradas",
    )

    rol = models.CharField(
        max_length=20,
        choices=Rol.choices,
    )

    fecha_inicio = models.DateTimeField()

    fecha_fin = models.DateTimeField()

    activo = models.BooleanField(
        default=True,
    )

    creado_por = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.PROTECT,
        related_name="accesos_parada_creados",
    )

    creado_en = models.DateTimeField(
        auto_now_add=True,
    )

    actualizado_en = models.DateTimeField(
        auto_now=True,
    )

    class Meta:
        ordering = ["-creado_en"]
        verbose_name = "Acceso temporal a parada"
        verbose_name_plural = "Accesos temporales a paradas"

        constraints = [
            models.UniqueConstraint(
                fields=[
                    "parada",
                    "usuario",
                    "rol",
                ],
                name="acceso_unico_usuario_parada_rol",
            )
        ]

    def __str__(self):
        return (
            f"{self.usuario.username} - "
            f"{self.rol} - "
            f"{self.parada}"
        )

    def esta_vigente(self, momento=None):
        from django.utils import timezone

        momento = momento or timezone.now()

        return (
            self.activo
            and self.fecha_inicio <= momento <= self.fecha_fin
        )


class HistorialAsignacionParada(models.Model):
    parada = models.ForeignKey(
        "inspecciones.Parada",
        on_delete=models.CASCADE,
        related_name="historial_asignaciones",
    )

    rol = models.CharField(
        max_length=20,
        choices=AccesoParada.Rol.choices,
    )

    usuario_anterior = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.PROTECT,
        related_name="asignaciones_parada_reemplazadas",
        null=True,
        blank=True,
    )

    usuario_nuevo = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.PROTECT,
        related_name="asignaciones_parada_recibidas",
        null=True,
        blank=True,
    )

    motivo = models.TextField(
        blank=True,
        default="",
    )

    cambiado_por = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.PROTECT,
        related_name="cambios_asignacion_parada",
    )

    fecha = models.DateTimeField(
        auto_now_add=True,
    )

    class Meta:
        ordering = ["-fecha"]
        verbose_name = "Historial de asignación de parada"
        verbose_name_plural = "Historial de asignaciones de parada"

    def __str__(self):
        return (
            f"{self.parada} - "
            f"{self.rol} - "
            f"{self.usuario_anterior} → {self.usuario_nuevo}"
        )