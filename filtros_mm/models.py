from django.conf import settings
from django.db import models


class ReporteFiltro(models.Model):
    """
    Reporte genérico para la parada de Filtros/MMG.

    Los campos variables del formulario se guardan en JSON para no crear
    un modelo diferente por cada familia (Carrilería, Tuberías, Skirting,
    Faja/Poleas, etc.).

    La presentación final se genera desde una plantilla Excel maestra.
    """

    class Estado(models.TextChoices):
        BORRADOR = "BORRADOR", "Borrador"
        EN_REVISION = "EN_REVISION", "En revisión"
        DEVUELTO = "DEVUELTO", "Devuelto"
        REVISADO = "REVISADO", "Revisado"
        APROBADO = "APROBADO", "Aprobado"
        PUBLICADO = "PUBLICADO", "Publicado"

    class Condicion(models.TextChoices):
        NO_MEDIDO = "NO_MEDIDO", "No medido"
        NORMAL = "NORMAL", "Normal"
        PRECAUCION = "PRECAUCION", "Precaución"
        TOLERABLE = "TOLERABLE", "Tolerable"
        CRITICO = "CRITICO", "Crítico"

    # Parada existente del sistema. Es nullable para permitir pruebas
    # antes de vincular el reporte a una parada real.
    parada = models.ForeignKey(
        "inspecciones.Parada",
        on_delete=models.PROTECT,
        related_name="reportes_filtros",
        null=True,
        blank=True,
    )

    # Identidad tomada del OVERVIEW / catalogo.py.
    codigo_catalogo = models.CharField(
        max_length=30,
        db_index=True,
        help_text="Ej.: FA0201, FA0302, CV2401.",
    )
    area = models.CharField(max_length=30, default="420")
    tag = models.CharField(max_length=80, db_index=True)
    componente = models.CharField(max_length=150)
    familia = models.CharField(max_length=50, db_index=True)
    tecnica = models.CharField(max_length=50, blank=True)

    # Código final del reporte. Puede cambiar por fecha/parada sin alterar
    # el código base del catálogo.
    codigo_reporte = models.CharField(
        max_length=120,
        blank=True,
        db_index=True,
    )

    fecha_programada = models.DateField(null=True, blank=True)
    fecha_inspeccion = models.DateField(null=True, blank=True)
    fecha_reporte = models.DateField(null=True, blank=True)

    condicion_general = models.CharField(
        max_length=20,
        choices=Condicion.choices,
        default=Condicion.NO_MEDIDO,
    )

    estado = models.CharField(
        max_length=20,
        choices=Estado.choices,
        default=Estado.BORRADOR,
        db_index=True,
    )

    # Responsables principales. Más adelante los permisos pueden apoyarse
    # también en AccesoParada, igual que el módulo actual.
    inspector = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.PROTECT,
        related_name="reportes_filtros_inspector",
        null=True,
        blank=True,
    )
    supervisor = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.PROTECT,
        related_name="reportes_filtros_supervisor",
        null=True,
        blank=True,
    )
    analista = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.PROTECT,
        related_name="reportes_filtros_analista",
        null=True,
        blank=True,
    )
    cliente = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.PROTECT,
        related_name="reportes_filtros_cliente",
        null=True,
        blank=True,
    )

    # Datos variables del formulario.
    #
    # Ejemplo:
    # {
    #   "circunstancias": "...",
    #   "antecedentes": "...",
    #   "observaciones": "...",
    #   "recomendaciones": "...",
    #   "campos": {...}
    # }
    datos = models.JSONField(default=dict, blank=True)

    # Mediciones variables.
    #
    # Ejemplo:
    # [
    #   {
    #       "seccion": "Carril 01",
    #       "punto": "A",
    #       "valor": "12.5",
    #       "unidad": "mm",
    #       "condicion": "NORMAL",
    #       "observacion": ""
    #   }
    # ]
    mediciones = models.JSONField(default=list, blank=True)

    comentario_revision = models.TextField(blank=True)
    motivo_devolucion = models.TextField(blank=True)

    fecha_envio_revision = models.DateTimeField(null=True, blank=True)
    fecha_revision_supervisor = models.DateTimeField(null=True, blank=True)
    fecha_aprobacion_analista = models.DateTimeField(null=True, blank=True)
    fecha_publicacion = models.DateTimeField(null=True, blank=True)

    creado_por = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.PROTECT,
        related_name="reportes_filtros_creados",
        null=True,
        blank=True,
    )

    creado_en = models.DateTimeField(auto_now_add=True)
    actualizado_en = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ["-fecha_programada", "tag", "codigo_catalogo", "id"]
        constraints = [
            models.UniqueConstraint(
                fields=["parada", "codigo_catalogo"],
                name="uniq_reporte_filtro_parada_catalogo",
            ),
        ]

    def __str__(self):
        return (
            f"{self.codigo_catalogo} | "
            f"{self.tag} | "
            f"{self.componente}"
        )


class FotoReporteFiltro(models.Model):
    """
    Fotografías capturadas por el inspector.

    'seccion' permite ubicarlas después en el lugar correcto del Excel
    maestro sin depender del HTML.
    """

    reporte = models.ForeignKey(
        ReporteFiltro,
        on_delete=models.CASCADE,
        related_name="fotografias",
    )
    seccion = models.CharField(max_length=100, blank=True)
    titulo = models.CharField(max_length=200, blank=True)
    imagen = models.ImageField(
        upload_to="filtros_mm/%Y/%m/",
    )
    orden = models.PositiveIntegerField(default=0)
    creado_en = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["seccion", "orden", "id"]

    def __str__(self):
        return f"Foto {self.reporte_id} - {self.seccion or 'General'}"


class HistorialEstadoFiltro(models.Model):
    """
    Trazabilidad del flujo:
    Inspector -> Supervisor -> Analista -> Cliente.
    """

    reporte = models.ForeignKey(
        ReporteFiltro,
        on_delete=models.CASCADE,
        related_name="historial_estados",
    )

    estado_anterior = models.CharField(
        max_length=20,
        choices=ReporteFiltro.Estado.choices,
        blank=True,
    )
    estado_nuevo = models.CharField(
        max_length=20,
        choices=ReporteFiltro.Estado.choices,
    )

    accion = models.CharField(max_length=50)
    comentario = models.TextField(blank=True)

    usuario = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.PROTECT,
        related_name="historial_estados_filtros",
    )

    creado_en = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["-creado_en", "-id"]

    def __str__(self):
        return (
            f"{self.reporte_id}: "
            f"{self.estado_anterior or '-'} -> {self.estado_nuevo}"
        )
