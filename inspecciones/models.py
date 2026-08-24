from __future__ import annotations

from django.conf import settings
from django.core.validators import MinValueValidator
from django.db import models

from .reportes.cvb0003.code_utils import (
    es_tag_cvb0003,
    generar_codigo_cvb0003,
    sincronizar_codigo_cvb0003,
)


class Faja(models.Model):
    class Estado(models.TextChoices):
        ACTIVA = "ACTIVA", "Activa"
        INACTIVA = "INACTIVA", "Inactiva"

    nombre = models.CharField(
        max_length=120,
        verbose_name="Nombre",
    )

    tag = models.CharField(
        max_length=30,
        unique=True,
        verbose_name="TAG",
    )

    proceso = models.CharField(
        max_length=80,
        default="Chancado",
    )

    descripcion = models.TextField(
        blank=True,
    )

    estado = models.CharField(
        max_length=20,
        choices=Estado.choices,
        default=Estado.ACTIVA,
    )

    creada_en = models.DateTimeField(
        auto_now_add=True,
    )

    actualizada_en = models.DateTimeField(
        auto_now=True,
    )

    class Meta:
        ordering = ["tag"]
        verbose_name = "Faja"
        verbose_name_plural = "Fajas"

class Parada(models.Model):
    class Estado(models.TextChoices):
        PROGRAMADA = "PROGRAMADA", "Programada"
        EN_CURSO = "EN_CURSO", "En curso"
        FINALIZADA = "FINALIZADA", "Finalizada"
        CANCELADA = "CANCELADA", "Cancelada"

    nombre = models.CharField(
        max_length=150,
        blank=True,
        default="",
    )

    planta = models.CharField(
        max_length=100,
        default="Chancado",
    )

    fecha_inicio = models.DateField()

    fecha_fin = models.DateField(
        null=True,
        blank=True,
    )

    estado = models.CharField(
        max_length=20,
        choices=Estado.choices,
        default=Estado.PROGRAMADA,
    )

    observaciones = models.TextField(
        blank=True,
        default="",
    )

    creado_por = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.PROTECT,
        related_name="paradas_creadas",
    )

    creada_en = models.DateTimeField(
        auto_now_add=True,
    )

    actualizada_en = models.DateTimeField(
        auto_now=True,
    )

    class Meta:
        ordering = ["-fecha_inicio", "-id"]
        verbose_name = "Parada"
        verbose_name_plural = "Paradas"

    def __str__(self):
        return self.nombre or f"{self.planta} - {self.fecha_inicio:%d/%m/%Y}"

class Inspeccion(models.Model):
    class Tipo(models.TextChoices):
        FAJA = "FAJA", "Faja / Top Cover"
        POLEAS = "POLEAS", "Poleas"
        LIFE_SHAFT = "LIFE_SHAFT", "Life Shaft"

    class Estado(models.TextChoices):
        BORRADOR = "BORRADOR", "Borrador"
        EN_REVISION = "EN_REVISION", "En revisión"
        DEVUELTO = "DEVUELTO", "Devuelto al inspector"
        REVISADO = "REVISADO", "Revisado por supervisor"
        APROBADO = "APROBADO", "Aprobado por analista"
        PUBLICADO = "PUBLICADO", "Publicado para cliente"

    class Condicion(models.TextChoices):
        NORMAL = "NORMAL", "Normal"
        TOLERABLE = "TOLERABLE", "Tolerable"
        PRECAUCION = "PRECAUCION", "Precaución"
        CRITICO = "CRITICO", "Crítico"
        NO_MEDIDO = "NO_MEDIDO", "No medido"

    parada = models.ForeignKey(
        Parada,
        on_delete=models.PROTECT,
        related_name="inspecciones",
        null=True,
        blank=True,
    )

    faja = models.ForeignKey(
        Faja,
        on_delete=models.PROTECT,
        related_name="inspecciones",
    )

    tipo = models.CharField(
        max_length=20,
        choices=Tipo.choices,
    )

    codigo_reporte = models.CharField(
        max_length=150,
        unique=True,
    )

    fecha_programada = models.DateField()

    fecha_inspeccion = models.DateField(
        null=True,
        blank=True,
    )

    fecha_reporte = models.DateField(
        null=True,
        blank=True,
    )

    inspector = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.PROTECT,
        related_name="inspecciones_asignadas",
    )

    supervisor = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.PROTECT,
        related_name="inspecciones_supervisadas",
    )

    analista = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.PROTECT,
        related_name="inspecciones_analizadas",
    )

    cliente = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.PROTECT,
        related_name="inspecciones_cliente",
        null=True,
        blank=True,
    )

    estado = models.CharField(
        max_length=30,
        choices=Estado.choices,
        default=Estado.BORRADOR,
    )

    condicion_general = models.CharField(
        max_length=20,
        choices=Condicion.choices,
        default=Condicion.NORMAL,
    )

    planta = models.CharField(
        max_length=100,
        default="Chancado",
    )

    proceso = models.CharField(
        max_length=120,
        default="Transporte de mineral",
    )

    etapa = models.CharField(
        max_length=100,
        default="Operaciones",
    )

    condicion_equipo = models.CharField(
        max_length=100,
        default="En uso",
    )

    inspector_campo_nombre = models.CharField(
        max_length=150,
        blank=True,
        default="",
        verbose_name="Inspector de campo",
    )

    supervisor_campo_nombre = models.CharField(
        max_length=150,
        blank=True,
        default="",
        verbose_name="Supervisor de campo",
    )

    analista_elabora_nombre = models.CharField(
        max_length=150,
        blank=True,
        default="",
        verbose_name="Analista que elabora",
    )

    analista_valida_nombre = models.CharField(
        max_length=150,
        blank=True,
        default="",
        verbose_name="Analista que valida",
    )

    circunstancias = models.TextField(blank=True)
    antecedentes = models.TextField(blank=True)
    observaciones = models.TextField(blank=True)
    recomendaciones = models.TextField(blank=True)
    comentarios_revision = models.TextField(blank=True)

    creado_por = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.PROTECT,
        related_name="inspecciones_creadas",
    )

    creada_en = models.DateTimeField(auto_now_add=True)
    actualizada_en = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ["-fecha_programada", "-id"]
        verbose_name = "Inspección"
        verbose_name_plural = "Inspecciones"

        permissions = [
            ("enviar_revision", "Puede enviar inspecciones a revisión"),
            ("devolver_inspector", "Puede devolver inspecciones al inspector"),
            ("revisar_inspeccion", "Puede revisar inspecciones"),
            ("aprobar_inspeccion", "Puede aprobar inspecciones"),
            ("publicar_inspeccion", "Puede publicar inspecciones"),
        ]

        constraints = [
            models.UniqueConstraint(
                fields=["faja", "tipo", "fecha_programada"],
                name="inspeccion_unica_faja_tipo_fecha",
            )
        ]

    def generar_codigo_reporte(self) -> str:
        """
        Genera el código usando la FECHA DE INSPECCIÓN.

        Ejemplo:
        04/08/2026 + CVB0004 + LIFE_SHAFT
        -> 20260804-VTUT-CVB0004-LIFE-SHAFT
        """
        if not self.fecha_inspeccion:
            return self.codigo_reporte

        fecha_codigo = self.fecha_inspeccion.strftime("%Y%m%d")
        tag = self.faja.tag

        if es_tag_cvb0003(tag):
            return generar_codigo_cvb0003(
                self.fecha_inspeccion,
                self.tipo,
            )

        if self.tipo == self.Tipo.LIFE_SHAFT:
            sufijo = "LIFE-SHAFT"
        elif self.tipo == self.Tipo.POLEAS:
            sufijo = "POLEAS"
        elif self.tipo == self.Tipo.FAJA:
            sufijo = "FAJA"
        else:
            return self.codigo_reporte

        return f"{fecha_codigo}-VTUT-{tag}-{sufijo}"

    def save(self, *args, **kwargs) -> None:
        """
        Cada vez que se guarda la inspección, si existe fecha_inspeccion,
        sincroniza automáticamente codigo_reporte con esa fecha.
        """
        codigo_gestionado_cvb0003 = sincronizar_codigo_cvb0003(self)

        if self.fecha_inspeccion and not codigo_gestionado_cvb0003:
            self.codigo_reporte = self.generar_codigo_reporte()

        # Si save(update_fields=[...]) fue llamado, asegurar que también
        # se persista codigo_reporte cuando cambia automáticamente.
        update_fields = kwargs.get("update_fields")
        if update_fields is not None and self.fecha_inspeccion:
            kwargs["update_fields"] = set(update_fields) | {"codigo_reporte"}

        super().save(*args, **kwargs)

    def __str__(self) -> str:
        return (
            f"{self.codigo_reporte} - "
            f"{self.faja.tag} - "
            f"{self.get_tipo_display()}"
        )


class Medicion(models.Model):
    inspeccion = models.ForeignKey(
        Inspeccion,
        on_delete=models.CASCADE,
        related_name="mediciones",
    )

    seccion = models.CharField(
        max_length=120,
    )

    punto = models.CharField(
        max_length=50,
        blank=True,
    )

    bastidor = models.CharField(
        max_length=50,
        blank=True,
    )

    lado = models.CharField(
        max_length=50,
        blank=True,
    )

    posicion = models.CharField(
        max_length=80,
        blank=True,
    )

    componente = models.CharField(
        max_length=120,
        blank=True,
    )

    espesor_nominal = models.DecimalField(
        max_digits=8,
        decimal_places=2,
        null=True,
        blank=True,
        validators=[MinValueValidator(0)],
    )

    a = models.DecimalField(
        max_digits=8,
        decimal_places=2,
        null=True,
        blank=True,
    )

    b = models.DecimalField(
        max_digits=8,
        decimal_places=2,
        null=True,
        blank=True,
    )

    c = models.DecimalField(
        max_digits=8,
        decimal_places=2,
        null=True,
        blank=True,
    )

    d = models.DecimalField(
        max_digits=8,
        decimal_places=2,
        null=True,
        blank=True,
    )

    e = models.DecimalField(
        max_digits=8,
        decimal_places=2,
        null=True,
        blank=True,
    )

    f = models.DecimalField(
        max_digits=8,
        decimal_places=2,
        null=True,
        blank=True,
    )

    g = models.DecimalField(
        max_digits=8,
        decimal_places=2,
        null=True,
        blank=True,
    )

    minimo = models.DecimalField(
        max_digits=8,
        decimal_places=2,
        null=True,
        blank=True,
    )

    promedio = models.DecimalField(
        max_digits=8,
        decimal_places=2,
        null=True,
        blank=True,
    )

    desgaste = models.DecimalField(
        max_digits=8,
        decimal_places=2,
        null=True,
        blank=True,
    )

    porcentaje_desgaste = models.DecimalField(
        max_digits=8,
        decimal_places=2,
        null=True,
        blank=True,
    )

    porcentaje_residual = models.DecimalField(
        max_digits=8,
        decimal_places=2,
        null=True,
        blank=True,
    )

    condicion = models.CharField(
        max_length=20,
        choices=Inspeccion.Condicion.choices,
        default=Inspeccion.Condicion.NORMAL,
    )

    observacion = models.TextField(
        blank=True,
    )

    orden = models.PositiveIntegerField(
        default=0,
    )

    es_resumen = models.BooleanField(
        default=False,
    )

    creada_en = models.DateTimeField(
        auto_now_add=True,
    )

    actualizada_en = models.DateTimeField(
        auto_now=True,
    )

    class Meta:
        ordering = ["orden", "id"]
        verbose_name = "Medición"
        verbose_name_plural = "Mediciones"

    def __str__(self) -> str:
        return (
            f"{self.inspeccion.faja.tag} - "
            f"{self.seccion} - "
            f"{self.punto or self.posicion}"
        )

    def valores_disponibles(self) -> list:
        return [
            valor
            for valor in [
                self.a,
                self.b,
                self.c,
                self.d,
                self.e,
                self.f,
                self.g,
            ]
            if valor is not None
        ]

    def calcular_resultados(self) -> None:
        valores = self.valores_disponibles()

        if not valores:
            self.minimo = None
            self.promedio = None
            self.desgaste = None
            self.porcentaje_desgaste = None
            self.porcentaje_residual = None
            return

        self.minimo = min(valores)

        self.promedio = round(
            sum(valores) / len(valores),
            2,
        )

        if self.espesor_nominal and self.espesor_nominal > 0:
            self.desgaste = round(
                self.espesor_nominal - self.minimo,
                2,
            )

            self.porcentaje_desgaste = round(
                self.desgaste / self.espesor_nominal * 100,
                2,
            )

            self.porcentaje_residual = round(
                self.minimo / self.espesor_nominal * 100,
                2,
            )
        else:
            self.desgaste = None
            self.porcentaje_desgaste = None
            self.porcentaje_residual = None

    def save(self, *args, **kwargs) -> None:
        self.calcular_resultados()
        super().save(*args, **kwargs)


def ruta_foto_inspeccion(
    instancia: "FotoInspeccion",
    nombre_archivo: str,
) -> str:
    return (
        f"inspecciones/"
        f"{instancia.inspeccion.id}/"
        f"{nombre_archivo}"
    )


class FotoInspeccion(models.Model):
    inspeccion = models.ForeignKey(
        Inspeccion,
        on_delete=models.CASCADE,
        related_name="fotografias",
    )

    imagen = models.ImageField(
        upload_to=ruta_foto_inspeccion,
    )

    codigo_dano = models.CharField(
        max_length=50,
        blank=True,
    )

    descripcion = models.TextField(
        blank=True,
    )

    subida_por = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.PROTECT,
        related_name="fotografias_subidas",
    )

    creada_en = models.DateTimeField(
        auto_now_add=True,
    )

    class Meta:
        ordering = ["creada_en"]
        verbose_name = "Fotografía"
        verbose_name_plural = "Fotografías"

    def __str__(self) -> str:
        return (
            f"Foto {self.id} - "
            f"{self.inspeccion.codigo_reporte}"
        )


class HistorialEstado(models.Model):
    class Accion(models.TextChoices):
        ENVIAR_A_REVISION = "ENVIAR_A_REVISION", "Enviar a revisión"
        DEVOLVER_SUPERVISOR = "DEVOLVER_SUPERVISOR", "Devolver supervisor"
        APROBAR_SUPERVISOR = "APROBAR_SUPERVISOR", "Aprobar supervisor"
        DEVOLVER_ANALISTA = "DEVOLVER_ANALISTA", "Devolver analista"
        APROBAR_ANALISTA = "APROBAR_ANALISTA", "Aprobar analista"
        PUBLICAR = "PUBLICAR", "Publicar"

    inspeccion = models.ForeignKey(
        Inspeccion,
        on_delete=models.CASCADE,
        related_name="historial",
    )

    estado_anterior = models.CharField(
        max_length=30,
        blank=True,
    )

    estado_nuevo = models.CharField(
        max_length=30,
        choices=Inspeccion.Estado.choices,
    )

    comentario = models.TextField(
        blank=True,
    )

    rol = models.CharField(
        max_length=30,
        blank=True,
        default="",
    )

    accion = models.CharField(
        max_length=40,
        choices=Accion.choices,
        blank=True,
        default="",
    )

    usuario = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.PROTECT,
        related_name="cambios_estado",
    )

    fecha = models.DateTimeField(
        auto_now_add=True,
    )

    class Meta:
        ordering = ["-fecha"]
        verbose_name = "Historial de estado"
        verbose_name_plural = "Historial de estados"

    def __str__(self) -> str:
        return (
            f"{self.inspeccion.codigo_reporte}: "
            f"{self.estado_anterior} → {self.estado_nuevo}"
        )

class TipoMedicionComponente(models.TextChoices):
    NORMAL = "NORMAL", "Normal"
    CAMPANA = "CAMPANA", "Campaña (Inicio + Fin)"


class FaseCampana(models.TextChoices):
    INICIO = "INICIO", "Inicio de campaña"
    FIN = "FIN", "Fin de campaña"


class PoleaInspeccion(models.Model):
    inspeccion = models.ForeignKey(Inspeccion, on_delete=models.CASCADE, related_name="poleas_inspeccionadas")
    numero = models.PositiveSmallIntegerField()
    nombre = models.CharField(max_length=100)
    tag = models.CharField(max_length=60, blank=True)
    ubicacion = models.CharField(max_length=120, blank=True)
    componente = models.CharField(max_length=120, default="Lagging de polea")
    material = models.CharField(max_length=100, blank=True)
    espesor_nominal = models.DecimalField(max_digits=8, decimal_places=2, null=True, blank=True)
    condicion = models.CharField(max_length=20, choices=Inspeccion.Condicion.choices, default=Inspeccion.Condicion.NO_MEDIDO)
    tipo_medicion = models.CharField(
        max_length=10,
        choices=TipoMedicionComponente.choices,
        default=TipoMedicionComponente.NORMAL,
    )
    observacion_visual = models.TextField(blank=True)
    observacion_medicion = models.TextField(blank=True)
    recomendaciones = models.TextField(blank=True)
    marca_equipo = models.CharField(max_length=80, default="OLYMPUS")
    modelo_equipo = models.CharField(max_length=80, default="EPOCH 6LT")
    frecuencia_mhz = models.CharField(max_length=30, default="5 MHz")
    rango_mm = models.CharField(max_length=30, blank=True)
    metodo_empleado = models.CharField(max_length=80, default="Pulso - eco")
    componente_calibracion = models.CharField(max_length=100, blank=True)
    acoplante = models.CharField(max_length=60, default="Echo gel")
    rectificacion = models.CharField(max_length=50, default="Full")
    velocidad_ms = models.CharField(max_length=30, default="6079")
    retardo_us = models.CharField(max_length=30, default="1.53")
    tipo_scan = models.CharField(max_length=50, default="A Scan")
    orden = models.PositiveIntegerField(default=0)
    creada_en = models.DateTimeField(auto_now_add=True)
    actualizada_en = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ["orden", "numero"]
        constraints = [models.UniqueConstraint(fields=["inspeccion", "numero"], name="polea_unica_por_inspeccion")]

    def __str__(self):
        return f"{self.inspeccion.faja.tag} - {self.nombre}"


class MedicionPolea(models.Model):
    polea = models.ForeignKey(PoleaInspeccion, on_delete=models.CASCADE, related_name="mediciones")
    punto = models.PositiveSmallIntegerField()
    posicion = models.CharField(max_length=100, blank=True)
    a = models.DecimalField(max_digits=8, decimal_places=2, null=True, blank=True)
    b = models.DecimalField(max_digits=8, decimal_places=2, null=True, blank=True)
    c = models.DecimalField(max_digits=8, decimal_places=2, null=True, blank=True)
    d = models.DecimalField(max_digits=8, decimal_places=2, null=True, blank=True)
    e = models.DecimalField(max_digits=8, decimal_places=2, null=True, blank=True)
    f = models.DecimalField(max_digits=8, decimal_places=2, null=True, blank=True)
    g = models.DecimalField(max_digits=8, decimal_places=2, null=True, blank=True)
    minimo = models.DecimalField(max_digits=8, decimal_places=2, null=True, blank=True)
    promedio = models.DecimalField(max_digits=8, decimal_places=2, null=True, blank=True)
    observacion = models.TextField(blank=True)
    orden = models.PositiveIntegerField(default=0)

    class Meta:
        ordering = ["orden", "punto"]
        constraints = [models.UniqueConstraint(fields=["polea", "punto"], name="punto_unico_por_polea")]

    def valores_disponibles(self):
        return [v for v in [self.a, self.b, self.c, self.d, self.e, self.f, self.g] if v is not None]

    def save(self, *args, **kwargs):
        valores = self.valores_disponibles()
        if valores:
            self.minimo = min(valores)
            self.promedio = round(sum(valores) / len(valores), 2)
        else:
            self.minimo = None
            self.promedio = None
        super().save(*args, **kwargs)


class MedicionPoleaCampana(models.Model):
    polea = models.ForeignKey(
        PoleaInspeccion,
        on_delete=models.CASCADE,
        related_name="mediciones_campana",
    )
    fase = models.CharField(max_length=10, choices=FaseCampana.choices)
    punto = models.PositiveSmallIntegerField()
    posicion = models.CharField(max_length=100, blank=True)
    a = models.DecimalField(max_digits=8, decimal_places=2, null=True, blank=True)
    b = models.DecimalField(max_digits=8, decimal_places=2, null=True, blank=True)
    c = models.DecimalField(max_digits=8, decimal_places=2, null=True, blank=True)
    d = models.DecimalField(max_digits=8, decimal_places=2, null=True, blank=True)
    e = models.DecimalField(max_digits=8, decimal_places=2, null=True, blank=True)
    f = models.DecimalField(max_digits=8, decimal_places=2, null=True, blank=True)
    g = models.DecimalField(max_digits=8, decimal_places=2, null=True, blank=True)
    minimo = models.DecimalField(max_digits=8, decimal_places=2, null=True, blank=True)
    promedio = models.DecimalField(max_digits=8, decimal_places=2, null=True, blank=True)
    observacion = models.TextField(blank=True)
    orden = models.PositiveIntegerField(default=0)

    class Meta:
        ordering = ["fase", "orden", "punto"]
        constraints = [
            models.UniqueConstraint(
                fields=["polea", "fase", "punto"],
                name="punto_unico_polea_fase_campana",
            )
        ]

    def valores_disponibles(self):
        return [
            valor
            for valor in (self.a, self.b, self.c, self.d, self.e, self.f, self.g)
            if valor is not None
        ]

    def save(self, *args, **kwargs):
        valores = self.valores_disponibles()
        self.minimo = min(valores) if valores else None
        self.promedio = round(sum(valores) / len(valores), 2) if valores else None
        super().save(*args, **kwargs)


class FotoPolea(models.Model):
    polea = models.ForeignKey(PoleaInspeccion, on_delete=models.CASCADE, related_name="fotografias")
    imagen = models.ImageField(upload_to="inspecciones/poleas/")
    codigo_dano = models.CharField(max_length=50, blank=True)
    descripcion = models.TextField(blank=True)
    subida_por = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.PROTECT, related_name="fotos_poleas_subidas")
    creada_en = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering=["creada_en"]


class LifeShaftInspeccion(models.Model):
    inspeccion = models.ForeignKey(Inspeccion, on_delete=models.CASCADE, related_name="life_shafts")
    numero = models.PositiveSmallIntegerField()
    nombre = models.CharField(max_length=100)
    tag = models.CharField(max_length=60, blank=True)
    ubicacion = models.CharField(max_length=120, blank=True)
    condicion = models.CharField(max_length=20, choices=Inspeccion.Condicion.choices, default=Inspeccion.Condicion.NO_MEDIDO)
    tipo_medicion = models.CharField(
        max_length=10,
        choices=TipoMedicionComponente.choices,
        default=TipoMedicionComponente.NORMAL,
    )
    observacion_visual = models.TextField(blank=True)
    observacion_medicion = models.TextField(blank=True)
    recomendaciones = models.TextField(blank=True)
    marca_equipo = models.CharField(max_length=80, default="OLYMPUS")
    tipo_haz = models.CharField(max_length=50, default="COMPLETA")
    frecuencia_mhz = models.CharField(max_length=30, default="5 MHz")
    ancho_banda = models.CharField(max_length=40, default="0.2-10 MHz")
    amortiguamiento = models.CharField(max_length=30, default="50")
    velocidad_ms = models.CharField(max_length=30, default="6079")
    retardo_us = models.CharField(max_length=30, default="1.53")
    orden = models.PositiveIntegerField(default=0)
    creada_en = models.DateTimeField(auto_now_add=True)
    actualizada_en = models.DateTimeField(auto_now=True)

    class Meta:
        ordering=["orden","numero"]
        constraints=[models.UniqueConstraint(fields=["inspeccion","numero"],name="lifeshaft_unico_por_inspeccion")]

    def __str__(self):
        return f"{self.inspeccion.faja.tag} - {self.nombre}"


class MedicionLifeShaft(models.Model):
    life_shaft = models.ForeignKey(LifeShaftInspeccion, on_delete=models.CASCADE, related_name="mediciones")
    punto = models.PositiveSmallIntegerField()
    ubicacion = models.CharField(max_length=100, default="PUNTOS SENTIDO RADIAL")
    a = models.DecimalField(max_digits=8, decimal_places=2, null=True, blank=True)
    b = models.DecimalField(max_digits=8, decimal_places=2, null=True, blank=True)
    c = models.DecimalField(max_digits=8, decimal_places=2, null=True, blank=True)
    d = models.DecimalField(max_digits=8, decimal_places=2, null=True, blank=True)
    e = models.DecimalField(max_digits=8, decimal_places=2, null=True, blank=True)
    f = models.DecimalField(max_digits=8, decimal_places=2, null=True, blank=True)
    g = models.DecimalField(max_digits=8, decimal_places=2, null=True, blank=True)
    promedio = models.DecimalField(max_digits=8, decimal_places=2, null=True, blank=True)
    minimo = models.DecimalField(max_digits=8, decimal_places=2, null=True, blank=True)
    observacion = models.TextField(blank=True)
    orden = models.PositiveIntegerField(default=0)

    class Meta:
        ordering=["orden","punto"]
        constraints=[models.UniqueConstraint(fields=["life_shaft","punto"],name="punto_unico_por_lifeshaft")]

    def valores_disponibles(self):
        return [v for v in [self.a,self.b,self.c,self.d,self.e,self.f,self.g] if v is not None]

    def save(self,*args,**kwargs):
        valores=self.valores_disponibles()
        if valores:
            self.minimo=min(valores)
            self.promedio=round(sum(valores)/len(valores),2)
        else:
            self.minimo=None
            self.promedio=None
        super().save(*args,**kwargs)


class MedicionLifeShaftCampana(models.Model):
    life_shaft = models.ForeignKey(
        LifeShaftInspeccion,
        on_delete=models.CASCADE,
        related_name="mediciones_campana",
    )
    fase = models.CharField(max_length=10, choices=FaseCampana.choices)
    punto = models.PositiveSmallIntegerField()
    ubicacion = models.CharField(max_length=100, default="PUNTOS SENTIDO RADIAL")
    a = models.DecimalField(max_digits=8, decimal_places=2, null=True, blank=True)
    b = models.DecimalField(max_digits=8, decimal_places=2, null=True, blank=True)
    c = models.DecimalField(max_digits=8, decimal_places=2, null=True, blank=True)
    d = models.DecimalField(max_digits=8, decimal_places=2, null=True, blank=True)
    e = models.DecimalField(max_digits=8, decimal_places=2, null=True, blank=True)
    f = models.DecimalField(max_digits=8, decimal_places=2, null=True, blank=True)
    g = models.DecimalField(max_digits=8, decimal_places=2, null=True, blank=True)
    promedio = models.DecimalField(max_digits=8, decimal_places=2, null=True, blank=True)
    minimo = models.DecimalField(max_digits=8, decimal_places=2, null=True, blank=True)
    observacion = models.TextField(blank=True)
    orden = models.PositiveIntegerField(default=0)

    class Meta:
        ordering = ["fase", "orden", "punto"]
        constraints = [
            models.UniqueConstraint(
                fields=["life_shaft", "fase", "punto"],
                name="punto_unico_lifeshaft_fase_campana",
            )
        ]

    def valores_disponibles(self):
        return [
            valor
            for valor in (self.a, self.b, self.c, self.d, self.e, self.f, self.g)
            if valor is not None
        ]

    def save(self, *args, **kwargs):
        valores = self.valores_disponibles()
        self.minimo = min(valores) if valores else None
        self.promedio = round(sum(valores) / len(valores), 2) if valores else None
        super().save(*args, **kwargs)


class FotoLifeShaft(models.Model):
    life_shaft = models.ForeignKey(
        LifeShaftInspeccion,
        on_delete=models.CASCADE,
        related_name="fotografias",
    )
    imagen = models.ImageField(upload_to="inspecciones/life_shaft/")
    codigo_dano = models.CharField(max_length=50, blank=True)
    descripcion = models.TextField(blank=True)
    subida_por = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.PROTECT,
        related_name="fotos_lifeshaft_subidas",
    )
    creada_en = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["creada_en"]

    def __str__(self) -> str:
        return (
            f"Foto Life Shaft {self.life_shaft.numero} - "
            f"{self.life_shaft.inspeccion.codigo_reporte}"
        )

class MedicionEmpalmeCVB0003(models.Model):
    inspeccion = models.ForeignKey(
        Inspeccion,
        on_delete=models.CASCADE,
        related_name="empalmes_cvb0003",
    )

    zona = models.CharField(max_length=50)
    empalme = models.CharField(max_length=20)
    bastidor_lado = models.CharField(max_length=80)
    posicion = models.CharField(max_length=20)

    espesor_nominal = models.DecimalField(
        max_digits=6,
        decimal_places=2,
        null=True,
        blank=True,
    )

    a = models.DecimalField(max_digits=6, decimal_places=2, null=True, blank=True)
    b = models.DecimalField(max_digits=6, decimal_places=2, null=True, blank=True)
    c = models.DecimalField(max_digits=6, decimal_places=2, null=True, blank=True)
    d = models.DecimalField(max_digits=6, decimal_places=2, null=True, blank=True)
    e = models.DecimalField(max_digits=6, decimal_places=2, null=True, blank=True)
    f = models.DecimalField(max_digits=6, decimal_places=2, null=True, blank=True)
    g = models.DecimalField(max_digits=6, decimal_places=2, null=True, blank=True)

    observacion = models.CharField(
        max_length=250,
        blank=True,
    )

    orden = models.PositiveIntegerField(default=0)

    @property
    def valores(self):
        return [
            valor
            for valor in [
                self.a,
                self.b,
                self.c,
                self.d,
                self.e,
                self.f,
                self.g,
            ]
            if valor is not None
        ]

    @property
    def minimo(self):
        return min(self.valores) if self.valores else None

    @property
    def promedio(self):
        if not self.valores:
            return None

        return sum(self.valores) / len(self.valores)

    @property
    def desgaste(self):
        if self.espesor_nominal is None or self.minimo is None:
            return None

        return self.espesor_nominal - self.minimo

    @property
    def porcentaje_desgaste(self):
        if not self.espesor_nominal or self.desgaste is None:
            return None

        return self.desgaste * 100 / self.espesor_nominal

    @property
    def porcentaje_residual(self):
        if self.porcentaje_desgaste is None:
            return None

        return 100 - self.porcentaje_desgaste

    class Meta:
        ordering = ["orden", "id"]

    def __str__(self):
        return f"{self.empalme} - {self.posicion}"


class MedicionTramoCVB0003(models.Model):
    class Tipo(models.TextChoices):
        CARGA = "CARGA", "Carga"
        RETORNO = "RETORNO", "Retorno"

    inspeccion = models.ForeignKey(
        Inspeccion,
        on_delete=models.CASCADE,
        related_name="tramos_cvb0003",
    )

    tipo = models.CharField(
        max_length=10,
        choices=Tipo.choices,
    )

    tramo = models.CharField(max_length=60)
    medicion = models.PositiveIntegerField()
    bastidor = models.CharField(max_length=20)

    espesor_nominal = models.DecimalField(
        max_digits=6,
        decimal_places=2,
        default=19,
    )

    a = models.DecimalField(max_digits=6, decimal_places=2, null=True, blank=True)
    b = models.DecimalField(max_digits=6, decimal_places=2, null=True, blank=True)
    c = models.DecimalField(max_digits=6, decimal_places=2, null=True, blank=True)
    d = models.DecimalField(max_digits=6, decimal_places=2, null=True, blank=True)
    e = models.DecimalField(max_digits=6, decimal_places=2, null=True, blank=True)
    f = models.DecimalField(max_digits=6, decimal_places=2, null=True, blank=True)
    g = models.DecimalField(max_digits=6, decimal_places=2, null=True, blank=True)

    observacion = models.CharField(
        max_length=250,
        blank=True,
    )

    orden = models.PositiveIntegerField(default=0)

    @property
    def valores(self):
        return [
            valor
            for valor in [
                self.a,
                self.b,
                self.c,
                self.d,
                self.e,
                self.f,
                self.g,
            ]
            if valor is not None
        ]

    @property
    def minimo(self):
        return min(self.valores) if self.valores else None

    @property
    def promedio(self):
        if not self.valores:
            return None

        return sum(self.valores) / len(self.valores)

    @property
    def desgaste(self):
        if self.espesor_nominal is None or self.minimo is None:
            return None

        return self.espesor_nominal - self.minimo

    @property
    def porcentaje_desgaste(self):
        if not self.espesor_nominal or self.desgaste is None:
            return None

        return self.desgaste * 100 / self.espesor_nominal

    @property
    def porcentaje_residual(self):
        if self.porcentaje_desgaste is None:
            return None

        return 100 - self.porcentaje_desgaste

    class Meta:
        ordering = ["tipo", "orden", "id"]

    def __str__(self):
        return f"{self.tipo} - {self.medicion} - {self.bastidor}"
    
class FotoFajaCVB0003(models.Model):
    class Seccion(models.TextChoices):
        EMPALMES = "EMPALMES", "Empalmes"
        CARGA = "CARGA", "Tramos de carga"
        RETORNO = "RETORNO", "Tramos de retorno"

    inspeccion = models.ForeignKey(
        Inspeccion,
        on_delete=models.CASCADE,
        related_name="fotografias_cvb0003",
    )

    seccion = models.CharField(
        max_length=20,
        choices=Seccion.choices,
    )

    imagen = models.ImageField(
        upload_to="inspecciones/faja/cvb0003/",
    )

    codigo_dano = models.CharField(
        max_length=100,
        blank=True,
    )

    descripcion = models.TextField(
        blank=True,
    )

    subida_por = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="fotos_faja_cvb0003_subidas",
    )

    creada_en = models.DateTimeField(
        auto_now_add=True,
    )

    class Meta:
        ordering = ["seccion", "creada_en", "id"]

    def __str__(self):
        return (
            f"{self.get_seccion_display()} - "
            f"{self.codigo_dano or 'Sin código'}"
        )


class CalibracionUTFajaCVB0003(models.Model):
    VALORES_INICIALES = (
        ("1636", "2.243"),
        ("1589", "1.652"),
        ("1656", "1.389"),
        ("1543", "2.016"),
        ("1624", "1.973"),
        ("1639", "1.182"),
        ("1671", "1.268"),
        ("1614", "0.967"),
    )

    inspeccion = models.ForeignKey(
        Inspeccion,
        on_delete=models.CASCADE,
        related_name="calibraciones_ut_faja_cvb0003",
    )
    numero = models.PositiveSmallIntegerField()
    marca_equipo = models.CharField(max_length=80, default="Olympus")
    modelo_equipo = models.CharField(max_length=80, default="Epoch 6Lt")
    frecuencia_mhz = models.CharField(max_length=30, default="1")
    rango_mm = models.CharField(max_length=30, default="30-90")
    metodo_empleado = models.CharField(max_length=80, default="Pulso - eco")
    acoplante = models.CharField(max_length=60, default="Echo gel")
    rectificacion = models.CharField(max_length=50, default="Full")
    velocidad_ms = models.CharField(max_length=30, blank=True)
    retardo_us = models.CharField(max_length=30, blank=True)
    tipo_scan = models.CharField(max_length=50, default="A Scan")

    class Meta:
        ordering = ["numero"]
        constraints = [
            models.UniqueConstraint(
                fields=["inspeccion", "numero"],
                name="calibracion_ut_faja_cvb0003_unica",
            )
        ]

    def __str__(self):
        return f"{self.inspeccion.codigo_reporte} - UT {self.numero:02d}"

    @classmethod
    def crear_estructura(cls, inspeccion):
        for numero, (velocidad, retardo) in enumerate(
            cls.VALORES_INICIALES,
            start=1,
        ):
            cls.objects.get_or_create(
                inspeccion=inspeccion,
                numero=numero,
                defaults={
                    "velocidad_ms": velocidad,
                    "retardo_us": retardo,
                },
            )
