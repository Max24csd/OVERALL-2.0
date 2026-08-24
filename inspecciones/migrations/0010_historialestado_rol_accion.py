from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("inspecciones", "0009_calibracion_ut_faja_cvb0003"),
    ]

    operations = [
        migrations.AddField(
            model_name="historialestado",
            name="rol",
            field=models.CharField(blank=True, default="", max_length=30),
        ),
        migrations.AddField(
            model_name="historialestado",
            name="accion",
            field=models.CharField(
                blank=True,
                choices=[
                    ("ENVIAR_A_REVISION", "Enviar a revisión"),
                    ("DEVOLVER_SUPERVISOR", "Devolver supervisor"),
                    ("APROBAR_SUPERVISOR", "Aprobar supervisor"),
                    ("DEVOLVER_ANALISTA", "Devolver analista"),
                    ("APROBAR_ANALISTA", "Aprobar analista"),
                    ("PUBLICAR", "Publicar"),
                ],
                default="",
                max_length=40,
            ),
        ),
    ]
