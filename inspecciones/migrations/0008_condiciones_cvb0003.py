from django.db import migrations, models


CONDICIONES = [
    ("NORMAL", "Normal"),
    ("TOLERABLE", "Tolerable"),
    ("PRECAUCION", "Precaución"),
    ("CRITICO", "Crítico"),
    ("NO_MEDIDO", "No medido"),
]


class Migration(migrations.Migration):
    dependencies = [
        ("inspecciones", "0007_tipo_medicion_campana"),
    ]

    operations = [
        migrations.AlterField(
            model_name="inspeccion",
            name="condicion_general",
            field=models.CharField(
                choices=CONDICIONES,
                default="NORMAL",
                max_length=20,
            ),
        ),
        migrations.AlterField(
            model_name="lifeshaftinspeccion",
            name="condicion",
            field=models.CharField(
                choices=CONDICIONES,
                default="NO_MEDIDO",
                max_length=20,
            ),
        ),
        migrations.AlterField(
            model_name="medicion",
            name="condicion",
            field=models.CharField(
                choices=CONDICIONES,
                default="NORMAL",
                max_length=20,
            ),
        ),
        migrations.AlterField(
            model_name="poleainspeccion",
            name="condicion",
            field=models.CharField(
                choices=CONDICIONES,
                default="NO_MEDIDO",
                max_length=20,
            ),
        ),
    ]
