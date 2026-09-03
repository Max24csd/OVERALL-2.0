from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("inspecciones", "0012_remove_fotoinspeccion_seccion_and_more"),
    ]

    operations = [
        migrations.AddField(
            model_name="fotoinspeccion",
            name="seccion",
            field=models.CharField(blank=True, max_length=120),
        ),
        migrations.AddField(
            model_name="fotoinspeccion",
            name="titulo",
            field=models.CharField(blank=True, max_length=200),
        ),
    ]
