from django.contrib import admin
from .models import (
    Faja, Inspeccion, Medicion, FotoInspeccion, HistorialEstado,
    PoleaInspeccion, MedicionPolea, FotoPolea,
    LifeShaftInspeccion, MedicionLifeShaft, FotoLifeShaft,
)

admin.site.register(Faja)
admin.site.register(Inspeccion)
admin.site.register(Medicion)
admin.site.register(FotoInspeccion)
admin.site.register(HistorialEstado)
admin.site.register(PoleaInspeccion)
admin.site.register(MedicionPolea)
admin.site.register(FotoPolea)
admin.site.register(LifeShaftInspeccion)
admin.site.register(MedicionLifeShaft)
admin.site.register(FotoLifeShaft)
