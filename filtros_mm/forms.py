from django import forms

from .models import ReporteFiltro


class ReporteFiltroCabeceraForm(forms.ModelForm):
    """
    Campos técnicos editables por el inspector.

    Los nombres de Inspector/Supervisor/Analista/Cliente NO se escriben
    manualmente: se toman de AccesoParada en views.py.
    """

    circunstancias = forms.CharField(
        required=False,
        widget=forms.Textarea(
            attrs={"rows": 4, "placeholder": "Circunstancias de la inspección..."}
        ),
    )
    antecedentes = forms.CharField(
        required=False,
        widget=forms.Textarea(
            attrs={"rows": 4, "placeholder": "Antecedentes..."}
        ),
    )
    observaciones = forms.CharField(
        required=False,
        widget=forms.Textarea(
            attrs={"rows": 5, "placeholder": "Observaciones técnicas..."}
        ),
    )
    recomendaciones = forms.CharField(
        required=False,
        widget=forms.Textarea(
            attrs={"rows": 5, "placeholder": "Conclusiones y recomendaciones..."}
        ),
    )
    observaciones_ruedas = forms.CharField(
        required=False,
        widget=forms.Textarea(
            attrs={
                "rows": 4,
                "placeholder": "Observaciones de la inspección de ruedas...",
            }
        ),
    )

    ut_lh_marca = forms.CharField(required=False, label="MARCA DEL EQUIPO")
    ut_lh_modelo = forms.CharField(required=False, label="MODELO")
    ut_lh_tipo_haz = forms.CharField(required=False, label="TIPO DE HAZ")
    ut_lh_ganancia = forms.CharField(required=False, label="GANANCIA (dB)")
    ut_lh_frecuencia = forms.CharField(required=False, label="FRECUENCIA (MHz)")
    ut_lh_velocidad = forms.CharField(required=False, label="VELOCIDAD (m/s)")
    ut_lh_ancho_banda = forms.CharField(required=False, label="ANCHO DE BANDA")
    ut_lh_retardo = forms.CharField(required=False, label="RETARDO (us)")
    ut_lh_amortiguamiento = forms.CharField(required=False, label="AMORTIGUAMIENTO")
    ut_lh_diametro = forms.CharField(required=False, label="DIAMETRO (mm)")

    ut_rh_marca = forms.CharField(required=False, label="MARCA DEL EQUIPO")
    ut_rh_modelo = forms.CharField(required=False, label="MODELO")
    ut_rh_tipo_haz = forms.CharField(required=False, label="TIPO DE HAZ")
    ut_rh_ganancia = forms.CharField(required=False, label="GANANCIA (dB)")
    ut_rh_frecuencia = forms.CharField(required=False, label="FRECUENCIA (MHz)")
    ut_rh_velocidad = forms.CharField(required=False, label="VELOCIDAD (m/s)")
    ut_rh_ancho_banda = forms.CharField(required=False, label="ANCHO DE BANDA")
    ut_rh_retardo = forms.CharField(required=False, label="RETARDO (us)")
    ut_rh_amortiguamiento = forms.CharField(required=False, label="AMORTIGUAMIENTO")
    ut_rh_diametro = forms.CharField(required=False, label="DIAMETRO (mm)")

    class Meta:
        model = ReporteFiltro
        fields = [
            "fecha_inspeccion",
            "fecha_reporte",
            "condicion_general",
        ]
        widgets = {
            "fecha_inspeccion": forms.DateInput(
                attrs={"type": "date"},
                format="%Y-%m-%d",
            ),
            "fecha_reporte": forms.DateInput(
                attrs={"type": "date"},
                format="%Y-%m-%d",
            ),
            "condicion_general": forms.Select(),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        datos = self.instance.datos or {}

        campos_texto = [
            "circunstancias",
            "antecedentes",
            "observaciones",
            "recomendaciones",
            "observaciones_ruedas",
        ]

        campos_ut = [
            "ut_lh_marca",
            "ut_lh_modelo",
            "ut_lh_tipo_haz",
            "ut_lh_ganancia",
            "ut_lh_frecuencia",
            "ut_lh_velocidad",
            "ut_lh_ancho_banda",
            "ut_lh_retardo",
            "ut_lh_amortiguamiento",
            "ut_lh_diametro",
            "ut_rh_marca",
            "ut_rh_modelo",
            "ut_rh_tipo_haz",
            "ut_rh_ganancia",
            "ut_rh_frecuencia",
            "ut_rh_velocidad",
            "ut_rh_ancho_banda",
            "ut_rh_retardo",
            "ut_rh_amortiguamiento",
            "ut_rh_diametro",
        ]

        defaults = {
            "ut_lh_marca": "OLYMPUS",
            "ut_lh_tipo_haz": "HAZ RECTO",
            "ut_lh_frecuencia": "5",
            "ut_lh_ancho_banda": "1.5 - 8 MHZ",
            "ut_lh_amortiguamiento": "100",
            "ut_rh_marca": "OLYMPUS",
            "ut_rh_tipo_haz": "HAZ RECTO",
            "ut_rh_frecuencia": "5",
            "ut_rh_ancho_banda": "1.5 - 8 MHZ",
            "ut_rh_amortiguamiento": "100",
        }

        for campo in campos_texto:
            self.fields[campo].initial = datos.get(campo, "")

        parametros_ut = datos.get("parametros_ut") or {}
        for campo in campos_ut:
            self.fields[campo].initial = parametros_ut.get(campo, defaults.get(campo, ""))

    def guardar_datos_texto(self, reporte):
        datos = dict(reporte.datos or {})

        for campo in [
            "circunstancias",
            "antecedentes",
            "observaciones",
            "recomendaciones",
            "observaciones_ruedas",
        ]:
            datos[campo] = self.cleaned_data.get(campo, "")

        datos["parametros_ut"] = {
            campo: self.cleaned_data.get(campo, "")
            for campo in [
                "ut_lh_marca",
                "ut_lh_modelo",
                "ut_lh_tipo_haz",
                "ut_lh_ganancia",
                "ut_lh_frecuencia",
                "ut_lh_velocidad",
                "ut_lh_ancho_banda",
                "ut_lh_retardo",
                "ut_lh_amortiguamiento",
                "ut_lh_diametro",
                "ut_rh_marca",
                "ut_rh_modelo",
                "ut_rh_tipo_haz",
                "ut_rh_ganancia",
                "ut_rh_frecuencia",
                "ut_rh_velocidad",
                "ut_rh_ancho_banda",
                "ut_rh_retardo",
                "ut_rh_amortiguamiento",
                "ut_rh_diametro",
            ]
        }

        reporte.datos = datos
        return reporte
