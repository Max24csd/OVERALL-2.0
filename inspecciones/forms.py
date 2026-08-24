from django import forms
from django.forms import inlineformset_factory

from django.contrib.auth import get_user_model
from .models import (
    CalibracionUTFajaCVB0003,
    FotoFajaCVB0003,
    FotoInspeccion,
    FotoLifeShaft,
    FotoPolea,
    Inspeccion,
    LifeShaftInspeccion,
    Medicion,
    MedicionEmpalmeCVB0003,
    MedicionLifeShaft,
    MedicionLifeShaftCampana,
    MedicionPolea,
    MedicionPoleaCampana,
    MedicionTramoCVB0003,
    PoleaInspeccion,
    Parada,
)


CONDICIONES_CVB0003 = (
    (Inspeccion.Condicion.NORMAL, "NORMAL"),
    (Inspeccion.Condicion.TOLERABLE, "TOLERABLE"),
    (Inspeccion.Condicion.PRECAUCION, "PRECAUCIÓN"),
    (Inspeccion.Condicion.CRITICO, "CRÍTICO"),
)

CONDICIONES_OTROS_REPORTES = (
    (Inspeccion.Condicion.NORMAL, "Normal"),
    (Inspeccion.Condicion.PRECAUCION, "Precaución"),
    (Inspeccion.Condicion.CRITICO, "Crítico"),
    (Inspeccion.Condicion.NO_MEDIDO, "No medido"),
)


def _es_cvb0003(instancia):
    if isinstance(instancia, Inspeccion):
        inspeccion = instancia
    else:
        inspeccion = getattr(instancia, "inspeccion", None)
    faja = getattr(inspeccion, "faja", None)
    return (getattr(faja, "tag", "") or "").upper().strip() in {
        "CVB0003",
        "0220-CVB-0003",
        "0220-CVB0003",
    }


def _configurar_condiciones(formulario, instancia, nombre_campo):
    formulario.fields[nombre_campo].choices = (
        CONDICIONES_CVB0003
        if _es_cvb0003(instancia)
        else CONDICIONES_OTROS_REPORTES
    )

User = get_user_model()


class NuevaParadaForm(forms.ModelForm):
    inspector = forms.ModelChoiceField(
        queryset=User.objects.none(),
        label="Inspector",
        widget=forms.Select(attrs={"class": "form-control"}),
    )

    supervisor = forms.ModelChoiceField(
        queryset=User.objects.none(),
        label="Supervisor",
        widget=forms.Select(attrs={"class": "form-control"}),
    )

    analista = forms.ModelChoiceField(
        queryset=User.objects.none(),
        label="Analista",
        widget=forms.Select(attrs={"class": "form-control"}),
    )

    cliente = forms.ModelChoiceField(
        queryset=User.objects.none(),
        label="Cliente",
        required=False,
        widget=forms.Select(attrs={"class": "form-control"}),
    )

    class Meta:
        model = Parada

        fields = [
            "nombre",
            "fecha_inicio",
            "fecha_fin",
            "observaciones",
        ]

        widgets = {
            "nombre": forms.TextInput(
                attrs={
                    "class": "form-control",
                    "placeholder": "Ej. Parada Chancado Agosto 2026",
                }
            ),
            "fecha_inicio": forms.DateInput(
                attrs={
                    "type": "date",
                    "class": "form-control",
                }
            ),
            "fecha_fin": forms.DateInput(
                attrs={
                    "type": "date",
                    "class": "form-control",
                }
            ),
            "observaciones": forms.Textarea(
                attrs={
                    "class": "form-control",
                    "rows": 3,
                    "placeholder": "Observaciones opcionales",
                }
            ),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        usuarios_activos = User.objects.filter(is_active=True)

        self.fields["inspector"].queryset = usuarios_activos.filter(
            groups__name="Inspector"
        ).distinct()

        self.fields["supervisor"].queryset = usuarios_activos.filter(
            groups__name="Supervisor"
        ).distinct()

        self.fields["analista"].queryset = usuarios_activos.filter(
            groups__name="Analista"
        ).distinct()

        self.fields["cliente"].queryset = usuarios_activos.filter(
            groups__name="Cliente"
        ).distinct()

    def clean(self):
        cleaned_data = super().clean()

        fecha_inicio = cleaned_data.get("fecha_inicio")
        fecha_fin = cleaned_data.get("fecha_fin")

        if fecha_inicio and fecha_fin and fecha_fin < fecha_inicio:
            self.add_error(
                "fecha_fin",
                "La fecha de fin no puede ser anterior a la fecha de inicio.",
            )

        return cleaned_data

class InspeccionForm(forms.ModelForm):

    class Meta:
        model = Inspeccion

        fields = [
            "fecha_inspeccion",
            "fecha_reporte",
            "inspector_campo_nombre",
            "supervisor_campo_nombre",
            "analista_elabora_nombre",
            "analista_valida_nombre",
            "condicion_general",
            "condicion_equipo",
            "circunstancias",
            "antecedentes",
            "observaciones",
            "recomendaciones",
        ]

        widgets = {
            "fecha_inspeccion": forms.DateInput(
                format="%Y-%m-%d",
                attrs={
                    "type": "date",
                    "class": "form-control",
                },
            ),
            "fecha_reporte": forms.DateInput(
                format="%Y-%m-%d",
                attrs={
                    "type": "date",
                    "class": "form-control",
                },
            ),
            "inspector_campo_nombre": forms.TextInput(
                attrs={
                    "class": "form-control",
                    "placeholder": "Nombre del inspector de campo",
                }
            ),
            "supervisor_campo_nombre": forms.TextInput(
                attrs={
                    "class": "form-control",
                    "placeholder": "Nombre del supervisor de campo",
                }
            ),
            "analista_elabora_nombre": forms.TextInput(
                attrs={
                    "class": "form-control",
                    "placeholder": "Nombre del analista que elabora",
                }
            ),
            "analista_valida_nombre": forms.TextInput(
                attrs={
                    "class": "form-control",
                    "placeholder": "Nombre del analista que valida",
                }
            ),
            "condicion_general": forms.Select(
                attrs={"class": "form-control"}
            ),
            "condicion_equipo": forms.TextInput(
                attrs={"class": "form-control"}
            ),
            "circunstancias": forms.Textarea(
                attrs={"class": "form-control", "rows": 4}
            ),
            "antecedentes": forms.Textarea(
                attrs={"class": "form-control", "rows": 4}
            ),
            "observaciones": forms.Textarea(
                attrs={"class": "form-control", "rows": 4}
            ),
            "recomendaciones": forms.Textarea(
                attrs={"class": "form-control", "rows": 4}
            ),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        _configurar_condiciones(self, self.instance, "condicion_general")

        self.fields["fecha_inspeccion"].input_formats = [
            "%Y-%m-%d"
        ]
        self.fields["fecha_reporte"].input_formats = [
            "%Y-%m-%d"
        ]

class MedicionForm(forms.ModelForm):
    class Meta:
        model = Medicion
        fields = ["seccion","punto","bastidor","lado","posicion","espesor_nominal","a","b","c","d","e","f","g","orden"]
        widgets = {
            "seccion": forms.HiddenInput(), "punto": forms.HiddenInput(), "bastidor": forms.HiddenInput(),
            "lado": forms.HiddenInput(), "posicion": forms.HiddenInput(), "orden": forms.HiddenInput(),
            "espesor_nominal": forms.NumberInput(attrs={"class":"measurement-input nominal-input","step":"0.01","min":"0"}),
            **{letra: forms.NumberInput(attrs={"class":"measurement-input value-input","step":"0.01"}) for letra in "abcdefg"},
        }

MedicionFormSet = inlineformset_factory(Inspeccion, Medicion, form=MedicionForm, extra=0, can_delete=False)


class FotoInspeccionForm(forms.ModelForm):
    class Meta:
        model = FotoInspeccion

        fields = [
            "imagen",
            "descripcion",
        ]

        widgets = {
            "imagen": forms.FileInput(
                attrs={
                    "accept": "image/*",
                    "class": "photo-input",
                }
            ),
            "descripcion": forms.Textarea(
                attrs={
                    "class": "form-control",
                    "rows": 3,
                    "placeholder": "Descripción u observación",
                }
            ),
        }

    def clean_imagen(self):
        imagen = self.cleaned_data.get("imagen")

        if not imagen:
            return imagen

        # Máximo 10 MB
        if imagen.size > 10 * 1024 * 1024:
            raise forms.ValidationError(
                "La fotografía no puede superar los 10 MB."
            )

        tipos_permitidos = [
            "image/jpeg",
            "image/png",
            "image/webp",
        ]

        tipo = getattr(
            imagen,
            "content_type",
            "",
        )

        if tipo and tipo not in tipos_permitidos:
            raise forms.ValidationError(
                "Utiliza una imagen JPG, PNG o WEBP."
            )

        return imagen

    def clean_imagen(self):
        imagen = self.cleaned_data.get("imagen")
        if not imagen:
            return imagen
        if imagen.size > 10 * 1024 * 1024:
            raise forms.ValidationError(
                "La fotografía no puede superar los 10 MB."
            )
        tipos_permitidos = ["image/jpeg", "image/png", "image/webp"]
        tipo = getattr(imagen, "content_type", "")
        if tipo and tipo not in tipos_permitidos:
            raise forms.ValidationError("Utiliza una imagen JPG, PNG o WEBP.")
        return imagen


FotoInspeccionFormSet = inlineformset_factory(
    Inspeccion,
    FotoInspeccion,
    form=FotoInspeccionForm,
    extra=5,
    can_delete=True,
)


class PoleaInspeccionForm(forms.ModelForm):
    class Meta:
        model = PoleaInspeccion
        fields = [
            "nombre","tag","ubicacion","componente","material","espesor_nominal","condicion","tipo_medicion",
            "observacion_visual","observacion_medicion","recomendaciones","marca_equipo","modelo_equipo",
            "frecuencia_mhz","rango_mm","metodo_empleado","componente_calibracion","acoplante",
            "rectificacion","velocidad_ms","retardo_us","tipo_scan",
        ]
        widgets = {
            **{f: forms.TextInput(attrs={"class":"form-control"}) for f in ["nombre","tag","ubicacion","componente","material","marca_equipo","modelo_equipo","frecuencia_mhz","rango_mm","metodo_empleado","componente_calibracion","acoplante","rectificacion","velocidad_ms","retardo_us","tipo_scan"]},
            "espesor_nominal": forms.NumberInput(attrs={"class":"form-control","step":"0.01"}),
            "condicion": forms.Select(attrs={"class":"form-control"}),
            "tipo_medicion": forms.Select(attrs={"class":"form-control tipo-medicion-selector"}),
            "observacion_visual": forms.Textarea(attrs={"class":"form-control","rows":3}),
            "observacion_medicion": forms.Textarea(attrs={"class":"form-control","rows":3}),
            "recomendaciones": forms.Textarea(attrs={"class":"form-control","rows":3}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        _configurar_condiciones(self, self.instance, "condicion")
        self.fields["tipo_medicion"].label = "TIPO DE MEDICIÓN"
        self.fields["tipo_medicion"].choices = (
            ("NORMAL", "NORMAL"),
            ("CAMPANA", "CAMPAÑA (INICIO + FIN)"),
        )


class MedicionPoleaForm(forms.ModelForm):
    class Meta:
        model = MedicionPolea
        fields = ["punto","posicion","a","b","c","d","e","f","g","observacion","orden"]
        widgets = {
            "punto": forms.HiddenInput(), "orden": forms.HiddenInput(),
            "posicion": forms.TextInput(attrs={"class":"compact-text"}),
            **{letra: forms.NumberInput(attrs={"class":"measurement-input value-input","step":"0.01"}) for letra in "abcdefg"},
            "observacion": forms.TextInput(attrs={"class":"compact-text"}),
        }

MedicionPoleaFormSet = inlineformset_factory(PoleaInspeccion, MedicionPolea, form=MedicionPoleaForm, extra=0, can_delete=False)


class MedicionPoleaCampanaForm(forms.ModelForm):
    class Meta:
        model = MedicionPoleaCampana
        fields = ["punto", "posicion", "a", "b", "c", "d", "e", "f", "g", "observacion", "orden"]
        widgets = {
            "punto": forms.HiddenInput(),
            "orden": forms.HiddenInput(),
            "posicion": forms.TextInput(attrs={"class": "compact-text"}),
            **{
                letra: forms.NumberInput(
                    attrs={"class": "measurement-input value-input", "step": "0.01"}
                )
                for letra in "abcdefg"
            },
            "observacion": forms.TextInput(attrs={"class": "compact-text"}),
        }


def crear_medicion_polea_campana_formset(extra=0):
    return inlineformset_factory(
        PoleaInspeccion,
        MedicionPoleaCampana,
        form=MedicionPoleaCampanaForm,
        extra=extra,
        can_delete=False,
    )


class FotoPoleaForm(forms.ModelForm):
    """
    Formulario para registrar una fotografía vinculada
    a una polea específica.
    """

    class Meta:
        model = FotoPolea

        fields = [
            "imagen",
            "descripcion",
        ]

        widgets = {
            "imagen": forms.ClearableFileInput(
                attrs={
                    "class": "polea-photo-input",
                    "accept": "image/*",
                }
            ),
            "descripcion": forms.Textarea(
                attrs={
                    "class": "form-control",
                    "rows": 3,
                    "placeholder": (
                        "Describe la condición encontrada."
                    ),
                }
            ),
        }

    def save(self, commit=True):
        instancia = super().save(commit=False)

        # El código de daño ya no se utiliza.
        instancia.codigo_dano = ""

        if commit:
            instancia.save()

        return instancia

    def clean_imagen(self):
        imagen = self.cleaned_data.get("imagen")

        if not imagen:
            return imagen

        limite_bytes = 10 * 1024 * 1024

        if imagen.size > limite_bytes:
            raise forms.ValidationError(
                "La fotografía no puede superar los 10 MB."
            )

        tipos_permitidos = [
            "image/jpeg",
            "image/png",
            "image/webp",
        ]

        tipo_contenido = getattr(
            imagen,
            "content_type",
            "",
        )

        if (
            tipo_contenido
            and tipo_contenido not in tipos_permitidos
        ):
            raise forms.ValidationError(
                (
                    "Formato no permitido. "
                    "Utiliza JPG, PNG o WEBP."
                )
            )

        return imagen


FotoPoleaFormSet = inlineformset_factory(
    parent_model=PoleaInspeccion,
    model=FotoPolea,
    form=FotoPoleaForm,
    extra=10,
    can_delete=True,
)


class LifeShaftInspeccionForm(forms.ModelForm):
    class Meta:
        model = LifeShaftInspeccion

        # Únicamente los campos que el inspector puede editar.
        # Los parámetros UT no se incluyen porque son valores fijos.
        fields = [
            "nombre",
            "tag",
            "ubicacion",
            "condicion",
            "tipo_medicion",
            "observacion_visual",
            "observacion_medicion",
            "recomendaciones",
            "marca_equipo",
            "tipo_haz",
            "frecuencia_mhz",
            "ancho_banda",
            "amortiguamiento",
            "velocidad_ms",
            "retardo_us",
        ]

        widgets = {
            "nombre": forms.TextInput(
                attrs={
                    "class": "form-control",
                }
            ),
            "tag": forms.TextInput(
                attrs={
                    "class": "form-control",
                }
            ),
            "ubicacion": forms.TextInput(
                attrs={
                    "class": "form-control",
                    "placeholder": "Ubicación del Life Shaft",
                }
            ),
            "condicion": forms.Select(
                attrs={
                    "class": "form-control",
                }
            ),
            "tipo_medicion": forms.Select(
                attrs={
                    "class": "form-control tipo-medicion-selector",
                }
            ),
            "observacion_visual": forms.Textarea(
                attrs={
                    "class": "form-control",
                    "rows": 3,
                }
            ),
            "observacion_medicion": forms.Textarea(
                attrs={
                    "class": "form-control",
                    "rows": 3,
                }
            ),
            "recomendaciones": forms.Textarea(
                attrs={
                    "class": "form-control",
                    "rows": 3,
                }
            ),
            "marca_equipo": forms.TextInput(attrs={"class": "form-control"}),
            "tipo_haz": forms.TextInput(attrs={"class": "form-control"}),
            "frecuencia_mhz": forms.TextInput(attrs={"class": "form-control"}),
            "ancho_banda": forms.TextInput(attrs={"class": "form-control"}),
            "amortiguamiento": forms.TextInput(attrs={"class": "form-control"}),
            "velocidad_ms": forms.TextInput(attrs={"class": "form-control"}),
            "retardo_us": forms.TextInput(attrs={"class": "form-control"}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        _configurar_condiciones(self, self.instance, "condicion")

        # Estos campos pueden quedar vacíos durante el borrador.
        self.fields["tipo_medicion"].label = "TIPO DE MEDICIÓN"
        self.fields["tipo_medicion"].choices = (
            ("NORMAL", "NORMAL"),
            ("CAMPANA", "CAMPAÑA (INICIO + FIN)"),
        )
        self.fields["ubicacion"].required = False
        self.fields["observacion_visual"].required = False
        self.fields["observacion_medicion"].required = False
        self.fields["recomendaciones"].required = False

        etiquetas_ut = {
            "marca_equipo": "Marca del equipo",
            "tipo_haz": "Tipo de haz",
            "frecuencia_mhz": "Frecuencia (MHz)",
            "ancho_banda": "Ancho de banda",
            "amortiguamiento": "Amortiguamiento",
            "velocidad_ms": "Velocidad (m/s)",
            "retardo_us": "Retardo (µs)",
        }
        for campo, etiqueta in etiquetas_ut.items():
            self.fields[campo].label = etiqueta


class MedicionLifeShaftForm(forms.ModelForm):
    class Meta:
        model = MedicionLifeShaft
        fields = ["punto","ubicacion","a","b","c","d","e","f","g","observacion","orden"]
        widgets = {
            "punto": forms.HiddenInput(), "orden": forms.HiddenInput(),
            "ubicacion": forms.TextInput(attrs={"class":"compact-text"}),
            **{letra: forms.NumberInput(attrs={"class":"measurement-input value-input","step":"0.01"}) for letra in "abcdefg"},
            "observacion": forms.TextInput(attrs={"class":"compact-text"}),
        }

MedicionLifeShaftFormSet = inlineformset_factory(LifeShaftInspeccion, MedicionLifeShaft, form=MedicionLifeShaftForm, extra=0, can_delete=False)


class MedicionLifeShaftCampanaForm(forms.ModelForm):
    class Meta:
        model = MedicionLifeShaftCampana
        fields = ["punto", "ubicacion", "a", "b", "c", "d", "e", "f", "g", "observacion", "orden"]
        widgets = {
            "punto": forms.HiddenInput(),
            "orden": forms.HiddenInput(),
            "ubicacion": forms.TextInput(attrs={"class": "compact-text"}),
            **{
                letra: forms.NumberInput(
                    attrs={"class": "measurement-input value-input", "step": "0.01"}
                )
                for letra in "abcdefg"
            },
            "observacion": forms.TextInput(attrs={"class": "compact-text"}),
        }


def crear_medicion_life_shaft_campana_formset(extra=0):
    return inlineformset_factory(
        LifeShaftInspeccion,
        MedicionLifeShaftCampana,
        form=MedicionLifeShaftCampanaForm,
        extra=extra,
        can_delete=False,
    )

class FotoLifeShaftForm(forms.ModelForm):
    class Meta:
        model = FotoLifeShaft

        fields = [
            "imagen",
            "codigo_dano",
            "descripcion",
        ]

        widgets = {
            "imagen": forms.ClearableFileInput(
                attrs={
                    "accept": "image/*",
                }
            ),
            "codigo_dano": forms.TextInput(),
            "descripcion": forms.Textarea(
                attrs={
                    "rows": 3,
                }
            ),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        # Permite guardar el formulario o borrar una foto
        # sin tener que seleccionar una imagen nueva.
        self.fields["imagen"].required = False
        self.fields["codigo_dano"].required = False
        self.fields["descripcion"].required = False

class MedicionEmpalmeCVB0003Form(forms.ModelForm):
    class Meta:
        model = MedicionEmpalmeCVB0003

        fields = [
            "posicion",
            "espesor_nominal",
            "a",
            "b",
            "c",
            "d",
            "e",
            "f",
            "g",
            "observacion",
        ]

        widgets = {
            "posicion": forms.TextInput(
                attrs={
                    "class": "measurement-input",
                    "readonly": "readonly",
                }
            ),
            "espesor_nominal": forms.NumberInput(
                attrs={
                    "class": "measurement-input",
                    "step": "0.01",
                }
            ),
            "a": forms.NumberInput(
                attrs={
                    "class": "measurement-input",
                    "step": "0.01",
                }
            ),
            "b": forms.NumberInput(
                attrs={
                    "class": "measurement-input",
                    "step": "0.01",
                }
            ),
            "c": forms.NumberInput(
                attrs={
                    "class": "measurement-input",
                    "step": "0.01",
                }
            ),
            "d": forms.NumberInput(
                attrs={
                    "class": "measurement-input",
                    "step": "0.01",
                }
            ),
            "e": forms.NumberInput(
                attrs={
                    "class": "measurement-input",
                    "step": "0.01",
                }
            ),
            "f": forms.NumberInput(
                attrs={
                    "class": "measurement-input",
                    "step": "0.01",
                }
            ),
            "g": forms.NumberInput(
                attrs={
                    "class": "measurement-input",
                    "step": "0.01",
                }
            ),
            "observacion": forms.TextInput(
                attrs={
                    "class": "measurement-input",
                    "placeholder": "Observación",
                }
            ),
        }
        
MedicionEmpalmeCVB0003FormSet = inlineformset_factory(
    Inspeccion,
    MedicionEmpalmeCVB0003,
    form=MedicionEmpalmeCVB0003Form,
    extra=0,
    can_delete=False,
)
class MedicionTramoCVB0003Form(forms.ModelForm):
    class Meta:
        model = MedicionTramoCVB0003

        fields = [
            "espesor_nominal",
            "a",
            "b",
            "c",
            "d",
            "e",
            "f",
            "g",
            "observacion",
        ]

        widgets = {
            "espesor_nominal": forms.NumberInput(
                attrs={
                    "class": "measurement-input",
                    "step": "0.01",
                }
            ),
            "a": forms.NumberInput(
                attrs={
                    "class": "measurement-input",
                    "step": "0.01",
                }
            ),
            "b": forms.NumberInput(
                attrs={
                    "class": "measurement-input",
                    "step": "0.01",
                }
            ),
            "c": forms.NumberInput(
                attrs={
                    "class": "measurement-input",
                    "step": "0.01",
                }
            ),
            "d": forms.NumberInput(
                attrs={
                    "class": "measurement-input",
                    "step": "0.01",
                }
            ),
            "e": forms.NumberInput(
                attrs={
                    "class": "measurement-input",
                    "step": "0.01",
                }
            ),
            "f": forms.NumberInput(
                attrs={
                    "class": "measurement-input",
                    "step": "0.01",
                }
            ),
            "g": forms.NumberInput(
                attrs={
                    "class": "measurement-input",
                    "step": "0.01",
                }
            ),
            "observacion": forms.TextInput(
                attrs={
                    "class": "measurement-input",
                    "placeholder": "Observación",
                }
            ),
        }
MedicionTramoCVB0003FormSet = inlineformset_factory(
    Inspeccion,
    MedicionTramoCVB0003,
    form=MedicionTramoCVB0003Form,
    extra=0,
    can_delete=False,
)


CAMPOS_CALIBRACION_UT_FAJA_CVB0003 = (
    "marca_equipo",
    "modelo_equipo",
    "frecuencia_mhz",
    "rango_mm",
    "metodo_empleado",
    "acoplante",
    "rectificacion",
    "velocidad_ms",
    "retardo_us",
    "tipo_scan",
)


class CalibracionUTFajaCVB0003Form(forms.ModelForm):
    class Meta:
        model = CalibracionUTFajaCVB0003
        fields = CAMPOS_CALIBRACION_UT_FAJA_CVB0003
        widgets = {
            campo: forms.TextInput(attrs={"class": "form-control"})
            for campo in CAMPOS_CALIBRACION_UT_FAJA_CVB0003
        }
        labels = {
            "marca_equipo": "Marca",
            "modelo_equipo": "Modelo",
            "frecuencia_mhz": "Frecuencia (MHz)",
            "rango_mm": "Rango (mm)",
            "metodo_empleado": "Método empleado",
            "acoplante": "Acoplante",
            "rectificacion": "Rectificación",
            "velocidad_ms": "Velocidad (m/s)",
            "retardo_us": "Retardo (µs)",
            "tipo_scan": "Tipo de scan",
        }


CalibracionUTFajaCVB0003FormSet = inlineformset_factory(
    Inspeccion,
    CalibracionUTFajaCVB0003,
    form=CalibracionUTFajaCVB0003Form,
    extra=0,
    can_delete=False,
)


class FotoFajaCVB0003Form(forms.ModelForm):
    class Meta:
        model = FotoFajaCVB0003

        fields = [
            "imagen",
            "codigo_dano",
            "descripcion",
        ]

        widgets = {
            "imagen": forms.ClearableFileInput(
                attrs={
                    "accept": "image/*",
                }
            ),
            "codigo_dano": forms.TextInput(
                attrs={
                    "placeholder": "Ejemplo: DAÑO-01",
                }
            ),
            "descripcion": forms.Textarea(
                attrs={
                    "rows": 2,
                    "placeholder": "Describe el daño observado",
                }
            ),
        }


FotoFajaCVB0003FormSet = inlineformset_factory(
    Inspeccion,
    FotoFajaCVB0003,
    form=FotoFajaCVB0003Form,
    extra=30,
    can_delete=True,
)
FotoLifeShaftFormSet = inlineformset_factory(
    LifeShaftInspeccion,
    FotoLifeShaft,
    form=FotoLifeShaftForm,
    extra=10,
    can_delete=True,
)
