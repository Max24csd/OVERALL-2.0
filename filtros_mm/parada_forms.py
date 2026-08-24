from django import forms
from django.contrib.auth import get_user_model

from inspecciones.models import Parada


User = get_user_model()


class UsuarioMultipleChoiceField(forms.ModelMultipleChoiceField):
    def label_from_instance(self, obj):
        nombre = obj.get_full_name().strip()

        if nombre:
            return f"{nombre} ({obj.username})"

        return obj.username


class UsuarioChoiceField(forms.ModelChoiceField):
    def label_from_instance(self, obj):
        nombre = obj.get_full_name().strip()

        if nombre:
            return f"{nombre} ({obj.username})"

        return obj.username


class NuevaParadaFiltrosForm(forms.Form):

    nombre = forms.CharField(
        max_length=150,
        initial="PARADA SEPTIEMBRE FILTROS",
        widget=forms.TextInput(
            attrs={
                "class": "form-control",
                "placeholder": "Ej.: PARADA SEPTIEMBRE FILTROS",
            }
        ),
    )

    fecha_inicio = forms.DateField(
        widget=forms.DateInput(
            attrs={
                "type": "date",
                "class": "form-control",
            }
        )
    )

    fecha_fin = forms.DateField(
        required=False,
        widget=forms.DateInput(
            attrs={
                "type": "date",
                "class": "form-control",
            }
        ),
    )

    observaciones = forms.CharField(
        required=False,
        widget=forms.Textarea(
            attrs={
                "rows": 3,
                "class": "form-control",
                "placeholder": "Observaciones de la parada...",
            }
        ),
    )

    # ==========================================================
    # INSPECTORES - PERMITE VARIOS
    # ==========================================================

    inspectores = UsuarioMultipleChoiceField(
        queryset=User.objects.none(),
        required=True,
        widget=forms.CheckboxSelectMultiple(
            attrs={
                "class": "usuarios-checkbox",
            }
        ),
    )

    # ==========================================================
    # SUPERVISORES - PERMITE VARIOS
    # ==========================================================

    supervisores = UsuarioMultipleChoiceField(
        queryset=User.objects.none(),
        required=True,
        widget=forms.CheckboxSelectMultiple(
            attrs={
                "class": "usuarios-checkbox",
            }
        ),
    )

    # ==========================================================
    # ANALISTAS - PERMITE VARIOS
    # ==========================================================

    analistas = UsuarioMultipleChoiceField(
        queryset=User.objects.none(),
        required=True,
        widget=forms.CheckboxSelectMultiple(
            attrs={
                "class": "usuarios-checkbox",
            }
        ),
    )

    # ==========================================================
    # CLIENTE - UNO SOLO
    # ==========================================================

    cliente = UsuarioChoiceField(
        queryset=User.objects.none(),
        required=True,
        empty_label="Seleccione cliente",
        widget=forms.Select(
            attrs={
                "class": "form-control",
            }
        ),
    )

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        activos = User.objects.filter(
            is_active=True
        )

        # ======================================================
        # INSPECTORES
        # ======================================================

        self.fields["inspectores"].queryset = (
            activos
            .filter(
                groups__name="Inspector"
            )
            .distinct()
            .order_by(
                "first_name",
                "last_name",
                "username",
            )
        )

        # ======================================================
        # SUPERVISORES
        # ======================================================

        self.fields["supervisores"].queryset = (
            activos
            .filter(
                groups__name="Supervisor"
            )
            .distinct()
            .order_by(
                "first_name",
                "last_name",
                "username",
            )
        )

        # ======================================================
        # ANALISTAS
        # ======================================================

        self.fields["analistas"].queryset = (
            activos
            .filter(
                groups__name="Analista"
            )
            .distinct()
            .order_by(
                "first_name",
                "last_name",
                "username",
            )
        )

        # ======================================================
        # CLIENTES
        # ======================================================

        self.fields["cliente"].queryset = (
            activos
            .filter(
                groups__name="Cliente"
            )
            .distinct()
            .order_by(
                "first_name",
                "last_name",
                "username",
            )
        )

    def clean(self):
        cleaned = super().clean()

        inicio = cleaned.get(
            "fecha_inicio"
        )

        fin = cleaned.get(
            "fecha_fin"
        )

        nombre = (
            cleaned.get("nombre")
            or ""
        ).strip()

        # ======================================================
        # VALIDAR FECHAS
        # ======================================================

        if (
            inicio
            and fin
            and fin < inicio
        ):
            self.add_error(
                "fecha_fin",
                (
                    "La fecha fin no puede ser "
                    "anterior a la fecha inicio."
                ),
            )

        # ======================================================
        # EVITAR DUPLICAR PARADA
        # ======================================================

        if nombre and inicio:

            existe = (
                Parada.objects
                .filter(
                    nombre__iexact=nombre,
                    fecha_inicio=inicio,
                    planta__iexact="Filtros",
                )
                .exists()
            )

            if existe:
                self.add_error(
                    "nombre",
                    (
                        "Ya existe una parada de "
                        "Filtros con ese nombre y fecha."
                    ),
                )

        return cleaned