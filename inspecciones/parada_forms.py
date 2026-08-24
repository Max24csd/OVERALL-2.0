from django import forms
from django.contrib.auth import get_user_model

from accounts.models import PerfilUsuario


User = get_user_model()


class GestionAsignacionParadaForm(forms.Form):
    ACCIONES = (
        ("AGREGAR", "Agregar integrante"),
        ("QUITAR", "Quitar integrante"),
        ("CAMBIAR_CLIENTE", "Cambiar cliente"),
    )

    ROLES = (
        ("Inspector", "Inspector"),
        ("Supervisor", "Supervisor"),
        ("Analista", "Analista"),
        ("Cliente", "Cliente"),
    )

    accion = forms.ChoiceField(
        label="Acción",
        choices=ACCIONES,
        widget=forms.Select(
            attrs={"class": "form-control"}
        ),
    )

    rol = forms.ChoiceField(
        label="Rol",
        choices=ROLES,
        widget=forms.Select(
            attrs={"class": "form-control"}
        ),
    )

    usuario = forms.ModelChoiceField(
        queryset=User.objects.none(),
        label="Usuario",
        widget=forms.Select(
            attrs={"class": "form-control"}
        ),
    )

    fecha_inicio = forms.DateTimeField(
        required=False,
        label="Inicio de acceso temporal",
        widget=forms.DateTimeInput(
            attrs={
                "type": "datetime-local",
                "class": "form-control",
            }
        ),
    )

    fecha_fin = forms.DateTimeField(
        required=False,
        label="Fin de acceso temporal",
        widget=forms.DateTimeInput(
            attrs={
                "type": "datetime-local",
                "class": "form-control",
            }
        ),
    )

    motivo = forms.CharField(
        label="Motivo",
        widget=forms.Textarea(
            attrs={
                "class": "form-control",
                "rows": 3,
                "placeholder": (
                    "Ej. Se incorpora un inspector adicional "
                    "para la parada."
                ),
            }
        ),
    )

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.fields["usuario"].queryset = (
            User.objects
            .filter(
                is_active=True,
                groups__name__in=[
                    "Inspector",
                    "Supervisor",
                    "Analista",
                    "Cliente",
                ],
            )
            .distinct()
            .order_by(
                "first_name",
                "last_name",
                "username",
            )
        )

    def clean(self):
        datos = super().clean()

        accion = datos.get("accion")
        rol = datos.get("rol")
        usuario = datos.get("usuario")
        fecha_inicio = datos.get("fecha_inicio")
        fecha_fin = datos.get("fecha_fin")

        if not usuario or not rol:
            return datos

        if not usuario.groups.filter(name=rol).exists():
            self.add_error(
                "usuario",
                (
                    "El usuario seleccionado no pertenece "
                    f"al rol {rol}."
                ),
            )

        if rol == "Cliente" and accion != "CAMBIAR_CLIENTE":
            self.add_error(
                "accion",
                (
                    "El Cliente es único. Para modificarlo "
                    "usa 'Cambiar cliente'."
                ),
            )

        if rol != "Cliente" and accion == "CAMBIAR_CLIENTE":
            self.add_error(
                "accion",
                (
                    "La acción 'Cambiar cliente' solamente "
                    "puede utilizarse con el rol Cliente."
                ),
            )

        try:
            perfil = usuario.perfil_sistema
        except PerfilUsuario.DoesNotExist:
            perfil = None

        if (
            accion in {"AGREGAR", "CAMBIAR_CLIENTE"}
            and perfil
            and perfil.tipo_vinculo
            == PerfilUsuario.TipoVinculo.INTERMITENTE
        ):
            if not fecha_inicio:
                self.add_error(
                    "fecha_inicio",
                    (
                        "El usuario intermitente necesita "
                        "fecha y hora de inicio."
                    ),
                )

            if not fecha_fin:
                self.add_error(
                    "fecha_fin",
                    (
                        "El usuario intermitente necesita "
                        "fecha y hora de fin."
                    ),
                )

        if (
            fecha_inicio
            and fecha_fin
            and fecha_fin <= fecha_inicio
        ):
            self.add_error(
                "fecha_fin",
                (
                    "La fecha de fin debe ser posterior "
                    "a la fecha de inicio."
                ),
            )

        return datos
