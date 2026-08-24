from django import forms
from django.contrib.auth.models import Group, User

from .models import PerfilUsuario


ROLES_PERMITIDOS = [
    "Administrador",
    "Inspector",
    "Supervisor",
    "Analista",
    "Cliente",
]


class UsuarioSistemaForm(forms.ModelForm):
    password1 = forms.CharField(
        label="Contraseña",
        required=False,
        widget=forms.PasswordInput(
            attrs={
                "placeholder": "Contraseña temporal",
                "autocomplete": "new-password",
            }
        ),
    )

    password2 = forms.CharField(
        label="Confirmar contraseña",
        required=False,
        widget=forms.PasswordInput(
            attrs={
                "placeholder": "Repite la contraseña",
                "autocomplete": "new-password",
            }
        ),
    )

    rol = forms.ModelChoiceField(
        label="Rol",
        queryset=Group.objects.none(),
        empty_label="Selecciona un rol",
    )

    tipo_vinculo = forms.ChoiceField(
        label="Tipo de vínculo",
        choices=PerfilUsuario.TipoVinculo.choices,
    )

    fecha_inicio_contrato = forms.DateField(
        label="Inicio de contrato",
        required=False,
        widget=forms.DateInput(
            attrs={"type": "date"}
        ),
    )

    fecha_fin_contrato = forms.DateField(
        label="Fin de contrato",
        required=False,
        widget=forms.DateInput(
            attrs={"type": "date"}
        ),
    )

    class Meta:
        model = User

        fields = [
            "first_name",
            "last_name",
            "username",
            "email",
            "is_active",
        ]

        labels = {
            "first_name": "Nombres",
            "last_name": "Apellidos",
            "username": "Usuario",
            "email": "Correo",
            "is_active": "Usuario activo",
        }

        widgets = {
            "first_name": forms.TextInput(
                attrs={"placeholder": "Nombres"}
            ),
            "last_name": forms.TextInput(
                attrs={"placeholder": "Apellidos"}
            ),
            "username": forms.TextInput(
                attrs={
                    "placeholder": "Ejemplo: maximo.sardon",
                    "autocomplete": "off",
                }
            ),
            "email": forms.EmailInput(
                attrs={"placeholder": "Correo opcional"}
            ),
        }

    def __init__(self, *args, **kwargs):
        self.usuario_editado = kwargs.get("instance")

        super().__init__(*args, **kwargs)

        self.fields["rol"].queryset = Group.objects.filter(
            name__in=ROLES_PERMITIDOS
        ).order_by("name")

        if self.usuario_editado and self.usuario_editado.pk:
            grupo = self.usuario_editado.groups.first()

            if grupo:
                self.fields["rol"].initial = grupo

            try:
                perfil = self.usuario_editado.perfil_sistema
            except PerfilUsuario.DoesNotExist:
                perfil = None

            if perfil:
                self.fields["tipo_vinculo"].initial = (
                    perfil.tipo_vinculo
                )
                self.fields["fecha_inicio_contrato"].initial = (
                    perfil.fecha_inicio_contrato
                )
                self.fields["fecha_fin_contrato"].initial = (
                    perfil.fecha_fin_contrato
                )
            else:
                # Compatibilidad con usuarios antiguos:
                # si aún no tienen perfil, se consideran permanentes
                # hasta que el administrador los clasifique.
                self.fields["tipo_vinculo"].initial = (
                    PerfilUsuario.TipoVinculo.PERMANENTE
                )
        else:
            self.fields["tipo_vinculo"].initial = (
                PerfilUsuario.TipoVinculo.CONTRATO
            )

        for field in self.fields.values():
            clase_actual = field.widget.attrs.get("class", "")
            field.widget.attrs["class"] = (
                f"{clase_actual} user-form-control"
            ).strip()

    def clean(self):
        datos = super().clean()

        password1 = datos.get("password1")
        password2 = datos.get("password2")
        tipo_vinculo = datos.get("tipo_vinculo")
        fecha_inicio = datos.get("fecha_inicio_contrato")
        fecha_fin = datos.get("fecha_fin_contrato")

        creando = not (
            self.usuario_editado
            and self.usuario_editado.pk
        )

        if creando and not password1:
            self.add_error(
                "password1",
                "La contraseña es obligatoria.",
            )

        if password1 or password2:
            if password1 != password2:
                self.add_error(
                    "password2",
                    "Las contraseñas no coinciden.",
                )

            if password1 and len(password1) < 8:
                self.add_error(
                    "password1",
                    "La contraseña debe tener al menos 8 caracteres.",
                )

        if (
            tipo_vinculo
            == PerfilUsuario.TipoVinculo.CONTRATO
        ):
            if not fecha_inicio:
                self.add_error(
                    "fecha_inicio_contrato",
                    "Indica la fecha de inicio del contrato.",
                )

            if not fecha_fin:
                self.add_error(
                    "fecha_fin_contrato",
                    "Indica la fecha de fin del contrato.",
                )

            if (
                fecha_inicio
                and fecha_fin
                and fecha_fin < fecha_inicio
            ):
                self.add_error(
                    "fecha_fin_contrato",
                    (
                        "La fecha de fin no puede ser "
                        "anterior a la fecha de inicio."
                    ),
                )

        return datos

    def save(self, commit=True):
        usuario = super().save(commit=False)

        password = self.cleaned_data.get("password1")

        if password:
            usuario.set_password(password)

        if commit:
            usuario.save()

            usuario.groups.clear()
            usuario.groups.add(
                self.cleaned_data["rol"]
            )

            tipo_vinculo = self.cleaned_data[
                "tipo_vinculo"
            ]

            fecha_inicio = self.cleaned_data.get(
                "fecha_inicio_contrato"
            )
            fecha_fin = self.cleaned_data.get(
                "fecha_fin_contrato"
            )

            if tipo_vinculo != PerfilUsuario.TipoVinculo.CONTRATO:
                fecha_inicio = None
                fecha_fin = None

            PerfilUsuario.objects.update_or_create(
                usuario=usuario,
                defaults={
                    "tipo_vinculo": tipo_vinculo,
                    "fecha_inicio_contrato": fecha_inicio,
                    "fecha_fin_contrato": fecha_fin,
                },
            )

        return usuario
