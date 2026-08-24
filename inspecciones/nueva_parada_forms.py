from django import forms
from django.contrib.auth import get_user_model

from accounts.models import PerfilUsuario
from .models import Parada


User = get_user_model()


class NuevaParadaForm(forms.ModelForm):
    inspectores = forms.ModelMultipleChoiceField(
        queryset=User.objects.none(),
        label="Inspectores",
        widget=forms.CheckboxSelectMultiple(),
    )

    supervisores = forms.ModelMultipleChoiceField(
        queryset=User.objects.none(),
        label="Supervisores",
        widget=forms.CheckboxSelectMultiple(),
    )

    analistas = forms.ModelMultipleChoiceField(
        queryset=User.objects.none(),
        label="Analistas",
        widget=forms.CheckboxSelectMultiple(),
    )

    cliente = forms.ModelChoiceField(
        queryset=User.objects.none(),
        label="Cliente",
        empty_label="Selecciona un cliente",
    )

    acceso_inicio = forms.DateTimeField(
        required=False,
        label="Inicio de acceso para personal intermitente",
        widget=forms.DateTimeInput(
            attrs={
                "type": "datetime-local",
                "class": "form-control",
            }
        ),
    )

    acceso_fin = forms.DateTimeField(
        required=False,
        label="Fin de acceso para personal intermitente",
        widget=forms.DateTimeInput(
            attrs={
                "type": "datetime-local",
                "class": "form-control",
            }
        ),
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
                attrs={"class": "form-control"}
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
                    "rows": 3,
                    "class": "form-control",
                }
            ),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        roles = {
            "inspectores": "Inspector",
            "supervisores": "Supervisor",
            "analistas": "Analista",
            "cliente": "Cliente",
        }

        for campo, rol in roles.items():
            self.fields[campo].queryset = (
                User.objects
                .filter(
                    is_active=True,
                    groups__name=rol,
                )
                .distinct()
                .order_by(
                    "first_name",
                    "last_name",
                    "username",
                )
            )

    def clean(self):
        cleaned_data = super().clean()

        fecha_inicio = cleaned_data.get("fecha_inicio")
        fecha_fin = cleaned_data.get("fecha_fin")
        acceso_inicio = cleaned_data.get("acceso_inicio")
        acceso_fin = cleaned_data.get("acceso_fin")

        if (
            fecha_inicio
            and fecha_fin
            and fecha_fin < fecha_inicio
        ):
            self.add_error(
                "fecha_fin",
                "La fecha final no puede ser anterior a la fecha inicial.",
            )

        usuarios = []

        for campo in (
            "inspectores",
            "supervisores",
            "analistas",
        ):
            seleccionados = cleaned_data.get(campo)
            if seleccionados is not None:
                usuarios.extend(list(seleccionados))

        cliente = cleaned_data.get("cliente")
        if cliente:
            usuarios.append(cliente)

        existe_intermitente = False

        for usuario in usuarios:
            try:
                perfil = usuario.perfil_sistema
            except PerfilUsuario.DoesNotExist:
                # Compatibilidad con usuarios antiguos.
                continue

            if (
                perfil.tipo_vinculo
                == PerfilUsuario.TipoVinculo.INTERMITENTE
            ):
                existe_intermitente = True

            elif (
                perfil.tipo_vinculo
                == PerfilUsuario.TipoVinculo.CONTRATO
                and fecha_inicio
                and not perfil.contrato_vigente(fecha_inicio)
            ):
                self.add_error(
                    None,
                    (
                        f"El contrato de "
                        f"{usuario.get_full_name() or usuario.username} "
                        "no estará vigente en la fecha de la parada."
                    ),
                )

        if existe_intermitente:
            if not acceso_inicio:
                self.add_error(
                    "acceso_inicio",
                    (
                        "Debes indicar la fecha y hora de inicio "
                        "para el personal intermitente."
                    ),
                )

            if not acceso_fin:
                self.add_error(
                    "acceso_fin",
                    (
                        "Debes indicar la fecha y hora de fin "
                        "para el personal intermitente."
                    ),
                )

        if (
            acceso_inicio
            and acceso_fin
            and acceso_fin <= acceso_inicio
        ):
            self.add_error(
                "acceso_fin",
                (
                    "El fin del acceso debe ser posterior "
                    "al inicio del acceso."
                ),
            )

        return cleaned_data


class CrearIntermitenteParadaForm(forms.Form):
    ROLES = (
        ("Inspector", "Inspector"),
        ("Supervisor", "Supervisor"),
        ("Analista", "Analista"),
        ("Cliente", "Cliente"),
    )

    rol_intermitente = forms.ChoiceField(
        label="Rol",
        choices=ROLES,
    )

    nombres_intermitente = forms.CharField(
        label="Nombres",
        max_length=150,
        widget=forms.TextInput(
            attrs={
                "placeholder": "Nombres",
            }
        ),
    )

    apellidos_intermitente = forms.CharField(
        label="Apellidos",
        max_length=150,
        required=False,
        widget=forms.TextInput(
            attrs={
                "placeholder": "Apellidos",
            }
        ),
    )

    username_intermitente = forms.CharField(
        label="Usuario",
        max_length=150,
        widget=forms.TextInput(
            attrs={
                "placeholder": "Ej. juan.perez",
                "autocomplete": "off",
            }
        ),
    )

    email_intermitente = forms.EmailField(
        label="Correo",
        required=False,
        widget=forms.EmailInput(
            attrs={
                "placeholder": "Correo opcional",
            }
        ),
    )

    password1_intermitente = forms.CharField(
        label="Contraseña",
        widget=forms.PasswordInput(
            attrs={
                "placeholder": "Contraseña temporal",
                "autocomplete": "new-password",
            }
        ),
    )

    password2_intermitente = forms.CharField(
        label="Confirmar contraseña",
        widget=forms.PasswordInput(
            attrs={
                "placeholder": "Repite la contraseña",
                "autocomplete": "new-password",
            }
        ),
    )

    def clean_username_intermitente(self):
        username = (
            self.cleaned_data["username_intermitente"]
            .strip()
        )

        if User.objects.filter(
            username__iexact=username
        ).exists():
            raise forms.ValidationError(
                "Ya existe un usuario con este nombre."
            )

        return username

    def clean(self):
        datos = super().clean()

        password1 = datos.get("password1_intermitente")
        password2 = datos.get("password2_intermitente")

        if password1 or password2:
            if password1 != password2:
                self.add_error(
                    "password2_intermitente",
                    "Las contraseñas no coinciden.",
                )

            if password1 and len(password1) < 8:
                self.add_error(
                    "password1_intermitente",
                    (
                        "La contraseña debe tener "
                        "al menos 8 caracteres."
                    ),
                )

        return datos

    def crear_usuario(self):
        if not self.is_valid():
            raise ValueError(
                "El formulario intermitente no es válido."
            )

        from django.contrib.auth.models import Group

        rol = self.cleaned_data["rol_intermitente"]

        grupo = Group.objects.filter(
            name=rol
        ).first()

        if not grupo:
            raise ValueError(
                f"No existe el grupo {rol}."
            )

        usuario = User.objects.create_user(
            username=self.cleaned_data[
                "username_intermitente"
            ],
            email=self.cleaned_data.get(
                "email_intermitente",
                "",
            ),
            password=self.cleaned_data[
                "password1_intermitente"
            ],
            first_name=self.cleaned_data[
                "nombres_intermitente"
            ],
            last_name=self.cleaned_data.get(
                "apellidos_intermitente",
                "",
            ),
            is_active=True,
        )

        usuario.groups.add(grupo)

        PerfilUsuario.objects.create(
            usuario=usuario,
            tipo_vinculo=(
                PerfilUsuario.TipoVinculo.INTERMITENTE
            ),
        )

        return usuario
