from django import forms

class ValidacionForm(forms.Form):
    nombre = forms.CharField(
        label="Nombre completo",
        max_length=200,
        widget=forms.TextInput(attrs={'class': 'form-control'})
    )
    documento = forms.CharField(
        label="No. Identificación",
        max_length=50,
        widget=forms.TextInput(attrs={'class': 'form-control'})
    )
    numero_contacto = forms.CharField(
        label="Número de contacto",
        max_length=20,
        required=False,
        widget=forms.TextInput(attrs={'class': 'form-control'})
    )
    cargo = forms.CharField(
        label="Cargo",
        max_length=100,
        required=False,
        widget=forms.TextInput(attrs={'class': 'form-control'})
    )
    categoria_salarial = forms.CharField(
        label="Categoría salarial",
        max_length=100,
        required=False,
        widget=forms.TextInput(attrs={'class': 'form-control'})
    )
    salario_basico = forms.CharField(
        label="Salario básico mensual",
        max_length=50,
        required=False,
        widget=forms.TextInput(attrs={'class': 'form-control'})
    )
    beneficio_alimentacion = forms.CharField(
        label="Beneficio alimentación",
        max_length=50,
        required=False,
        widget=forms.TextInput(attrs={'class': 'form-control'})
    )
    beneficio_pension = forms.CharField(
        label="Beneficio pensión voluntaria",
        max_length=50,
        required=False,
        widget=forms.TextInput(attrs={'class': 'form-control'})
    )
    auxilio_localizacion = forms.CharField(
        label="Auxilio de localización",
        max_length=50,
        required=False,
        widget=forms.TextInput(attrs={'class': 'form-control'})
    )
    auxilio_vivienda = forms.CharField(
        label="Auxilio de vivienda",
        max_length=100,
        required=False,
        widget=forms.TextInput(attrs={'class': 'form-control'})
    )
    condicion_descanso = forms.CharField(
        label="Condición de descanso",
        max_length=50,
        required=False,
        widget=forms.TextInput(attrs={'class': 'form-control'})
    )
    tipo_contrato = forms.CharField(
        label="Tipo de contrato",
        max_length=100,
        required=False,
        widget=forms.TextInput(attrs={'class': 'form-control'})
    )
    fecha_inicio = forms.DateField(
        label="Fecha de inicio",
        required=False,
        widget=forms.DateInput(attrs={"type": "date", "class": "form-control"})
    )
    proyecto = forms.CharField(
        label="Proyecto / Centro de costos",
        max_length=200,
        required=False,
        widget=forms.TextInput(attrs={'class': 'form-control'})
    )
    aux_desplazamiento = forms.CharField(
        label="Auxilio de Desplazamiento",
        required=False,
        widget=forms.TextInput(attrs={'class': 'form-control'})
    )