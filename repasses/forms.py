from django import forms


class ImportarMedPlusForm(forms.Form):
    arquivo = forms.FileField(
        label='Relatório da MedPlus',
        help_text='Arquivo .xls (ou .xlsx) exportado em "Procedimentos pela Agenda".',
        widget=forms.ClearableFileInput(attrs={'accept': '.xls,.xlsx'}),
    )

    def clean_arquivo(self):
        arquivo = self.cleaned_data['arquivo']
        nome = (arquivo.name or '').lower()
        if not nome.endswith(('.xls', '.xlsx')):
            raise forms.ValidationError('Envie um arquivo Excel (.xls ou .xlsx) exportado da MedPlus.')
        return arquivo
