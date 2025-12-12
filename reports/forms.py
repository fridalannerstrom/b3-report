from django import forms
from django.core.exceptions import ValidationError

class ExcelUploadForm(forms.Form):
    file = forms.FileField(label="Ladda upp testresultat (Excel)")

    def clean_file(self):
        f = self.cleaned_data["file"]
        if not f.name.lower().endswith((".xlsx", ".xls")):
            raise ValidationError("Endast Excel-filer (.xlsx eller .xls) är tillåtna.")
        return f