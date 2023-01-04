from django import forms
from django.forms import ModelForm
from myapp.models import MyModel


class MyModelForm(ModelForm):
    files = forms.FileField(widget=forms.ClearableFileInput(attrs={'multiple': True}))

    class Meta:
        model = MyModel
       # files = MultiFileField(min_num=1, max_num=3, max_file_size=1024*1024*5)
        fields = [
            'provider',
            'files',
            'user_name'
        ]