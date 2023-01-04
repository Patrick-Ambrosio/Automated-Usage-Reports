from django.db import models

# Create your models here.
PROVIDERS = (
    ('adobe','Adobe'),
    ('liveramp', 'LiveRamp'),
    ('neustar', 'Neustar'),
    ('oracle blue kai', 'Oracle Blue Kai'),
    ('salesforce (krux)', 'Salesforce (Krux)'),
    ('eyeota', 'Eyeota'),
    ('fyllo', 'Fyllo'),
    ('icx', 'ICX'),
    ('dstillery', 'DStillery'),
    ('comscoretv', 'Comscore TV'),
    ('comscorepa', 'Comscore PA'),
    ('nielsen', 'Nielsen')
    
)

class MyModel(models.Model):
    files = models.FileField(upload_to='attachments')
    user_name = models.CharField(max_length = 50)
    provider = models.CharField(
        max_length = 50, 
        choices=PROVIDERS
    )