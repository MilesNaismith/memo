from django.db import models

# Create your models here.
class Company(models.Model):
    company_name = models.CharField(max_length=200)

class Budget(models.Model):
    budget_item = models.CharField(max_length=200)

    