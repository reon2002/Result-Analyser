from django.db import models

# Create your models here.
# class Result(models.Model):
#     id=models.CharField(primary_key=True)
#     score=models.IntegerField(default=0)
class MyModel(models.Model):
    file=models.FileField(upload_to='uploads/')

class Student(models.Model):
    stu_id = models.CharField(max_length=50)
    sem = models.CharField(max_length=50)
    sgpa = models.FloatField()

    class Meta:
        unique_together = [['stu_id', 'sem']]