# Generated by Django 4.2.1 on 2023-06-04 09:47

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('result_app', '0003_delete_demo_result_reon'),
    ]

    operations = [
        migrations.RemoveField(
            model_name='result',
            name='reon',
        ),
        migrations.RemoveField(
            model_name='result',
            name='score',
        ),
        migrations.AddField(
            model_name='result',
            name='c1',
            field=models.CharField(default=0),
        ),
        migrations.AddField(
            model_name='result',
            name='c2',
            field=models.CharField(default=49),
        ),
        migrations.AddField(
            model_name='result',
            name='c3',
            field=models.CharField(default=0),
        ),
        migrations.AddField(
            model_name='result',
            name='c4',
            field=models.CharField(default=49),
        ),
        migrations.AddField(
            model_name='result',
            name='c5',
            field=models.CharField(default=0),
        ),
        migrations.AddField(
            model_name='result',
            name='c6',
            field=models.CharField(default=49),
        ),
        migrations.AddField(
            model_name='result',
            name='c7',
            field=models.CharField(default=0),
        ),
        migrations.AddField(
            model_name='result',
            name='c8',
            field=models.CharField(default=49),
        ),
    ]