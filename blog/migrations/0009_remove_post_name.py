# Generated by Django 2.2.2 on 2019-06-27 15:23

from django.db import migrations


class Migration(migrations.Migration):

    dependencies = [
        ('blog', '0008_auto_20190627_1121'),
    ]

    operations = [
        migrations.RemoveField(
            model_name='post',
            name='name',
        ),
    ]
