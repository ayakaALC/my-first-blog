# Generated by Django 2.2.2 on 2019-07-09 19:43

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('blog', '0014_auto_20190709_1038'),
    ]

    operations = [
        migrations.RemoveField(
            model_name='post',
            name='cover',
        ),
        migrations.AddField(
            model_name='post',
            name='photo',
            field=models.ImageField(blank=True, upload_to='images/'),
        ),
    ]
