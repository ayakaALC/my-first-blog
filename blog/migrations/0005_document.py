# Generated by Django 2.2.2 on 2019-06-26 13:30

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('blog', '0004_auto_20190613_0903'),
    ]

    operations = [
        migrations.CreateModel(
            name='Document',
            fields=[
                ('id', models.AutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('docfile', models.FileField(upload_to='documents/%Y/%m/%d')),
            ],
        ),
    ]
