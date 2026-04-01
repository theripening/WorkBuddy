from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('tickets', '0011_add_cloudnote'),
    ]

    operations = [
        migrations.AddField(
            model_name='todoitem',
            name='cloud_id',
            field=models.IntegerField(blank=True, db_index=True, null=True),
        ),
        migrations.AddField(
            model_name='waitingon',
            name='cloud_id',
            field=models.IntegerField(blank=True, db_index=True, null=True),
        ),
    ]
