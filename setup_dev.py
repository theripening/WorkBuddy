"""
Dev setup script — creates superuser and sample assignees.
Run with: python setup_dev.py
"""
import os
import django

os.environ.setdefault("DJANGO_SETTINGS_MODULE", "workbuddy.settings")
django.setup()

from django.contrib.auth import get_user_model
from tickets.models import Assignee

User = get_user_model()

# --- Superuser ---
if not User.objects.filter(username="admin").exists():
    User.objects.create_superuser(username="admin", password="admin", email="admin@example.com")
    print("Created superuser: admin / admin")
else:
    print("Superuser 'admin' already exists, skipping.")

# --- Assignees ---
assignees = [
    {"name": "Jane Smith", "email": "jane.smith@example.com"},
    {"name": "Bob Johnson", "email": "bob.johnson@example.com"},
]

for data in assignees:
    obj, created = Assignee.objects.get_or_create(email=data["email"], defaults={"name": data["name"]})
    if created:
        print(f"Created assignee: {obj.name} <{obj.email}>")
    else:
        print(f"Assignee '{obj.email}' already exists, skipping.")

print("Done.")
