"""
WSGI entrypoint for production (e.g. gunicorn wsgi:app or gunicorn wsgi:application).
"""
from app.app import app

application = app
