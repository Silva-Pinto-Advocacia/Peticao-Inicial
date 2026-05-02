web: gunicorn app:app --worker-class gthread --workers 1 --threads 4 --timeout 600 --graceful-timeout 600 --bind 0.0.0.0:$PORT
