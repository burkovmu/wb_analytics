from django.core.wsgi import get_wsgi_application
import os
import sys

# Добавляем путь к проекту
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Устанавливаем настройки Django
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'wb_analytics_project.settings')

# Получаем WSGI приложение Django
django_app = get_wsgi_application()

def application(environ, start_response):
    """WSGI wrapper для Django"""
    try:
        return django_app(environ, start_response)
    except Exception as e:
        # Простая обработка ошибок
        status = '500 Internal Server Error'
        headers = [('Content-Type', 'text/plain')]
        start_response(status, headers)
        return [f'Error: {str(e)}'.encode()]