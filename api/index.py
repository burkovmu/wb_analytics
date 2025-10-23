from django.http import HttpResponse
from django.core.wsgi import get_wsgi_application
import os
import sys

# Добавляем путь к проекту
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Устанавливаем настройки Django
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'wb_analytics_project.settings')

# Получаем WSGI приложение
application = get_wsgi_application()

def handler(request):
    """Обработчик для Vercel"""
    return application(request.environ, lambda *args: None)
