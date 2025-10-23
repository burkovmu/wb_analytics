import os
import sys
from django.core.wsgi import get_wsgi_application

# Добавляем путь к проекту
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Устанавливаем настройки Django
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'wb_analytics_project.settings')

# Получаем WSGI приложение
application = get_wsgi_application()

# Экспортируем для Vercel
handler = application
