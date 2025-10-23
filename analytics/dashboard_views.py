from django.shortcuts import render
from django.http import JsonResponse
from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.http import require_http_methods
from django.db.models import Sum, Avg
from analytics.services import CalculationService
from analytics.models import SalesAnalysis, FinancialReport
from analytics.serializers import SalesAnalysisSerializer


def dashboard_view(request):
    """Главная страница дашборда"""
    return render(request, 'analytics/dashboard.html')


@csrf_exempt
@require_http_methods(["GET"])
def dashboard_overview(request):
    """API для получения обзорных данных дашборда"""
    try:
        # Получаем сводные данные
        summary = CalculationService.calculate_summary_data()
        
        # Получаем топ товары по марже
        top_products = SalesAnalysis.objects.filter(
            margin_after_tax__gt=0
        ).order_by('-margin_after_tax')[:10]
        
        # Статистика по брендам
        brand_stats = SalesAnalysis.objects.values('brand').annotate(
            total_sales=Sum('sales_minus_returns'),
            total_margin=Sum('margin_after_tax'),
            avg_margin_percent=Avg('margin_percent')
        ).order_by('-total_margin')[:5]
        
        return JsonResponse({
            'summary': {
                'sales_minus_returns_count': summary.sales_minus_returns_count,
                'purchase_percentage': float(summary.purchase_percentage),
                'average_check_after_spp': float(summary.average_check_after_spp),
                'margin_after_all_expenses': float(summary.margin_after_all_expenses),
                'margin_percent_with_spp': float(summary.margin_percent_with_spp),
            },
            'top_products': SalesAnalysisSerializer(top_products, many=True).data,
            'brand_stats': list(brand_stats)
        })
    except Exception as e:
        return JsonResponse({'error': str(e)}, status=500)


@csrf_exempt
@require_http_methods(["GET"])
def dashboard_charts_data(request):
    """API для получения данных для графиков"""
    try:
        # Данные для графика продаж по дням
        sales_by_day = FinancialReport.objects.extra(
            select={'day': 'DATE(sale_date)'}
        ).values('day').annotate(
            total_sales=Sum('wb_sold_product'),
            total_quantity=Sum('quantity')
        ).order_by('day')
        
        # Данные для графика маржинальности по товарам
        margin_data = SalesAnalysis.objects.filter(
            margin_percent__gt=0
        ).values('brand', 'article').annotate(
            margin_percent=Avg('margin_percent'),
            total_margin=Sum('margin_after_tax')
        ).order_by('-total_margin')[:20]
        
        return JsonResponse({
            'sales_by_day': list(sales_by_day),
            'margin_by_product': list(margin_data)
        })
    except Exception as e:
        return JsonResponse({'error': str(e)}, status=500)
