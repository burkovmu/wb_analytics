from rest_framework import viewsets, status
from rest_framework.decorators import action
from rest_framework.response import Response
from django.shortcuts import get_object_or_404
from django.db.models import Sum, Avg
from django.http import JsonResponse
import json
import os
from django.conf import settings
from analytics.models import (
    SalesAnalysis, FinancialReport, Nomenclature, 
    StockBalance, SummaryData, PurchasePlan
)
from analytics.services import CalculationService
from analytics.serializers import (
    SalesAnalysisSerializer, FinancialReportSerializer, 
    NomenclatureSerializer, StockBalanceSerializer,
    SummaryDataSerializer, PurchasePlanSerializer
)


def load_static_data(filename):
    """Загрузка статических данных из JSON файлов"""
    try:
        static_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'static', 'data', filename)
        with open(static_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"Ошибка загрузки {filename}: {e}")
        return []


class SalesAnalysisViewSet(viewsets.ModelViewSet):
    """API для анализа продаж"""
    queryset = SalesAnalysis.objects.all()
    serializer_class = SalesAnalysisSerializer
    
    def list(self, request):
        """Возвращает статические данные анализа продаж"""
        data = load_static_data('sales_analysis.json')
        return Response(data)
    
    @action(detail=False, methods=['post'])
    def recalculate(self, request):
        """Пересчет всех показателей анализа продаж"""
        return Response({'message': 'Расчеты обновлены (статический режим)'}, status=status.HTTP_200_OK)
    
    @action(detail=False, methods=['get'])
    def summary(self, request):
        """Получение сводных данных"""
        data = load_static_data('summary.json')
        return Response(data[0] if data else {})


class FinancialReportViewSet(viewsets.ModelViewSet):
    """API для финансовых отчетов"""
    queryset = FinancialReport.objects.all()
    serializer_class = FinancialReportSerializer
    
    def list(self, request):
        """Возвращает статические данные финансовых отчетов"""
        data = load_static_data('financial_reports.json')
        return Response(data)
    
    @action(detail=False, methods=['post'])
    def bulk_import(self, request):
        """Массовый импорт данных из Excel"""
        return Response({
            'message': 'Импорт недоступен в статическом режиме'
        }, status=status.HTTP_400_BAD_REQUEST)


class NomenclatureViewSet(viewsets.ModelViewSet):
    """API для номенклатуры"""
    queryset = Nomenclature.objects.all()
    serializer_class = NomenclatureSerializer
    
    def list(self, request):
        """Возвращает статические данные номенклатуры"""
        data = load_static_data('nomenclature.json')
        return Response(data)
    
    @action(detail=False, methods=['get'])
    def by_brand(self, request):
        """Получение номенклатуры по бренду"""
        brand = request.query_params.get('brand')
        data = load_static_data('nomenclature.json')
        if brand:
            filtered_data = [item for item in data if brand.lower() in item.get('brand', '').lower()]
            return Response(filtered_data)
        return Response(data)


class StockBalanceViewSet(viewsets.ModelViewSet):
    """API для остатков на складах"""
    queryset = StockBalance.objects.all()
    serializer_class = StockBalanceSerializer
    
    def list(self, request):
        """Возвращает статические данные остатков"""
        data = load_static_data('stock_balance.json')
        return Response(data)


class SummaryDataViewSet(viewsets.ReadOnlyModelViewSet):
    """API для сводных данных (только чтение)"""
    queryset = SummaryData.objects.all()
    serializer_class = SummaryDataSerializer
    
    def list(self, request):
        """Возвращает статические сводные данные"""
        data = load_static_data('summary.json')
        return Response(data)
    
    @action(detail=False, methods=['post'])
    def refresh(self, request):
        """Обновление сводных данных"""
        data = load_static_data('summary.json')
        return Response(data[0] if data else {})


class PurchasePlanViewSet(viewsets.ModelViewSet):
    """API для планов по выкупам"""
    queryset = PurchasePlan.objects.all()
    serializer_class = PurchasePlanSerializer


class DashboardViewSet(viewsets.ViewSet):
    """API для дашборда"""
    
    @action(detail=False, methods=['get'])
    def overview(self, request):
        """Обзор основных показателей"""
        try:
            data = load_static_data('dashboard_overview.json')
            return Response(data)
        except Exception as e:
            return Response({'error': str(e)}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)
    
    @action(detail=False, methods=['get'])
    def charts_data(self, request):
        """Данные для графиков"""
        try:
            data = load_static_data('charts_data.json')
            return Response(data)
        except Exception as e:
            return Response({'error': str(e)}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)