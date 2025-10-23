from django.urls import path, include
from rest_framework.routers import DefaultRouter
from analytics.views import (
    SalesAnalysisViewSet, FinancialReportViewSet, 
    NomenclatureViewSet, StockBalanceViewSet,
    SummaryDataViewSet, PurchasePlanViewSet, DashboardViewSet
)
from analytics.dashboard_views import (
    dashboard_view, dashboard_overview, dashboard_charts_data
)

router = DefaultRouter()
router.register(r'sales-analysis', SalesAnalysisViewSet)
router.register(r'financial-reports', FinancialReportViewSet)
router.register(r'nomenclature', NomenclatureViewSet)
router.register(r'stock-balance', StockBalanceViewSet)
router.register(r'summary', SummaryDataViewSet)
router.register(r'purchase-plan', PurchasePlanViewSet)
router.register(r'dashboard', DashboardViewSet, basename='dashboard')

urlpatterns = [
    path('', dashboard_view, name='dashboard'),
    path('api/', include(router.urls)),
    path('api/dashboard/overview/', dashboard_overview, name='dashboard_overview'),
    path('api/dashboard/charts_data/', dashboard_charts_data, name='dashboard_charts'),
]
