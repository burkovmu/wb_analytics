from rest_framework import serializers
from analytics.models import (
    SalesAnalysis, FinancialReport, Nomenclature, 
    StockBalance, SummaryData, PurchasePlan
)


class NomenclatureSerializer(serializers.ModelSerializer):
    class Meta:
        model = Nomenclature
        fields = '__all__'


class FinancialReportSerializer(serializers.ModelSerializer):
    class Meta:
        model = FinancialReport
        fields = '__all__'


class StockBalanceSerializer(serializers.ModelSerializer):
    class Meta:
        model = StockBalance
        fields = '__all__'


class SalesAnalysisSerializer(serializers.ModelSerializer):
    class Meta:
        model = SalesAnalysis
        fields = '__all__'


class SummaryDataSerializer(serializers.ModelSerializer):
    class Meta:
        model = SummaryData
        fields = '__all__'


class PurchasePlanSerializer(serializers.ModelSerializer):
    class Meta:
        model = PurchasePlan
        fields = '__all__'
