from decimal import Decimal, ROUND_HALF_UP
from django.db.models import Sum, Count, Avg
from analytics.models import (
    SalesAnalysis, FinancialReport, Nomenclature, 
    StockBalance, SummaryData, PurchasePlan
)
from analytics.excel_formulas import ExcelFormulasService


class CalculationService:
    """Сервис для выполнения всех расчетов из Excel файла"""
    
    @staticmethod
    def safe_divide(numerator, denominator, default=0):
        """Безопасное деление с обработкой ошибок"""
        if denominator == 0 or denominator is None:
            return Decimal(str(default))
        return Decimal(str(numerator)) / Decimal(str(denominator))
    
    @staticmethod
    def calculate_sales_analysis():
        """Пересчет всех показателей анализа продаж"""
        # Получаем все записи анализа продаж
        sales_records = SalesAnalysis.objects.all()
        
        for record in sales_records:
            # Расчет процента выкупа
            if record.orders > 0:
                record.purchase_percentage = CalculationService.safe_divide(
                    record.sales, record.orders
                )
            
            # Расчет продаж минус возвраты
            record.sales_minus_returns = record.sales - record.returns
            
            # Расчет среднего чека
            if record.sales_minus_returns > 0:
                record.average_check = CalculationService.safe_divide(
                    record.sales_with_spp, record.sales_minus_returns
                )
            
            # Расчет комиссии на единицу
            if record.sales_minus_returns > 0:
                record.commission_per_unit = CalculationService.safe_divide(
                    record.commission, record.sales_minus_returns
                )
            
            # Расчет логистики на единицу
            if record.sales_minus_returns > 0:
                record.logistics_per_unit = CalculationService.safe_divide(
                    record.logistics, record.sales_minus_returns
                )
            
            # Расчет себестоимости проданного товара
            record.sold_goods_cost = record.cost_per_unit * record.sales_minus_returns
            
            # Расчет маржи до налога
            record.margin_before_tax = record.sales_minus_commission - record.sold_goods_cost
            
            # Расчет налога 6%
            record.tax_6_percent = record.margin_before_tax * Decimal('0.06')
            
            # Расчет маржи после налога
            record.margin_after_tax = record.margin_before_tax - record.tax_6_percent
            
            # Расчет маржи на единицу
            if record.sales_minus_returns > 0:
                record.margin_per_unit = CalculationService.safe_divide(
                    record.margin_after_tax, record.sales_minus_returns
                )
            
            # Расчет маржинальности
            if record.sales_minus_commission > 0:
                record.margin_percent = CalculationService.safe_divide(
                    record.margin_after_tax, record.sales_minus_commission
                )
            
            # Расчет ROI от себестоимости
            if record.sold_goods_cost > 0:
                record.roi_from_cost = CalculationService.safe_divide(
                    record.margin_after_tax, record.sold_goods_cost
                )
            
            # Расчет GMROI
            record.gmroi = record.roi_from_cost
            
            # Расчет доли в выручке
            total_revenue = SalesAnalysis.objects.aggregate(
                total=Sum('sales_minus_commission')
            )['total'] or 0
            
            if total_revenue > 0:
                record.revenue_share = CalculationService.safe_divide(
                    record.sales_minus_commission, total_revenue
                )
            
            # Расчет доли в марже
            total_margin = SalesAnalysis.objects.aggregate(
                total=Sum('margin_after_tax')
            )['total'] or 0
            
            if total_margin > 0:
                record.margin_share = CalculationService.safe_divide(
                    record.margin_after_tax, total_margin
                )
            
            # Расчет денег в товаре
            record.money_in_goods = record.cost_per_unit * record.total_stock_wb
            
            record.save()
    
    @staticmethod
    def calculate_summary_data():
        """Расчет сводных данных используя ВСЕ формулы из Excel файла"""
        
        # Используем новые формулы из Excel файла
        all_formulas = ExcelFormulasService.calculate_all_formulas()
        
        # Получаем данные из формул листа "Сводный"
        svodny = all_formulas['svodny']
        
        # Получаем данные из формул листа "Анализ продаж"
        analysis = all_formulas['analysis']
        
        # Создаем или обновляем запись сводных данных
        summary_data, created = SummaryData.objects.get_or_create(
            id=1,
            defaults={
                'sales_minus_returns_count': svodny['c3'],  # C3: ='Анализ продаж'!O3
                'purchase_percentage': analysis['p3'],  # P3: =IFERROR(O3/K3, )
                'average_check_after_spp': analysis['z3'],  # Z3: =IFERROR(Y3/S3, )
                'commission_per_unit': svodny['c5'],  # C5: =IFERROR(F17/C3, )
                'logistics_per_unit': svodny['c6'],  # C6: =IFERROR(F3/C3, )
                'storage_per_unit': Decimal('0'),  # Пока нет данных о хранении
                'sold_goods_cost': svodny['f20'],  # F20: ='Анализ продаж'!AL3
                'average_cost_per_unit': svodny['c7'],  # C7: =IFERROR(F4/C3, )
                'average_margin_per_unit': svodny['c8'],  # C8: =IFERROR(F5/C3, )
                'margin_without_ads': svodny['c9'],  # C9: ='Анализ продаж'!AJ3
                'margin_per_unit_without_ads': svodny['c10'],  # C10: =IFERROR(C9/C3, )
                'margin_after_all_expenses': svodny['c14'],  # C14: =F19-C9-F20
                'total_commission': svodny['f17'],  # F17: ='Анализ продаж'!T3
                'commission_percent': svodny['c16'],  # C16: =IFERROR(C14/F17, )
                'total_logistics': svodny['f18'],  # F18: ='Анализ продаж'!U3+F12+F13
                'logistics_percent': svodny['c17'],  # C17: =IFERROR( C14/F18, )
                'total_storage': Decimal('0'),
                'storage_percent': Decimal('0'),
                'additional_payments': svodny['f13'],  # F13: ='Анализ продаж'!AF3
                'fines': Decimal('0'),  # Пока нет данных о штрафах
                'paid_reception': Decimal('0'),  # Пока нет данных о платной приемке
                'advertising_deductions': Decimal('0'),  # Пока нет данных о рекламе
                'advertising_percent': Decimal('0'),
                'transit_deductions': Decimal('0'),  # Пока нет данных о транзите
                'acquiring': svodny['f11'],  # F11: ='Анализ продаж'!AC3
                'defect_compensation': svodny['f12'],  # F12: Сложная формула с SUMIFS
                'substitution_compensation': Decimal('0'),  # Пока нет данных о замене
                'margin_percent_before_spp': svodny['c6'],  # C6: =IFERROR(F3/C3, )
                'margin_percent_with_spp': svodny['c7'],  # C7: =IFERROR(F4/C3, )
                'margin_percent_after_commission': svodny['c8'],  # C8: =IFERROR(F5/C3, )
                'sales_before_spp': svodny['f3'],  # F3: ='Анализ продаж'!Y3
                'sales_with_spp': svodny['c4'],  # C4: ='Анализ продаж'!P3
                'sales_minus_commission': svodny['f16'],  # F16: ='Анализ продаж'!S3
                'payment_to_account': svodny['f19'],  # F19: =F16-F3-F4-F5+F6-F7-F8-F9-F10+F12+F13
                'tax_6_percent': svodny['c14'] * Decimal('0.06'),  # НДС 6% от маржи
            }
        )
        
        if not created:
            # Обновляем существующую запись
            summary_data.sales_minus_returns_count = svodny['c3']
            summary_data.purchase_percentage = analysis['p3']
            summary_data.average_check_after_spp = analysis['z3']
            summary_data.commission_per_unit = svodny['c5']
            summary_data.logistics_per_unit = svodny['c6']
            summary_data.storage_per_unit = Decimal('0')
            summary_data.sold_goods_cost = svodny['f20']
            summary_data.average_cost_per_unit = svodny['c7']
            summary_data.average_margin_per_unit = svodny['c8']
            summary_data.margin_without_ads = svodny['c9']
            summary_data.margin_per_unit_without_ads = svodny['c10']
            summary_data.margin_after_all_expenses = svodny['c14']
            summary_data.total_commission = svodny['f17']
            summary_data.commission_percent = svodny['c16']
            summary_data.total_logistics = svodny['f18']
            summary_data.logistics_percent = svodny['c17']
            summary_data.total_storage = Decimal('0')
            summary_data.storage_percent = Decimal('0')
            summary_data.additional_payments = svodny['f13']
            summary_data.fines = Decimal('0')
            summary_data.paid_reception = Decimal('0')
            summary_data.advertising_deductions = Decimal('0')
            summary_data.advertising_percent = Decimal('0')
            summary_data.transit_deductions = Decimal('0')
            summary_data.acquiring = svodny['f11']
            summary_data.defect_compensation = svodny['f12']
            summary_data.substitution_compensation = Decimal('0')
            summary_data.margin_percent_before_spp = svodny['c6']
            summary_data.margin_percent_with_spp = svodny['c7']
            summary_data.margin_percent_after_commission = svodny['c8']
            summary_data.sales_before_spp = svodny['f3']
            summary_data.sales_with_spp = svodny['c4']
            summary_data.sales_minus_commission = svodny['f16']
            summary_data.payment_to_account = svodny['f19']
            summary_data.tax_6_percent = svodny['c14'] * Decimal('0.06')
            summary_data.save()
        
        return summary_data
    
    @staticmethod
    def recalculate_all():
        """Пересчет всех данных"""
        CalculationService.calculate_sales_analysis()
        CalculationService.calculate_summary_data()