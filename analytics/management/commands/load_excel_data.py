from django.core.management.base import BaseCommand
import pandas as pd
from analytics.models import (
    SalesAnalysis, FinancialReport, Nomenclature, 
    StockBalance, SummaryData, PurchasePlan
)
from analytics.services import CalculationService


class Command(BaseCommand):
    help = 'Загрузка данных из Excel файла'

    def add_arguments(self, parser):
        parser.add_argument('file_path', type=str, help='Путь к Excel файлу')

    def handle(self, *args, **options):
        file_path = options['file_path']
        
        try:
            self.stdout.write('Начинаем загрузку данных из Excel файла...')
            
            # Загружаем данные из листа "Номенклатуры (вставить)"
            self.load_nomenclature(file_path)
            
            # Загружаем данные из листа "Финотчеты (вставить)"
            self.load_financial_reports(file_path)
            
            # Загружаем данные из листа "Анализ продаж"
            self.load_sales_analysis(file_path)
            
            # Загружаем данные из листа "План по выкупам"
            self.load_purchase_plan(file_path)
            
            # Пересчитываем все показатели
            self.stdout.write('Пересчитываем все показатели...')
            CalculationService.recalculate_all()
            
            self.stdout.write(
                self.style.SUCCESS('Данные успешно загружены и обработаны!')
            )
            
        except Exception as e:
            self.stdout.write(
                self.style.ERROR(f'Ошибка при загрузке данных: {str(e)}')
            )

    def load_nomenclature(self, file_path):
        """Загрузка номенклатуры"""
        self.stdout.write('Загружаем номенклатуру...')
        
        df = pd.read_excel(file_path, sheet_name='Номенклатуры (вставить)', header=1)
        
        for _, row in df.iterrows():
            if pd.notna(row['Бренд']) and pd.notna(row['Артикул продавца']):
                Nomenclature.objects.get_or_create(
                    brand=row['Бренд'],
                    supplier_article=row['Артикул продавца'],
                    size=row['Размер'] if pd.notna(row['Размер']) else '0',
                    defaults={
                        'subject': row['Предмет'] if pd.notna(row['Предмет']) else '',
                        'size_code': row['Код размера (chrt_id)'] if pd.notna(row['Код размера (chrt_id)']) else '',
                        'wb_article': row['Артикул WB'] if pd.notna(row['Артикул WB']) else '',
                        'barcode': row['Баркод'] if pd.notna(row['Баркод']) else '',
                        'equipment': row['Комплектация'] if pd.notna(row['Комплектация']) else 0,
                        'composition': row['Состав'] if pd.notna(row['Состав']) else '',
                        'cost_price': row['Себестоимость'] if pd.notna(row['Себестоимость']) else 0,
                    }
                )
        
        self.stdout.write(f'Загружено {len(df)} записей номенклатуры')

    def load_financial_reports(self, file_path):
        """Загрузка финансовых отчетов"""
        self.stdout.write('Загружаем финансовые отчеты...')
        
        df = pd.read_excel(file_path, sheet_name='Финотчеты (вставить)', header=1)
        
        for _, row in df.iterrows():
            if pd.notna(row['№']) and pd.notna(row['Бренд']):
                FinancialReport.objects.get_or_create(
                    number=int(row['№']),
                    defaults={
                        'delivery_number': row['Номер поставки'] if pd.notna(row['Номер поставки']) else '',
                        'subject': row['Предмет'] if pd.notna(row['Предмет']) else '',
                        'nomenclature_code': row['Код номенклатуры'] if pd.notna(row['Код номенклатуры']) else '',
                        'brand': row['Бренд'] if pd.notna(row['Бренд']) else '',
                        'supplier_article': row['Артикул поставщика'] if pd.notna(row['Артикул поставщика']) else '',
                        'name': row['Название'] if pd.notna(row['Название']) else '',
                        'size': row['Размер'] if pd.notna(row['Размер']) else '',
                        'barcode': row['Баркод'] if pd.notna(row['Баркод']) else '',
                        'document_type': row['Тип документа'] if pd.notna(row['Тип документа']) else '',
                        'payment_basis': row['Обоснование для оплаты'] if pd.notna(row['Обоснование для оплаты']) else '',
                        'order_date': row['Дата заказа покупателем'] if pd.notna(row['Дата заказа покупателем']) else None,
                        'sale_date': row['Дата продажи'] if pd.notna(row['Дата продажи']) else None,
                        'quantity': row['Кол-во'] if pd.notna(row['Кол-во']) else 0,
                        'retail_price': row['Цена розничная'] if pd.notna(row['Цена розничная']) else 0,
                        'wb_sold_product': row['Вайлдберриз реализовал Товар (Пр)'] if pd.notna(row['Вайлдберриз реализовал Товар (Пр)']) else 0,
                        'storage': row['Хранение'] if pd.notna(row['Хранение']) else 0,
                        'deductions': row['Удержания'] if pd.notna(row['Удержания']) else 0,
                    }
                )
        
        self.stdout.write(f'Загружено {len(df)} записей финансовых отчетов')

    def load_sales_analysis(self, file_path):
        """Загрузка анализа продаж"""
        self.stdout.write('Загружаем анализ продаж...')
        
        df = pd.read_excel(file_path, sheet_name='Анализ продаж', header=1)
        
        for _, row in df.iterrows():
            if pd.notna(row['Бренд']) and pd.notna(row['Артикул']):
                SalesAnalysis.objects.get_or_create(
                    brand=row['Бренд'],
                    article=row['Артикул'],
                    size=row['Размер'] if pd.notna(row['Размер']) else '0',
                    defaults={
                        'subject': row['Предмет'] if pd.notna(row['Предмет']) else '',
                        'barcode': row['Баркод'] if pd.notna(row['Баркод']) else '',
                        'in_transit_to_client': row['В пути до клиента'] if pd.notna(row['В пути до клиента']) else 0,
                        'in_transit_from_client': row['В пути от клиента'] if pd.notna(row['В пути от клиента']) else 0,
                        'in_warehouses': row['На складах'] if pd.notna(row['На складах']) else 0,
                        'total_stock_wb': row['ИТОГО остаток на ВБ'] if pd.notna(row['ИТОГО остаток на ВБ']) else 0,
                        'orders': row['Заказы, шт'] if pd.notna(row['Заказы, шт']) else 0,
                        'rejections': row['Отказы'] if pd.notna(row['Отказы']) else 0,
                        'sales': row['Продажи, шт'] if pd.notna(row['Продажи, шт']) else 0,
                        'returns': row['Возвраты, шт'] if pd.notna(row['Возвраты, шт']) else 0,
                        'sales_minus_returns': row['Продажи минус возвраты'] if pd.notna(row['Продажи минус возвраты']) else 0,
                        'purchase_percentage': row['Процент выкупа'] if pd.notna(row['Процент выкупа']) else 0,
                        'sales_before_spp': row['Продажи по ценам до СПП'] if pd.notna(row['Продажи по ценам до СПП']) else 0,
                        'sales_with_spp': row['Продажи по ценам с СПП'] if pd.notna(row['Продажи по ценам с СПП']) else 0,
                        'sales_minus_commission': row['Продажи за вычетом комиссии'] if pd.notna(row['Продажи за вычетом комиссии']) else 0,
                        'commission': row['Комиссия, руб'] if pd.notna(row['Комиссия, руб']) else 0,
                        'commission_percent': row['Комиссия %'] if pd.notna(row['Комиссия %']) else 0,
                        'logistics': row['Логистика '] if pd.notna(row['Логистика ']) else 0,
                        'logistics_per_unit': row['Логистика на 1 продажу'] if pd.notna(row['Логистика на 1 продажу']) else 0,
                        'acquiring': row['Эквайринг'] if pd.notna(row['Эквайринг']) else 0,
                        'fine': row['Штраф'] if pd.notna(row['Штраф']) else 0,
                        'additional_payments': row['Доплаты'] if pd.notna(row['Доплаты']) else 0,
                        'substitution_compensation': row['Компенсация подмен'] if pd.notna(row['Компенсация подмен']) else 0,
                        'defect_compensation': row['Возмещение брака'] if pd.notna(row['Возмещение брака']) else 0,
                        'average_check': row['Средний чек '] if pd.notna(row['Средний чек ']) else 0,
                        'cost_per_unit': row['Себестомость 1 шт'] if pd.notna(row['Себестомость 1 шт']) else 0,
                        'sold_goods_cost': row['Себестоимость проданного товара'] if pd.notna(row['Себестоимость проданного товара']) else 0,
                        'margin_before_tax': row['Маржа до налогов'] if pd.notna(row['Маржа до налогов']) else 0,
                        'tax_6_percent': row['Налог 6%'] if pd.notna(row['Налог 6%']) else 0,
                        'margin_after_tax': row['Маржа после налогов, руб'] if pd.notna(row['Маржа после налогов, руб']) else 0,
                        'margin_per_unit': row['Маржа на 1 продажу, руб'] if pd.notna(row['Маржа на 1 продажу, руб']) else 0,
                        'margin_percent': row['Маржинальность, %'] if pd.notna(row['Маржинальность, %']) else 0,
                        'roi_from_cost': row['ROI от себестоимости, %'] if pd.notna(row['ROI от себестоимости, %']) else 0,
                        'gmroi': row['GMROI, %'] if pd.notna(row['GMROI, %']) else 0,
                        'revenue_share': row['Доля от общей выручки, %'] if pd.notna(row['Доля от общей выручки, %']) else 0,
                        'margin_share': row['Доля от общей маржи, %'] if pd.notna(row['Доля от общей маржи, %']) else 0,
                        'abc_by_cost': row['По себестоимости'] if pd.notna(row['По себестоимости']) else '',
                        'abc_by_price': row['По цене за вычетом комиссии '] if pd.notna(row['По цене за вычетом комиссии ']) else '',
                        'abc_by_margin': row['По средней марже'] if pd.notna(row['По средней марже']) else '',
                        'money_in_goods': row.get('Деньги в товаре', 0) if pd.notna(row.get('Деньги в товаре', 0)) else 0,
                    }
                )
        
        self.stdout.write(f'Загружено {len(df)} записей анализа продаж')

    def load_purchase_plan(self, file_path):
        """Загрузка плана по выкупам"""
        self.stdout.write('Загружаем план по выкупам...')
        
        df = pd.read_excel(file_path, sheet_name='План по выкупам', header=1)
        
        for _, row in df.iterrows():
            if pd.notna(row['выкупы']):
                PurchasePlan.objects.get_or_create(
                    position=int(row['Unnamed: 2']) if pd.notna(row['Unnamed: 2']) else 0,
                    defaults={
                        'purchases': int(row['выкупы']) if pd.notna(row['выкупы']) else 0,
                        'initial_position': int(row['изначальная поз']) if pd.notna(row['изначальная поз']) else 0,
                        'total_orders': int(row['всего заказзаов']) if pd.notna(row['всего заказзаов']) else 0,
                    }
                )
        
        self.stdout.write(f'Загружено {len(df)} записей плана по выкупам')
