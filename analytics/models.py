from django.db import models
from django.core.validators import MinValueValidator
from decimal import Decimal


class Nomenclature(models.Model):
    """Модель для листа 'Номенклатуры (вставить)'"""
    brand = models.CharField(max_length=255, verbose_name="Бренд")
    subject = models.CharField(max_length=255, verbose_name="Предмет")
    size_code = models.CharField(max_length=255, verbose_name="Код размера (chrt_id)")
    supplier_article = models.CharField(max_length=255, verbose_name="Артикул продавца")
    wb_article = models.CharField(max_length=255, verbose_name="Артикул WB")
    size = models.CharField(max_length=100, verbose_name="Размер")
    barcode = models.CharField(max_length=255, verbose_name="Баркод")
    equipment = models.DecimalField(max_digits=10, decimal_places=2, verbose_name="Комплектация")
    composition = models.TextField(blank=True, null=True, verbose_name="Состав")
    cost_price = models.DecimalField(max_digits=10, decimal_places=2, verbose_name="Себестоимость")
    
    class Meta:
        verbose_name = "Номенклатура"
        verbose_name_plural = "Номенклатуры"
        unique_together = ['brand', 'supplier_article', 'size']
    
    def __str__(self):
        return f"{self.brand} - {self.supplier_article}"


class FinancialReport(models.Model):
    """Модель для листа 'Финотчеты (вставить)'"""
    number = models.IntegerField(verbose_name="№")
    delivery_number = models.CharField(max_length=255, verbose_name="Номер поставки")
    subject = models.CharField(max_length=255, verbose_name="Предмет")
    nomenclature_code = models.CharField(max_length=255, verbose_name="Код номенклатуры")
    brand = models.CharField(max_length=255, verbose_name="Бренд")
    supplier_article = models.CharField(max_length=255, verbose_name="Артикул поставщика")
    name = models.CharField(max_length=500, verbose_name="Название")
    size = models.CharField(max_length=100, verbose_name="Размер")
    barcode = models.CharField(max_length=255, verbose_name="Баркод")
    document_type = models.CharField(max_length=255, blank=True, null=True, verbose_name="Тип документа")
    payment_basis = models.CharField(max_length=255, verbose_name="Обоснование для оплаты")
    order_date = models.DateTimeField(verbose_name="Дата заказа покупателем")
    sale_date = models.DateTimeField(verbose_name="Дата продажи")
    quantity = models.IntegerField(verbose_name="Кол-во")
    retail_price = models.DecimalField(max_digits=10, decimal_places=2, verbose_name="Цена розничная")
    wb_sold_product = models.DecimalField(max_digits=10, decimal_places=2, verbose_name="Вайлдберриз реализовал Товар (Пр)")
    agreed_product_discount = models.DecimalField(max_digits=5, decimal_places=2, default=0, verbose_name="Согласованный продуктовый дисконт, %")
    promo_code_percent = models.DecimalField(max_digits=5, decimal_places=2, default=0, verbose_name="Промокод %")
    total_agreed_discount = models.DecimalField(max_digits=5, decimal_places=2, default=0, verbose_name="Итоговая согласованная скидка")
    retail_price_with_discount = models.DecimalField(max_digits=10, decimal_places=2, default=0, verbose_name="Цена розничная с учетом согласованной скидки")
    kvv_reduction_rating = models.DecimalField(max_digits=5, decimal_places=2, default=0, verbose_name="Размер снижения кВВ из-за рейтинга, %")
    kvv_reduction_promotion = models.DecimalField(max_digits=5, decimal_places=2, default=0, verbose_name="Размер снижения кВВ из-за акции, %")
    spp_discount = models.DecimalField(max_digits=5, decimal_places=2, default=0, verbose_name="Скидка постоянного Покупателя (СПП)")
    kvv_size = models.DecimalField(max_digits=5, decimal_places=2, default=0, verbose_name="Размер кВВ, %")
    kvv_base_without_vat = models.DecimalField(max_digits=5, decimal_places=2, default=0, verbose_name="Размер кВВ без НДС, % Базовый")
    kvv_final_without_vat = models.DecimalField(max_digits=5, decimal_places=2, default=0, verbose_name="Итоговый кВВ без НДС, %")
    reward_before_services = models.DecimalField(max_digits=10, decimal_places=2, default=0, verbose_name="Вознаграждение с продаж до вычета услуг поверенного, без НДС")
    compensation_pickup_return = models.DecimalField(max_digits=10, decimal_places=2, default=0, verbose_name="Возмещение за выдачу и возврат товаров на ПВЗ")
    acquiring_commission = models.DecimalField(max_digits=10, decimal_places=2, default=0, verbose_name="Эквайринг/Комиссии за организацию платежей")
    acquiring_commission_percent = models.DecimalField(max_digits=5, decimal_places=2, default=0, verbose_name="Размер комиссии за эквайринг/Комиссии за организацию платежей, %")
    acquiring_payment_type = models.CharField(max_length=255, default='', verbose_name="Тип платежа за Эквайринг/Комиссии за организацию платежей")
    wb_reward = models.DecimalField(max_digits=10, decimal_places=2, default=0, verbose_name="Вознаграждение Вайлдберриз (ВВ), без НДС")
    vat_wb_reward = models.DecimalField(max_digits=10, decimal_places=2, default=0, verbose_name="НДС с Вознаграждения Вайлдберриз")
    payment_to_seller = models.DecimalField(max_digits=10, decimal_places=2, default=0, verbose_name="К перечислению Продавцу за реализованный Товар")
    delivery_count = models.IntegerField(default=0, verbose_name="Количество доставок")
    return_count = models.IntegerField(default=0, verbose_name="Количество возврата")
    delivery_services = models.DecimalField(max_digits=10, decimal_places=2, default=0, verbose_name="Услуги по доставке товара покупателю")
    fixation_start_date = models.DateTimeField(blank=True, null=True, verbose_name="Дата начала действия фиксации")
    fixation_end_date = models.DateTimeField(blank=True, null=True, verbose_name="Дата конца действия фиксации")
    paid_delivery_flag = models.CharField(max_length=255, blank=True, null=True, verbose_name="Признак услуги платной доставки")
    total_fines = models.DecimalField(max_digits=10, decimal_places=2, default=0, verbose_name="Общая сумма штрафов")
    additional_payments = models.DecimalField(max_digits=10, decimal_places=2, default=0, verbose_name="Доплаты")
    logistics_types = models.CharField(max_length=255, default='', verbose_name="Виды логистики, штрафов и доплат")
    mp_sticker = models.CharField(max_length=255, default='', verbose_name="Стикер МП")
    acquirer_bank = models.CharField(max_length=255, default='', verbose_name="Наименование банка-эквайера")
    office_number = models.CharField(max_length=255, default='', verbose_name="Номер офиса")
    delivery_office = models.CharField(max_length=255, default='', verbose_name="Наименование офиса доставки")
    partner_inn = models.CharField(max_length=255, default='', verbose_name="ИНН партнера")
    partner = models.CharField(max_length=255, default='', verbose_name="Партнер")
    warehouse = models.CharField(max_length=255, default='', verbose_name="Склад")
    country = models.CharField(max_length=255, default='', verbose_name="Страна")
    box_type = models.CharField(max_length=255, default='', verbose_name="Тип коробов")
    customs_declaration = models.CharField(max_length=255, blank=True, null=True, verbose_name="Номер таможенной декларации")
    assembly_task = models.CharField(max_length=255, blank=True, null=True, verbose_name="Номер сборочного задания")
    marking_code = models.CharField(max_length=255, default='', verbose_name="Код маркировки")
    sku = models.CharField(max_length=255, default='', verbose_name="ШК")
    srid = models.CharField(max_length=255, default='', verbose_name="Srid")
    transport_compensation = models.DecimalField(max_digits=10, decimal_places=2, default=0, verbose_name="Возмещение издержек по перевозке/по складским операциям с товаром")
    transport_organizer = models.CharField(max_length=255, blank=True, null=True, verbose_name="Организатор перевозки")
    storage = models.DecimalField(max_digits=10, decimal_places=2, default=0, verbose_name="Хранение")
    deductions = models.DecimalField(max_digits=10, decimal_places=2, default=0, verbose_name="Удержания")
    paid_reception = models.DecimalField(max_digits=10, decimal_places=2, default=0, verbose_name="Платная приемка")
    chrt_id = models.CharField(max_length=255, default='', verbose_name="chrtId")
    warehouse_coefficient = models.CharField(max_length=255, default='', verbose_name="Фиксированный коэффициент склада по поставке")
    
    class Meta:
        verbose_name = "Финансовый отчет"
        verbose_name_plural = "Финансовые отчеты"
    
    def __str__(self):
        return f"{self.brand} - {self.supplier_article} - {self.sale_date}"


class StockBalance(models.Model):
    """Модель для листа 'Остатки (вставить)'"""
    brand = models.CharField(max_length=255, verbose_name="Бренд")
    subject = models.CharField(max_length=255, verbose_name="Предмет")
    supplier_article = models.CharField(max_length=255, verbose_name="Артикул продавца")
    item_size = models.CharField(max_length=100, verbose_name="Размер вещи")
    in_transit_to_client = models.IntegerField(verbose_name="В пути до клиента")
    in_transit_from_client = models.IntegerField(verbose_name="В пути от клиента")
    total_in_warehouses = models.IntegerField(verbose_name="Итого по складам")
    
    class Meta:
        verbose_name = "Остаток на складе"
        verbose_name_plural = "Остатки на складах"
        unique_together = ['brand', 'supplier_article', 'item_size']
    
    def __str__(self):
        return f"{self.brand} - {self.supplier_article} - {self.item_size}"


class SalesAnalysis(models.Model):
    """Модель для листа 'Анализ продаж' - основной расчетный лист"""
    brand = models.CharField(max_length=255, verbose_name="Бренд")
    subject = models.CharField(max_length=255, verbose_name="Предмет")
    article = models.CharField(max_length=255, verbose_name="Артикул")
    size = models.CharField(max_length=100, verbose_name="Размер")
    barcode = models.CharField(max_length=255, verbose_name="Баркод")
    
    # Остатки товара
    in_transit_to_client = models.IntegerField(default=0, verbose_name="В пути до клиента")
    in_transit_from_client = models.IntegerField(default=0, verbose_name="В пути от клиента")
    in_warehouses = models.IntegerField(default=0, verbose_name="На складах")
    total_stock_wb = models.IntegerField(default=0, verbose_name="ИТОГО остаток на ВБ")
    turnover_days = models.IntegerField(blank=True, null=True, verbose_name="Оборачиваемость, дней")
    
    # Движение товаров в шт
    orders = models.IntegerField(default=0, verbose_name="Заказы, шт")
    rejections = models.IntegerField(default=0, verbose_name="Отказы")
    sales = models.IntegerField(default=0, verbose_name="Продажи, шт")
    returns = models.IntegerField(default=0, verbose_name="Возвраты, шт")
    sales_minus_returns = models.IntegerField(default=0, verbose_name="Продажи минус возвраты")
    purchase_percentage = models.DecimalField(max_digits=5, decimal_places=4, default=0, verbose_name="Процент выкупа")
    
    # Возврат поставщика
    supplier_return = models.IntegerField(default=0, verbose_name="Возврат, шт")
    logistics_for_returns = models.DecimalField(max_digits=10, decimal_places=2, default=0, verbose_name="Логистика за возвраты")
    
    # 3 уровня выручки
    sales_before_spp = models.DecimalField(max_digits=10, decimal_places=2, default=0, verbose_name="Продажи по ценам до СПП")
    sales_with_spp = models.DecimalField(max_digits=10, decimal_places=2, default=0, verbose_name="Продажи по ценам с СПП")
    sales_minus_commission = models.DecimalField(max_digits=10, decimal_places=2, default=0, verbose_name="Продажи за вычетом комиссии")
    sales_minus_commission_no_returns = models.DecimalField(max_digits=10, decimal_places=2, default=0, verbose_name="Продажи за вычетом комиссии (без учета возвратов)")
    
    # Детализация по возвратам
    return_amount = models.DecimalField(max_digits=10, decimal_places=2, default=0, verbose_name="Возвраты")
    sales_minus_returns_amount = models.DecimalField(max_digits=10, decimal_places=2, default=0, verbose_name="Продажи минус возвраты")
    
    # Расходы маркетплейса
    commission = models.DecimalField(max_digits=10, decimal_places=2, default=0, verbose_name="Комиссия, руб")
    commission_percent = models.DecimalField(max_digits=5, decimal_places=4, default=0, verbose_name="Комиссия %")
    logistics = models.DecimalField(max_digits=10, decimal_places=2, default=0, verbose_name="Логистика")
    logistics_per_unit = models.DecimalField(max_digits=10, decimal_places=2, default=0, verbose_name="Логистика на 1 продажу")
    acquiring = models.DecimalField(max_digits=10, decimal_places=2, default=0, verbose_name="Эквайринг")
    fine = models.DecimalField(max_digits=10, decimal_places=2, default=0, verbose_name="Штраф")
    additional_payments = models.DecimalField(max_digits=10, decimal_places=2, default=0, verbose_name="Доплаты")
    substitution_compensation = models.DecimalField(max_digits=10, decimal_places=2, default=0, verbose_name="Компенсация подмен")
    defect_compensation = models.DecimalField(max_digits=10, decimal_places=2, default=0, verbose_name="Возмещение брака")
    
    # Финансовые показатели
    average_check = models.DecimalField(max_digits=10, decimal_places=2, default=0, verbose_name="Средний чек")
    cost_per_unit = models.DecimalField(max_digits=10, decimal_places=2, default=0, verbose_name="Себестомость 1 шт")
    sold_goods_cost = models.DecimalField(max_digits=10, decimal_places=2, default=0, verbose_name="Себестоимость проданного товара")
    margin_before_tax = models.DecimalField(max_digits=10, decimal_places=2, default=0, verbose_name="Маржа до налогов")
    tax_6_percent = models.DecimalField(max_digits=10, decimal_places=2, default=0, verbose_name="Налог 6%")
    margin_after_tax = models.DecimalField(max_digits=10, decimal_places=2, default=0, verbose_name="Маржа после налогов, руб")
    margin_per_unit = models.DecimalField(max_digits=10, decimal_places=2, default=0, verbose_name="Маржа на 1 продажу, руб")
    margin_percent = models.DecimalField(max_digits=5, decimal_places=4, default=0, verbose_name="Маржинальность, %")
    
    # Оценка эффективности товаров
    roi_from_cost = models.DecimalField(max_digits=5, decimal_places=4, default=0, verbose_name="ROI от себестоимости, %")
    gmroi = models.DecimalField(max_digits=5, decimal_places=4, default=0, verbose_name="GMROI, %")
    revenue_share = models.DecimalField(max_digits=5, decimal_places=4, default=0, verbose_name="Доля от общей выручки, %")
    margin_share = models.DecimalField(max_digits=5, decimal_places=4, default=0, verbose_name="Доля от общей маржи, %")
    
    # АВС анализ
    abc_by_cost = models.CharField(max_length=1, default='', verbose_name="По себестоимости")
    abc_by_price = models.CharField(max_length=1, default='', verbose_name="По цене за вычетом комиссии")
    abc_by_margin = models.CharField(max_length=1, default='', verbose_name="По средней марже")
    
    # Деньги в товаре
    money_in_goods = models.DecimalField(max_digits=10, decimal_places=2, default=0, verbose_name="Деньги в товаре")
    
    class Meta:
        verbose_name = "Анализ продаж"
        verbose_name_plural = "Анализ продаж"
        unique_together = ['brand', 'article', 'size']
    
    def __str__(self):
        return f"{self.brand} - {self.article} - {self.size}"


class PurchasePlan(models.Model):
    """Модель для листа 'План по выкупам'"""
    position = models.IntegerField(verbose_name="Позиция")
    purchases = models.IntegerField(verbose_name="Выкупы")
    initial_position = models.IntegerField(verbose_name="Изначальная позиция")
    total_orders = models.IntegerField(verbose_name="Всего заказов")
    
    class Meta:
        verbose_name = "План по выкупам"
        verbose_name_plural = "Планы по выкупам"
    
    def __str__(self):
        return f"Позиция {self.position}: {self.purchases} выкупов"


class SummaryData(models.Model):
    """Модель для листа 'Сводный' - итоговые показатели"""
    # Основные показатели
    sales_minus_returns_count = models.IntegerField(verbose_name="Продаж за вычетом возвратов, шт")
    purchase_percentage = models.DecimalField(max_digits=5, decimal_places=4, verbose_name="Процент выкупа")
    average_check_after_spp = models.DecimalField(max_digits=10, decimal_places=2, verbose_name="Средний чек для покупателя после СПП")
    commission_per_unit = models.DecimalField(max_digits=10, decimal_places=2, verbose_name="Комиссия на 1 продажу, руб")
    logistics_per_unit = models.DecimalField(max_digits=10, decimal_places=2, verbose_name="Логистика на 1 продажу, руб")
    storage_per_unit = models.DecimalField(max_digits=10, decimal_places=6, verbose_name="Хранение на 1 продажу, руб")
    sold_goods_cost = models.DecimalField(max_digits=10, decimal_places=2, verbose_name="Себестоимость проданных товаров")
    average_cost_per_unit = models.DecimalField(max_digits=10, decimal_places=2, verbose_name="Средняя себестоимость 1 продажи, руб")
    average_margin_per_unit = models.DecimalField(max_digits=10, decimal_places=2, verbose_name="Средняя маржа с 1 продажи, руб")
    margin_without_ads = models.DecimalField(max_digits=10, decimal_places=2, verbose_name="Маржа за период БЕЗ учета рекламы, руб")
    margin_per_unit_without_ads = models.DecimalField(max_digits=10, decimal_places=2, verbose_name="Средняя маржа с 1 продажи, руб (без рекламы)")
    margin_after_all_expenses = models.DecimalField(max_digits=10, decimal_places=2, verbose_name="Маржа за период после всех расходов МП и налогов, руб")
    
    # Расходы маркетплейса
    total_commission = models.DecimalField(max_digits=10, decimal_places=2, verbose_name="Комиссия")
    commission_percent = models.DecimalField(max_digits=5, decimal_places=4, verbose_name="Комиссия %")
    total_logistics = models.DecimalField(max_digits=10, decimal_places=2, verbose_name="Логистика")
    logistics_percent = models.DecimalField(max_digits=5, decimal_places=4, verbose_name="Логистика %")
    total_storage = models.DecimalField(max_digits=10, decimal_places=2, verbose_name="Хранение")
    storage_percent = models.DecimalField(max_digits=8, decimal_places=6, verbose_name="Хранение %")
    additional_payments = models.DecimalField(max_digits=10, decimal_places=2, default=0, verbose_name="Доплаты")
    fines = models.DecimalField(max_digits=10, decimal_places=2, default=0, verbose_name="Штрафы")
    paid_reception = models.DecimalField(max_digits=10, decimal_places=2, default=0, verbose_name="Платная приемка")
    advertising_deductions = models.DecimalField(max_digits=10, decimal_places=2, verbose_name="Удержания (реклама)")
    advertising_percent = models.DecimalField(max_digits=5, decimal_places=4, verbose_name="Удержания (реклама) %")
    transit_deductions = models.DecimalField(max_digits=10, decimal_places=2, default=0, verbose_name="Удержания (транзит и др)")
    acquiring = models.DecimalField(max_digits=10, decimal_places=2, verbose_name="Эквайринг")
    defect_compensation = models.DecimalField(max_digits=10, decimal_places=2, default=0, verbose_name="Компенсация брака")
    substitution_compensation = models.DecimalField(max_digits=10, decimal_places=2, default=0, verbose_name="Компенсация подмен (продажи)")
    
    # Маржинальность
    margin_percent_before_spp = models.DecimalField(max_digits=5, decimal_places=4, verbose_name="Маржинальность от продаж до СПП, %")
    margin_percent_with_spp = models.DecimalField(max_digits=5, decimal_places=4, verbose_name="Маржинальность от продаж с СПП, %")
    margin_percent_after_commission = models.DecimalField(max_digits=5, decimal_places=4, verbose_name="Маржинальность от продаж за вычетом комиссии, %")
    
    # Проверочные суммы
    sales_before_spp = models.DecimalField(max_digits=10, decimal_places=2, verbose_name="Продажи по ценам до СПП")
    sales_with_spp = models.DecimalField(max_digits=10, decimal_places=2, verbose_name="Продажи по ценам с СПП")
    sales_minus_commission = models.DecimalField(max_digits=10, decimal_places=2, verbose_name="Продажи за вычетом комиссии")
    payment_to_account = models.DecimalField(max_digits=10, decimal_places=2, verbose_name="К перечислению на р/с")
    tax_6_percent = models.DecimalField(max_digits=10, decimal_places=2, verbose_name="Налог 6%")
    
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    
    class Meta:
        verbose_name = "Сводные данные"
        verbose_name_plural = "Сводные данные"
    
    def __str__(self):
        return f"Сводные данные от {self.created_at.strftime('%d.%m.%Y')}"