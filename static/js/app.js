// ===== СОВРЕМЕННЫЙ WB ANALYTICS DASHBOARD =====

// Глобальные переменные
let salesChart, categoryChart;
let columnVisibility = {};
let currentZoom = 1;

// ===== ИНИЦИАЛИЗАЦИЯ ПРИЛОЖЕНИЯ =====

document.addEventListener('DOMContentLoaded', function() {
    initializeApp();
});

function initializeApp() {
    // Инициализируем современную навигацию
    initializeModernNavigation();
    
    // Загружаем данные дашборда
    loadDashboardData();
    
    // Инициализируем графики
    initializeCharts();
    
    // Настраиваем обработчики событий
    setupEventListeners();
}

// ===== СОВРЕМЕННАЯ НАВИГАЦИЯ =====

function initializeModernNavigation() {
    // Обработка переключения сайдбара
    const navToggle = document.getElementById('navToggle');
    const sidebar = document.getElementById('sidebar');
    const mainContent = document.getElementById('mainContent');
    
    if (navToggle && sidebar && mainContent) {
        navToggle.addEventListener('click', function() {
            sidebar.classList.toggle('collapsed');
            mainContent.classList.toggle('expanded');
            
            // Сохраняем состояние в localStorage
            const isCollapsed = sidebar.classList.contains('collapsed');
            localStorage.setItem('sidebarCollapsed', isCollapsed);
        });
        
        // Восстанавливаем состояние сайдбара
        const isCollapsed = localStorage.getItem('sidebarCollapsed') === 'true';
        if (isCollapsed) {
            sidebar.classList.add('collapsed');
            mainContent.classList.add('expanded');
        }
    }
    
    // Обработка навигационных ссылок
    const navLinks = document.querySelectorAll('.nav-link[data-tab]');
    navLinks.forEach(link => {
        link.addEventListener('click', function(e) {
            e.preventDefault();
            
            const targetTab = this.getAttribute('data-tab');
            switchTab(targetTab);
            
            // Обновляем активное состояние
            navLinks.forEach(l => l.classList.remove('active'));
            this.classList.add('active');
        });
    });
}

function switchTab(tabName) {
    // Скрываем все вкладки
    const allTabs = document.querySelectorAll('.tab-content');
    allTabs.forEach(tab => {
        tab.style.display = 'none';
        tab.classList.remove('fade-in-up');
    });
    
    // Показываем выбранную вкладку
    const targetTab = document.getElementById(tabName);
    if (targetTab) {
        targetTab.style.display = 'block';
        setTimeout(() => {
            targetTab.classList.add('fade-in-up');
        }, 10);
    }
    
    // Обновляем заголовок страницы
    const pageTitle = document.getElementById('pageTitle');
    if (pageTitle) {
        const titles = {
            'dashboard': 'Главная',
            'sales-analysis': 'Анализ продаж',
            'financial-reports': 'Финотчеты',
            'nomenclature': 'Номенклатуры',
            'stock-balance': 'Остатки',
            'summary': 'Сводные данные'
        };
        pageTitle.textContent = titles[tabName] || 'Главная';
    }
    
    // Загружаем данные для вкладки
    loadTabData(tabName);
}

function loadTabData(tabName) {
    console.log('Загружаем данные для вкладки:', tabName);
    switch(tabName) {
        case 'dashboard':
            loadDashboardData();
            break;
        case 'sales-analysis':
            loadSalesAnalysisData();
            break;
        case 'financial-reports':
            loadFinancialReportsData();
            break;
        case 'nomenclature':
            loadNomenclatureData();
            break;
        case 'stock-balance':
            loadStockBalanceData();
            break;
        case 'summary':
            loadSummaryData();
            break;
        default:
            console.log('Неизвестная вкладка:', tabName);
    }
}

// ===== ЗАГРУЗКА ДАННЫХ ДАШБОРДА =====

async function loadDashboardData() {
    try {
        showLoading(true);
        console.log('Начинаем загрузку данных дашборда...');
        
        // Загружаем обзорные данные
        const overviewResponse = await fetch('/api/dashboard/overview/');
        const overviewData = await overviewResponse.json();
        console.log('Данные обзора загружены:', overviewData);
        
        // Обновляем статистические карточки
        updateStatsCards(overviewData);
        
        // Загружаем данные для графиков
        const chartsResponse = await fetch('/api/dashboard/charts_data/');
        const chartsData = await chartsResponse.json();
        console.log('Данные графиков загружены:', chartsData);
        
        // Обновляем графики
        updateCharts(chartsData);
        
        // Загружаем последнюю активность (временно отключено)
        // loadRecentActivity();
        
        console.log('Данные дашборда успешно загружены');
        
    } catch (error) {
        console.error('Ошибка загрузки данных дашборда:', error);
        showAlert('Ошибка загрузки данных', 'error');
    } finally {
        showLoading(false);
    }
}

function updateStatsCards(data) {
    console.log('Обновляем статистические карточки с данными:', data);
    
    // Общая выручка (используем данные из summary)
    const totalRevenue = document.getElementById('totalRevenue');
    if (totalRevenue && data.summary) {
        const revenue = parseFloat(data.summary.average_check_after_spp) * parseInt(data.summary.sales_minus_returns_count);
        totalRevenue.textContent = formatCurrency(revenue);
        console.log('Обновлена выручка:', revenue);
    }
    
    // Общее количество заказов
    const totalOrders = document.getElementById('totalOrders');
    if (totalOrders && data.summary) {
        totalOrders.textContent = formatNumber(data.summary.sales_minus_returns_count || 0);
        console.log('Обновлено количество заказов:', data.summary.sales_minus_returns_count);
    }
    
    // Средняя маржа
    const avgMargin = document.getElementById('avgMargin');
    if (avgMargin && data.summary) {
        avgMargin.textContent = formatPercentage(data.summary.margin_percent_with_spp || 0);
        console.log('Обновлена средняя маржа:', data.summary.margin_percent_with_spp);
    }
    
    // Общее количество товаров (используем количество уникальных товаров)
    const totalProducts = document.getElementById('totalProducts');
    if (totalProducts && data.top_products) {
        totalProducts.textContent = formatNumber(data.top_products.length || 0);
        console.log('Обновлено количество товаров:', data.top_products.length);
    }
}

function updateCharts(data) {
    // График продаж
    updateSalesChart(data.sales_by_day);
    
    // График категорий
    updateCategoryChart(data.margin_by_product);
}

function updateSalesChart(chartData) {
    const ctx = document.getElementById('salesChart');
    if (!ctx || !chartData) return;
    
    if (salesChart) {
        salesChart.destroy();
    }
    
    // Подготавливаем данные для графика
    const labels = chartData.map(item => {
        const date = new Date(item.day);
        return date.toLocaleDateString('ru-RU', { month: 'short', day: 'numeric' });
    });
    
    const salesData = chartData.map(item => parseFloat(item.total_sales) || 0);
    
    salesChart = new Chart(ctx, {
        type: 'line',
        data: {
            labels: labels,
            datasets: [{
                label: 'Выручка',
                data: salesData,
                borderColor: '#6366f1',
                backgroundColor: 'rgba(99, 102, 241, 0.1)',
                borderWidth: 3,
                fill: true,
                tension: 0.4,
                pointBackgroundColor: '#6366f1',
                pointBorderColor: '#ffffff',
                pointBorderWidth: 2,
                pointRadius: 6,
                pointHoverRadius: 8
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            aspectRatio: 2,
            plugins: {
                legend: {
                    display: false
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    grid: {
                        color: '#e2e8f0',
                        drawBorder: false
                    },
                    ticks: {
                        color: '#64748b',
                        font: {
                            family: 'Inter, sans-serif'
                        },
                        callback: function(value) {
                            return '₽' + value.toLocaleString('ru-RU');
                        }
                    }
                },
                x: {
                    grid: {
                        display: false
                    },
                    ticks: {
                        color: '#64748b',
                        font: {
                            family: 'Inter, sans-serif'
                        }
                    }
                }
            },
            interaction: {
                intersect: false,
                mode: 'index'
            }
        }
    });
}

function updateCategoryChart(chartData) {
    const ctx = document.getElementById('categoryChart');
    if (!ctx || !chartData) return;
    
    if (categoryChart) {
        categoryChart.destroy();
    }
    
    // Подготавливаем данные для графика
    const labels = chartData.map(item => item.brand || 'Неизвестный бренд');
    const data = chartData.map(item => parseFloat(item.total_margin) || 0);
    
    categoryChart = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: labels,
            datasets: [{
                data: data,
                backgroundColor: [
                    '#6366f1',
                    '#10b981',
                    '#f59e0b',
                    '#ef4444',
                    '#06b6d4',
                    '#8b5cf6',
                    '#f97316',
                    '#84cc16',
                    '#ec4899',
                    '#8b5cf6'
                ],
                borderWidth: 0,
                cutout: '60%'
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            aspectRatio: 1,
            plugins: {
                legend: {
                    position: 'bottom',
                    labels: {
                        color: '#64748b',
                        font: {
                            family: 'Inter, sans-serif',
                            size: 12
                        },
                        padding: 20,
                        usePointStyle: true
                    }
                }
            }
        }
    });
}

async function loadRecentActivity() {
    try {
        const response = await fetch('/api/dashboard/recent_activity/');
        const data = await response.json();
        
        const tbody = document.getElementById('recentActivityTable');
        if (!tbody) return;
        
        if (data.length === 0) {
            tbody.innerHTML = '<tr><td colspan="4" class="text-center text-muted">Нет данных о последней активности</td></tr>';
            return;
        }
        
        tbody.innerHTML = data.map(activity => `
            <tr>
                <td>${formatDateTime(activity.timestamp)}</td>
                <td>${activity.action}</td>
                <td>${activity.details}</td>
                <td><span class="badge bg-${getStatusColor(activity.status)}">${activity.status}</span></td>
            </tr>
        `).join('');
        
    } catch (error) {
        console.error('Ошибка загрузки последней активности:', error);
    }
}

// ===== УПРАВЛЕНИЕ СТОЛБЦАМИ ТАБЛИЦЫ АНАЛИЗА ПРОДАЖ =====

function setupColumnControls() {
    console.log('Настраиваем контролы столбцов...');
    const table = document.getElementById('sales-analysis-table');
    if (!table) {
        console.error('Таблица sales-analysis-table не найдена');
        return;
    }
    
    // Проверяем, не создан ли уже контейнер
    let controlsContainer = document.getElementById('column-controls');
    if (controlsContainer) {
        console.log('Контролы столбцов уже созданы, пропускаем...');
        return;
    }
    
    // Создаем контейнер для контролов столбцов
    controlsContainer = document.createElement('div');
    controlsContainer.id = 'column-controls';
    controlsContainer.className = 'mb-3 p-3 bg-light rounded';
    controlsContainer.innerHTML = '<h6>Управление видимостью столбцов</h6>';
    
    // Вставляем перед таблицей
    const tableContainer = table.closest('.table-responsive');
    if (tableContainer) {
        tableContainer.parentNode.insertBefore(controlsContainer, tableContainer);
    }
    
    // Получаем все заголовки таблицы
    const headers = table.querySelectorAll('thead th');
    console.log(`Найдено заголовков: ${headers.length}`);
    const checkboxes = [];
    
    headers.forEach((header, index) => {
        console.log(`Создаем чекбокс для столбца ${index}: ${header.textContent.trim()}`);
        const checkbox = document.createElement('div');
        checkbox.className = 'form-check form-check-inline';
        checkbox.innerHTML = `
            <input class="form-check-input" type="checkbox" id="col-${index}" checked>
            <label class="form-check-label" for="col-${index}">${header.textContent.trim()}</label>
        `;
        
        // Добавляем обработчик события
        const input = checkbox.querySelector('input');
        input.addEventListener('change', function() {
            console.log(`Чекбокс столбца ${index} изменен на ${this.checked}`);
            toggleColumn(index, this.checked);
        });
        
        controlsContainer.appendChild(checkbox);
        checkboxes.push(input);
    });
    
    // Функция для переключения видимости столбца
    window.toggleColumn = function(columnIndex, visible) {
        console.log(`Переключаем столбец ${columnIndex}, видимый: ${visible}`);
        const table = document.getElementById('sales-analysis-table');
        if (!table) {
            console.error('Таблица не найдена');
            return;
        }
        
        // Используем более надежный селектор
        const columnNumber = columnIndex + 1;
        
        // Скрываем/показываем заголовок
        const headers = table.querySelectorAll('thead th');
        if (headers[columnIndex]) {
            headers[columnIndex].style.display = visible ? '' : 'none';
            console.log(`Заголовок столбца ${columnIndex} ${visible ? 'показан' : 'скрыт'}`);
        } else {
            console.error(`Заголовок столбца ${columnIndex} не найден`);
        }
        
        // Скрываем/показываем все ячейки в этом столбце
        const rows = table.querySelectorAll('tbody tr');
        let hiddenCells = 0;
        rows.forEach((row, rowIndex) => {
            const cells = row.querySelectorAll('td');
            if (cells[columnIndex]) {
                cells[columnIndex].style.display = visible ? '' : 'none';
                if (!visible) hiddenCells++;
            }
        });
        
        console.log(`Скрыто ячеек в столбце ${columnIndex}: ${hiddenCells}`);
    };
    
    // Тестовая функция для проверки работы
    window.testColumnToggle = function() {
        console.log('Тестируем переключение столбца 0...');
        toggleColumn(0, false);
        setTimeout(() => {
            console.log('Показываем столбец 0...');
            toggleColumn(0, true);
        }, 2000);
    };
}

// ===== ЗАГРУЗКА ДАННЫХ АНАЛИЗА ПРОДАЖ =====

async function loadSalesAnalysisData() {
    try {
        console.log('Загружаем данные анализа продаж...');
        showLoading(true);
        
        const response = await fetch('/api/sales-analysis/');
        const responseData = await response.json();
        const data = responseData.results || responseData; // Поддержка пагинации DRF
        console.log('Данные анализа продаж загружены:', data.length, 'записей');
        
        const tbody = document.getElementById('sales-analysis-tbody');
        if (!tbody) {
            console.error('Элемент sales-analysis-tbody не найден');
            return;
        }
        
        if (data.length === 0) {
            tbody.innerHTML = '<tr><td colspan="48" class="text-center text-muted">Нет данных для анализа продаж</td></tr>';
            return;
        }
        
        tbody.innerHTML = data.map(item => `
            <tr>
                <td class="text-cell" title="${item.brand}">${item.brand}</td>
                <td class="text-cell" title="${item.subject}">${item.subject}</td>
                <td class="text-cell" title="${item.article}">${item.article}</td>
                <td class="text-cell" title="${item.size}">${item.size}</td>
                <td class="text-cell" title="${item.barcode}">${item.barcode}</td>
                <td class="number-cell">${formatNumber(item.in_transit_to_client)}</td>
                <td class="number-cell">${formatNumber(item.in_transit_from_client)}</td>
                <td class="number-cell">${formatNumber(item.in_warehouses)}</td>
                <td class="number-cell">${formatNumber(item.total_stock_wb)}</td>
                <td class="number-cell">${formatNumber(item.turnover_days)}</td>
                <td class="number-cell">${formatNumber(item.orders)}</td>
                <td class="number-cell">${formatNumber(item.rejections)}</td>
                <td class="number-cell">${formatNumber(item.sales)}</td>
                <td class="number-cell">${formatNumber(item.returns)}</td>
                <td class="number-cell">${formatNumber(item.sales_minus_returns)}</td>
                <td class="percent-cell">${formatPercentage(item.purchase_percentage)}</td>
                <td class="number-cell">${formatNumber(item.supplier_return)}</td>
                <td class="currency-cell">${formatCurrency(item.logistics_for_returns)}</td>
                <td class="currency-cell">${formatCurrency(item.sales_before_spp)}</td>
                <td class="currency-cell">${formatCurrency(item.sales_with_spp)}</td>
                <td class="currency-cell">${formatCurrency(item.sales_minus_commission)}</td>
                <td class="currency-cell">${formatCurrency(item.sales_minus_commission_no_returns)}</td>
                <td class="currency-cell">${formatCurrency(item.return_amount)}</td>
                <td class="currency-cell">${formatCurrency(item.sales_minus_returns_amount)}</td>
                <td class="currency-cell">${formatCurrency(item.commission)}</td>
                <td class="percent-cell">${formatPercentage(item.commission_percent)}</td>
                <td class="currency-cell">${formatCurrency(item.logistics)}</td>
                <td class="currency-cell">${formatCurrency(item.logistics_per_unit)}</td>
                <td class="currency-cell">${formatCurrency(item.acquiring)}</td>
                <td class="currency-cell">${formatCurrency(item.fine)}</td>
                <td class="currency-cell">${formatCurrency(item.additional_payments)}</td>
                <td class="currency-cell">${formatCurrency(item.substitution_compensation)}</td>
                <td class="currency-cell">${formatCurrency(item.defect_compensation)}</td>
                <td class="currency-cell">${formatCurrency(item.average_check)}</td>
                <td class="currency-cell">${formatCurrency(item.cost_per_unit)}</td>
                <td class="currency-cell">${formatCurrency(item.sold_goods_cost)}</td>
                <td class="currency-cell">${formatCurrency(item.margin_before_tax)}</td>
                <td class="currency-cell">${formatCurrency(item.tax_6_percent)}</td>
                <td class="currency-cell">${formatCurrency(item.margin_after_tax)}</td>
                <td class="currency-cell">${formatCurrency(item.margin_per_unit)}</td>
                <td class="percent-cell">${formatPercentage(item.margin_percent)}</td>
                <td class="percent-cell">${formatPercentage(item.roi_from_cost)}</td>
                <td class="percent-cell">${formatPercentage(item.gmroi)}</td>
                <td class="percent-cell">${formatPercentage(item.revenue_share)}</td>
                <td class="percent-cell">${formatPercentage(item.margin_share)}</td>
                <td class="text-cell" title="${item.abc_by_cost}">${item.abc_by_cost}</td>
                <td class="text-cell" title="${item.abc_by_price}">${item.abc_by_price}</td>
                <td class="text-cell" title="${item.abc_by_margin}">${item.abc_by_margin}</td>
                <td class="currency-cell">${formatCurrency(item.money_in_goods)}</td>
            </tr>
        `).join('');
        
        // Инициализируем управление столбцами после загрузки данных
        setTimeout(() => {
            setupColumnControls();
        }, 500); // Увеличиваем задержку
        
    } catch (error) {
        console.error('Ошибка загрузки данных анализа продаж:', error);
        showAlert('Ошибка загрузки данных анализа продаж', 'error');
    } finally {
        showLoading(false);
    }
}

// ===== ЗАГРУЗКА ДАННЫХ ФИНАНСОВЫХ ОТЧЕТОВ =====

async function loadFinancialReportsData() {
    try {
        console.log('Загружаем финансовые отчеты...');
        showLoading(true);
        
        const response = await fetch('/api/financial-reports/');
        const responseData = await response.json();
        const data = responseData.results || responseData; // Поддержка пагинации DRF
        console.log('Финансовые отчеты загружены:', data.length, 'записей');
        
        const tbody = document.getElementById('financial-reports-tbody');
        if (!tbody) {
            console.error('Элемент financial-reports-tbody не найден');
            return;
        }
        
        if (data.length === 0) {
            tbody.innerHTML = '<tr><td colspan="11" class="text-center text-muted">Нет данных финансовых отчетов</td></tr>';
            return;
        }
        
        tbody.innerHTML = data.map(item => `
            <tr>
                <td>${item.number}</td>
                <td>${item.delivery_number}</td>
                <td>${item.subject}</td>
                <td>${item.brand}</td>
                <td>${item.supplier_article}</td>
                <td>${item.name}</td>
                <td>${item.size}</td>
                <td>${formatNumber(item.quantity)}</td>
                <td>${formatCurrency(item.retail_price)}</td>
                <td>${formatDateTime(item.sale_date)}</td>
                <td>${formatCurrency(item.payment_to_seller)}</td>
            </tr>
        `).join('');
        
    } catch (error) {
        console.error('Ошибка загрузки финансовых отчетов:', error);
        showAlert('Ошибка загрузки финансовых отчетов', 'error');
    } finally {
        showLoading(false);
    }
}

// ===== ЗАГРУЗКА ДАННЫХ НОМЕНКЛАТУР =====

async function loadNomenclatureData() {
    try {
        console.log('Загружаем номенклатуры...');
        showLoading(true);
        
        const response = await fetch('/api/nomenclature/');
        const responseData = await response.json();
        const data = responseData.results || responseData; // Поддержка пагинации DRF
        console.log('Номенклатуры загружены:', data.length, 'записей');
        
        const tbody = document.getElementById('nomenclature-tbody');
        if (!tbody) {
            console.error('Элемент nomenclature-tbody не найден');
            return;
        }
        
        if (data.length === 0) {
            tbody.innerHTML = '<tr><td colspan="9" class="text-center text-muted">Нет данных номенклатур</td></tr>';
            return;
        }
        
        tbody.innerHTML = data.map(item => `
            <tr>
                <td>${item.brand}</td>
                <td>${item.subject}</td>
                <td>${item.size_code}</td>
                <td>${item.supplier_article}</td>
                <td>${item.wb_article}</td>
                <td>${item.size}</td>
                <td>${item.barcode}</td>
                <td>${formatCurrency(item.equipment)}</td>
                <td>${formatCurrency(item.cost_price)}</td>
            </tr>
        `).join('');
        
    } catch (error) {
        console.error('Ошибка загрузки номенклатур:', error);
        showAlert('Ошибка загрузки номенклатур', 'error');
    } finally {
        showLoading(false);
    }
}

// ===== ЗАГРУЗКА ДАННЫХ ОСТАТКОВ =====

async function loadStockBalanceData() {
    try {
        console.log('Загружаем остатки...');
        showLoading(true);
        
        const response = await fetch('/api/stock-balance/');
        const responseData = await response.json();
        const data = responseData.results || responseData; // Поддержка пагинации DRF
        console.log('Остатки загружены:', data.length, 'записей');
        
        const tbody = document.getElementById('stock-balance-tbody');
        if (!tbody) {
            console.error('Элемент stock-balance-tbody не найден');
            return;
        }
        
        if (data.length === 0) {
            tbody.innerHTML = '<tr><td colspan="8" class="text-center text-muted">Нет данных об остатках</td></tr>';
            return;
        }
        
        tbody.innerHTML = data.map(item => `
            <tr>
                <td>${item.brand}</td>
                <td>${item.subject}</td>
                <td>${item.article}</td>
                <td>${item.size}</td>
                <td>${formatNumber(item.in_transit_to_client)}</td>
                <td>${formatNumber(item.in_transit_from_client)}</td>
                <td>${formatNumber(item.in_warehouses)}</td>
                <td>${formatNumber(item.total_stock_wb)}</td>
            </tr>
        `).join('');
        
    } catch (error) {
        console.error('Ошибка загрузки остатков:', error);
        showAlert('Ошибка загрузки остатков', 'error');
    } finally {
        showLoading(false);
    }
}

// ===== ЗАГРУЗКА СВОДНЫХ ДАННЫХ =====

async function loadSummaryData() {
    try {
        console.log('Загружаем сводные данные...');
        showLoading(true);
        
        const response = await fetch('/api/summary/');
        const responseData = await response.json();
        const data = responseData.results || responseData; // Поддержка пагинации DRF
        console.log('Сводные данные загружены:', data.length, 'записей');
        
        const tbody = document.getElementById('summary-tbody');
        if (!tbody) {
            console.error('Элемент summary-tbody не найден');
            return;
        }
        
        if (data.length === 0) {
            tbody.innerHTML = '<tr><td colspan="4" class="text-center text-muted">Нет сводных данных</td></tr>';
            return;
        }
        
        tbody.innerHTML = data.map(item => `
            <tr>
                <td>${item.metric_name}</td>
                <td>${item.value}</td>
                <td>${item.unit}</td>
                <td>${formatDateTime(item.updated_at)}</td>
            </tr>
        `).join('');
        
    } catch (error) {
        console.error('Ошибка загрузки сводных данных:', error);
        showAlert('Ошибка загрузки сводных данных', 'error');
    } finally {
        showLoading(false);
    }
}

// ===== ФУНКЦИИ УПРАВЛЕНИЯ СТОЛБЦАМИ =====

function initializeColumnControls() {
    const table = document.getElementById('sales-analysis-table');
    if (!table) return;
    
    const headers = table.querySelectorAll('thead th');
    const checkboxesContainer = document.getElementById('column-checkboxes');
    
    // Очищаем контейнер
    checkboxesContainer.innerHTML = '';
    
    // Создаем чекбоксы для каждого столбца
    headers.forEach((header, index) => {
        const columnName = header.textContent.trim();
        const columnId = `col-${index}`;
        
        // Инициализируем видимость (все столбцы видимы по умолчанию)
        columnVisibility[columnId] = true;
        
        // Создаем чекбокс
        const checkboxDiv = document.createElement('div');
        checkboxDiv.className = 'col-md-3 col-sm-4 col-6 column-checkbox';
        checkboxDiv.innerHTML = `
            <div class="form-check">
                <input class="form-check-input" type="checkbox" id="${columnId}" checked onchange="updateColumnVisibility('${columnId}', this.checked)">
                <label class="form-check-label" for="${columnId}">
                    ${columnName}
                </label>
            </div>
        `;
        
        checkboxesContainer.appendChild(checkboxDiv);
    });
}

function toggleColumnVisibility() {
    const controls = document.getElementById('column-controls');
    if (controls.style.display === 'none') {
        controls.style.display = 'block';
        initializeColumnControls();
    } else {
        controls.style.display = 'none';
    }
}

function updateColumnVisibility(columnId, isVisible) {
    columnVisibility[columnId] = isVisible;
}

function applyColumnVisibility() {
    const table = document.getElementById('sales-analysis-table');
    if (!table) return;
    
    const headers = table.querySelectorAll('thead th');
    const rows = table.querySelectorAll('tbody tr');
    
    headers.forEach((header, index) => {
        const columnId = `col-${index}`;
        const isVisible = columnVisibility[columnId];
        
        // Показываем/скрываем заголовок
        if (isVisible) {
            header.classList.remove('hidden');
        } else {
            header.classList.add('hidden');
        }
        
        // Показываем/скрываем ячейки в строках
        rows.forEach(row => {
            const cell = row.children[index];
            if (cell) {
                if (isVisible) {
                    cell.classList.remove('hidden');
                } else {
                    cell.classList.add('hidden');
                }
            }
        });
    });
    
    // Скрываем панель управления
    document.getElementById('column-controls').style.display = 'none';
    
    showAlert('Настройки столбцов применены', 'success');
}

function showAllColumns() {
    Object.keys(columnVisibility).forEach(columnId => {
        columnVisibility[columnId] = true;
        const checkbox = document.getElementById(columnId);
        if (checkbox) {
            checkbox.checked = true;
        }
    });
}

function hideAllColumns() {
    Object.keys(columnVisibility).forEach(columnId => {
        columnVisibility[columnId] = false;
        const checkbox = document.getElementById(columnId);
        if (checkbox) {
            checkbox.checked = false;
        }
    });
}

// ===== ФУНКЦИИ МАСШТАБИРОВАНИЯ ТАБЛИЦЫ =====

function zoomTable(factor) {
    currentZoom *= factor;
    currentZoom = Math.max(0.5, Math.min(2, currentZoom)); // Ограничиваем диапазон
    
    updateTableZoom(currentZoom);
}

function updateTableZoom(zoomValue) {
    currentZoom = parseFloat(zoomValue);
    const table = document.getElementById('sales-analysis-table');
    const zoomPercentage = document.getElementById('zoom-percentage');
    
    if (table && zoomPercentage) {
        // Применяем масштаб к таблице
        table.style.transform = `scale(${currentZoom})`;
        table.style.transformOrigin = 'top left';
        
        // Обновляем отображение процента
        zoomPercentage.textContent = Math.round(currentZoom * 100) + '%';
        
        // Добавляем контейнер для масштабирования если его нет
        const tableContainer = table.closest('.table-responsive');
        if (!tableContainer.classList.contains('table-zoom-container')) {
            tableContainer.classList.add('table-zoom-container');
        }
    }
}

function resetTableZoom() {
    currentZoom = 1;
    updateTableZoom(1);
    showAlert('Масштаб сброшен', 'info');
}

// ===== ПЕРЕСЧЕТ ДАННЫХ =====

async function recalculateSales() {
    try {
        showLoading(true);
        
        const response = await fetch('/api/sales-analysis/recalculate/', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            }
        });
        
        if (response.ok) {
            showAlert('Данные успешно пересчитаны', 'success');
            // Перезагружаем данные
            await loadSalesAnalysisData();
        } else {
            throw new Error('Ошибка пересчета данных');
        }
        
    } catch (error) {
        console.error('Ошибка пересчета:', error);
        showAlert('Ошибка пересчета данных', 'error');
    } finally {
        showLoading(false);
    }
}

// ===== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ =====

function setupEventListeners() {
    // Здесь можно добавить дополнительные обработчики событий
}

function showLoading(show) {
    const loadingElements = document.querySelectorAll('.loading');
    loadingElements.forEach(el => {
        if (show) {
            el.classList.add('loading');
        } else {
            el.classList.remove('loading');
        }
    });
}

function showAlert(message, type = 'info') {
    // Создаем уведомление
    const alertDiv = document.createElement('div');
    alertDiv.className = `alert alert-${type === 'error' ? 'danger' : type} alert-dismissible fade show position-fixed`;
    alertDiv.style.cssText = 'top: 20px; right: 20px; z-index: 9999; min-width: 300px;';
    alertDiv.innerHTML = `
        ${message}
        <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
    `;
    
    document.body.appendChild(alertDiv);
    
    // Автоматически удаляем через 5 секунд
    setTimeout(() => {
        if (alertDiv.parentNode) {
            alertDiv.parentNode.removeChild(alertDiv);
        }
    }, 5000);
}

function formatCurrency(value) {
    if (value === null || value === undefined) return '₽0';
    return '₽' + parseFloat(value).toLocaleString('ru-RU', {
        minimumFractionDigits: 0,
        maximumFractionDigits: 0
    });
}

function formatNumber(value) {
    if (value === null || value === undefined) return '0';
    return parseFloat(value).toLocaleString('ru-RU');
}

function formatPercentage(value) {
    if (value === null || value === undefined) return '0%';
    return (parseFloat(value) * 100).toFixed(1) + '%';
}

function formatDateTime(value) {
    if (!value) return '-';
    const date = new Date(value);
    return date.toLocaleDateString('ru-RU') + ' ' + date.toLocaleTimeString('ru-RU', {
        hour: '2-digit',
        minute: '2-digit'
    });
}

function getStatusColor(status) {
    const colors = {
        'success': 'success',
        'error': 'danger',
        'warning': 'warning',
        'info': 'info',
        'pending': 'secondary'
    };
    return colors[status] || 'secondary';
}

// ===== ИНИЦИАЛИЗАЦИЯ ГРАФИКОВ =====

function initializeCharts() {
    // Графики будут инициализированы при загрузке данных
    console.log('Графики готовы к инициализации');
}

// ===== ЭКСПОРТ ФУНКЦИЙ ДЛЯ ГЛОБАЛЬНОГО ДОСТУПА =====

window.toggleColumnVisibility = toggleColumnVisibility;
window.updateColumnVisibility = updateColumnVisibility;
window.applyColumnVisibility = applyColumnVisibility;
window.showAllColumns = showAllColumns;
window.hideAllColumns = hideAllColumns;
window.zoomTable = zoomTable;
window.updateTableZoom = updateTableZoom;
window.resetTableZoom = resetTableZoom;
window.recalculateSales = recalculateSales;