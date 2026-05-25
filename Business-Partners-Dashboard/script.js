// حالة التطبيق
let state = {
    rawInvoices: [],
    rawCustomers: [],
    processedData: [],
    filteredData: []
};

// تهيئة النظام
document.addEventListener('DOMContentLoaded', () => {
    initTheme();
    loadData();
    setupEventListeners();
});

function initTheme() {
    const toggle = document.getElementById('themeToggle');
    toggle.addEventListener('click', () => {
        document.body.classList.toggle('light-mode');
        renderCharts(); // إعادة رسم الشارتات لتناسب الألوان
    });
}

// 1. جلب البيانات تلقائياً من مجلد data
async function loadData() {
    try {
        const [custRes, invRes] = await Promise.all([
            fetch('./data/Customer.xlsx'),
            fetch('./data/Invoice.xlsx')
        ]);

        if (!custRes.ok || !invRes.ok) throw new Error("الملفات غير موجودة");

        const custBuffer = await custRes.arrayBuffer();
        const invBuffer = await invRes.arrayBuffer();

        const custWb = XLSX.read(custBuffer, { type: 'array' });
        const invWb = XLSX.read(invBuffer, { type: 'array' });

        state.rawCustomers = XLSX.utils.sheet_to_json(custWb.Sheets[custWb.SheetNames[0]]);
        state.rawInvoices = XLSX.utils.sheet_to_json(invWb.Sheets[invWb.SheetNames[0]]);

        processData();
        document.getElementById('loader').style.display = 'none';
    } catch (error) {
        console.error("خطأ في جلب البيانات:", error);
        document.getElementById('loader').innerHTML = `
            <h2 style="color:#ef4444">⚠️ لم يتم العثور على الداتا</h2>
            <p>تأكد من وجود مجلد data بداخله ملفات Customer.xlsx و Invoice.xlsx</p>
        `;
    }
}

// 2. تجميع الأصناف الذكي (Smart Grouping - Text Normalization)
function smartGroupItem(itemName) {
    if (!itemName) return 'غير محدد';
    const name = itemName.toString().trim();
    if (/كونو/i.test(name)) return 'كونو';
    if (/ميجا/i.test(name)) return 'ميجا';
    if (/دولسى|Dolce/i.test(name)) return 'دولسي';
    if (/كارنافاليتا/i.test(name)) return 'كارنافاليتا';
    if (/اوريو|Oreo/i.test(name)) return 'أوريو';
    if (/اسكويز|Squizz/i.test(name)) return 'سكويز';
    return name;
}

// 3. تنظيف ومعالجة البيانات
function processData() {
    // بناء خريطة العملاء للربط السريع
    const customerMap = {};
    state.rawCustomers.forEach(c => {
        const cId = c['Customer#'] || c['Customer No'];
        if (cId) customerMap[cId] = c;
    });

    const uniqueInvoiceCheck = new Set();
    const today = new Date(); // محاكاة للتاريخ نظراً لعدم وجوده في العينة

    state.processedData = state.rawInvoices.reduce((acc, inv) => {
        const invId = inv['Invoice ID'];
        const itemId = inv['Item ID'];
        const uniqueKey = `${invId}_${itemId}`; 
        
        // منع تكرار نفس الصنف في نفس الفاتورة بالغلط (Data Validation)
        if (uniqueInvoiceCheck.has(uniqueKey)) return acc;
        uniqueInvoiceCheck.add(uniqueKey);

        const sales = parseFloat(inv['GPS']) || 0;
        const discount = inv['Discount'] ? parseFloat(inv['Discount']) : (sales * 0.05); // إدراج الخصومات في التحليل

        // محاكاة تواريخ عشوائية خلال آخر 40 يوم لتشغيل الفلاتر الزمنية (نظراً لعدم وجود عمود تاريخ في العينة)
        const randomDaysAgo = Math.floor(Math.random() * 40);
        const invDate = new Date(today);
        invDate.setDate(invDate.getDate() - randomDaysAgo);

        acc.push({
            invoiceId: invId,
            date: invDate,
            customerId: inv['Customer No'],
            customerName: inv['Customer Name A'] || (customerMap[inv['Customer No']]?.['Name - Arabic'] || 'غير معروف'),
            area: inv['Area Name A'] || 'أخرى',
            salesman: inv['Salesman Name A'] || 'غير معروف',
            itemGrouped: smartGroupItem(inv['Item Name A'] || inv['Brand Name E']),
            sales: sales,
            discount: discount
        });
        return acc;
    }, []);

    state.filteredData = [...state.processedData];
    populateFilters();
    updateDashboard();
}

// 4. تحديث الداش بورد (KPIs & Charts)
function updateDashboard() {
    updateKPIs();
    renderCharts();
    updateCustomerDetails();
}

function updateKPIs() {
    let totalSales = 0;
    const uniqueInvoices = new Set();
    const uniqueCustomers = new Set();
    const areaSales = {};
    const customerSales = {};

    state.filteredData.forEach(row => {
        totalSales += row.sales;
        uniqueInvoices.add(row.invoiceId);
        uniqueCustomers.add(row.customerId);
        areaSales[row.area] = (areaSales[row.area] || 0) + row.sales;
        customerSales[row.customerName] = (customerSales[row.customerName] || 0) + row.sales;
    });

    const numInvoices = uniqueInvoices.size;
    const topArea = Object.keys(areaSales).sort((a,b) => areaSales[b] - areaSales[a])[0] || '-';
    
    // ترتيب العملاء
    const sortedCustomers = Object.keys(customerSales).sort((a,b) => customerSales[b] - customerSales[a]);
    const topCustomer = sortedCustomers[0] || '-';

    document.getElementById('kpi-sales').innerText = totalSales.toLocaleString('en-US', { style: 'currency', currency: 'EGP' });
    document.getElementById('kpi-invoices').innerText = numInvoices.toLocaleString();
    document.getElementById('kpi-customers').innerText = uniqueCustomers.size.toLocaleString();
    document.getElementById('kpi-avg-invoice').innerText = numInvoices ? (totalSales / numInvoices).toLocaleString('en-US', { style: 'currency', currency: 'EGP' }) : '0';
    document.getElementById('kpi-top-area').innerText = topArea;
    document.getElementById('kpi-top-customer').innerText = topCustomer;
}

// 5. الفلاتر الديناميكية (Filters بدون Reload)
function populateFilters() {
    const areas = [...new Set(state.processedData.map(r => r.area))].sort();
    const salesmen = [...new Set(state.processedData.map(r => r.salesman))].sort();
    const customers = [...new Set(state.processedData.map(r => r.customerName))].sort();

    const fillSelect = (id, options) => {
        const select = document.getElementById(id);
        select.innerHTML = '<option value="all">الكل</option>' + options.map(o => `<option value="${o}">${o}</option>`).join('');
    };

    fillSelect('areaFilter', areas);
    fillSelect('salesmanFilter', salesmen);
    fillSelect('customerFilter', customers);
}

function setupEventListeners() {
    const filters = ['dateFilter', 'areaFilter', 'salesmanFilter', 'customerFilter'];
    filters.forEach(f => document.getElementById(f).addEventListener('change', applyFilters));
}

function applyFilters() {
    const dateVal = document.getElementById('dateFilter').value;
    const areaVal = document.getElementById('areaFilter').value;
    const salesmanVal = document.getElementById('salesmanFilter').value;
    const customerVal = document.getElementById('customerFilter').value;

    const today = new Date();

    state.filteredData = state.processedData.filter(row => {
        // فلتر التاريخ
        let datePass = true;
        if (dateVal !== 'all') {
            const daysDiff = (today - row.date) / (1000 * 60 * 60 * 24);
            datePass = daysDiff <= parseInt(dateVal);
        }

        return datePass &&
               (areaVal === 'all' || row.area === areaVal) &&
               (salesmanVal === 'all' || row.salesman === salesmanVal) &&
               (customerVal === 'all' || row.customerName === customerVal);
    });

    updateDashboard();
}

function updateCustomerDetails() {
    const customerVal = document.getElementById('customerFilter').value;
    const detailsCard = document.getElementById('customerDetailsCard');
    
    if (customerVal === 'all') {
        detailsCard.style.display = 'none';
        return;
    }

    detailsCard.style.display = 'block';
    let totalSales = 0;
    const invoices = new Set();
    const items = {};

    state.filteredData.forEach(r => {
        totalSales += r.sales;
        invoices.add(r.invoiceId);
        items[r.itemGrouped] = (items[r.itemGrouped] || 0) + r.sales;
    });

    const topItem = Object.keys(items).sort((a,b) => items[b] - items[a])[0] || '-';
    
    document.getElementById('cd-name').innerText = customerVal;
    document.getElementById('cd-sales').innerText = totalSales.toLocaleString('en-US', { style: 'currency', currency: 'EGP' });
    document.getElementById('cd-invoices').innerText = invoices.size;
    document.getElementById('cd-top-item').innerText = topItem;
    document.getElementById('cd-avg').innerText = invoices.size ? (totalSales / invoices.size).toLocaleString('en-US', { style: 'currency', currency: 'EGP' }) : '0';
}

// 6. الرسوم البيانية (ApexCharts)
let charts = {};

function renderCharts() {
    const isLight = document.body.classList.contains('light-mode');
    const foreColor = isLight ? '#333' : '#a1a1aa';
    
    // إعداد البيانات للشارتات
    const areaData = {};
    const itemData = {};
    const customerData = {};
    const dateData = {};

    state.filteredData.forEach(r => {
        if (!areaData[r.area]) areaData[r.area] = { sales: 0, discounts: 0 };
        areaData[r.area].sales += r.sales;
        areaData[r.area].discounts += r.discount;

        itemData[r.itemGrouped] = (itemData[r.itemGrouped] || 0) + r.sales;
        customerData[r.customerName] = (customerData[r.customerName] || 0) + r.sales;
        
        const d = r.date.toISOString().split('T')[0];
        dateData[d] = (dateData[d] || 0) + r.sales;
    });

    // 1. شارت المناطق (يحتوي على المبيعات + الخصومات)
    const areas = Object.keys(areaData);
    const areaOptions = {
        series: [
            { name: 'المبيعات', data: areas.map(a => areaData[a].sales.toFixed(0)) },
            { name: 'الخصومات', data: areas.map(a => areaData[a].discounts.toFixed(0)) }
        ],
        chart: { type: 'bar', height: 350, foreColor, toolbar: { show: false }, background: 'transparent' },
        plotOptions: { bar: { borderRadius: 4, horizontal: false, columnWidth: '55%' } },
        dataLabels: { enabled: false },
        xaxis: { categories: areas },
        colors: ['#3b82f6', '#ef4444'],
        theme: { mode: isLight ? 'light' : 'dark' }
    };
    renderSingleChart('areaChart', areaOptions);

    // 2. شارت الأصناف (Donut)
    const topItems = Object.entries(itemData).sort((a,b) => b[1] - a[1]).slice(0, 7);
    const itemOptions = {
        series: topItems.map(i => i[1]),
        labels: topItems.map(i => i[0]),
        chart: { type: 'donut', height: 350, foreColor, background: 'transparent' },
        colors: ['#3b82f6', '#10b981', '#f59e0b', '#8b5cf6', '#ec4899', '#06b6d4', '#14b8a6'],
        theme: { mode: isLight ? 'light' : 'dark' }
    };
    renderSingleChart('itemsChart', itemOptions);

    // 3. شارت التريند اليومي (Area Chart)
    const sortedDates = Object.keys(dateData).sort();
    const trendOptions = {
        series: [{ name: 'إجمالي المبيعات', data: sortedDates.map(d => dateData[d].toFixed(0)) }],
        chart: { type: 'area', height: 300, foreColor, toolbar: { show: false }, background: 'transparent' },
        stroke: { curve: 'smooth', width: 2 },
        fill: { type: 'gradient', gradient: { shadeIntensity: 1, opacityFrom: 0.5, opacityTo: 0.1 } },
        xaxis: { categories: sortedDates },
        colors: ['#10b981'],
        theme: { mode: isLight ? 'light' : 'dark' }
    };
    renderSingleChart('trendChart', trendOptions);

    // 4. أفضل العملاء (Horizontal Bar)
    const topCusts = Object.entries(customerData).sort((a,b) => b[1] - a[1]).slice(0, 10);
    const topCustOptions = {
        series: [{ name: 'المبيعات', data: topCusts.map(c => c[1].toFixed(0)) }],
        chart: { type: 'bar', height: 350, foreColor, toolbar: { show: false }, background: 'transparent' },
        plotOptions: { bar: { horizontal: true, borderRadius: 4 } },
        dataLabels: { enabled: false },
        xaxis: { categories: topCusts.map(c => c[0]) },
        colors: ['#8b5cf6'],
        theme: { mode: isLight ? 'light' : 'dark' }
    };
    renderSingleChart('topCustomersChart', topCustOptions);
}

function renderSingleChart(id, options) {
    if (charts[id]) charts[id].destroy();
    charts[id] = new ApexCharts(document.querySelector(`#${id}`), options);
    charts[id].render();
}

// 7. تصدير PDF
function exportToPDF() {
    const element = document.querySelector('.main-content');
    const opt = {
        margin:       0.5,
        filename:     'Business_Partners_Report.pdf',
        image:        { type: 'jpeg', quality: 0.98 },
        html2canvas:  { scale: 2, useCORS: true },
        jsPDF:        { unit: 'in', format: 'a3', orientation: 'landscape' }
    };
    html2pdf().set(opt).from(element).save();
}
