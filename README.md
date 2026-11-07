<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>نظام إدارة ومناقلة مخزون الفروع - موارد الحجاز</title>
    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- FontAwesome Icons -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <!-- SheetJS (XLSX Parser) -->
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700;800;900&display=swap');
        body {
            font-family: 'Tajawal', sans-serif;
            background-color: #f8fafc;
        }
        .tab-active {
            border-bottom: 3px solid #2563eb;
            color: #2563eb;
            font-weight: 700;
            background-color: rgba(37, 99, 235, 0.04);
        }
        .badge-warning { background-color: #fef3c7; color: #92400e; border: 1px solid #fde68a; }
        .badge-danger { background-color: #fee2e2; color: #991b1b; border: 1px solid #fca5a5; }
        .badge-success { background-color: #dcfce7; color: #166534; border: 1px solid #86efac; }
        
        ::-webkit-scrollbar {
            width: 6px;
            height: 6px;
        }
        ::-webkit-scrollbar-track { background: #f1f5f9; }
        ::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 4px; }
        ::-webkit-scrollbar-thumb:hover { background: #94a3b8; }
    </style>
</head>
<body class="pb-24 md:pb-8 min-h-screen flex flex-col text-gray-800 select-none">

    <!-- Header -->
    <header class="bg-gradient-to-r from-blue-950 via-blue-900 to-indigo-900 text-white shadow-md sticky top-0 z-50">
        <div class="max-w-6xl mx-auto px-4 py-3.5 flex justify-between items-center">
            <div class="flex items-center space-x-3 space-x-reverse">
                <div class="bg-blue-600/90 p-2.5 rounded-xl shadow-md border border-blue-400/30">
                    <i class="fa-solid fa-boxes-stacked text-2xl text-white"></i>
                </div>
                <div>
                    <div class="flex items-center gap-2">
                        <h1 class="text-xl font-black tracking-tight leading-tight">موارد الحجاز</h1>
                        <span class="bg-emerald-500/20 text-emerald-300 text-[10px] px-2 py-0.5 rounded-md font-bold border border-emerald-500/30">إصدار الفروع المزدوج</span>
                    </div>
                    <p class="text-xs text-blue-200/90 font-medium">نظام إعادة توازن المخزون اللوجستي وتنسيق السائقين والمفاضلة</p>
                </div>
            </div>
            <div class="flex items-center gap-2">
                <span id="header-branch-count" class="bg-blue-800/80 text-blue-100 text-xs px-3.5 py-1.5 rounded-full font-bold border border-blue-700/60 shadow-sm flex items-center gap-1.5">
                    <i class="fa-solid fa-store text-emerald-400"></i>
                    <span>6 فروع معتمدة</span>
                </span>
            </div>
        </div>
    </header>

    <!-- Main Navigation Bar -->
    <nav class="bg-white border-b border-gray-200 shadow-sm sticky top-[65px] z-40">
        <div class="max-w-6xl mx-auto flex justify-around text-xs sm:text-sm font-medium text-gray-600">
            <button onclick="switchTab('uploadTab')" id="btn-uploadTab" class="tab-btn tab-active flex-1 py-3.5 px-2 text-center flex items-center justify-center gap-2 transition-all">
                <i class="fa-solid fa-file-excel text-base sm:text-lg"></i>
                <span>رفع وتأمين المخزون</span>
            </button>
            <button onclick="switchTab('branchesTab')" id="btn-branchesTab" class="tab-btn flex-1 py-3.5 px-2 text-center flex items-center justify-center gap-2 transition-all">
                <i class="fa-solid fa-store text-base sm:text-lg"></i>
                <span>إدارة الفروع المعتمدة</span>
            </button>
            <button onclick="switchTab('driversTab')" id="btn-driversTab" class="tab-btn flex-1 py-3.5 px-2 text-center flex items-center justify-center gap-2 transition-all relative">
                <i class="fa-solid fa-truck-fast text-base sm:text-lg"></i>
                <span>بوابة السائقين</span>
                <span id="pending-count" class="bg-amber-500 text-white text-[11px] font-black rounded-full h-5 w-5 flex items-center justify-center mr-1 shadow-sm">0</span>
            </button>
        </div>
    </nav>

    <!-- Main Container -->
    <main class="max-w-6xl mx-auto px-4 py-6 flex-grow w-full">

        <!-- TAB 1: UPLOAD & ANALYSIS -->
        <section id="uploadTab" class="tab-content space-y-6">
            
            <!-- Upload Mode Selector -->
            <div class="bg-white p-4 rounded-2xl shadow-sm border border-gray-100 flex flex-wrap items-center justify-between gap-3">
                <div class="flex items-center gap-2">
                    <i class="fa-solid fa-sliders text-blue-600"></i>
                    <span class="text-xs sm:text-sm font-extrabold text-gray-800">طريقة رفع بيانات الأرصدة للمفاضلة:</span>
                </div>
                <div class="flex items-center bg-gray-100 p-1 rounded-xl text-xs font-bold">
                    <button id="mode-single-btn" onclick="setUploadMode('single')" class="px-4 py-2 rounded-lg bg-white text-blue-700 shadow-sm transition-all">
                        <i class="fa-solid fa-file-lines ml-1"></i> ملف موحد لجميع الفروع
                    </button>
                    <button id="mode-multi-btn" onclick="setUploadMode('multi')" class="px-4 py-2 rounded-lg text-gray-600 transition-all">
                        <i class="fa-solid fa-folder-tree ml-1"></i> رفع ملف لكل فرع منفصل
                    </button>
                </div>
            </div>

            <!-- Single Upload Mode Section -->
            <div id="singleUploadSection" class="bg-white p-6 rounded-2xl shadow-sm border border-gray-100 text-center">
                <div class="border-2 border-dashed border-blue-200 hover:border-blue-500 transition-all rounded-2xl p-6 sm:p-8 bg-blue-50/40 hover:bg-blue-50/80 cursor-pointer group" onclick="document.getElementById('excelFileInput').click()">
                    <div class="w-16 h-16 bg-blue-100 text-blue-600 rounded-full flex items-center justify-center mx-auto mb-3 group-hover:scale-110 transition-transform shadow-inner">
                        <i class="fa-solid fa-cloud-arrow-up text-3xl"></i>
                    </div>
                    <h3 class="text-base font-extrabold text-gray-800 mb-1">إرفاق ملف Excel الموحد للأرصدة والفروع</h3>
                    <p class="text-xs text-gray-500 max-w-md mx-auto mb-4 leading-relaxed">يتضمن الحقول: <span class="text-blue-700 font-bold">GTIN، كود الصنف، اسم الصنف، الرصيد، الاستهلاك، كود الفرع</span></p>
                    <input type="file" id="excelFileInput" accept=".xlsx, .xls, .csv" class="hidden" onchange="handleSingleFileSelect(event)">
                    
                    <button class="bg-blue-600 hover:bg-blue-700 active:scale-95 text-white text-xs font-bold px-5 py-2.5 rounded-xl shadow-md transition-all inline-flex items-center gap-2">
                        <i class="fa-solid fa-folder-open"></i>
                        <span>رفع الملف الموحد</span>
                    </button>
                </div>
                
                <div class="mt-4 flex flex-wrap items-center justify-center gap-3">
                    <span class="text-xs text-gray-400 font-medium">أو يمكنك التجربة فوراً ببيانات نموذجية:</span>
                    <button onclick="loadDemoData()" class="bg-emerald-600 hover:bg-emerald-700 active:scale-95 text-white text-xs font-extrabold px-4 py-2.5 rounded-xl shadow-md transition-all inline-flex items-center gap-2">
                        <i class="fa-solid fa-bolt text-amber-300"></i>
                        <span>تشغيل ملف محاكاة آلي (بيانات نموذجية)</span>
                    </button>
                </div>
            </div>

            <!-- Multi Upload Mode Section (File per Branch) -->
            <div id="multiUploadSection" class="bg-white p-6 rounded-2xl shadow-sm border border-gray-100 hidden space-y-4">
                <div class="flex items-center justify-between border-b border-gray-100 pb-3">
                    <div>
                        <h3 class="text-sm sm:text-base font-black text-gray-800">رفع ملفات الإكسل لكل فرع على حدة</h3>
                        <p class="text-xs text-gray-500">اختر ملف إكسل الخاص بكل فرع بشكل منفصل، وسيقوم النظام بدمج الأرصدة والمفاضلة بينها.</p>
                    </div>
                    <button onclick="processMultiBranchAnalysis()" class="bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-extrabold px-4 py-2.5 rounded-xl shadow-md transition-all flex items-center gap-2">
                        <i class="fa-solid fa-calculator"></i>
                        <span>بدء تحليل الأرصدة والمفاضلة</span>
                    </button>
                </div>

                <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3" id="branchUploadCards">
                    <!-- Dynamic Multi Upload Cards per Branch -->
                </div>
            </div>

            <!-- Analysis Summary Cards -->
            <div id="summaryCards" class="grid grid-cols-2 sm:grid-cols-4 gap-3.5 hidden">
                <div class="bg-white p-4 rounded-2xl border border-gray-100 shadow-sm text-center">
                    <div class="w-8 h-8 bg-gray-100 text-gray-600 rounded-lg flex items-center justify-center mx-auto mb-2">
                        <i class="fa-solid fa-boxes-stacked text-sm"></i>
                    </div>
                    <p class="text-xs text-gray-500 font-semibold mb-0.5">إجمالي الأصناف</p>
                    <p id="stat-total-items" class="text-2xl font-black text-gray-800">0</p>
                </div>
                <div class="bg-white p-4 rounded-2xl border border-gray-100 shadow-sm text-center">
                    <div class="w-8 h-8 bg-red-100 text-red-600 rounded-lg flex items-center justify-center mx-auto mb-2">
                        <i class="fa-solid fa-triangle-exclamation text-sm"></i>
                    </div>
                    <p class="text-xs text-gray-500 font-semibold mb-0.5">فروع بها عجز (&lt;15 يوم)</p>
                    <p id="stat-deficit" class="text-2xl font-black text-red-600">0</p>
                </div>
                <div class="bg-white p-4 rounded-2xl border border-gray-100 shadow-sm text-center">
                    <div class="w-8 h-8 bg-amber-100 text-amber-600 rounded-lg flex items-center justify-center mx-auto mb-2">
                        <i class="fa-solid fa-chart-line text-sm"></i>
                    </div>
                    <p class="text-xs text-gray-500 font-semibold mb-0.5">فروع بها فائض (&gt;60 يوم)</p>
                    <p id="stat-surplus" class="text-2xl font-black text-amber-600">0</p>
                </div>
                <div class="bg-white p-4 rounded-2xl border border-gray-100 shadow-sm text-center">
                    <div class="w-8 h-8 bg-blue-100 text-blue-600 rounded-lg flex items-center justify-center mx-auto mb-2">
                        <i class="fa-solid fa-truck-ramp-box text-sm"></i>
                    </div>
                    <p class="text-xs text-gray-500 font-semibold mb-0.5">أوامر النقل الصادرة</p>
                    <p id="stat-orders" class="text-2xl font-black text-blue-600">0</p>
                </div>
            </div>

            <!-- Inventory Days of Coverage Table -->
            <div id="analysisSection" class="bg-white rounded-2xl shadow-sm border border-gray-100 overflow-hidden hidden">
                <div class="p-4 border-b border-gray-100 flex flex-wrap justify-between items-center gap-2 bg-gray-50/70">
                    <h3 class="font-black text-gray-800 text-sm sm:text-base flex items-center gap-2">
                        <i class="fa-solid fa-chart-pie text-blue-600"></i>
                        <span>نتائج تحليل التغطية وأيام المخزون بين الفروع</span>
                    </h3>
                    <span class="text-xs text-blue-700 bg-blue-50 border border-blue-100 font-bold px-3 py-1 rounded-full">معادلة التغطية: (الرصيد × 30) / الاستهلاك</span>
                </div>
                <div class="overflow-x-auto">
                    <table class="w-full text-right text-xs sm:text-sm">
                        <thead class="bg-gray-100/80 text-gray-700 font-extrabold border-b border-gray-200">
                            <tr>
                                <th class="p-3.5">الصنف وGTIN</th>
                                <th class="p-3.5">الفرع</th>
                                <th class="p-3.5">الرصيد الحالي</th>
                                <th class="p-3.5">الاستهلاك الشهري</th>
                                <th class="p-3.5">أيام التغطية</th>
                                <th class="p-3.5">حالة المخزون</th>
                            </tr>
                        </thead>
                        <tbody id="inventoryTableBody" class="divide-y divide-gray-100">
                            <!-- Dynamic Content -->
                        </tbody>
                    </table>
                </div>
            </div>

            <!-- Generated Transfer Orders -->
            <div id="transferOrdersSection" class="bg-white rounded-2xl shadow-sm border border-gray-100 p-5 hidden">
                <div class="mb-4 flex flex-wrap justify-between items-center gap-2">
                    <div>
                        <h3 class="font-black text-gray-800 text-sm sm:text-base flex items-center gap-2">
                            <i class="fa-solid fa-truck-arrow-right text-emerald-600"></i>
                            <span>أوامر التحويل والتنسيق الآلي الموصى بها</span>
                        </h3>
                        <p class="text-xs text-gray-500 mt-0.5">توليد تلقائي لسحب الفائض لتغطية الفروع ذات العجز مع حماية رصيد الأمان</p>
                    </div>
                    <span class="text-xs bg-emerald-50 text-emerald-700 border border-emerald-200 px-3 py-1 rounded-full font-bold flex items-center gap-1.5">
                        <i class="fa-solid fa-shield-halved"></i>
                        <span>رصيد الأمان مفعل</span>
                    </span>
                </div>
                <div id="generatedOrdersContainer" class="space-y-3">
                    <!-- Dynamic Orders -->
                </div>
            </div>

        </section>

        <!-- TAB 2: BRANCHES (MANAGEMENT) -->
        <section id="branchesTab" class="tab-content hidden space-y-4">
            <div class="bg-white p-5 rounded-2xl shadow-sm border border-gray-100 flex flex-wrap items-center justify-between gap-3">
                <div>
                    <h3 class="font-black text-gray-800 text-base mb-0.5">إدارة أفرع موارد الحجاز المعتمدة</h3>
                    <p class="text-xs text-gray-500">يمكنك تعديل بيانات الفروع الحالية، أو إضافة فرع جديد بسهولة للنظام.</p>
                </div>
                <button onclick="openAddBranchModal()" class="bg-blue-600 hover:bg-blue-700 active:scale-95 text-white text-xs font-bold px-4 py-2.5 rounded-xl shadow-md transition-all flex items-center gap-2">
                    <i class="fa-solid fa-plus"></i>
                    <span>إضافة فرع جديد</span>
                </button>
            </div>

            <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4" id="branchesList">
                <!-- Dynamic Branch Cards -->
            </div>
        </section>

        <!-- TAB 3: DRIVERS PORTAL -->
        <section id="driversTab" class="tab-content hidden space-y-4">
            <div class="bg-amber-50/80 border border-amber-200/80 p-4 rounded-2xl flex items-start gap-3 shadow-sm">
                <div class="bg-amber-500 text-white p-2 rounded-xl text-sm mt-0.5 shadow-sm">
                    <i class="fa-solid fa-hand-pointer"></i>
                </div>
                <div class="text-xs text-amber-900 leading-relaxed">
                    <p class="font-extrabold text-sm mb-0.5">بوابة السائقين لطلب الاستلام الفوري (Claim-based System)</p>
                    <p>الطلبات المعروضة أدناه معلقة وفي انتظار استلامها من أول سائق متاح بالفرع لمباشرة التوصيل.</p>
                </div>
            </div>

            <div class="bg-white p-5 rounded-2xl shadow-sm border border-gray-100">
                <div class="flex items-center justify-between mb-4 border-b border-gray-100 pb-3">
                    <h3 class="font-black text-gray-800 text-base flex items-center gap-2">
                        <i class="fa-solid fa-list-check text-blue-600"></i>
                        <span>قائمة الطلبات المعلقة (Driver Order Pool)</span>
                    </h3>
                    <span id="active-orders-badge" class="bg-red-100 text-red-700 text-xs px-3 py-1 rounded-full font-bold border border-red-200">0 طلبات معلقة</span>
                </div>

                <div id="driverOrdersContainer" class="space-y-3">
                    <!-- Dynamic Driver Orders -->
                </div>
            </div>
        </section>

    </main>

    <!-- Modal for Add/Edit Branch -->
    <div id="branchModal" class="fixed inset-0 bg-black/50 z-50 hidden flex items-center justify-center p-4 backdrop-blur-sm">
        <div class="bg-white rounded-2xl max-w-md w-full p-6 shadow-2xl space-y-4">
            <div class="flex justify-between items-center border-b border-gray-100 pb-3">
                <h3 id="modalTitle" class="font-black text-gray-800 text-base">إضافة فرع جديد</h3>
                <button onclick="closeBranchModal()" class="text-gray-400 hover:text-gray-600"><i class="fa-solid fa-xmark text-lg"></i></button>
            </div>
            <input type="hidden" id="editIndex" value="-1">
            <div class="space-y-3 text-xs">
                <div>
                    <label class="block text-gray-700 font-bold mb-1">كود الفرع (مثال: BR-008)</label>
                    <input type="text" id="branchCodeInput" class="w-full border border-gray-200 rounded-xl p-2.5 focus:outline-none focus:border-blue-500">
                </div>
                <div>
                    <label class="block text-gray-700 font-bold mb-1">اسم الفرع</label>
                    <input type="text" id="branchNameInput" class="w-full border border-gray-200 rounded-xl p-2.5 focus:outline-none focus:border-blue-500">
                </div>
                <div>
                    <label class="block text-gray-700 font-bold mb-1">مسؤول الفرع</label>
                    <input type="text" id="branchManagerInput" class="w-full border border-gray-200 rounded-xl p-2.5 focus:outline-none focus:border-blue-500">
                </div>
                <div>
                    <label class="block text-gray-700 font-bold mb-1">رقم الجوال</label>
                    <input type="text" id="branchPhoneInput" class="w-full border border-gray-200 rounded-xl p-2.5 focus:outline-none focus:border-blue-500">
                </div>
                <div>
                    <label class="block text-gray-700 font-bold mb-1">رابط Google Maps</label>
                    <input type="text" id="branchMapInput" class="w-full border border-gray-200 rounded-xl p-2.5 focus:outline-none focus:border-blue-500">
                </div>
            </div>
            <div class="flex justify-end gap-2 pt-3 border-t border-gray-100">
                <button onclick="closeBranchModal()" class="bg-gray-100 hover:bg-gray-200 text-gray-700 font-bold text-xs px-4 py-2.5 rounded-xl transition-all">إلغاء</button>
                <button onclick="saveBranch()" class="bg-blue-600 hover:bg-blue-700 text-white font-bold text-xs px-5 py-2.5 rounded-xl shadow-md transition-all">حفظ الفرع</button>
            </div>
        </div>
    </div>

    <!-- Footer -->
    <footer class="bg-white border-t border-gray-200 mt-auto py-4 text-center text-xs text-gray-500">
        <div class="max-w-6xl mx-auto px-4 flex flex-col sm:flex-row justify-between items-center gap-2">
            <p class="font-medium">نظام إدارة ومناقلة مخزون الفروع © موارد الحجاز</p>
            <p class="text-gray-400">الإصدار التفاعلي المتقدم 3.0</p>
        </div>
    </footer>

    <!-- Global Toast / Notification -->
    <div id="toastNotification" class="fixed bottom-6 left-1/2 -translate-x-1/2 bg-gray-950/90 text-white text-xs sm:text-sm font-bold px-5 py-3 rounded-2xl shadow-2xl flex items-center gap-3 transition-all duration-300 opacity-0 pointer-events-none z-50 border border-gray-700/50 backdrop-blur-md">
        <i id="toastIcon" class="fa-solid fa-circle-check text-emerald-400 text-base sm:text-lg"></i>
        <span id="toastMessage">تمت العملية بنجاح</span>
    </div>

    <script>
        // Data Store - Branches Data
        let branches = [
            { id: 'BR-001', name: 'الحمدانية 1', mapUrl: 'https://maps.google.com/?q=Al-Hamdaniya+Jeddah', code: 'BR-001', manager: 'أحمد الغامدي', phone: '0501112233' },
            { id: 'BR-003', name: 'النزهه حراء 3', mapUrl: 'https://maps.google.com/?q=Al-Nuzha+Hira+Jeddah', code: 'BR-003', manager: 'محمد الشهري', phone: '0502223344' },
            { id: 'BR-004', name: 'فرع الفيصلية 4', mapUrl: 'https://maps.google.com/?q=Al-Faisaliah+Jeddah', code: 'BR-004', manager: 'خالد الحربي', phone: '0503334455' },
            { id: 'BR-005', name: 'المروه 5', mapUrl: 'https://maps.google.com/?q=Al-Marwa+Jeddah', code: 'BR-005', manager: 'سعيد الزهراني', phone: '0504445566' },
            { id: 'BR-006', name: 'الصفا 6', mapUrl: 'https://maps.google.com/?q=Al-Safa+Jeddah', code: 'BR-006', manager: 'عمر القحطاني', phone: '0505556677' },
            { id: 'BR-007', name: 'الصالحيه 7', mapUrl: 'https://maps.google.com/?q=Al-Salihiyah+Jeddah', code: 'BR-007', manager: 'فهد المطيري', phone: '0506667788' }
        ];

        let uploadMode = 'single'; // 'single' or 'multi'
        let multiBranchData = {}; // Stores items per branch
        let inventoryItems = [];
        let transferOrders = [];

        function setUploadMode(mode) {
            uploadMode = mode;
            if (mode === 'single') {
                document.getElementById('mode-single-btn').className = 'px-4 py-2 rounded-lg bg-white text-blue-700 shadow-sm transition-all';
                document.getElementById('mode-multi-btn').className = 'px-4 py-2 rounded-lg text-gray-600 transition-all';
                document.getElementById('singleUploadSection').classList.remove('hidden');
                document.getElementById('multiUploadSection').classList.add('hidden');
            } else {
                document.getElementById('mode-multi-btn').className = 'px-4 py-2 rounded-lg bg-white text-blue-700 shadow-sm transition-all';
                document.getElementById('mode-single-btn').className = 'px-4 py-2 rounded-lg text-gray-600 transition-all';
                document.getElementById('singleUploadSection').classList.add('hidden');
                document.getElementById('multiUploadSection').classList.remove('hidden');
                renderMultiBranchCards();
            }
        }

        function switchTab(tabId) {
            document.querySelectorAll('.tab-content').forEach(el => el.classList.add('hidden'));
            document.querySelectorAll('.tab-btn').forEach(el => el.classList.remove('tab-active'));
            
            document.getElementById(tabId).classList.remove('hidden');
            document.getElementById('btn-' + tabId).classList.add('tab-active');
        }

        function renderBranches() {
            const container = document.getElementById('branchesList');
            document.getElementById('header-branch-count').innerHTML = `<i class="fa-solid fa-store text-emerald-400"></i> <span>${branches.length} فروع معتمدة</span>`;
            
            container.innerHTML = branches.map((b, index) => `
                <div class="bg-white p-4 rounded-2xl border border-gray-100 shadow-sm hover:shadow-md transition-shadow flex flex-col justify-between">
                    <div>
                        <div class="flex items-center justify-between mb-3">
                            <span class="bg-blue-50 text-blue-700 border border-blue-100 px-2.5 py-1 rounded-lg font-black text-xs">
                                ${b.id}
                            </span>
                            <div class="flex items-center gap-1">
                                <button onclick="editBranch(${index})" class="text-gray-400 hover:text-blue-600 p-1 rounded-lg transition-colors" title="تعديل الفرع"><i class="fa-solid fa-pen-to-square"></i></button>
                                <button onclick="deleteBranch(${index})" class="text-gray-400 hover:text-red-600 p-1 rounded-lg transition-colors" title="حذف الفرع"><i class="fa-solid fa-trash-can"></i></button>
                            </div>
                        </div>
                        <h4 class="font-black text-gray-800 text-base mb-1">${b.name}</h4>
                        <div class="text-xs text-gray-500 space-y-1 mb-4">
                            <p class="flex items-center gap-1.5"><i class="fa-solid fa-user-tie text-gray-400"></i> <span>المسؤول: <strong>${b.manager}</strong></span></p>
                            <p class="flex items-center gap-1.5"><i class="fa-solid fa-phone text-gray-400"></i> <span dir="ltr">${b.phone}</span></p>
                        </div>
                    </div>
                    <a href="${b.mapUrl}" target="_blank" class="w-full bg-blue-50 hover:bg-blue-100 text-blue-700 text-xs font-bold py-2.5 px-3 rounded-xl transition-colors flex items-center justify-center gap-2 border border-blue-100">
                        <i class="fa-solid fa-location-dot"></i>
                        <span>عرض الفرع على الخريطة</span>
                    </a>
                </div>
            `).join('');
            
            if (uploadMode === 'multi') renderMultiBranchCards();
        }

        function renderMultiBranchCards() {
            const container = document.getElementById('branchUploadCards');
            container.innerHTML = branches.map(b => {
                const isLoaded = multiBranchData[b.name] && multiBranchData[b.name].length > 0;
                return `
                    <div class="p-3.5 rounded-xl border ${isLoaded ? 'border-emerald-200 bg-emerald-50/30' : 'border-gray-200 bg-gray-50/50'} flex items-center justify-between">
                        <div>
                            <span class="text-[10px] font-bold text-gray-400">${b.id}</span>
                            <h4 class="font-bold text-xs text-gray-800">${b.name}</h4>
                            <p class="text-[11px] ${isLoaded ? 'text-emerald-700 font-bold' : 'text-gray-400'}">
                                ${isLoaded ? `✓ تم رفع ${multiBranchData[b.name].length} صنف` : 'لم يتم اختيار ملف'}
                            </p>
                        </div>
                        <input type="file" id="file-${b.id}" accept=".xlsx, .xls, .csv" class="hidden" onchange="handleMultiFileSelect(event, '${b.name}')">
                        <button onclick="document.getElementById('file-${b.id}').click()" class="bg-white border border-gray-200 hover:bg-blue-50 hover:border-blue-300 text-gray-700 text-xs font-bold px-3 py-2 rounded-lg transition-all flex items-center gap-1.5 shadow-sm">
                            <i class="fa-solid fa-file-arrow-up text-blue-600"></i>
                            <span>${isLoaded ? 'تغيير' : 'إرفاق'}</span>
                        </button>
                    </div>
                `;
            }).join('');
        }

        function openAddBranchModal() {
            document.getElementById('modalTitle').innerText = 'إضافة فرع جديد';
            document.getElementById('editIndex').value = '-1';
            document.getElementById('branchCodeInput').value = `BR-00${branches.length + 1}`;
            document.getElementById('branchNameInput').value = '';
            document.getElementById('branchManagerInput').value = '';
            document.getElementById('branchPhoneInput').value = '';
            document.getElementById('branchMapInput').value = '';
            document.getElementById('branchModal').classList.remove('hidden');
        }

        function editBranch(index) {
            const b = branches[index];
            document.getElementById('modalTitle').innerText = 'تعديل بيانات الفرع';
            document.getElementById('editIndex').value = index;
            document.getElementById('branchCodeInput').value = b.id;
            document.getElementById('branchNameInput').value = b.name;
            document.getElementById('branchManagerInput').value = b.manager;
            document.getElementById('branchPhoneInput').value = b.phone;
            document.getElementById('branchMapInput').value = b.mapUrl;
            document.getElementById('branchModal').classList.remove('hidden');
        }

        function deleteBranch(index) {
            const b = branches[index];
            delete multiBranchData[b.name];
            branches.splice(index, 1);
            renderBranches();
            showToast('تم حذف الفرع بنجاح');
        }

        function closeBranchModal() {
            document.getElementById('branchModal').classList.add('hidden');
        }

        function saveBranch() {
            const index = parseInt(document.getElementById('editIndex').value);
            const id = document.getElementById('branchCodeInput').value.trim();
            const name = document.getElementById('branchNameInput').value.trim();
            const manager = document.getElementById('branchManagerInput').value.trim();
            const phone = document.getElementById('branchPhoneInput').value.trim();
            const mapUrl = document.getElementById('branchMapInput').value.trim();

            if (!id || !name) {
                showToast('يرجى إدخال كود الفرع واسمه على الأقل', true);
                return;
            }

            const branchData = { id, name, manager: manager || 'غير محدد', phone: phone || '-', mapUrl: mapUrl || 'https://maps.google.com' };

            if (index === -1) {
                branches.push(branchData);
                showToast('تم إضافة الفرع الجديد بنجاح!');
            } else {
                branches[index] = branchData;
                showToast('تم تحديث بيانات الفرع بنجاح!');
            }

            closeBranchModal();
            renderBranches();
        }

        function handleSingleFileSelect(event) {
            const file = event.target.files[0];
            if (!file) return;

            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    const jsonData = XLSX.utils.sheet_to_json(firstSheet);

                    if (jsonData.length > 0) {
                        inventoryItems = jsonData.map(row => ({
                            gtin: String(row['GTIN'] || row['رمز GTIN'] || row['gtin'] || '628110099999'),
                            code: String(row['ItemCode'] || row['كود الصنف'] || row['code'] || 'MED-000'),
                            name: String(row['ItemName'] || row['اسم الصنف'] || row['name'] || 'صنف طبي'),
                            branch: String(row['BranchCode'] || row['الفرع'] || row['branch'] || 'فرع الفيصلية 4'),
                            stock: parseInt(row['Stock'] || row['الرصيد'] || row['stock'] || 0),
                            consumption: parseInt(row['MonthlyConsumption'] || row['الاستهلاك الشهري'] || row['consumption'] || 1)
                        }));
                        
                        generateOrdersFromUpload();
                        processAnalysis();
                        showToast('تم قراءة الملف الموحد وقابليته للمفاضلة بنجاح!');
                    }
                } catch(err) {
                    showToast('حدث خطأ أثناء قراءة الملف، يرجى التأكد من التنسيق.', true);
                }
            };
            reader.readAsArrayBuffer(file);
        }

        function handleMultiFileSelect(event, branchName) {
            const file = event.target.files[0];
            if (!file) return;

            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    const jsonData = XLSX.utils.sheet_to_json(firstSheet);

                    multiBranchData[branchName] = jsonData.map(row => ({
                        gtin: String(row['GTIN'] || row['رمز GTIN'] || row['gtin'] || '628110099999'),
                        code: String(row['ItemCode'] || row['كود الصنف'] || row['code'] || 'MED-000'),
                        name: String(row['ItemName'] || row['اسم الصنف'] || row['name'] || 'صنف طبي'),
                        branch: branchName,
                        stock: parseInt(row['Stock'] || row['الرصيد'] || row['stock'] || 0),
                        consumption: parseInt(row['MonthlyConsumption'] || row['الاستهلاك الشهري'] || row['consumption'] || 1)
                    }));

                    renderMultiBranchCards();
                    showToast(`تم إرفاق ملف ${branchName} بنجاح!`);
                } catch(err) {
                    showToast(`خطأ في قراءة ملف الفرع ${branchName}`, true);
                }
            };
            reader.readAsArrayBuffer(file);
        }

        function processMultiBranchAnalysis() {
            inventoryItems = [];
            Object.values(multiBranchData).forEach(items => {
                inventoryItems.push(...items);
            });

            if (inventoryItems.length === 0) {
                showToast('يرجى إرفاق ملف إكسل لفرع واحد على الأقل قبل التحليل.', true);
                return;
            }

            generateOrdersFromUpload();
            processAnalysis();
            showToast('تم دمج ملفات الفروع وإجراء المفاضلة بنجاح!');
        }

        function loadDemoData() {
            inventoryItems = [
                { gtin: '6281100123451', code: 'MED-101', name: 'مضاد حيوي أوجمنتين 1 غرام', branch: 'فرع الفيصلية 4', stock: 180, consumption: 20 },
                { gtin: '6281100123451', code: 'MED-101', name: 'مضاد حيوي أوجمنتين 1 غرام', branch: 'المروه 5', stock: 4, consumption: 25 },
                { gtin: '6281100123452', code: 'MED-102', name: 'مسكن بانادول أدفانس 500 مغ', branch: 'الحمدانية 1', stock: 350, consumption: 40 },
                { gtin: '6281100123452', code: 'MED-102', name: 'مسكن بانادول أدفانس 500 مغ', branch: 'النزهه حراء 3', stock: 6, consumption: 30 },
                { gtin: '6281100123453', code: 'MED-103', name: 'فيتامين سي فوار 1000 مغ', branch: 'الصفا 6', stock: 240, consumption: 30 },
                { gtin: '6281100123453', code: 'MED-103', name: 'فيتامين سي فوار 1000 مغ', branch: 'الصالحيه 7', stock: 10, consumption: 28 }
            ];

            transferOrders = [
                { id: 'ORD-901', gtin: '6281100123451', itemName: 'مضاد حيوي أوجمنتين 1 غرام', from: 'فرع الفيصلية 4', to: 'المروه 5', qty: 45, status: 'معلق' },
                { id: 'ORD-902', gtin: '6281100123452', itemName: 'مسكن بانادول أدفانس 500 مغ', from: 'الحمدانية 1', to: 'النزهه حراء 3', qty: 60, status: 'معلق' },
                { id: 'ORD-903', gtin: '6281100123453', itemName: 'فيتامين سي فوار 1000 مغ', from: 'الصفا 6', to: 'الصالحيه 7', qty: 35, status: 'معلق' }
            ];

            processAnalysis();
            showToast('تم إجراء المحاكاة وتوليد المفاضلة الآلية بنجاح!');
        }

        function generateOrdersFromUpload() {
            transferOrders = [];
            let count = 1;
            inventoryItems.forEach(item => {
                const days = Math.round((item.stock / (item.consumption || 1)) * 30);
                if (days < 15) {
                    const surplus = inventoryItems.find(i => i.gtin === item.gtin && ((i.stock / (i.consumption||1))*30) > 45);
                    if (surplus) {
                        transferOrders.push({
                            id: `ORD-${900 + count++}`,
                            gtin: item.gtin,
                            itemName: item.name,
                            from: surplus.branch,
                            to: item.branch,
                            qty: Math.round(surplus.stock * 0.35),
                            status: 'معلق'
                        });
                    }
                }
            });
        }

        function processAnalysis() {
            document.getElementById('summaryCards').classList.remove('hidden');
            document.getElementById('analysisSection').classList.remove('hidden');
            document.getElementById('transferOrdersSection').classList.remove('hidden');

            let totalDeficit = 0;
            let totalSurplus = 0;

            const tbody = document.getElementById('inventoryTableBody');
            tbody.innerHTML = inventoryItems.map(item => {
                const days = Math.round((item.stock / (item.consumption || 1)) * 30);
                let badgeClass = 'badge-success';
                let statusText = 'مخزون طبيعي';
                
                if (days < 15) {
                    badgeClass = 'badge-danger';
                    statusText = 'عجز حاد (< 15 يوم)';
                    totalDeficit++;
                } else if (days > 60) {
                    badgeClass = 'badge-warning';
                    statusText = 'فائض مرتفع (> 60 يوم)';
                    totalSurplus++;
                }

                return `
                    <tr class="hover:bg-blue-50/40 transition-colors">
                        <td class="p-3.5">
                            <div class="font-extrabold text-gray-800">${item.name}</div>
                            <div class="text-[11px] text-gray-400 font-medium">GTIN: ${item.gtin} | كود: ${item.code}</div>
                        </td>
                        <td class="p-3.5 font-bold text-gray-700">${item.branch}</td>
                        <td class="p-3.5 font-extrabold text-gray-900">${item.stock.toLocaleString()}</td>
                        <td class="p-3.5 text-gray-600 font-medium">${item.consumption.toLocaleString()} / شهر</td>
                        <td class="p-3.5 font-black text-sm ${days < 15 ? 'text-red-600' : (days > 60 ? 'text-amber-600' : 'text-emerald-600')}">${days} يوم</td>
                        <td class="p-3.5">
                            <span class="px-2.5 py-1 rounded-full text-[11px] font-bold inline-block ${badgeClass}">${statusText}</span>
                        </td>
                    </tr>
                `;
            }).join('');

            document.getElementById('stat-total-items').innerText = inventoryItems.length;
            document.getElementById('stat-deficit').innerText = totalDeficit;
            document.getElementById('stat-surplus').innerText = totalSurplus;
            document.getElementById('stat-orders').innerText = transferOrders.length;

            renderGeneratedOrders();
            renderDriverOrders();
        }

        function renderGeneratedOrders() {
            const container = document.getElementById('generatedOrdersContainer');
            if (transferOrders.length === 0) {
                container.innerHTML = `<p class="text-xs text-gray-400 text-center py-6">لا توجد حاجة لمناقلات حالياً، المخزون متوازن بكافة الفروع.</p>`;
                return;
            }

            container.innerHTML = transferOrders.map(o => `
                <div class="p-4 border border-blue-100 bg-blue-50/30 rounded-2xl flex flex-wrap items-center justify-between gap-3 hover:border-blue-200 transition-colors">
                    <div>
                        <div class="flex items-center gap-2 mb-1">
                            <span class="font-black text-blue-900 bg-blue-100 px-2.5 py-0.5 rounded-md text-xs">${o.id}</span>
                            <span class="text-xs font-bold text-gray-800">${o.itemName}</span>
                        </div>
                        <div class="text-xs text-gray-600 flex items-center gap-2">
                            <span>من (فائض): <strong class="text-gray-800">${o.from}</strong></span>
                            <i class="fa-solid fa-arrow-left text-blue-500 text-[10px]"></i>
                            <span>إلى (عجز): <strong class="text-gray-800">${o.to}</strong></span>
                        </div>
                    </div>
                    <div class="text-left flex items-center gap-3">
                        <span class="bg-blue-600 text-white font-black text-xs px-3 py-1.5 rounded-xl shadow-sm inline-block">${o.qty} قطعة</span>
                    </div>
                </div>
            `).join('');
        }

        function renderDriverOrders() {
            const container = document.getElementById('driverOrdersContainer');
            const pendingOrders = transferOrders.filter(o => o.status === 'معلق');
            
            document.getElementById('pending-count').innerText = pendingOrders.length;
            document.getElementById('active-orders-badge').innerText = `${pendingOrders.length} طلبات معلقة`;

            if (transferOrders.length === 0) {
                container.innerHTML = `<p class="text-xs text-gray-400 text-center py-8">قم برفع ملفات الأرصدة أو تشغيل المحاكاة أولاً لتوليد الطلبات للسائقين.</p>`;
                return;
            }

            container.innerHTML = transferOrders.map(o => `
                <div class="p-4 border border-gray-200 rounded-2xl bg-white shadow-sm hover:border-blue-300 transition-all">
                    <div class="flex justify-between items-start mb-3">
                        <div>
                            <span class="text-xs font-black text-blue-700 bg-blue-50 border border-blue-100 px-2.5 py-0.5 rounded-md">${o.id}</span>
                            <h4 class="font-extrabold text-gray-800 text-sm mt-1.5">${o.itemName}</h4>
                        </div>
                        <span class="text-xs font-bold px-3 py-1 rounded-full ${o.status === 'معلق' ? 'bg-amber-100 text-amber-800 border border-amber-200' : 'bg-emerald-100 text-emerald-800 border border-emerald-200'}">
                            ${o.status === 'معلق' ? '⏳ معلق (في انتظار سائق)' : '🚚 قيد التوصيل (تم الاستلام)'}
                        </span>
                    </div>

                    <div class="bg-gray-50 p-3 rounded-xl text-xs text-gray-600 space-y-1.5 mb-3 border border-gray-100">
                        <div class="flex items-center justify-between">
                            <span class="flex items-center gap-1.5"><i class="fa-solid fa-circle-arrow-up text-blue-500"></i> فرع المصدر (فائض):</span>
                            <strong class="text-gray-800 font-bold">${o.from}</strong>
                        </div>
                        <div class="flex items-center justify-between">
                            <span class="flex items-center gap-1.5"><i class="fa-solid fa-circle-arrow-down text-emerald-500"></i> فرع الوجهة (عجز):</span>
                            <strong class="text-gray-800 font-bold">${o.to}</strong>
                        </div>
                        <div class="flex items-center justify-between border-t border-gray-200/60 pt-1.5 mt-1.5">
                            <span class="flex items-center gap-1.5"><i class="fa-solid fa-box text-amber-500"></i> الكمية المطلوبة للمناقلة:</span>
                            <strong class="text-blue-900 font-black text-sm">${o.qty} قطعة</strong>
                        </div>
                    </div>

                    ${o.status === 'معلق' ? `
                        <button onclick="claimOrder('${o.id}')" class="w-full bg-amber-500 hover:bg-amber-600 active:scale-95 text-white font-black text-xs py-2.5 rounded-xl transition-all shadow-md flex items-center justify-center gap-2">
                            <i class="fa-solid fa-hand-pointer"></i>
                            <span>حجز واستلام الطلب فوراً (Claim Order)</span>
                        </button>
                    ` : `
                        <div class="text-xs text-emerald-800 font-bold bg-emerald-50 border border-emerald-200 p-2.5 rounded-xl text-center flex items-center justify-center gap-2">
                            <i class="fa-solid fa-circle-check text-emerald-600"></i>
                            <span>تم حجز المهمة بنجاح، جاري النقل والتوصيل للفرع المستلم</span>
                        </div>
                    `}
                </div>
            `).join('');
        }

        function claimOrder(orderId) {
            const order = transferOrders.find(o => o.id === orderId);
            if (order) {
                order.status = 'مقبول';
                renderDriverOrders();
                showToast(`تم حجز واستلام الطلب (${orderId}) بنجاح!`);
            }
        }

        function showToast(msg, isError = false) {
            const toast = document.getElementById('toastNotification');
            const toastMsg = document.getElementById('toastMessage');
            const toastIcon = document.getElementById('toastIcon');

            toastMsg.innerText = msg;
            toastIcon.className = isError ? "fa-solid fa-circle-xmark text-red-400 text-lg" : "fa-solid fa-circle-check text-emerald-400 text-lg";
            
            toast.classList.remove('opacity-0', 'pointer-events-none');
            setTimeout(() => {
                toast.classList.add('opacity-0', 'pointer-events-none');
            }, 3500);
        }

        window.onload = function() {
            renderBranches();
            loadDemoData();
        };
    </script>
</body>
</html>
