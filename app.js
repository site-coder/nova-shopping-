/**
 * ==========================================================================
 * اسم المشروع: متجر Nova Shopping الإلكتروني الفاخر
 * الملف: app.js
 * الوصف: المحرك البرمجي المسؤول عن إدارة حالة التطبيق وعرض تفاصيل المنتجات
 * وجلب ومعالجة المنتجات من ملف Excel وإدارة واجهة الشحن والتوصيل الجزائري.
 * المطور: مصمم ومطور واجهات المستخدم الفاخرة
 * الإصدار: 1.3.0
 * ==========================================================================
 */

"use strict";

const AppState = {
    config: {
        storeName: "Nova Shopping",
        currency: "د.ج",
        excelFilePath: "products.xlsx", // ملف Excel الخاص بك (ضعه بنفس المجلد)
        sheetJsUrl: "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"
    },
    state: {
        products: [], // قائمة المنتجات المجلوبة من الـ Excel أو السجل الاحتياطي
        currentModalProductId: null, // المنتج المعروض حالياً داخل نافذة التفاصيل
        order: null // المنتج المحدد حالياً للطلب (بدون نظام سلة)
    },
    // قاعدة بيانات الولايات الـ 58 مع أسعار التوصيل المحدّثة للمكتب والمنزل بالدينار الجزائري
    wilayas: [
        { code: "01", name: "01 - أدرار (Adrar)", home: 1200, office: 600 },
        { code: "02", name: "02 - الشلف (Chlef)", home: 700, office: 400 },
        { code: "03", name: "03 - الأغواط (Laghouat)", home: 800, office: 400 },
        { code: "04", name: "04 - أم البواقي (Oum El Bouaghi)", home: 700, office: 400 },
        { code: "05", name: "05 - باتنة (Batna)", home: 700, office: 400 },
        { code: "06", name: "06 - بجاية (Bejaia)", home: 700, office: 400 },
        { code: "07", name: "07 - بسكرة (Biskra)", home: 800, office: 400 },
        { code: "08", name: "08 - بشار (Bechar)", home: 900, office: 600 },
        { code: "09", name: "09 - البليدة (Blida)", home: 500, office: 400 },
        { code: "10", name: "10 - البويرة (Bouira)", home: 700, office: 400 },
        { code: "11", name: "11 - تمنراست (Tamanrasset)", home: 1300, office: 800 },
        { code: "12", name: "12 - تبسة (Tebessa)", home: 700, office: 400 },
        { code: "13", name: "13 - تلمسان (Tlemcen)", home: 700, office: 400 },
        { code: "14", name: "14 - تيارت (Tiaret)", home: 700, office: 400 },
        { code: "15", name: "15 - تيزي وزو (Tizi Ouzou)", home: 700, office: 400 },
        { code: "16", name: "16 - الجزائر العاصمة (Alger)", home: 500, office: 500 },
        { code: "17", name: "17 - الجلفة (Djelfa)", home: 800, office: 500 },
        { code: "18", name: "18 - جيجل (Jijel)", home: 700, office: 400 },
        { code: "19", name: "19 - سطيف (Setif)", home: 700, office: 400 },
        { code: "20", name: "20 - سعيدة (Saida)", home: 700, office: 400 },
        { code: "21", name: "21 - سكيكدة (Skikda)", home: 700, office: 400 },
        { code: "22", name: "22 - سيدي بلعباس (Sidi Bel Abbes)", home: 700, office: 400 },
        { code: "23", name: "23 - عنابة (Annaba)", home: 700, office: 400 },
        { code: "24", name: "24 - قالمة (Guelma)", home: 700, office: 400 },
        { code: "25", name: "25 - قسنطينة (Constantine)", home: 700, office: 400 },
        { code: "26", name: "26 - المدية (Medea)", home: 700, office: 400 },
        { code: "27", name: "27 - مستغانم (Mostaganem)", home: 700, office: 400 },
        { code: "28", name: "28 - المسيلة (M'Sila)", home: 700, office: 400 },
        { code: "29", name: "29 - معسكر (Mascara)", home: 700, office: 400 },
        { code: "30", name: "30 - ورقلة (Ouargla)", home: 900, office: 500 },
        { code: "31", name: "31 - وهران (Oran)", home: 700, office: 400 },
        { code: "32", name: "32 - البيض (El Bayadh)", home: 1000, office: 600 },
        { code: "33", name: "33 - إليزي (Illizi)", home: 1100, office: 700 },
        { code: "34", name: "34 - برج بوعريريج (Bordj Bou Arreridj)", home: 700, office: 400 },
        { code: "35", name: "35 - بومرداس (Boumerdes)", home: 500, office: 400 },
        { code: "36", name: "36 - الطارف (El Tarf)", home: 700, office: 400 },
        { code: "37", name: "37 - تيندوف (Tindouf)", home: 1200, office: 700 },
        { code: "38", name: "38 - تسمسيلت (Tissemsilt)", home: 700, office: 400 },
        { code: "39", name: "39 - الوادي (El Oued)", home: 900, office: 600 },
        { code: "40", name: "40 - خنشلة (Khenchela)", home: 700, office: 400 },
        { code: "41", name: "41 - سوق أهراس (Souk Ahras)", home: 700, office: 400 },
        { code: "42", name: "42 - تيبازة (Tipaza)", home: 500, office: 400 },
        { code: "43", name: "43 - ميلة (Mila)", home: 700, office: 400 },
        { code: "44", name: "44 - عين الدفلى (Ain Defla)", home: 700, office: 400 },
        { code: "45", name: "45 - النعامة (Naama)", home: 900, office: 600 },
        { code: "46", name: "46 - عين تموشنت (Ain Temouchent)", home: 600, office: 400 },
        { code: "47", name: "47 - غرداية (Ghardaia)", home: 800, office: 500 },
        { code: "48", name: "48 - غليزان (Relizane)", home: 700, office: 400 },
        { code: "49", name: "49 - المغير (El M'ghair)", home: 900, office: 900 },
        { code: "50", name: "50 - المنيعة (El Meniaa)", home: 900, office: 900 },
        { code: "51", name: "51 - أولاد جلال (Ouled Djellal)", home: 800, office: 600 },
        { code: "52", name: "52 - برج باجي مختار (Bordj Baji Mokhtar)", home: 1200, office: 800 },
        { code: "53", name: "53 - بني عباس (Beni Abbes)", home: 1000, office: 1000 },
        { code: "54", name: "54 - تيميمون (Timimoun)", home: 1200, office: 700 },
        { code: "55", name: "55 - تقرت (Touggourt)", home: 800, office: 600 },
        { code: "56", name: "56 - جانت (Djanet)", home: 1200, office: 800 },
        { code: "57", name: "57 - عين صالح (In Salah)", home: 1300, office: 900 },
        { code: "58", name: "58 - عين قزام (In Guezzam)", home: 1200, office: 800 }
    ],
    // المنتجات الاحتياطية لضمان عمل الواجهة مباشرة في حال غياب ملف Excel
    fallbackProducts: [
        { id: "P001", category: "women", name: "طقم فولاذ مقاوم للصدأ - هلال لؤلؤي", price: 2100, description: "طقم رائع ومقاوم للصدأ مزين باللؤلؤ يمنحكِ جاذبية فريدة. مصنوع من خامات عالية الجودة لا تسبب الحساسية، ومقاوم للصدأ والاهتراء مع الاستخدام اليومي.", images: ["https://i.ibb.co/6Jp8GJgw/IMG-20260629-232925.png"] },
        { id: "P002", category: "offers", name: "بوكس الهدايا المتكامل والمميز للعلاقات", price: 4800, description: "بوكس فاخر يحتوي على أرقى المجموعات العطرية بأسعار استثنائية. تغليف أنيق يجعله هدية مثالية لكل المناسبات الخاصة.", images: ["https://i.ibb.co/yFpxy58c/IMG-20260629-233027.png"] },
        { id: "P003", category: "home", name: "طقم منظمات المطبخ الفاخر والذكي", price: 3200, description: "كل ما تحتاجه سيدة المنزل لتنظيم مطبخ عصري ومرتب بلمسة ذهبية. يتحمل الاستخدام اليومي المكثف ويسهل تنظيفه.", images: ["https://i.ibb.co/9k2Pw9f1/IMG-20260629-233200.png"] },
        { id: "P004", category: "variety", name: "مكواة البخار المحمولة Sokany 800W", price: 2700, description: "مكواة البخار العمودية والأقوى لإزالة سريعة للتجاعيد أثناء السفر. خفيفة الوزن وسهلة الحمل مع خزان مياه كبير السعة.", images: ["https://i.ibb.co/PssyncH9/IMG-20260629-233300.png"] }
    ]
};

/* ==========================================================================
   Helpers & Custom Notifications
   ========================================================================== */
const Helpers = {
    showToast(message, type = 'success') {
        let toastContainer = document.querySelector('.luxury-toast-container');
        if (!toastContainer) {
            toastContainer = document.createElement('div');
            toastContainer.className = 'luxury-toast-container';
            Object.assign(toastContainer.style, {
                position: 'fixed',
                bottom: '24px',
                left: '50%',
                transform: 'translateX(-50%)',
                zIndex: '9999',
                display: 'flex',
                flexDirection: 'column',
                gap: '8px',
                width: 'calc(100% - 32px)',
                maxWidth: '400px',
                pointerEvents: 'none'
            });
            document.body.appendChild(toastContainer);
        }

        const toast = document.createElement('div');
        toast.className = `luxury-toast toast-${type}`;
        
        let iconHtml = '<i class="fa-solid fa-circle-info"></i>';
        if (type === 'success') iconHtml = '<i class="fa-solid fa-circle-check"></i>';
        if (type === 'warning') iconHtml = '<i class="fa-solid fa-triangle-exclamation"></i>';

        toast.innerHTML = `
            <div class="toast-icon-wrapper" style="color: ${type === 'success' ? '#25D366' : '#F0D27A'}; font-size: 16px;">${iconHtml}</div>
            <div class="toast-message" style="font-family: 'Cairo', sans-serif; font-size: 13px; font-weight: 600;">${message}</div>
        `;

        Object.assign(toast.style, {
            background: 'rgba(17, 17, 17, 0.95)',
            border: `1px solid ${type === 'success' ? 'rgba(37, 211, 102, 0.3)' : 'rgba(212, 175, 55, 0.3)'}`,
            borderRadius: '16px',
            padding: '12px 18px',
            display: 'flex',
            alignItems: 'center',
            gap: '12px',
            color: '#FFFFFF',
            boxShadow: '0 8px 30px rgba(0,0,0,0.5)',
            backdropFilter: 'blur(10px)',
            transition: 'all 0.5s cubic-bezier(0.16, 1, 0.3, 1)',
            transform: 'translateY(20px) scale(0.95)',
            opacity: '0',
            pointerEvents: 'auto'
        });

        toastContainer.appendChild(toast);

        requestAnimationFrame(() => {
            toast.style.transform = 'translateY(0) scale(1)';
            toast.style.opacity = '1';
        });

        setTimeout(() => {
            toast.style.transform = 'translateY(-20px) scale(0.9)';
            toast.style.opacity = '0';
            setTimeout(() => {
                toast.remove();
                if (toastContainer.children.length === 0) {
                    toastContainer.remove();
                }
            }, 500);
        }, 3500);
    }
};

/* ==========================================================================
   Product Manager (بدون نظام سلة - عرض تفاصيل وطلب مباشر لمنتج واحد)
   ========================================================================== */
class ProductManager {
    constructor() {
        AppState.state.products = AppState.fallbackProducts;
        this.currentModalImageIndex = 0;
    }

    async init() {
        try {
            await this.fetchAndParseJSON();
        } catch (error) {
            console.warn("جاري تشغيل المتجر بالمنتجات الاحتياطية المدمجة:", error);
        }

        this.renderProductsGrid();
        this.setupEventListeners();
    }

    async fetchAndParseJSON() {
        try {
            const response = await fetch("products.json");
            if (!response.ok) {
                console.error(`❌ Error loading products: HTTP ${response.status}`);
                throw new Error("لم يتم العثور على ملف المنتجات.");
            }
            
            const data = await response.json();
            if (!Array.isArray(data)) {
                throw new Error("Invalid products data format");
            }
            
            AppState.state.products = data;
            console.log(`📦 تم تحميل ${AppState.state.products.length} منتجاً من الملف.`);
        } catch (error) {
            console.error("فشل في تحميل المنتجات:", error);
            throw error;
        }
    }

    /* ==========================================================================
       UI Renderer & Event Handlers
       ========================================================================== */
    renderProductsGrid() {
        const grid = document.getElementById('products-dynamic-grid');
        if (!grid) return;

        const products = AppState.state.products;

        if (products.length === 0) {
            grid.innerHTML = `<p style="grid-column: span 2; text-align: center; color: var(--text-secondary); font-size: 12px; py-6;">لا توجد منتجات متوفرة حالياً.</p>`;
            return;
        }

        grid.innerHTML = products.map((product, index) => {
            const isAvailable = product.available !== false; // المنتج متوفر افتراضياً ما لم يُحدد خلاف ذلك
            const availabilityBadge = isAvailable
                ? `<span class="availability-badge available"><i class="fa-solid fa-circle"></i> متوفر</span>`
                : `<span class="availability-badge unavailable"><i class="fa-solid fa-circle"></i> غير متوفر</span>`;
            const orderBtn = isAvailable
                ? `<button class="product-order-btn" onclick="event.stopPropagation(); window.ProductRepository.buyDirectly('${product.id}')">
                        <span>طلب</span>
                        <i class="fa-solid fa-arrow-left"></i>
                    </button>`
                : `<button class="product-order-btn disabled" disabled onclick="event.stopPropagation();">
                        <span>طلب</span>
                        <i class="fa-solid fa-arrow-left"></i>
                    </button>`;

            return `
            <div class="product-item-card" data-id="${product.id}" onclick="window.ProductRepository.openProductModal('${product.id}')">
                <span class="product-index-badge">${index + 1}</span>
                <div class="product-img-wrapper">
                    <img src="${product.images[0] || 'https://via.placeholder.com/300'}" alt="${product.name}">
                </div>
                <div class="product-info-wrapper">
                    <span class="product-category">${product.category}</span>
                    <h3>${product.name}</h3>
                    <p class="product-price">${product.price.toLocaleString('ar-DZ')} د.ج</p>
                </div>
                <div class="product-action-row">
                    ${orderBtn}
                    ${availabilityBadge}
                </div>
            </div>`;
        }).join('');
    }

    /* ==========================================================================
       Product Details Modal (نافذة تفاصيل المنتج)
       ========================================================================== */
    openProductModal(productId) {
        const product = AppState.state.products.find(p => p.id === String(productId));
        if (!product) return;

        AppState.state.currentModalProductId = product.id;
        this.currentModalImageIndex = 0;

        const overlay = document.getElementById('product-modal-overlay');
        const box = document.getElementById('product-modal-box');
        if (!overlay || !box) return;

        const images = (product.images && product.images.length > 0) ? product.images : ['https://via.placeholder.com/400'];
        const isAvailable = product.available !== false;
        const availabilityBadge = isAvailable
            ? `<span class="availability-badge available"><i class="fa-solid fa-circle"></i> متوفر</span>`
            : `<span class="availability-badge unavailable"><i class="fa-solid fa-circle"></i> غير متوفر</span>`;
        const modalOrderBtn = isAvailable
            ? `<button class="product-order-btn modal-order-btn" onclick="window.ProductRepository.buyDirectly('${product.id}')">
                    <span>اطلب الآن</span>
                    <i class="fa-solid fa-arrow-left"></i>
                </button>`
            : `<button class="product-order-btn modal-order-btn disabled" disabled>
                    <span>غير متوفر حالياً</span>
                </button>`;

        box.innerHTML = `
            <button class="modal-close-btn" onclick="window.ProductRepository.closeProductModal()" aria-label="إغلاق">
                <i class="fa-solid fa-xmark"></i>
            </button>

            <div class="modal-gallery">
                <div class="modal-gallery-main">
                    <img id="modal-main-image" src="${images[0]}" alt="${product.name}">
                    ${images.length > 1 ? `
                        <button class="modal-gallery-nav prev" onclick="window.ProductRepository.changeModalImage(-1)" aria-label="الصورة السابقة">
                            <i class="fa-solid fa-chevron-right"></i>
                        </button>
                        <button class="modal-gallery-nav next" onclick="window.ProductRepository.changeModalImage(1)" aria-label="الصورة التالية">
                            <i class="fa-solid fa-chevron-left"></i>
                        </button>
                        <span class="modal-gallery-counter" id="modal-gallery-counter">1 / ${images.length}</span>
                    ` : ''}
                </div>
                ${images.length > 1 ? `
                    <div class="modal-gallery-thumbs">
                        ${images.map((img, idx) => `
                            <div class="modal-thumb ${idx === 0 ? 'active' : ''}" data-idx="${idx}" onclick="window.ProductRepository.setModalImage(${idx})">
                                <img src="${img}" alt="صورة ${idx + 1}">
                            </div>
                        `).join('')}
                    </div>
                ` : ''}
            </div>

            <div class="modal-details">
                <span class="product-category">${product.category}</span>
                <h2 class="modal-product-name">${product.name}</h2>
                <div class="modal-price-row">
                    <p class="modal-product-price">${product.price.toLocaleString('ar-DZ')} د.ج</p>
                    ${availabilityBadge}
                </div>
                <div class="modal-product-description">
                    <h4>وصف المنتج</h4>
                    <p>${(product.description || 'لا يوجد وصف متاح لهذا المنتج حالياً.').replace(/\n/g, '<br>')}</p>
                </div>
            </div>

            <div class="modal-footer">
                ${modalOrderBtn}
            </div>
        `;

        overlay.classList.add('active');
        document.body.style.overflow = 'hidden';
    }

    closeProductModal() {
        const overlay = document.getElementById('product-modal-overlay');
        if (!overlay) return;
        overlay.classList.remove('active');
        document.body.style.overflow = '';
        AppState.state.currentModalProductId = null;
    }

    changeModalImage(direction) {
        const product = AppState.state.products.find(p => p.id === AppState.state.currentModalProductId);
        if (!product) return;
        const images = product.images && product.images.length > 0 ? product.images : ['https://via.placeholder.com/400'];

        this.currentModalImageIndex = (this.currentModalImageIndex + direction + images.length) % images.length;
        this.updateModalImage(images);
    }

    setModalImage(idx) {
        const product = AppState.state.products.find(p => p.id === AppState.state.currentModalProductId);
        if (!product) return;
        const images = product.images && product.images.length > 0 ? product.images : ['https://via.placeholder.com/400'];

        this.currentModalImageIndex = idx;
        this.updateModalImage(images);
    }

    updateModalImage(images) {
        const mainImg = document.getElementById('modal-main-image');
        const counter = document.getElementById('modal-gallery-counter');
        if (mainImg) mainImg.src = images[this.currentModalImageIndex];
        if (counter) counter.textContent = `${this.currentModalImageIndex + 1} / ${images.length}`;

        document.querySelectorAll('.modal-thumb').forEach((thumb, idx) => {
            thumb.classList.toggle('active', idx === this.currentModalImageIndex);
        });
    }

    /* ==========================================================================
       Direct Order (طلب منتج واحد مباشر - بدون سلة)
       ========================================================================== */
    buyDirectly(productId) {
        const product = AppState.state.products.find(p => p.id === String(productId));
        if (!product) return;

        if (product.available === false) {
            Helpers.showToast("عذراً، هذا المنتج غير متوفر حالياً.", "warning");
            return;
        }

        try {
            sessionStorage.setItem("nova_shopping_order", JSON.stringify({
                id: product.id,
                name: product.name,
                price: product.price,
                image: product.images[0] || "",
                category: product.category,
                quantity: 1
            }));
        } catch (e) {
            console.error("فشل حفظ بيانات الطلب:", e);
        }

        window.location.href = "order.html";
    }

    setupEventListeners() {
        // زر ابدأ التسوق: تمرير سلس مباشرة لقسم المنتجات
        const heroCtaBtn = document.getElementById('hero-cta-btn');
        if (heroCtaBtn) {
            heroCtaBtn.addEventListener('click', () => {
                const productsSection = document.getElementById('featured-products-section');
                if (productsSection) productsSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
            });
        }

        // إغلاق النافذة بالضغط خارج الصندوق أو بمفتاح Escape
        const overlay = document.getElementById('product-modal-overlay');
        if (overlay) {
            overlay.addEventListener('click', (e) => {
                if (e.target === overlay) this.closeProductModal();
            });
        }
        document.addEventListener('keydown', (e) => {
            if (e.key === 'Escape') this.closeProductModal();
        });
    }

    /* ==========================================================================
       Checkout Integration (لوجيك صفحة الدفع المستقلة - منتج واحد بدون سلة)
       ========================================================================== */
    initCheckoutPage() {
        const checkoutList = document.getElementById('checkout-products-list');
        if (!checkoutList) return;

        this.loadOrderFromStorage();

        // 1. بناء منسدلة الولايات الجزائرية الـ 58
        const wilayaSelect = document.getElementById('wilaya-select');
        if (wilayaSelect && wilayaSelect.children.length <= 1) {
            AppState.wilayas.forEach(wilaya => {
                const opt = document.createElement('option');
                opt.value = wilaya.name;
                opt.textContent = wilaya.name;
                wilayaSelect.appendChild(opt);
            });
            wilayaSelect.addEventListener('change', () => this.recalculateCheckoutBill());
        }

        // 2. التحقق من وجود منتج محدد للطلب
        if (!AppState.state.order) {
            checkoutList.innerHTML = `
                <div id="empty-checkout-placeholder" class="empty-cart-placeholder">
                    <div class="empty-cart-icon">
                        <i class="fa-solid fa-box-open"></i>
                    </div>
                    <div>
                        <p class="empty-cart-text">لم يتم تحديد أي منتج للطلب</p>
                    </div>
                    <a href="index.html" class="empty-cart-btn">العودة للتسوق</a>
                </div>
            `;
            this.recalculateCheckoutBill();
            return;
        }

        // 3. بناء بطاقة المنتج المطلوب مع التحكم بالكمية
        const item = AppState.state.order;
        checkoutList.innerHTML = `
            <div class="cart-item-card">
                <div class="cart-item-img">
                    <img src="${item.image}" alt="${item.name}">
                </div>
                <div class="cart-item-info">
                    <div>
                        <h4 class="cart-item-name">${item.name}</h4>
                    </div>
                    <div class="cart-item-bottom-row">
                        <span class="cart-item-price">${(item.price * item.quantity).toLocaleString('ar-DZ')} د.ج</span>
                        <div class="cart-item-qty-control">
                            <button onclick="window.ProductRepository.changeOrderQty(${item.quantity - 1})">-</button>
                            <span class="cart-item-qty-value">${item.quantity}</span>
                            <button onclick="window.ProductRepository.changeOrderQty(${item.quantity + 1})">+</button>
                        </div>
                    </div>
                </div>
            </div>
        `;

        this.recalculateCheckoutBill();
    }

    loadOrderFromStorage() {
        try {
            const stored = sessionStorage.getItem("nova_shopping_order");
            AppState.state.order = stored ? JSON.parse(stored) : null;
        } catch (e) {
            AppState.state.order = null;
        }
    }

    saveOrderToStorage() {
        try {
            if (AppState.state.order) {
                sessionStorage.setItem("nova_shopping_order", JSON.stringify(AppState.state.order));
            }
        } catch (e) {
            console.error("فشل حفظ بيانات الطلب:", e);
        }
    }

    changeOrderQty(newQty) {
        if (!AppState.state.order) return;
        AppState.state.order.quantity = Math.max(1, parseInt(newQty) || 1);
        this.saveOrderToStorage();
        this.initCheckoutPage();
    }

    clearOrder() {
        AppState.state.order = null;
        try {
            sessionStorage.removeItem("nova_shopping_order");
        } catch (e) { /* تجاهل */ }
    }

    recalculateCheckoutBill() {
        const item = AppState.state.order;
        const subtotal = item ? item.price * item.quantity : 0;
        const itemsCount = item ? item.quantity : 0;
        const wilayaSelect = document.getElementById('wilaya-select');
        const deliveryTypeInput = document.querySelector('input[name="delivery-type"]:checked');

        let deliveryCost = 0;
        let selectedWilayaName = wilayaSelect ? wilayaSelect.value : "";
        let deliveryType = deliveryTypeInput ? deliveryTypeInput.value : "office";

        if (selectedWilayaName) {
            const wilayaObj = AppState.wilayas.find(w => w.name === selectedWilayaName);
            if (wilayaObj) {
                deliveryCost = deliveryType === 'home' ? wilayaObj.home : wilayaObj.office;
                
                const officePreview = document.getElementById('office-price-preview');
                const homePreview = document.getElementById('home-price-preview');
                if (officePreview) officePreview.textContent = `${wilayaObj.office} د.ج`;
                if (homePreview) homePreview.textContent = `${wilayaObj.home} د.ج`;
            }
        }

        const totalCost = subtotal + deliveryCost;

        // تحديث الـ DOM
        const itemsCountBadge = document.getElementById('items-count-badge');
        const summaryItemsCount = document.getElementById('summary-items-count');
        const summarySubtotal = document.getElementById('summary-subtotal');
        const summaryDeliveryCost = document.getElementById('summary-delivery-cost');
        const summaryTotal = document.getElementById('summary-total');

        if (itemsCountBadge) itemsCountBadge.textContent = `${itemsCount} منتجات`;
        if (summaryItemsCount) summaryItemsCount.textContent = itemsCount;
        if (summarySubtotal) summarySubtotal.textContent = `${subtotal.toLocaleString('ar-DZ')} د.ج`;
        if (summaryDeliveryCost) summaryDeliveryCost.textContent = selectedWilayaName ? `${deliveryCost.toLocaleString('ar-DZ')} د.ج` : "حدد الولاية";
        if (summaryTotal) summaryTotal.textContent = `${totalCost.toLocaleString('ar-DZ')} د.ج`;
    }
}

// تسجيل المحرك محلياً وعالمياً
window.ProductRepository = new ProductManager();

document.addEventListener("DOMContentLoaded", async () => {
    // تفعيل المحرك التلقائي بناءً على الصفحة المفتوحة حالياً
    if (document.getElementById('products-dynamic-grid')) {
        await window.ProductRepository.init();
    } else if (document.getElementById('checkout-products-list')) {
        await window.ProductRepository.initCheckoutPage();
    }
});