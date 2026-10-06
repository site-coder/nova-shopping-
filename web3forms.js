/**
 * ==========================================================================
 * اسم المشروع: متجر Nova Shopping الإلكتروني الفاخر
 * الملف: web3forms.js
 * الوصف: المحرك السحابي المسؤول عن معالجة وإرسال طلبات العملاء إلى خدمة
 * Web3Forms API مع الحفاظ على بقاء العميل والتحقق الصارم من الحقول.
 * المطور: مصمم ومطور واجهات المستخدم الفاخرة
 * الإصدار: 1.1.0
 * ==========================================================================
 */

"use strict";

const WEB3FORMS_ACCESS_KEY = "de50c293-b2a7-4fab-b879-b0ca1e17d4e5";

class Web3FormsHandler {
    constructor() {
        this.submitButton = document.getElementById('submit-order-btn');
        this.form = document.getElementById('luxury-checkout-form');
        this.isSubmitting = false;
    }

    init() {
        if (!this.submitButton) {
            console.warn("لم يتم العثور على زر الإرسال في هذه الصفحة.");
            return;
        }

        this.submitButton.addEventListener('click', (e) => {
            e.preventDefault();
            this.handleOrderSubmission();
        });
    }

    async handleOrderSubmission() {
        if (this.isSubmitting) return;

        // 1. جلب بيانات المنتج المطلوب للتحقق منها
        const orderItem = this.getOrderData();
        if (!orderItem) {
            this.showToast("لم يتم تحديد أي منتج للطلب! يرجى اختيار منتج من المتجر أولاً.", "warning");
            return;
        }

        // 2. جلب بيانات الاستمارة وفحصها
        const formData = this.getFormData();
        if (!this.validateData(formData)) {
            return; // إيقاف الإرسال لوجود خطأ
        }

        // 3. قفل زر التفاعل وإظهار مؤشر الشحن الذهبي الفخم
        this.setLoadingState(true);

        try {
            // 4. بناء الحمولة البرمجية المنسقة
            const payload = this.buildPayload(formData, orderItem);

            // 5. الإرسال الآمن عبر Fetch API
            const response = await fetch("https://api.web3forms.com/submit", {
                method: "POST",
                headers: {
                    "Content-Type": "application/json",
                    "Accept": "application/json"
                },
                body: JSON.stringify(payload)
            });

            const result = await response.json();

            if (response.ok && result.success) {
                this.showToast("تم إرسال طلبك بنجاح! سنتواصل معك هاتفياً قريباً لتأكيد الشحن والاستلام.", "success");
                this.resetFormAndOrder();
                // إعادة العميل للرئيسية بعد ثانيتين
                setTimeout(() => {
                    window.location.href = "index.html";
                }, 2000);
            } else {
                throw new Error(result.message || "فشل معالجة الطلب على الخادم.");
            }

        } catch (error) {
            console.error("خطأ أثناء الإرسال لـ Web3Forms:", error);
            this.showToast("حدث خطأ غير متوقع أثناء إرسال الطلب. يرجى مراجعة الاتصال والمحاولة مجدداً.", "warning");
        } finally {
            this.setLoadingState(false);
        }
    }

    getOrderData() {
        if (window.ProductRepository && window.ProductRepository.constructor && AppState.state.order) {
            return AppState.state.order;
        }
        try {
            const stored = sessionStorage.getItem("nova_shopping_order");
            return stored ? JSON.parse(stored) : null;
        } catch (e) {
            return null;
        }
    }

    getFormData() {
        const deliveryTypeInput = document.querySelector('input[name="delivery-type"]:checked');
        return {
            firstName: document.getElementById('first-name')?.value.trim() || "",
            lastName: document.getElementById('last-name')?.value.trim() || "",
            phone: document.getElementById('customer-phone')?.value.trim() || "",
            wilaya: document.getElementById('wilaya-select')?.value || "",
            deliveryType: deliveryTypeInput ? deliveryTypeInput.value : "office",
            municipality: document.getElementById('municipality')?.value.trim() || "",
            address: document.getElementById('address-details')?.value.trim() || ""
        };
    }

    validateData(data) {
        if (!data.firstName || !data.lastName) {
            this.showToast("يرجى إدخال الاسم واللقب بالكامل.", "warning");
            return false;
        }

        if (!data.phone) {
            this.showToast("يرجى إدخال رقم الهاتف للتواصل.", "warning");
            return false;
        }

        // فحص الهاتف الجزائري (الموبيليس، الجازي، أوريدو)
        const algerianPhoneRegex = /^(0)(5|6|7)[0-9]{8}$/;
        const cleanedPhone = data.phone.replace(/\s+/g, '');
        if (!algerianPhoneRegex.test(cleanedPhone)) {
            this.showToast("رقم الهاتف غير صحيح! يجب أن يبدأ بـ 05 أو 06 أو 07 متبوعاً بـ 8 أرقام.", "warning");
            return false;
        }

        if (!data.wilaya) {
            this.showToast("يرجى تحديد ولاية الشحن.", "warning");
            return false;
        }

        if (data.deliveryType === 'home') {
            if (!data.municipality) {
                this.showToast("يرجى إدخال اسم البلدية للتوصيل المنزلي.", "warning");
                return false;
            }
            if (!data.address || data.address.length < 5) {
                this.showToast("يرجى كتابة عنوان المنزل بالتفصيل وبما لا يقل عن 5 أحرف.", "warning");
                return false;
            }
        }

        return true;
    }

    buildPayload(data, orderItem) {
        const productsDetailsText = `[المنتج: ${orderItem.name}] - [الكمية: ${orderItem.quantity}] - [سعر القطعة: ${orderItem.price.toLocaleString('ar-DZ')} د.ج]`;

        const subtotalText = document.getElementById('summary-subtotal')?.textContent || "0.00 د.ج";
        const deliveryCostText = document.getElementById('summary-delivery-cost')?.textContent || "0.00 د.ج";
        const totalText = document.getElementById('summary-total')?.textContent || "0.00 د.ج";

        const algerianTime = new Date().toLocaleString('ar-DZ', {
            timeZone: 'Africa/Algiers',
            hour12: true
        });

        return {
            access_key: WEB3FORMS_ACCESS_KEY,
            subject: `طلب شراء جديد من: ${data.firstName} ${data.lastName} (${data.wilaya})`,
            from_name: "متجر Nova Shopping الفاخر",
            "الاسم الأول": data.firstName,
            "اللقب": data.lastName,
            "رقم الهاتف": data.phone,
            "الولاية": data.wilaya,
            "نوع التوصيل": data.deliveryType === 'home' ? "توصيل للمنزل" : "توصيل للمكتب",
            "البلدية": data.deliveryType === 'home' ? data.municipality : "غير مطلوب (توصيل للمكتب)",
            "العنوان بالتفصيل": data.deliveryType === 'home' ? data.address : "غير مطلوب (توصيل للمكتب)",
            "تفاصيل السلع المشتراة": productsDetailsText,
            "إجمالي المنتجات": subtotalText,
            "تكلفة التوصيل": deliveryCostText,
            "المبلغ الإجمالي للدفع عند الاستلام": totalText,
            "تاريخ ووقت الطلب بالجزائر": algerianTime
        };
    }

    setLoadingState(isLoading) {
        this.isSubmitting = isLoading;
        if (!this.submitButton) return;

        if (isLoading) {
            this.submitButton.disabled = true;
            this.submitButton.style.opacity = "0.75";
            this.submitButton.style.cursor = "not-allowed";
            this.submitButton.innerHTML = `
                <i class="fa-solid fa-spinner"></i>
                <span>جاري إرسال وتأكيد طلبك بأمان...</span>
            `;
        } else {
            this.submitButton.disabled = false;
            this.submitButton.style.opacity = "1";
            this.submitButton.style.cursor = "pointer";
            this.submitButton.innerHTML = `
                <i class="fa-solid fa-circle-check"></i>
                <span>تأكيد طلب الشراء الآن</span>
            `;
        }
    }

    showToast(message, type = 'success') {
        if (window.Helpers && typeof window.Helpers.showToast === 'function') {
            window.Helpers.showToast(message, type);
        } else {
            console.log(`[Notification ${type}]: ${message}`);
        }
    }

    resetFormAndOrder() {
        if (this.form) this.form.reset();
        if (window.ProductRepository && typeof window.ProductRepository.clearOrder === 'function') {
            window.ProductRepository.clearOrder();
        } else {
            try { sessionStorage.removeItem("nova_shopping_order"); } catch (e) { /* تجاهل */ }
        }
    }
}

// البدء التلقائي للمحرك عند اكتمال الجاهزية
document.addEventListener("DOMContentLoaded", () => {
    window.LuxuryWeb3Forms = new Web3FormsHandler();
    window.LuxuryWeb3Forms.init();
});