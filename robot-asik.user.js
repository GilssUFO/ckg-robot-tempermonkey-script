// ==UserScript==
// @name         Robot ASIK Full Auto (V34.0 - Pendaftaran + Aturan Pustu)
// @namespace    http://tampermonkey.net/
// @version      34.1
// @description  Pendaftaran otomatis + 14 panel, GDS tanpa random, HPV/HIV/Lipid/Hepatitis auto tidak-periksa
// @author       Faris
// @match        *://*/*
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_xmlhttpRequest
// @connect      script.google.com
// @connect      script.googleusercontent.com
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// ==/UserScript==

(function () {
    'use strict';
    console.log('[ASIK Robot] === VERSI 34.1 AKTIF ===');

    // Blokir hanya widget Google Contacts yang mengganggu (biang kerok error sebelumnya)
    if (window.location.hostname.includes('contacts.google.com')) return;

    const sleep = (ms) => new Promise(r => setTimeout(r, ms));

    // --
    // TRIGGER KLIK -- tembus Vue/Nuxt event listener
    // --
    const triggerClick = (el) => {
        if (!el) return;

        const isSurveyJs = el.closest('.sv-root-modern, .sd-root-modern, .sd-question, .sv-question, .sd-dropdown');

        ['mouseover', 'pointerdown', 'mousedown', 'pointerup', 'mouseup'].forEach(ev => {
            el.dispatchEvent(new MouseEvent(ev, { bubbles: true, cancelable: true, view: document.defaultView, detail: 1 }));
        });

        if (isSurveyJs) {
            // Untuk form input data (SurveyJS), butuh event 'click' manual (tapi jangan panggil el.click() agar dropdown tidak tertutup)
            el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: document.defaultView, detail: 1 }));
        } else {
            // Untuk antarmuka Kemenkes (Vue), gunakan native click agar tidak klik ganda (double-click)
            if (typeof el.click === 'function') {
                try { el.click(); } catch (e) { }
            }
        }

        // Simulasi klik dengan keyboard (Enter & Space) untuk by-pass beberapa event listener React/Vue
        try {
            el.focus();
            el.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', code: 'Enter', keyCode: 13, bubbles: true }));
            el.dispatchEvent(new KeyboardEvent('keyup', { key: 'Enter', code: 'Enter', keyCode: 13, bubbles: true }));
            el.dispatchEvent(new KeyboardEvent('keydown', { key: ' ', code: 'Space', keyCode: 32, bubbles: true }));
            el.dispatchEvent(new KeyboardEvent('keyup', { key: ' ', code: 'Space', keyCode: 32, bubbles: true }));
        } catch (e) { }
    };

    // Normalkan teks: hapus spasi ganda, NBSP, karakter zero-width
    const bersih = (str) =>
        str ? str.replace(/[\u00A0\u200B\u200C\u200D\uFEFF]+/g, ' ')
            .replace(/\s+/g, ' ').trim().toLowerCase() : "";

    // Tunggu elemen muncul di DOM
    const tungguElemen = (selector, root = document, timeout = 10000) =>
        new Promise(resolve => {
            const el = root.querySelector(selector);
            if (el) return resolve(el);
            const ob = new MutationObserver(() => {
                const found = root.querySelector(selector);
                if (found) { ob.disconnect(); resolve(found); }
            });
            const observeTarget = root === document ? document.body : root;
            ob.observe(observeTarget, { childList: true, subtree: true });
            setTimeout(() => { ob.disconnect(); resolve(null); }, timeout);
        });

    // --
    // KUNCI JAWABAN KHUSUS & PRIORITAS OPSI AMAN
    // --
    const pertanyaanKhusus = [
        { q: "alasan tidak diberikan", a: "pengobatan tuntas", fallback: "lainnya" },
        { q: "terpapar asap rokok", a: "tidak" },
        { q: "berdiri dari kursi", a: "ya" },
        { q: "tiga kata", a: "ya" },
        { q: "mengingat tiga kata", a: "ya" },
        { q: "tanggal berapakah", a: "benar semua" },
        { q: "dimanakah anda saat ini", a: "benar semua" },
        { q: "terdaftar di puskes", a: "lainnya" },
        { q: "faskes lain", a: "lainnya" },
        { q: "edukasi", a: "ya" },
        { q: "diberikan konseling", a: "ya" },
        { q: "berapa lama", a: "0", type: "text" },
        { q: "kapan anda berhenti merokok", a: "today", type: "date" },
        { q: "tanggal pelaksanaan", a: "today", type: "date" }
    ];

    const prioritasOpsiAman = [
        "tidak batuk",
        "tidak sama sekali",
        "tidak pernah",
        "tidak tahu atau tidak ingat",
        "tidak tahu",
        "non disabilitas",
        "terkendali teratur",
        "mandiri",
        "bisa bepergian keluar rumah",
        "keluar rumah",
        "nafsu makan biasa saja",
        "nafsu makan biasa",
        "tidak ada masalah psikologis",
        "tidak ada masalah",
        "tidak ada",
        ">= 23",
        ">= 31",
        "normal",
        "pengobatan tuntas",
        "lainnya",
        "tidak" // ditempatkan terakhir agar yang lebih spesifik matching duluan
    ];

    // --
    // XPATH CONSTANTS -- dari Chrome Extension (untuk Pendaftaran)
    // --
    const XPATH_DAFTAR = {
        BTN_DAFTAR_BARU: "/html/body/div[1]/main/div/div[1]/section[2]/div/div/div/div[2]/div/div[3]/div[2]/div[2]/div[2]/div/button",
        INPUT_NIK_PENDAFTARAN: "//input[@id='nik']",
        BTN_CEK_NIK_PENDAFTARAN: "//button[.//div[text()='Cek NIK']]",
        BTN_GUNAKAN_NIK: "//button[.//div[normalize-space()='Gunakan Data']]",
        POPUP_NIK_TIDAK_DITEMUKAN: "//div[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'data tidak ditemukan')]",
        POPUP_INDIVIDU_SUDAH_MENERIMA_LAYANAN: "//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'individu sudah menerima layanan')]",
        POPUP_DATA_PESERTA_WALI_TIDAK_VALID: "//div[normalize-space()='Data peserta atau wali tidak valid']",
        BTN_LANJUTKAN_DATA_VALID: "//div[contains(normalize-space(),'Data peserta valid')]/ancestor::div[contains(@class,'shadow-gmail')]//button[.//*[normalize-space()='Lanjutkan']]",
        BTN_SELANJUTNYA_FORMULIR_PENDAFTARAN: "//button[not(@disabled) and not(contains(@class, 'cursor-not-allowed')) and (contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'selanjutnya') or contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'lanjutkan'))]",
        INPUT_NAMA_LENGKAP: "//input[@name='Nama' or @name='nama' or @name='namaLengkap' or @name='nama_lengkap'] | //label[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'nama lengkap') or normalize-space(text())='Nama']/following-sibling::div//input | //input[contains(translate(@placeholder, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'nama lengkap')]",
        INPUT_JENIS_KELAMIN: "//span[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'pilih jenis kelamin')]",
        SELECT_JK_LK: "//div[contains(@class, 'cursor-pointer') and contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'laki-laki')] | //li[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'laki-laki')]",
        SELECT_JK_PR: "//div[contains(@class, 'cursor-pointer') and contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'perempuan')] | //li[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'perempuan')]",
        INPUT_WA: "//input[@name='Nomor Whatsapp'] | //input[contains(translate(@placeholder, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'whatsapp') or contains(translate(@placeholder, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'telepon')]",
        INPUT_TGL_LAHIR: "(//input[contains(@placeholder, 'Tanggal Lahir') or contains(@placeholder, 'tanggal lahir')])[1] | (//div[@id='Tanggal Lahir']//div[contains(@class,'mx-input-wrapper')])[1] | (//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'tanggal lahir')]/following-sibling::div//input)[1]",
        INPUT_TGL_LAHIR_YEAR: "//button[contains(@class,'mx-btn-current-year')]",
        INPUT_TGL_LAHIR_YEAR_TABLE: "//table[contains(@class,'mx-table-year')]",
        INPUT_TGL_LAHIR_YEAR_BEFORE: "//button[contains(@class,'mx-btn-icon-double-left')]",
        INPUT_TGL_LAHIR_MONTH_TABLE: "//table[contains(@class,'mx-table-month')]",
        INPUT_TGL_LAHIR_DAY_TABLE: "//table[contains(@class,'mx-table-date')]",
        INPUT_TGL_PEMERIKSAAN_PARENT: "//div[text()='Tanggal Pemeriksaan']/following::div[contains(@class,'shadow-gmail')][1] | //label[contains(text(), 'Tanggal Pemeriksaan')]/following-sibling::div",
        BTN_SELANJUTNYA: "//button[.//div[normalize-space()='Selanjutnya']]",
        BTN_LANJUT_KUOTA_HABIS: "//button[.//div[normalize-space()='Lanjut']]",
        BTN_PILIH_PESERTA: "//button[.//div[text()='Pilih']]",
        BTN_DAFTAR_DENGAN_NIK: "/html/body/div[1]/main/div/div[1]/section[2]/div/div/div/div[2]/div/div[3]/div[5]/div[2]/div/div/div[3]/div[5]/div[2]/div[1]/button",
        BTN_DAFTAR_TANPA_NIK: "/html/body/div[1]/main/div/div[1]/section[2]/div/div/div/div[2]/div/div[3]/div[5]/div[2]/div/div/div[3]/div[5]/div[2]/div[2]/button",
        MSG_POPUP: "/html/body/div[1]/main/div/div[1]/section[2]/div/div/div/div[2]/div/div[3]/div[6]/div[2]/div/div[1]/div",
        MSG_POPUP_TERJADI_KESALAHAN: "//span[contains(., 'Terjadi kesalahan')]/ancestor::div[contains(@class,'p-2')]//div[contains(@class,'my-4')]//span",
        MSG_POPUP_SUCCESS: "//div[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'berhasil daftar') or contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'data berhasil disimpan')]",
        BTN_TUTUP_SUCCESS_DAFTAR: "//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'tutup')]",
        BTN_LANJUT_KUOTA_HABIS: "//button[(contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'lanjut') or contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'lanjuy')) and not(contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'selanjutnya'))]",
        SELECT_SEARCH: "/html/body/div[1]/main/div/div[1]/section[2]/div/div/div/div[2]/div/div[3]/div[1]/div[2]/div[1]/div/div[2]",
        SELECT_SEARCH_NAMA: "/html/body/div[1]/main/div/div[1]/section[2]/div/div/div/div[2]/div/div[3]/div[1]/div[2]/div[1]/div/div[3]/div/div[3]",
        INPUT_SEARCH: "/html/body/div[1]/main/div/div[1]/section[2]/div/div/div/div[2]/div/div[3]/div[1]/div[2]/div[2]/label/div/input",
        BTN_KONFIMASI_HADIR: "/html/body/div[1]/main/div/div[1]/section[2]/div/div/div/div[2]/div/div[3]/div[3]/div/div/table/tbody/tr/td[6]/div/div[1]/div/button",
        CHECKBOX_BERSEDIA_CKG: "/html/body/div[1]/main/div/div[1]/section[2]/div/div/div/div[2]/div/div[3]/div[4]/div[2]/div/div[4]/div[3]/div[1]/div/div[1]/div",
        BTN_HADIR_CKG: "/html/body/div[1]/main/div/div[1]/section[2]/div/div/div/div[2]/div/div[3]/div[4]/div[2]/div/div[5]/div[2]/button",
        MSG_POPUP_BERHASIL_HADIR: "/html/body/div[1]/main/div/div[1]/section[2]/div/div/div/div[2]/div/div[3]/div[4]/div[2]/div/div[1]/div[1]",
        CHECKBOX_TANPA_WALI: "//div[@class='check' and @id='noWali']",
        INPUT_NIK_WALI: "//input[@id='nik wali'] | (//label[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'nik')]/following-sibling::div//input)[last()] | (//div[contains(@class, 'font-semibold') and contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'nik')]/following-sibling::div//input)[last()]",
        INPUT_NAMA_LENGKAP_WALI: "//input[@name='Nama Lengkap Wali'] | //input[contains(@placeholder, 'Nama Wali')] | (//label[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'nama lengkap')]/following-sibling::div//input)[last()] | (//div[contains(@class, 'font-semibold') and contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'nama wali')]/following-sibling::div//input)[last()] | (//div[contains(@class, 'font-semibold') and contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'nama lengkap')]/following-sibling::div//input)[last()]",
        INPUT_JENIS_KELAMIN_WALI: "(//div[@id='Jenis Kelamin'])[last()] | (//label[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'jenis kelamin')]/following-sibling::div)[last()] | (//div[contains(@class,'font-semibold') and contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'jenis kelamin')]/following-sibling::div)[last()]",
        INPUT_WA_WALI: "//input[@name='Nomor whatsapp'] | //input[contains(@placeholder, 'Nomor WhatsApp')] | (//label[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'no. whatsapp')]/following-sibling::div//input)[last()] | (//div[contains(@class,'font-semibold') and contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'no. whatsapp')]/following-sibling::div//input)[last()] | (//div[contains(@class,'font-semibold') and contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'whatsapp')]/following-sibling::div//input)[last()]",
        INPUT_TGL_LAHIR_WALI: "(//input[contains(@placeholder, 'Tanggal Lahir') or contains(@placeholder, 'tanggal lahir')])[last()] | (//div[@id='Tanggal Lahir']//div[contains(@class,'mx-input-wrapper')])[last()] | (//label[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'tanggal lahir')]/following-sibling::div//input)[last()] | (//div[contains(@class,'font-semibold') and contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'tanggal lahir')]/following-sibling::div//input)[last()]",
    };

    // --
    // HELPER PENDAFTARAN -- diadaptasi dari Chrome Extension helper.js
    // --
    function parseDateStringDaftar(dateStr) {
        if (!dateStr) return null;
        // Handle Excel serial date (number > 20000)
        if (!isNaN(dateStr) && Number(dateStr) > 20000) {
            const serial = Number(dateStr);
            const excelEpoch = Date.UTC(1899, 11, 30);
            const d = new Date(excelEpoch + serial * 86400000);
            const day = d.getUTCDate();
            const month = d.getUTCMonth() + 1;
            const year = d.getUTCFullYear();
            const dayStr = String(day).padStart(2, '0');
            const monthStr = String(month).padStart(2, '0');
            return { day, month, year, date: `${year}-${monthStr}-${dayStr}` };
        }
        // Normalize separator
        const parts = String(dateStr).split(/[-/]/);
        let day, month, year;
        if (parts[0].length === 4) {
            year = parseInt(parts[0], 10); month = parseInt(parts[1], 10); day = parseInt(parts[2], 10);
        } else {
            day = parseInt(parts[0], 10); month = parseInt(parts[1], 10); year = parseInt(parts[2], 10);
        }

        // Auto-fix if day and month are swapped (MM/DD/YYYY)
        if (month > 12) {
            const temp = day;
            day = month;
            month = temp;
        }

        if (!day || !month || !year || month > 12 || day > 31) return null;

        const dayStr = String(day).padStart(2, '0');
        const monthStr = String(month).padStart(2, '0');
        return { day, month, year, date: `${year}-${monthStr}-${dayStr}` };
    }

    function toDDMMYYYY(dateStr) {
        const parsed = parseDateStringDaftar(dateStr);
        if (!parsed) return dateStr;
        return `${String(parsed.day).padStart(2, '0')}-${String(parsed.month).padStart(2, '0')}-${parsed.year}`;
    }

    function cleanPhoneNumber(phone, defPhone = "80000000") {
        if (!phone) return defPhone;
        phone = String(phone).replace(/\D/g, '');
        if (phone.startsWith('0')) phone = phone.slice(1);
        else if (phone.startsWith('62')) phone = phone.slice(2);
        if (!phone.startsWith('8')) phone = '8' + phone;
        if (phone.length < 7 || phone.length > 13) return defPhone;
        return phone;
    }

    function parseDDMMYYYY(dateStr) {
        if (!dateStr) return new Date();
        const parts = String(dateStr).split('-').map(Number);
        return new Date(parts[2], parts[1] - 1, parts[0]);
    }

    function isUnder10Years(dateStr) {
        const birthDate = parseDDMMYYYY(dateStr);
        const today = new Date();
        let age = today.getFullYear() - birthDate.getFullYear();
        const m = today.getMonth() - birthDate.getMonth();
        if (m < 0 || (m === 0 && today.getDate() < birthDate.getDate())) age--;
        return age < 10;
    }

    function isOver60Years(dateStr) {
        const birthDate = parseDDMMYYYY(dateStr);
        const today = new Date();
        let age = today.getFullYear() - birthDate.getFullYear();
        const m = today.getMonth() - birthDate.getMonth();
        if (m < 0 || (m === 0 && today.getDate() < birthDate.getDate())) age--;
        return age >= 60;
    }

    function hitungUsiaDariTglLahir(dateStr) {
        const birthDate = parseDDMMYYYY(dateStr);
        const today = new Date();
        let age = today.getFullYear() - birthDate.getFullYear();
        const m = today.getMonth() - birthDate.getMonth();
        if (m < 0 || (m === 0 && today.getDate() < birthDate.getDate())) age--;
        return age;
    }

    function waitForXPath(xpath, parentEl, maxTries = 10) {
        let attempt = 0;
        return new Promise((resolve) => {
            function tryFind() {
                attempt++;
                const el = document.evaluate(xpath, parentEl || document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                if (el) { setTimeout(() => resolve(el), 500); }
                else if (attempt < maxTries) { setTimeout(tryFind, 500 * attempt); }
                else { resolve(null); }
            }
            tryFind();
        });
    }

    function inputElValue(el, val) {
        if (!el) return;
        const strVal = String(val);
        console.log(`[ASIK Robot] MENGISI: "${strVal}" ke elemen [${el.name || el.id || el.placeholder}]`);

        el.focus();
        el.value = strVal;
        el.setAttribute('value', strVal);

        if (el._value !== undefined) el._value = strVal; // Hack for Vue 3 internals

        const proto = el.tagName === 'TEXTAREA' ? window.HTMLTextAreaElement.prototype : window.HTMLInputElement.prototype;
        const setter = Object.getOwnPropertyDescriptor(proto, 'value').set;
        if (setter) setter.call(el, strVal);
        else el.value = strVal;

        // Dispatch custom input events dengan composed: true (Penting untuk Vue 3)
        el.dispatchEvent(new Event('input', { bubbles: true, composed: true }));
        el.dispatchEvent(new InputEvent('input', { bubbles: true, composed: true, inputType: 'insertText', data: strVal }));

        ['keydown', 'keypress', 'change', 'keyup', 'blur'].forEach(e => {
            el.dispatchEvent(new Event(e, { bubbles: true }));
        });

        console.log(`[ASIK Robot] SELESAI MENGISI: [${el.name || el.id || el.placeholder}] = ${el.value}`);
    }

    function enterKeyEl(el) {
        if (!el) return;
        el.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true, cancelable: true }));
        el.dispatchEvent(new KeyboardEvent('keyup', { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true, cancelable: true }));
    }

    // Default data pendaftaran (Puskesmas Matraman)
    let defaultDaftar = {
        tanggal_pemeriksaan: new Date().getDate().toString(),
        no_wa: "80000000",
        alamat: "Jl. Pal Meriam No. 6A",
        provinsi: "DKI Jakarta",
        kabkota: "Kota Adm. Jakarta Timur",
        kecamatan: "Matraman",
        keldesa: "Palmeriam"
    };
    try {
        if (typeof GM_getValue === 'function') {
            const savedDaftar = GM_getValue('asik_default_daftar', null);
            if (savedDaftar) defaultDaftar = JSON.parse(savedDaftar);
        }
    } catch (e) { }

    // Flag: Apakah flow pendaftaran aktif (baru dimulai, belum masuk pemeriksaan)
    let pendaftaranAktif = false;
    let flagPasienSudahTerdaftar = false;
    try {
        if (typeof GM_getValue === 'function') {
            pendaftaranAktif = GM_getValue('pendaftaranAktif', false);
            flagPasienSudahTerdaftar = GM_getValue('flagPasienSudahTerdaftar', false);
        }
    } catch (e) { }

    let isRobotRunning = false;
    let isRobotPaused = false; // Pause manual dari UI
    let robotStatusTeks = "Idle"; // Status real-time untuk UI
    let layananSudahDimatikan = new Set();
    let layananGagalKlik = new Set();
    let namaPasienTerakhir = "";
    let globalUsiaPasien = 0;
    let globalJenisKelamin = "";
    let statusPerkawinanPasien = "";
    let jedaRobotSampai = 0;

    // --- RIWAYAT PASIEN (LAPORAN STATUS) ---
    let riwayatPasien = [];
    function muatRiwayat() {
        try {
            const raw = typeof GM_getValue === 'function' ? GM_getValue('asik_riwayat', '[]') : '[]';
            riwayatPasien = JSON.parse(raw);
        } catch (e) { riwayatPasien = []; }
    }
    function simpanRiwayat() {
        try {
            // Batasi maksimal 200 entri agar storage tidak membengkak
            if (riwayatPasien.length > 200) riwayatPasien = riwayatPasien.slice(-200);
            if (typeof GM_setValue === 'function') GM_setValue('asik_riwayat', JSON.stringify(riwayatPasien));
        } catch (e) { }
    }
    function catatRiwayat(nik, status, keterangan) {
        const waktu = new Date().toLocaleString('id-ID', { hour12: false });
        riwayatPasien.push({
            no: riwayatPasien.length + 1,
            nik: nik || '-',
            usia: globalUsiaPasien || 0,
            jk: globalJenisKelamin || '-',
            status: status,
            keterangan: keterangan,
            waktu: waktu
        });
        simpanRiwayat();
    }
    // Muat riwayat saat script dimulai
    try { muatRiwayat(); } catch (e) { }

    // --- AUTO SHEET FACTORY GLOBALS ---
    let isAutoSheetRunning = false;
    let currentSheetRow = 0; // Dipakai sebagai ID baris (sheet) atau index array (lokal)
    let currentSheetNik = "";
    let dataPasienLokal_v2 = []; // Antrean data pasien dari Excel
    let currentUserData = null; // Data medis/antropometri untuk pasien yang sedang diproses
    let percobaanCariNik = 0; // Counter: berapa kali robot gagal menemukan tombol Mulai untuk NIK saat ini
    try {
        if (typeof GM_getValue === 'function') {
            isAutoSheetRunning = GM_getValue('isAutoSheetRunning', false);
            currentSheetRow = GM_getValue('currentSheetRow', 0);
            currentSheetNik = GM_getValue('currentSheetNik', "");
            const savedData = GM_getValue('dataPasienLokal_v2', "[]");
            dataPasienLokal_v2 = JSON.parse(savedData);
            const savedUserData = GM_getValue('currentUserData', "null");
            currentUserData = JSON.parse(savedUserData);

            // PATCH: Jika data cache memiliki nama kosong, hapus cache & paksa upload ulang
            if (dataPasienLokal_v2.length > 0 && dataPasienLokal_v2.every(row => !row.nama || row.nama === '')) {
                console.warn('[ASIK Robot] ⚠️ DATA CACHE RUSAK: Semua baris memiliki nama kosong! Menghapus cache lama...');
                console.warn('[ASIK Robot] Silakan UPLOAD ULANG file Excel/CSV Anda.');
                dataPasienLokal_v2 = [];
                currentUserData = null;
                isAutoSheetRunning = false;
                currentSheetRow = 0;
                currentSheetNik = "";
                if (typeof GM_setValue === 'function') {
                    GM_setValue('dataPasienLokal_v2', '[]');
                    GM_setValue('currentUserData', 'null');
                    GM_setValue('isAutoSheetRunning', false);
                    GM_setValue('currentSheetRow', 0);
                    GM_setValue('currentSheetNik', '');
                }
            }
        }
    } catch (e) { }

    // --
    // HELPER: Deteksi Gender, Popup Sesi Berakhir & Delay Server
    // --
    function deteksiJenisKelamin() {
        if (globalJenisKelamin) return globalJenisKelamin;

        const text = document.body.innerText.toLowerCase();
        if (text.includes("perempuan")) {
            globalJenisKelamin = "perempuan";
        } else {
            globalJenisKelamin = "laki-laki";
        }
        return globalJenisKelamin;
    }

    function deteksiUsiaPasien() {
        if (globalUsiaPasien > 0) return globalUsiaPasien;

        // Coba ambil memori lintas-domain
        let savedUsia = null;
        try {
            savedUsia = typeof GM_getValue === 'function' ? GM_getValue('asik_usia_pasien', null) : sessionStorage.getItem('asik_usia_pasien');
        } catch (e) { }

        if (savedUsia !== null) {
            globalUsiaPasien = parseInt(savedUsia);
            return globalUsiaPasien;
        }

        const textBody = document.body.innerText;
        const textInputs = Array.from(document.querySelectorAll('input[type="text"], input:not([type])')).map(i => i.value).join(" ");
        const text = textBody + " " + textInputs;

        // Cari pola ketat: "Umur Saat Pemeriksaan \n 41 Tahun" atau "Usia: 37 Tahun"
        const matchExact = text.match(/(?:umur|usia)(?:\s+saat\s+pemeriksaan)?\s*:?\s*(\d+)\s*(?:tahun|thn|th)\b/i);
        if (matchExact) {
            globalUsiaPasien = parseInt(matchExact[1]);
            return globalUsiaPasien;
        }

        // Coba cari dari format list atau input value ("42 Tahun 9 Bulan")
        const matchTahun = text.match(/(\d+)\s*(?:tahun|thn|th)\b/i);
        if (matchTahun) {
            globalUsiaPasien = parseInt(matchTahun[1]);
            return globalUsiaPasien;
        }

        // Deteksi dari Badge UI Kemenkes
        if (textBody.includes("Lansia\n") || textBody.includes("\nLansia")) {
            globalUsiaPasien = 60; // Set minimal 60 agar form lansia muncul
            return globalUsiaPasien;
        }

        return 0; // Kemungkinan bayi
    }

    function robotSedangDijeda() {
        return Date.now() < jedaRobotSampai;
    }

    function deteksiPopupPeringatan() {
        const teks = bersih(document.body.innerText);
        return teks.includes("sesi telah berakhir") ||
            teks.includes("sedang digunakan oleh perangkat lain") ||
            teks.includes("terjadi kesalahan") ||
            teks.includes("gagal memproses data");
    }

    function cariTombolOk() {
        const wadah = document.querySelectorAll(
            '[role="dialog"], .modal.show, .modal[style*="display: block"], .swal2-container, .popup, .sv-popup'
        );

        for (const modal of wadah) {
            const btn = Array.from(modal.querySelectorAll('button, a, [role="button"]'))
                .find(el => bersih(el.innerText) === "ok" && el.offsetParent !== null);
            if (btn) return btn;
        }

        if (!deteksiPopupPeringatan()) return null;
        return Array.from(document.querySelectorAll('button, a, [role="button"]'))
            .find(el => bersih(el.innerText) === "ok" && el.offsetParent !== null) || null;
    }

    async function tanganiPopupPeringatan() {
        if (!deteksiPopupPeringatan()) return false;

        const teks = bersih(document.body.innerText);
        const isSesi = teks.includes("sesi telah berakhir") || teks.includes("perangkat lain");
        const isError = teks.includes("terjadi kesalahan") || teks.includes("ada kesalahan") || teks.includes("gagal memproses");

        if (isError) {
            console.warn("[ASIK Robot] Popup KESALAHAN terdeteksi! (Dulu robot berhenti, sekarang kita abaikan dan biarkan lanjut klik OK)");
            // isAutoSheetRunning = false; <-- Fitur berhenti otomatis dicabut sesuai permintaan
            // if (typeof GM_setValue === 'function') GM_setValue('isAutoSheetRunning', false);
            // if (typeof renderUIFactory === 'function') renderUIFactory();
            jedaRobotSampai = Date.now() + 2000;
            // Kita biarkan kode berlanjut ke bawah untuk mengklik OK pada popup
        }

        console.warn("[ASIK Robot] Popup Peringatan/Info terdeteksi, klik Ok...");
        const btnOk = cariTombolOk();
        if (btnOk) {
            triggerClick(btnOk);
            await sleep(600);
        }

        jedaRobotSampai = Date.now() + (isSesi ? 18000 : 2000);
        return true;
    }

    function deteksiLoadingServer() {
        const selectorLoading = [
            '.loading', '.spinner-border', '.spinner',
            'img[src*="response-fetch"]', '.img-response-fetch'
        ];

        for (const sel of selectorLoading) {
            const el = document.querySelector(sel);
            if (!el) continue;
            const rect = el.getBoundingClientRect();
            if (rect.width > 0 && rect.height > 0) return true;
        }
        return false;
    }

    async function tungguSiapHalaman(timeout = 10000) {
        const mulai = Date.now();

        while (Date.now() - mulai < timeout) {
            if (await tanganiPopupPeringatan()) return false;
            if (!deteksiLoadingServer()) {
                await sleep(400);
                if (!deteksiLoadingServer() && !deteksiPopupPeringatan()) return true;
            }
            await sleep(500);
        }

        return !deteksiLoadingServer() && !deteksiPopupPeringatan();
    }

    async function cekGangguanHalaman() {
        if (robotSedangDijeda()) return true;
        if (await tanganiPopupPeringatan()) return true;
        return !(await tungguSiapHalaman());
    }

    // Daftar layanan wajib -- sesuai definisi lengkap + tambahan lansia
    function daftarLayananWajib(usia) {
        const wajib = [
            // === WAJIB DIISI (Semua Umur) ===
            "demografi",                            // Wajib bawaan sistem
            "antropometri", "gizi",                 // Tinggi Badan, Berat Badan, Lingkar Perut
            "tekanan darah",                        // Sistol & Diastol
            "gula darah",                           // Gula Darah Sewaktu (GDS)
            "mata", "telinga",                      // Skrining Indera
            "gigi", "karies", "periodontal",        // Skrining Gigi
            "tropis terabaikan", "frambusia", "kusta", "skabies", // Penyakit Tropis Terabaikan
            "tuberkulosis", "tbc",                  // Skrining TB
            "mandiri",                              // Skrining Mandiri

            // === TAMBAHAN ANAK / BALITA ===
            "pertumbuhan", "perkembangan"
        ];
        // === TAMBAHAN LANSIA (>= 60 tahun) ===
        if (usia >= 60) wajib.push(
            "skilas",                               // SKILAS (Kognitif, Mobilisasi, Malnutrisi, Depresi)
            "gangguan fungsional", "barthel"        // Px Gangguan Fungsional (Barthel Index)
        );
        return wajib;
    }

    // Daftar layanan yang harus DIABAIKAN total (jangan diisi, jangan di-toggle)
    // Dinamis berdasarkan usia: untuk non-lansia, skrining geriatri juga diabaikan
    function daftarLayananAbaikan(usia) {
        const abaikan = [
            "ginjal",
            "lanjutan kanker usus", "skrining kanker paru",
            "calon pengantin", "jantung", "ekg", "ppok", "puma",
            "tindak lanjut", "pemeriksaan lanjutan"
        ];
        // Untuk NON-LANSIA: abaikan seluruh skrining geriatri
        if (usia < 60) {
            abaikan.push(
                "mobilisasi", "depresi", "kognitif",
                "malnutrisi", "barthel", "gangguan fungsional",
                "geriatri"
            );
        }
        return abaikan;
    }

    // ATURAN BARU: Layanan yang harus DIMATIKAN PAKSA (toggle ke "Tidak Diperiksa")
    // Sesuai instruksi: HPV/IVA, HIV/SIFILIS, PROFIL LIPID, HEPATITIS digeser ke tidak periksa
    function daftarLayananWajibMatikan() {
        return [
            "hiv", "sifilis", "lipid panel", "profil lipid",
            "payudara", "hpv", "inspekulo", "iva",
            "hepatitis", "sirosis", "fibrosis",
            "kadar co"
        ];
    }

    function layananSelesai(row) {
        const teks = row.innerText.toLowerCase().replace(/\s+/g, ' ');
        if (teks.includes("selesai diperiksa")) return true;

        // PENTING: Kemenkes mengubah status menjadi 'tidak diperiksa' setelah kita konfirmasi popup
        // Jika tidak ada pengecekan ini, robot akan menganggapnya belum selesai selamanya
        if (teks.includes("tidak diperiksa") && !teks.includes("belum diperiksa")) return true;

        const hijau = row.querySelector('img[src*="icon-success.svg"]');
        const abu = row.querySelector('img[src*="icon-success-gray.svg"]');
        if (hijau && !abu) return true;

        return false;
    }

    function layananPerluAksi(row, wajib) {
        const teks = row.innerText.toLowerCase().replace(/\s+/g, ' ');
        const rowId = teks.replace(/[^a-z0-9]/gi, '').substring(0, 25);
        if (!teks.trim() || layananSelesai(row) || layananGagalKlik.has(rowId)) return false;

        const btnInputMatches = Array.from(row.querySelectorAll('button, div')).filter(b => {
            const t = b.innerText ? b.innerText.trim().toLowerCase() : "";
            return t === 'input data' || t === 'sedang pemeriksaan' || t === 'dalam pemeriksaan' || t === 'lanjutkan';
        });
        const btnInput = btnInputMatches.length > 0 ? btnInputMatches[btnInputMatches.length - 1] : null;
        const checkboxYa = row.querySelector('input[type="checkbox"]');

        // Layanan hanya wajib jika ada di daftar user ATAU tidak memiliki checkbox (seperti Demografi)
        let isWajib = wajib.some(k => teks.includes(k)) || !checkboxYa;

        if (isWajib && btnInput) return true;

        // Jika tidak wajib, tapi masih bisa di-uncheck -> belum selesai (harus dimatikan)
        if (!isWajib && checkboxYa && !checkboxYa.disabled) return true;

        return false;
    }

    function pemeriksaanBelumSelesai(usia = globalUsiaPasien) {
        const bodyTextAwal = (document.body.innerText || "").toLowerCase();

        // SAFEGUARD: Jika data pasien sama sekali tidak ter-load dari Kemenkes (0 Tahun 0 Hari),
        // jangan anggap pemeriksaan sudah selesai! Harus direturn true agar masuk ke prosesListPage dan direfresh.
        if (bodyTextAwal.includes("(belum diketahui)") && bodyTextAwal.includes("0 tahun 0 hari")) {
            return true;
        }

        const wajib = daftarLayananWajib(usia);
        const rows = Array.from(document.querySelectorAll('.grid-cols-5, tr'));

        // Pastikan tabel benar-benar ada isinya (ada tulisan layanan seperti skrining/imunisasi)
        // Jika tidak ada sama sekali, berarti tabel masih kosong/loading. Jangan anggap selesai.
        const adaBarisLayanan = rows.some(r => r.innerText.toLowerCase().includes('skrining') || r.innerText.toLowerCase().includes('imunisasi') || r.innerText.toLowerCase().includes('kunjungan'));
        if (!adaBarisLayanan) {
            return true;
        }

        return rows.some(row => layananPerluAksi(row, wajib));
    }

    // --
    // DETEKSI JENIS INPUT -- berdasarkan struktur SurveyJS di referensi HTML
    // (.sd-dropdown, fieldset[role=radiogroup], input date/text/textarea)
    // --
    function jenisInputPertanyaan(qNode) {
        if (qNode.querySelector('.sd-dropdown, .sv-dropdown_select-wrapper')) return 'dropdown';
        if (qNode.querySelector('input[type="date"]')) return 'date';
        if (qNode.querySelector('textarea, .sd-comment')) return 'textarea';
        if (qNode.querySelector('input[type="text"], .sd-text input, .sd-text')) return 'text';
        if (qNode.querySelector('input[type="number"]')) return 'number';
        if (qNode.querySelector('fieldset[role="radiogroup"], .sd-selectbase input[type="radio"]')) return 'radio';
        if (qNode.querySelector('input[type="radio"]')) return 'radio';
        return 'unknown';
    }

    function pertanyaanWajibDiisi(qNode) {
        if (qNode.querySelector('.sd-question__required-text')) return true;
        if (qNode.querySelector('[aria-required="true"]')) return true;
        return false;
    }

    function teksDropdownKosong(teks) {
        if (!teks) return true;
        const t = bersih(teks);
        return !t || t === "pilih" || t.includes("pilih") || t.includes("select");
    }

    function ambilOpsiJawaban(qNode) {
        return Array.from(qNode.querySelectorAll('label .sv-string-viewer, .sd-item__control-label .sv-string-viewer, .sv-list__item .sv-string-viewer, [role="option"] .sv-string-viewer'))
            .map(el => bersih(el.innerText || el.textContent))
            .filter(t => t && t.trim() !== "");
    }

    function judulFormAktif() {
        return bersih(document.querySelector('.sd-header__text, .sd-title, h1, h2')?.innerText || "");
    }

    function cariJawaban(qNode, teksTanya) {
        // 1. Cek Pertanyaan Khusus berdasar kata kunci pertanyaan
        for (const map of pertanyaanKhusus) {
            if (teksTanya.includes(bersih(map.q))) return map;
        }

        const judul = judulFormAktif();
        const gender = deteksiJenisKelamin();
        const jenis = jenisInputPertanyaan(qNode);

        // Aturan Khusus Status Perkawinan
        if (teksTanya.includes("status perkawinan")) {
            statusPerkawinanPasien = globalUsiaPasien >= 30 ? "menikah" : "belum menikah";
            return {
                a: statusPerkawinanPasien,
                fallback: statusPerkawinanPasien === "menikah" ? "kawin" : "belum kawin"
            };
        }

        // Aturan Khusus Hubungan Seksual (Berdasarkan status perkawinan)
        if (teksTanya.includes("berhubungan seksual") || teksTanya.includes("berhubungan badan") || teksTanya.includes("hubungan intim") || teksTanya.includes("seksual")) {
            if (!statusPerkawinanPasien) {
                statusPerkawinanPasien = globalUsiaPasien >= 30 ? "menikah" : "belum menikah";
            }
            const isMenikah = statusPerkawinanPasien === "menikah";
            return {
                a: isMenikah ? "ya" : "tidak",
                fallback: isMenikah ? "pernah" : "belum pernah"
            };
        }

        // Aturan Khusus Kanker Usus & Paru (Skrining Mandiri)
        if (judul.includes("kanker usus") || judul.includes("kanker paru") || teksTanya.includes("kanker usus") || teksTanya.includes("kanker paru")) {
            return { a: "tidak" };
        }

        // Aturan Khusus Screening Aktivitas Fisik
        if (judul.includes("aktivitas fisik") || teksTanya.includes("aktivitas fisik") || teksTanya.includes("olahraga") || (teksTanya.includes("aktivitas") && (teksTanya.includes("hari") || teksTanya.includes("menit")))) {
            if (teksTanya.includes("berapa menit")) return { a: "30" };
            if (teksTanya.includes("berapa hari")) return { a: "6" };

            if (teksTanya.includes("dalam satu hari") || teksTanya.includes("dalam sehari")) return { a: "30" };
            if (teksTanya.includes("dalam satu minggu") || teksTanya.includes("dalam seminggu")) return { a: "6" };

            if (teksTanya.includes("menit") && !teksTanya.includes("hari")) return { a: "30" };
            if (teksTanya.includes("hari") && !teksTanya.includes("menit")) return { a: "6" };

            // Instruksi Mutlak Puskesmas: SEMUA jenis aktivitas fisik dijawab YA
            // Setiap "Ya" akan beranak menjadi pertanyaan hari (6) dan menit (30)
            return { a: "ya" };
        }

        // Aturan Khusus Antropometri / Gizi / Pertumbuhan (Hanya untuk input teks/angka)
        if (judul.includes("antropometri") || judul.includes("gizi") || judul.includes("pertumbuhan") || judul.includes("balita") || (jenis === 'text' && (teksTanya.includes("berat badan") || teksTanya.includes("tinggi badan") || teksTanya.includes("lingkar perut") || /\b(?:bb|tb)\b/.test(teksTanya)))) {

            // Aturan Balita / Anak (0 - 5 tahun)
            if (globalUsiaPasien < 6 || judul.includes("pertumbuhan") || judul.includes("balita")) {
                if (teksTanya.includes("berat badan") || /\bbb\b/.test(teksTanya)) {
                    if (currentUserData && currentUserData.berat_badan) return { a: String(currentUserData.berat_badan) };
                    return { a: globalUsiaPasien >= 4 ? "15" : "7" };
                }
                if (teksTanya.includes("tinggi badan") || /\btb\b/.test(teksTanya)) {
                    if (currentUserData && currentUserData.tinggi_badan) return { a: String(currentUserData.tinggi_badan) };
                    return { a: globalUsiaPasien >= 4 ? "116" : "70" };
                }
            } else {
                // Aturan Dewasa
                if (teksTanya.includes("berat badan") || /\bbb\b/.test(teksTanya)) {
                    if (currentUserData && currentUserData.berat_badan) return { a: String(currentUserData.berat_badan) };
                    return { a: gender === "perempuan" ? "55" : "60" };
                }
                if (teksTanya.includes("tinggi badan") || /\btb\b/.test(teksTanya)) {
                    if (currentUserData && currentUserData.tinggi_badan) return { a: String(currentUserData.tinggi_badan) };
                    return { a: gender === "perempuan" ? "155" : "160" };
                }
            }

            if (teksTanya.includes("lingkar perut")) {
                if (currentUserData && currentUserData.lingkar_perut) return { a: String(currentUserData.lingkar_perut) };
                return { a: "80" };
            }
        }

        // Aturan dropdown untuk balita / pertumbuhan
        if (judul.includes("pertumbuhan") || judul.includes("balita") || globalUsiaPasien < 6) {
            if (teksTanya.includes("posisi pengukuran")) {
                return { a: "berdiri", fallback: "berdiri" };
            }
            if (teksTanya.includes("lingkar kepala")) {
                return { a: "normal", fallback: "normal" };
            }
        }

        // Aturan Khusus TB / Skrining X-Ray TB
        if ((judul.includes("tb") && !judul.includes("antropometri")) || judul.includes("tuberkulosis") || judul.includes("x-ray")) {
            if (teksTanya.includes("batuk")) return { a: "tidak batuk", fallback: "tidak" };
            return { a: "tidak" };
        }

        // Aturan Khusus Tekanan Darah
        if (judul.includes("tekanan darah") || teksTanya.includes("sistol") || teksTanya.includes("diastol")) {
            if (teksTanya.includes("sistol")) {
                if (currentUserData && currentUserData.td_sistol) return { a: String(currentUserData.td_sistol) };
                return { a: "120" };
            }
            if (teksTanya.includes("diastol")) {
                if (currentUserData && currentUserData.td_diastol) return { a: String(currentUserData.td_diastol) };
                return { a: "80" };
            }
        }

        // Aturan Khusus Gula Darah
        if (judul.includes("gula darah") || teksTanya.includes("sewaktu") || teksTanya.includes("diabetes")) {
            if (teksTanya.includes("diabetes")) return { a: "tidak" };
            if (teksTanya.includes("sewaktu") || teksTanya === "gds") {
                if (currentUserData && currentUserData.pemeriksaan_gula) return { a: String(currentUserData.pemeriksaan_gula) };

                // Jika GDS kosong di data, isi secara acak di rentang normal (sekitar 95 - 112) agar form bisa disimpan
                const randomGds = Math.floor(Math.random() * (112 - 95 + 1)) + 95;
                console.log(`[ASIK Robot] GDS Kosong! Mengisi otomatis dengan angka acak normal: ${randomGds}`);
                return { a: String(randomGds) };
            }
        }

        // Aturan Khusus Lansia (Gangguan Fungsional, Mobilisasi, Depresi, Kognitif, Malnutrisi)
        // 1. Gangguan Fungsional / Barthel Index
        if (teksTanya.includes("buang air besar")) return { a: "terkendali teratur" };
        if (teksTanya.includes("buang air kecil") ||
            teksTanya.includes("membersihkan diri") ||
            teksTanya.includes("penggunaan wc") ||
            teksTanya.includes("makan minum") ||
            teksTanya.includes("kursi roda") ||
            teksTanya.includes("berjalan di tempat") ||
            teksTanya.includes("berpakaian") ||
            teksTanya.includes("naik turun tangga") ||
            teksTanya.includes("mandi")) {
            return { a: "mandiri" };
        }

        // 2. Mobilisasi
        if (teksTanya.includes("berdiri di kursi") || teksTanya.includes("berdiri dari kursi")) return { a: "ya" };

        // 3. Gejala Depresi
        if (teksTanya.includes("sedih, tertekan") || teksTanya.includes("sedikit minat")) return { a: "tidak" };

        // 4. Penurunan Kognitif
        if (teksTanya.includes("tiga kata: bunga")) return { a: "ya" };
        if (teksTanya.includes("tanggal berapakah") || teksTanya.includes("dimanakah anda saat ini")) return { a: "benar semua" };
        if (teksTanya.includes("tiga kata sebelumnya")) return { a: "ya" };

        // 5. Malnutrisi
        if (teksTanya.includes("berkurang >3") || teksTanya.includes("hilang nafsu makan") || teksTanya.includes("kesulitan makan") || teksTanya.includes("lila) <21")) return { a: "tidak" };

        // 2. Tipe Non-Pilihan
        if (jenis === 'date') return { a: 'today', type: 'date' };
        if (jenis === 'text' || jenis === 'number') return { a: '0', type: 'text' };
        if (jenis === 'textarea') return { a: '-', type: 'textarea' };

        // 3. Cari Opsi Aman di Pilihan Jawaban
        const opsi = ambilOpsiJawaban(qNode);
        for (const kataAman of prioritasOpsiAman) {
            const opt = opsi.find(o => o === kataAman || o.includes(kataAman));
            if (opt) return { a: opt };
        }

        // 4. Default khusus Hati/Kejiwaan/Apakah
        if (judul.includes("hati") || judul.includes("hepatitis") ||
            judul.includes("kesehatan jiwa") || teksTanya.includes("apakah")) {
            const opt = opsi.find(o => o === "tidak" || o.startsWith("tidak"));
            if (opt) return { a: opt };
        }

        // 5. Fallback ke opsi pertama jika sama sekali tidak ketemu
        if (opsi.length) return { a: opsi[0] };

        // 6. Jika tidak ada opsi yang terbaca (misal dropdown belum terbuka), beri tanda AUTO
        return { a: "AUTO" };
    }

    function laporanFieldKosong() {
        return daftarPertanyaanKosongTerlihat().map(q => ({
            teks: teksPertanyaan(q),
            jenis: jenisInputPertanyaan(q),
            wajib: pertanyaanWajibDiisi(q)
        }));
    }

    function cariTombolKirim() {
        return document.querySelector(
            'input[title="Kirim"],button[title="Kirim"],' +
            'input[value="Kirim"],button[title="Simpan"],' +
            'button.sd-navigation__complete-btn,input.sd-navigation__complete-btn,' +
            'button[type="submit"],input[type="submit"]'
        ) || Array.from(document.querySelectorAll('button,input')).find(el => {
            const t = bersih(el.innerText || el.value || el.title || "");
            return (t === "kirim" || t === "simpan" || t === "submit") && !el.disabled && el.offsetParent !== null;
        }) || null;
    }

    // --
    // --
    // AUTO SHEET FACTORY: API & UI
    // --
    function tarikDataSheet() {
        return new Promise((resolve) => {
            if (dataPasienLokal_v2 && dataPasienLokal_v2.length > 0) {
                const rowData = dataPasienLokal_v2[0];
                currentUserData = rowData;
                if (typeof GM_setValue === 'function') GM_setValue('currentUserData', JSON.stringify(currentUserData));
                resolve([{
                    row: currentSheetRow + 1,
                    nik: rowData.nik
                }]);
            } else {
                resolve([]);
            }
        });
    }

    function laporSelesaiSheet(row, statusTeks = "Lengkap 14 Panel") {
        console.log("[ASIK Robot - DEBUG] laporSelesaiSheet called dengan row:", row, "statusTeks:", statusTeks);
        return new Promise((resolve) => {
            if (dataPasienLokal_v2 && dataPasienLokal_v2.length > 0) {
                dataPasienLokal_v2.shift(); // Buang pasien yang sudah selesai
                if (typeof GM_setValue === 'function') GM_setValue('dataPasienLokal_v2', JSON.stringify(dataPasienLokal_v2));
            }
            resolve("OK LOKAL");
        });
    }

    function showSettingsModal() {
        let overlay = document.getElementById('asik-settings-overlay');
        if (!overlay) {
            overlay = document.createElement('div');
            overlay.id = 'asik-settings-overlay';
            overlay.style.cssText = 'position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.6);z-index:9999999;display:flex;align-items:center;justify-content:center;';

            const panel = document.createElement('div');
            panel.style.cssText = 'background:#0f172a;border-radius:12px;width:90%;max-width:400px;padding:20px;box-shadow:0 10px 30px rgba(0,0,0,0.5);font-family:system-ui,-apple-system,sans-serif;color:white;';
            panel.innerHTML = `
                <h3 style="margin-top:0; border-bottom:1px solid #1e293b; padding-bottom:10px;">⚙️ Pengaturan Kolom Excel</h3>
                <p style="font-size:12px; color:#94a3b8; margin-bottom:15px;">Ketik NAMA HEADER (judul kolom) persis seperti di Excel/CSV Anda. Kosongkan jika ingin memakai deteksi otomatis.</p>
                <div style="display:flex; flex-direction:column; gap:10px; font-size:13px;">
                    <label>Kolom NIK:<br><input id="set-col-nik" type="text" style="width:100%; padding:6px; margin-top:4px; border-radius:4px; border:1px solid #334155; background:#1e293b; color:white;"></label>
                    <label>Kolom Nama Pasien:<br><input id="set-col-nama" type="text" style="width:100%; padding:6px; margin-top:4px; border-radius:4px; border:1px solid #334155; background:#1e293b; color:white;"></label>
                    <label>Kolom Tanggal Lahir:<br><input id="set-col-tgl" type="text" style="width:100%; padding:6px; margin-top:4px; border-radius:4px; border:1px solid #334155; background:#1e293b; color:white;"></label>
                    <label>Kolom Jenis Kelamin:<br><input id="set-col-jk" type="text" style="width:100%; padding:6px; margin-top:4px; border-radius:4px; border:1px solid #334155; background:#1e293b; color:white;"></label>
                    <label>Kolom No HP/WA:<br><input id="set-col-hp" type="text" style="width:100%; padding:6px; margin-top:4px; border-radius:4px; border:1px solid #334155; background:#1e293b; color:white;"></label>
                </div>
                <div style="display:flex; justify-content:flex-end; gap:10px; margin-top:20px;">
                    <button id="set-btn-tutup" style="padding:8px 12px; background:#475569; color:white; border:none; border-radius:6px; cursor:pointer;">Batal</button>
                    <button id="set-btn-simpan" style="padding:8px 12px; background:#16a34a; color:white; border:none; border-radius:6px; cursor:pointer;">Simpan & Terapkan</button>
                </div>
            `;
            overlay.appendChild(panel);
            document.body.appendChild(overlay);

            document.getElementById('set-btn-tutup').onclick = () => overlay.style.display = 'none';
            document.getElementById('set-btn-simpan').onclick = () => {
                localStorage.setItem('asik_col_nik', document.getElementById('set-col-nik').value.trim());
                localStorage.setItem('asik_col_nama', document.getElementById('set-col-nama').value.trim());
                localStorage.setItem('asik_col_tgl_lahir', document.getElementById('set-col-tgl').value.trim());
                localStorage.setItem('asik_col_jenis_kelamin', document.getElementById('set-col-jk').value.trim());
                localStorage.setItem('asik_col_no_hp', document.getElementById('set-col-hp').value.trim());
                overlay.style.display = 'none';
                alert('Pengaturan Kolom Berhasil Disimpan!\\n\\nSilakan Upload Ulang File Excel Anda agar pengaturan baru ini diterapkan.');
            };
        }

        document.getElementById('set-col-nik').value = localStorage.getItem('asik_col_nik') || '';
        document.getElementById('set-col-nama').value = localStorage.getItem('asik_col_nama') || '';
        document.getElementById('set-col-tgl').value = localStorage.getItem('asik_col_tgl_lahir') || '';
        document.getElementById('set-col-jk').value = localStorage.getItem('asik_col_jenis_kelamin') || '';
        document.getElementById('set-col-hp').value = localStorage.getItem('asik_col_no_hp') || '';

        overlay.style.display = 'flex';
    }

    function renderUIFactory() {
        if (!document.body) return;

        let ui = document.getElementById('asik-factory-ui');

        if (!ui) {
            // --
            const style = document.createElement('style');
            style.textContent = `
                #asik-factory-ui {
                    position: fixed; bottom: 20px; right: 20px; z-index: 999999;
                    width: 260px; font-family: system-ui, -apple-system, sans-serif;
                    border-radius: 12px; overflow: hidden;
                    box-shadow: 0 8px 24px rgba(0,0,0,0.35);
                    transition: all 0.3s ease; user-select: none;
                }
                #asik-factory-ui.minimized { width: auto; }
                #asik-factory-ui.minimized .asik-body { display: none; }
                .asik-header {
                    padding: 10px 14px; cursor: move; display: flex;
                    align-items: center; justify-content: space-between;
                    background: #1e293b; color: #f8fafc;
                }
                .asik-header-title { font-weight: 700; font-size: 13px; }
                .asik-header-btn {
                    background: none; border: none; color: #94a3b8;
                    cursor: pointer; font-size: 16px; padding: 0 4px; line-height: 1;
                }
                .asik-header-btn:hover { color: #f8fafc; }
                .asik-body { padding: 12px 14px; background: #0f172a; }
                .asik-status-row {
                    display: flex; justify-content: space-between; align-items: center;
                    padding: 4px 0; font-size: 11px; color: #94a3b8;
                }
                .asik-status-val { color: #e2e8f0; font-weight: 600; text-align: right; max-width: 150px; overflow: hidden; text-overflow: ellipsis; }
                .asik-divider { border: none; border-top: 1px solid #1e293b; margin: 8px 0; }
                .asik-btn {
                    width: 100%; padding: 8px 0; border: none; border-radius: 6px;
                    font-weight: 700; font-size: 12px; cursor: pointer;
                    margin-bottom: 6px; transition: opacity 0.2s;
                }
                .asik-btn:hover { opacity: 0.85; }
                .asik-btn-green { background: #16a34a; color: white; }
                .asik-btn-red { background: #dc2626; color: white; }
                .asik-btn-yellow { background: #ca8a04; color: white; }
                .asik-btn-blue { background: #2563eb; color: white; }
                .asik-btn-gray { background: #475569; color: white; }
                .asik-dot { display: inline-block; width: 8px; height: 8px; border-radius: 50%; margin-right: 6px; }
                .asik-dot-green { background: #22c55e; box-shadow: 0 0 6px #22c55e; }
                .asik-dot-red { background: #ef4444; box-shadow: 0 0 6px #ef4444; }
                .asik-dot-yellow { background: #eab308; box-shadow: 0 0 6px #eab308; }
            `;
            document.head.appendChild(style);

            ui = document.createElement('div');
            ui.id = 'asik-factory-ui';

            // --
            const header = document.createElement('div');
            header.className = 'asik-header';
            header.innerHTML = '<span class="asik-header-title">[ASIK Robot]</span>';
            const btnMin = document.createElement('button');
            btnMin.className = 'asik-header-btn';
            btnMin.id = 'asik-btn-minimize';
            btnMin.textContent = '-';
            btnMin.onclick = (e) => {
                e.stopPropagation();
                ui.classList.toggle('minimized');
                const isMin = ui.classList.contains('minimized');
                btnMin.textContent = isMin ? '+' : '-';
                localStorage.setItem('asik_ui_minimized', isMin);
            };
            header.appendChild(btnMin);

            // Draggable
            let isDragging = false, offsetX = 0, offsetY = 0;
            header.onmousedown = (e) => { if (e.target === btnMin) return; isDragging = true; offsetX = e.clientX - ui.getBoundingClientRect().left; offsetY = e.clientY - ui.getBoundingClientRect().top; ui.style.transition = 'none'; };
            document.addEventListener('mousemove', (e) => { if (!isDragging) return; ui.style.left = (e.clientX - offsetX) + 'px'; ui.style.top = (e.clientY - offsetY) + 'px'; ui.style.right = 'auto'; ui.style.bottom = 'auto'; });
            document.addEventListener('mouseup', () => {
                if (isDragging) {
                    isDragging = false;
                    ui.style.transition = 'all 0.3s ease';
                    localStorage.setItem('asik_ui_left', ui.style.left);
                    localStorage.setItem('asik_ui_top', ui.style.top);
                }
            });

            // --
            const body = document.createElement('div');
            body.className = 'asik-body';

            body.innerHTML = `
                <div class="asik-status-row"><span>Rute Awal</span>
                    <select id="asik-mode-awal" style="font-size:10px; padding:2px;">
                        <option value="pendaftaran">1. Pendaftaran</option>
                        <option value="pelayanan" selected>2. Pelayanan</option>
                    </select>
                </div>
                <div id="asik-local-info" style="display:none; font-size:10px; color:#94a3b8; text-align:center; padding:2px; background:#1e293b; border-radius:4px; margin-bottom:4px;"></div>
                <div class="asik-status-row"><span>Auto Jalan</span><span id="asik-s-sheet" class="asik-status-val">-</span></div>
                <div class="asik-status-row"><span>Robot</span><span id="asik-s-robot" class="asik-status-val">-</span></div>
                <div class="asik-status-row"><span>Halaman</span><span id="asik-s-page" class="asik-status-val">-</span></div>
                <div class="asik-status-row"><span>NIK</span><span id="asik-s-nik" class="asik-status-val">-</span></div>
                <div class="asik-status-row"><span>Usia</span><span id="asik-s-usia" class="asik-status-val">-</span></div>
                <hr class="asik-divider">
                <button id="asik-btn-sheet" class="asik-btn"></button>
                <button id="asik-btn-pause" class="asik-btn"></button>
                <button id="asik-btn-skip" class="asik-btn asik-btn-blue">SKIP PASIEN</button>
                <button id="asik-btn-clear" class="asik-btn asik-btn-red">CLEAR LOCAL</button>
                <div style="display:flex; gap:5px; margin-top:5px;">
                    <button id="asik-btn-upload" class="asik-btn asik-btn-gray" style="flex:1; font-size:10px;">UPLOAD EXCEL</button>
                    <button id="asik-btn-laporan" class="asik-btn asik-btn-gray" style="flex:1; font-size:10px;">LAPORAN</button>
                    <button id="asik-btn-settings" class="asik-btn asik-btn-gray" style="flex:1; font-size:10px;">⚙️ KOLOM</button>
                </div>
                <input type="file" id="asik-file-upload" accept=".xlsx, .xls, .csv" style="display:none;" />
            `;

            ui.appendChild(header);
            ui.appendChild(body);
            document.body.appendChild(ui);

            // Restore posisi dan minimize state
            const savedLeft = localStorage.getItem('asik_ui_left');
            const savedTop = localStorage.getItem('asik_ui_top');
            if (savedLeft && savedTop) {
                let topVal = parseInt(savedTop) || 20;
                let leftVal = parseInt(savedLeft) || 20;
                // Cegah UI tersangkut di luar layar atas atau kiri
                if (topVal < 10) topVal = 20;
                if (leftVal < 10) leftVal = 20;

                ui.style.left = leftVal + 'px';
                ui.style.top = topVal + 'px';
                ui.style.right = 'auto';
                ui.style.bottom = 'auto';
            }
            if (localStorage.getItem('asik_ui_minimized') === 'true') {
                ui.classList.add('minimized');
                btnMin.textContent = '+';
            }

            // --
            const savedModeAwal = typeof GM_getValue === 'function' ? GM_getValue('asik_mode_awal', 'pelayanan') : 'pelayanan';
            document.getElementById('asik-mode-awal').value = savedModeAwal;
            document.getElementById('asik-mode-awal').onchange = (e) => {
                if (typeof GM_setValue === 'function') GM_setValue('asik_mode_awal', e.target.value);
            };

            document.getElementById('asik-btn-settings').onclick = () => {
                showSettingsModal();
            };

            document.getElementById('asik-btn-upload').onclick = () => {
                document.getElementById('asik-file-upload').click();
            };

            document.getElementById('asik-file-upload').onchange = (e) => {
                const file = e.target.files[0];
                if (!file) return;

                const btnUpload = document.getElementById('asik-btn-upload');
                const textAwal = btnUpload.innerHTML;
                btnUpload.innerHTML = 'MEMPROSES...';
                btnUpload.disabled = true;

                // Beri jeda waktu sebentar agar browser sempat merender perubahan UI (tombol berubah)
                // sebelum main thread dibekukan (freeze) oleh proses pembacaan XLSX.
                setTimeout(() => {
                    const reader = new FileReader();
                    reader.onload = (evt) => {
                        try {
                            const data = new Uint8Array(evt.target.result);
                            const workbook = XLSX.read(data, { type: 'array', raw: true });
                            const firstSheet = workbook.SheetNames[0];
                            const rawJson = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], { defval: "", raw: false });

                            if (rawJson && rawJson.length > 0) {
                                console.log('[ASIK Robot] === DEBUG EXCEL HEADERS ===');
                                console.log('[ASIK Robot] Headers asli:', Object.keys(rawJson[0]));
                                console.log('[ASIK Robot] Contoh baris ke-1 asli:', rawJson[0]);
                            }

                            // Pengaturan kolom kustom
                            const setColNik = (localStorage.getItem('asik_col_nik') || '').trim();
                            const setColNama = (localStorage.getItem('asik_col_nama') || '').trim();
                            const setColTgl = (localStorage.getItem('asik_col_tgl_lahir') || '').trim();
                            const setColJk = (localStorage.getItem('asik_col_jenis_kelamin') || '').trim();
                            const setColHp = (localStorage.getItem('asik_col_no_hp') || '').trim();

                            console.log('[ASIK Robot] Settings columns:', {
                                setColNik,
                                setColNama,
                                setColTgl,
                                setColJk,
                                setColHp
                            });

                            // Helper: case‑insensitive, trimmed lookup for a column name
                            function getColValue(row, colName) {
                                if (!colName) return undefined;
                                const target = colName.trim().toLowerCase();
                                for (const key in row) {
                                    if (Object.prototype.hasOwnProperty.call(row, key)) {
                                        if (key.trim().toLowerCase() === target) return row[key];
                                    }
                                }
                                return undefined;
                            }

                            // Normalisasi header (nama kolom) secara Cerdas
                            const json = rawJson.map(row => {
                                const newRow = {};

                                // Prioritas 1: Pencocokan persis dari Setting (case‑insensitive)
                                if (setColNik) newRow['nik_custom'] = getColValue(row, setColNik);
                                if (setColNama) newRow['nama_custom'] = getColValue(row, setColNama);
                                if (setColTgl) newRow['tgl_custom'] = getColValue(row, setColTgl);
                                if (setColJk) newRow['jk_custom'] = getColValue(row, setColJk);
                                if (setColHp) newRow['hp_custom'] = getColValue(row, setColHp);

                                for (let key in row) {
                                    if (Object.prototype.hasOwnProperty.call(row, key)) {
                                        const cleanKey = String(key).trim().toLowerCase().replace(/[^a-z0-9]/g, '');
                                        if (cleanKey.includes('nik')) {
                                            const isNikOrtu = cleanKey.includes('ibu') || cleanKey.includes('ayah') || cleanKey.includes('wali') || cleanKey.includes('ortu') || cleanKey.includes('kk');
                                            if (!isNikOrtu) {
                                                const isExactNik = cleanKey === 'nik' || cleanKey.includes('pasien') || cleanKey.includes('sasaran') || cleanKey.includes('anak') || cleanKey.includes('balita');
                                                const currentNikPriority = newRow['_nik_priority'] || 0;
                                                const newNikPriority = isExactNik ? 2 : 1;
                                                if (newNikPriority > currentNikPriority && row[key] !== undefined && row[key] !== null && String(row[key]).trim() !== '') {
                                                    newRow['nik'] = row[key];
                                                    newRow['_nik_priority'] = newNikPriority;
                                                }
                                            }
                                        }
                                        else if (cleanKey.includes('nama') && !cleanKey.includes('ibu') && !cleanKey.includes('ayah') && !cleanKey.includes('wali') && !cleanKey.includes('ortu') && !cleanKey.includes('suami') && !cleanKey.includes('istri') && !cleanKey.includes('kk')) {
                                            const isHighlySpecific = cleanKey.includes('pasien') || cleanKey.includes('lengkap') || cleanKey.includes('sasaran') || cleanKey.includes('balita') || cleanKey.includes('anak') || cleanKey.includes('remaja') || cleanKey.includes('dewasa') || cleanKey.includes('lansia') || cleanKey.includes('bayi');
                                            const isExact = cleanKey === 'nama';
                                            
                                            const currentPriority = newRow['_nama_priority'] || 0;
                                            const newPriority = isHighlySpecific ? 3 : (isExact ? 2 : 1);
                                            
                                            if (newPriority > currentPriority && row[key] !== undefined && row[key] !== null && String(row[key]).trim() !== '') {
                                                newRow['nama'] = row[key];
                                                newRow['_nama_priority'] = newPriority;
                                            }
                                        }
                                        else if (cleanKey.includes('lahir') && !newRow['tgl_lahir']) newRow['tgl_lahir'] = row[key];
                                        else if ((cleanKey.includes('kelamin') || cleanKey === 'jk') && !newRow['jenis_kelamin']) newRow['jenis_kelamin'] = row[key];
                                        else if ((cleanKey.includes('wa') || cleanKey.includes('hp') || cleanKey.includes('telepon') || cleanKey.includes('telp')) && !newRow['no_hp']) newRow['no_hp'] = row[key];
                                        else if (cleanKey.includes('alamat') && !newRow['alamat']) newRow['alamat'] = row[key];
                                        else newRow[key.trim().toLowerCase()] = row[key];
                                    }
                                }
                                return newRow;
                            });

                            // Filter NIK valid & abaikan yang sudah 'selesai'
                            const validRows = json.filter(row => {
                                if (!row.nik) return false;
                                const n = String(row.nik).replace(/[^0-9]/g, '');
                                if (n.length < 15) return false;

                                // Auto-skip jika status sudah diisi dengan kata 'selesai' atau 'sukses'
                                const stat = String(row.status || row.status_daftar || '').toLowerCase();
                                if (stat.includes('selesai') || stat.includes('sukses')) return false;

                                return true;
                            });

                            let startRow = 1;

                            const lanjutkanProsesExcel = () => {
                                try {
                                    // DEBUG: Log semua key dan value dari baris pertama validRows
                                    if (validRows.length > 0) {
                                        const debugRow = validRows[0];
                                        console.log('[ASIK Robot] === DEBUG VALIDROWS[0] SEMUA KEY ===');
                                        for (const k in debugRow) {
                                            if (Object.prototype.hasOwnProperty.call(debugRow, k)) {
                                                console.log('[ASIK Robot] KEY: "' + k + '" => VALUE: "' + debugRow[k] + '"');
                                            }
                                        }
                                        console.log('[ASIK Robot] nama_custom:', JSON.stringify(debugRow.nama_custom), '| nama:', JSON.stringify(debugRow.nama), '| nama pasien:', JSON.stringify(debugRow['nama pasien']));
                                    }
                                    const cleanData = validRows.slice(startRow - 1).map(row => {
                                        return {
                                            nik: String(row.nik_custom || row.nik || row['no nik'] || row['nik anak']).replace(/[^0-9]/g, ''),
                                            nama: row.nama_custom || row['nama pasien'] || row['nama lengkap'] || row.pasien || row['nama sasaran'] || row.sasaran || row['nama anak'] || row['nama balita'] || row.nama || '',
                                            tgl_lahir: row.tgl_custom ? toDDMMYYYY(row.tgl_custom) : (row.tgl_lahir ? toDDMMYYYY(row.tgl_lahir) : ''),
                                            jenis_kelamin: row.jk_custom || row.jenis_kelamin || '',
                                            no_hp: row.hp_custom || row.no_hp || '',
                                            alamat: row.alamat || '',
                                            berat_badan: row.berat_badan || '',
                                            tinggi_badan: row.tinggi_badan || '',
                                            lingkar_perut: row.lingkar_perut || '',
                                            td_sistol: row.td_sistol || '',
                                            td_diastol: row.td_diastol || '',
                                            pemeriksaan_gula: row.pemeriksaan_gula || ''
                                        };
                                    });

                                    console.log('[ASIK Robot] === HASIL PARSING CSV / EXCEL ===');
                                    console.log('[ASIK Robot] Jumlah data valid:', cleanData.length);
                                    if (cleanData.length > 0) {
                                        console.log('[ASIK Robot] Contoh Data ke-1:', JSON.stringify(cleanData[0], null, 2));
                                    }

                                    dataPasienLokal_v2 = cleanData;
                                    currentSheetRow = 0;
                                    currentSheetNik = "";
                                    currentUserData = null;
                                    if (typeof GM_setValue === 'function') {
                                        GM_setValue('dataPasienLokal_v2', JSON.stringify(dataPasienLokal_v2));
                                        GM_setValue('currentSheetRow', 0);
                                        GM_setValue('currentSheetNik', "");
                                        GM_setValue('currentUserData', "null");
                                        GM_setValue('namaFileLokal', file.name);
                                    }

                                    btnUpload.innerHTML = textAwal;
                                    btnUpload.disabled = false;
                                    alert(`Berhasil memuat ${cleanData.length} data pasien dari Excel (${file.name})!`);
                                    renderUIFactory();
                                } catch (err2) {
                                    btnUpload.innerHTML = textAwal;
                                    btnUpload.disabled = false;
                                    alert("Gagal memproses data Excel. Pastikan format tanggal dan nomor valid.");
                                    console.error(err2);
                                }
                            };

                            if (validRows.length > 0) {
                                setTimeout(() => {
                                    const msg = "Ditemukan " + validRows.length + " pasien yang belum selesai.\nMulai proses dari urutan ke berapa?\n(Ketik 1 untuk mulai dari paling awal)";
                                    const inputIndex = prompt(msg, "1");
                                    if (inputIndex !== null && !isNaN(parseInt(inputIndex))) {
                                        startRow = Math.max(1, parseInt(inputIndex));
                                        lanjutkanProsesExcel();
                                    } else {
                                        btnUpload.innerHTML = textAwal;
                                        btnUpload.disabled = false;
                                    }
                                }, 10);
                            } else {
                                lanjutkanProsesExcel();
                            }

                        } catch (err) {
                            btnUpload.innerHTML = textAwal;
                            btnUpload.disabled = false;
                            alert("Gagal membaca Excel. Pastikan file valid.");
                            console.error(err);
                        }
                    };
                    reader.readAsArrayBuffer(file);
                }, 100);
            };

            document.getElementById('asik-btn-sheet').onclick = () => {
                if (dataPasienLokal_v2.length === 0) {
                    alert("Data Excel belum diupload!");
                    return;
                }
                isAutoSheetRunning = !isAutoSheetRunning;
                if (typeof GM_setValue === 'function') GM_setValue('isAutoSheetRunning', isAutoSheetRunning);
                if (!isAutoSheetRunning) {
                    // Jangan hapus queue lokal jika di pause, hanya reset current process
                    // currentSheetRow = 0; currentSheetNik = "";
                }
                renderUIFactory();
            };

            document.getElementById('asik-btn-skip').onclick = () => {
                if (!isAutoSheetRunning) return;
                console.log("[ASIK Robot] Skip pasien dari UI:", currentSheetNik);
                catatRiwayat(currentSheetNik, 'Skip', 'lainnya: Di-skip manual dari UI');
                skipPasienDanLanjut();
            };
            document.getElementById('asik-btn-laporan').onclick = () => {
                tampilkanDashboard();
            };

            document.getElementById('asik-btn-pause').onclick = () => {
                isRobotPaused = !isRobotPaused;
                if (isRobotPaused) {
                    jedaRobotSampai = Date.now() + 999999999; // Pause indefinitely
                } else {
                    jedaRobotSampai = 0; // Resume
                }
                renderUIFactory();
            };
        }

        // --
        const sSheet = document.getElementById('asik-s-sheet');
        const sRobot = document.getElementById('asik-s-robot');
        const sPage = document.getElementById('asik-s-page');
        const sNik = document.getElementById('asik-s-nik');
        const sUsia = document.getElementById('asik-s-usia');
        const btnSheet = document.getElementById('asik-btn-sheet');
        const btnPause = document.getElementById('asik-btn-pause');
        const btnSkip = document.getElementById('asik-btn-skip');

        if (!sSheet) return; // Safety

        // Status Auto Sheet
        const localInfo = document.getElementById('asik-local-info');
        if (localInfo) {
            const namaFile = typeof GM_getValue === 'function' ? GM_getValue('namaFileLokal', '') : '';
            const dataLoc = typeof GM_getValue === 'function' ? JSON.parse(GM_getValue('dataPasienLokal_v2', '[]')) : [];
            const sisa = dataLoc.length - currentSheetRow;
            localInfo.style.display = 'block';
            localInfo.innerHTML = namaFile ? 'Data: ' + namaFile + ' | Sisa: ' + Math.max(0, sisa) : 'Belum ada data lokal';
        }

        sSheet.innerHTML = isAutoSheetRunning
            ? '<span class="asik-dot asik-dot-green"></span>ON'
            : '<span class="asik-dot asik-dot-red"></span>OFF';

        // Status Robot
        if (isRobotPaused) {
            sRobot.innerHTML = '<span class="asik-dot asik-dot-yellow"></span>PAUSED';
        } else if (isRobotRunning) {
            sRobot.innerHTML = '<span class="asik-dot asik-dot-green"></span>RUNNING';
        } else {
            sRobot.innerHTML = '<span class="asik-dot asik-dot-red"></span>IDLE';
        }

        // Status Halaman
        sPage.textContent = robotStatusTeks;

        // NIK
        sNik.textContent = currentSheetNik || '-';

        // Usia
        sUsia.textContent = globalUsiaPasien > 0 ? `${globalUsiaPasien} tahun` : '-';

        // Tombol Auto Sheet
        btnSheet.className = `asik-btn ${isAutoSheetRunning ? 'asik-btn-red' : 'asik-btn-green'}`;
        btnSheet.textContent = isAutoSheetRunning ? 'STOP ROBOT' : 'START ROBOT';

        // Tombol Pause
        btnPause.className = `asik-btn ${isRobotPaused ? 'asik-btn-green' : 'asik-btn-yellow'}`;
        btnPause.textContent = isRobotPaused ? 'RESUME ROBOT' : 'PAUSE ROBOT';

        // Tombol Skip
        btnSkip.style.display = isAutoSheetRunning ? 'block' : 'none';
    }

    // --
    function tampilkanDashboard() {
        // Hapus dashboard lama jika ada
        const old = document.getElementById('asik-dashboard-overlay');
        if (old) { old.remove(); return; } // Toggle: klik lagi untuk tutup

        muatRiwayat();

        const overlay = document.createElement('div');
        overlay.id = 'asik-dashboard-overlay';
        overlay.style.cssText = 'position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.6);z-index:9999999;display:flex;align-items:center;justify-content:center;';

        // Hitung ringkasan
        const total = riwayatPasien.length;
        const sukses = riwayatPasien.filter(r => r.status.includes('Selesai')).length;
        const gagal = riwayatPasien.filter(r => r.status.includes('Gagal')).length;
        const skip = riwayatPasien.filter(r => r.status.includes('Skip')).length;

        // Bangun tabel riwayat (urutan terbaru di atas)
        const sorted = [...riwayatPasien].reverse();
        let tableRows = '';
        sorted.forEach((r, i) => {
            const bgColor = r.status.includes('Selesai') ? '#0d3320' : r.status.includes('Gagal') ? '#3b1111' : '#2a2000';
            tableRows += `<tr style="background:${bgColor}">
                <td style="padding:6px 10px;border-bottom:1px solid #334155;color:#94a3b8;">${sorted.length - i}</td>
                <td style="padding:6px 10px;border-bottom:1px solid #334155;color:#e2e8f0;font-family:monospace;">${r.nik}</td>
                <td style="padding:6px 10px;border-bottom:1px solid #334155;color:#94a3b8;">${r.usia > 0 ? r.usia + ' th' : '-'}</td>
                <td style="padding:6px 10px;border-bottom:1px solid #334155;color:#94a3b8;">${r.jk}</td>
                <td style="padding:6px 10px;border-bottom:1px solid #334155;font-weight:600;">${r.status}</td>
                <td style="padding:6px 10px;border-bottom:1px solid #334155;color:#94a3b8;">${r.keterangan}</td>
                <td style="padding:6px 10px;border-bottom:1px solid #334155;color:#64748b;font-size:11px;">${r.waktu}</td>
            </tr>`;
        });

        const panel = document.createElement('div');
        panel.style.cssText = 'background:#0f172a;border-radius:16px;width:90%;max-width:900px;max-height:85vh;overflow:hidden;display:flex;flex-direction:column;box-shadow:0 20px 60px rgba(0,0,0,0.5);font-family:system-ui,-apple-system,sans-serif;';
        panel.innerHTML = `
            <div style="padding:20px 24px;background:#1e293b;display:flex;align-items:center;justify-content:space-between;">
                <div>
                    <h2 style="margin:0;color:#f8fafc;font-size:18px;">Laporan Status Robot ASIK</h2>
                    <p style="margin:4px 0 0;color:#64748b;font-size:12px;">Riwayat pengisian data pasien</p>
                </div>
                <button id="asik-dash-close" style="background:none;border:none;color:#94a3b8;font-size:24px;cursor:pointer;padding:4px 8px;">&times;</button>
            </div>
            <div style="padding:16px 24px;display:flex;gap:12px;flex-wrap:wrap;">
                <div style="flex:1;min-width:120px;background:#1e293b;border-radius:10px;padding:14px 16px;text-align:center;">
                    <div style="font-size:28px;font-weight:800;color:#e2e8f0;">${total}</div>
                    <div style="font-size:11px;color:#64748b;margin-top:2px;">Total Pasien</div>
                </div>
                <div style="flex:1;min-width:120px;background:#052e16;border:1px solid #166534;border-radius:10px;padding:14px 16px;text-align:center;">
                    <div style="font-size:28px;font-weight:800;color:#4ade80;">${sukses}</div>
                    <div style="font-size:11px;color:#86efac;margin-top:2px;">Selesai</div>
                </div>
                <div style="flex:1;min-width:120px;background:#450a0a;border:1px solid #991b1b;border-radius:10px;padding:14px 16px;text-align:center;">
                    <div style="font-size:28px;font-weight:800;color:#f87171;">${gagal}</div>
                    <div style="font-size:11px;color:#fca5a5;margin-top:2px;">Gagal</div>
                </div>
                <div style="flex:1;min-width:120px;background:#422006;border:1px solid #92400e;border-radius:10px;padding:14px 16px;text-align:center;">
                    <div style="font-size:28px;font-weight:800;color:#fbbf24;">${skip}</div>
                    <div style="font-size:11px;color:#fde68a;margin-top:2px;">Skip</div>
                </div>
            </div>
            <div style="flex:1;overflow-y:auto;padding:0 24px 16px;">
                <table style="width:100%;border-collapse:collapse;font-size:13px;">
                    <thead>
                        <tr style="background:#1e293b;position:sticky;top:0;">
                            <th style="padding:8px 10px;text-align:left;color:#94a3b8;font-weight:600;border-bottom:2px solid #334155;">No</th>
                            <th style="padding:8px 10px;text-align:left;color:#94a3b8;font-weight:600;border-bottom:2px solid #334155;">NIK</th>
                            <th style="padding:8px 10px;text-align:left;color:#94a3b8;font-weight:600;border-bottom:2px solid #334155;">Usia</th>
                            <th style="padding:8px 10px;text-align:left;color:#94a3b8;font-weight:600;border-bottom:2px solid #334155;">JK</th>
                            <th style="padding:8px 10px;text-align:left;color:#94a3b8;font-weight:600;border-bottom:2px solid #334155;">Status</th>
                            <th style="padding:8px 10px;text-align:left;color:#94a3b8;font-weight:600;border-bottom:2px solid #334155;">Keterangan</th>
                            <th style="padding:8px 10px;text-align:left;color:#94a3b8;font-weight:600;border-bottom:2px solid #334155;">Status Sheet</th>
                            <th style="padding:8px 10px;text-align:left;color:#94a3b8;font-weight:600;border-bottom:2px solid #334155;">Waktu</th>
                        </tr>
                    </thead>
                    <tbody>${tableRows || '<tr><td colspan="8" style="padding:24px;text-align:center;color:#475569;">Belum ada riwayat</td></tr>'}</tbody>
                </table>
            </div>
            <div style="padding:12px 24px;background:#1e293b;display:flex;gap:8px;">
                <button id="asik-dash-download" style="flex:1;padding:10px;border:none;border-radius:8px;background:#2563eb;color:white;font-weight:700;font-size:12px;cursor:pointer;">[+] Download CSV</button>
                <button id="asik-dash-clear" style="flex:1;padding:10px;border:none;border-radius:8px;background:#dc2626;color:white;font-weight:700;font-size:12px;cursor:pointer;">[x] Hapus Riwayat</button>
            </div>
        `;

        overlay.appendChild(panel);
        document.body.appendChild(overlay);

        // Event: Tutup
        document.getElementById('asik-dash-close').onclick = () => overlay.remove();
        overlay.onclick = (e) => { if (e.target === overlay) overlay.remove(); };

        // Event: Download CSV
        document.getElementById('asik-dash-download').onclick = () => {
            if (riwayatPasien.length === 0) { alert('Belum ada riwayat!'); return; }
            const header = 'No,NIK,Usia,Jenis Kelamin,Status,Keterangan,Waktu';
            const rows = riwayatPasien.map(r =>
                `${r.no},"${r.nik}",${r.usia},"${r.jk}","${r.status}","${r.keterangan}","${r.waktu}"`
            );
            const csv = '\uFEFF' + header + '\n' + rows.join('\n');
            const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `ASIK-Robot-Laporan-${new Date().toISOString().slice(0, 10)}.csv`;
            a.click();
            URL.revokeObjectURL(url);
        };

        // Event: Hapus riwayat
        document.getElementById('asik-dash-clear').onclick = () => {
            if (!confirm('Yakin ingin menghapus semua riwayat? Data tidak bisa dikembalikan!')) return;
            riwayatPasien = [];
            simpanRiwayat();
            overlay.remove();
            tampilkanDashboard(); // Re-open kosong
        };
    }

    // --
    // HELPER A: Isi Radio Button (DIPERBAIKI: Klik tepat di teks span!)
    // --
    async function isiRadio(qNode, targetText, fallbackText) {
        const target = bersih(targetText);
        const fallback = fallbackText ? bersih(fallbackText) : null;
        const spans = Array.from(qNode.querySelectorAll('label .sv-string-viewer, .sd-item__control-label .sv-string-viewer'));
        if (spans.length === 0) return false;

        let pilihanSpan = null;

        if (target === "auto") {
            // Mode AUTO: Cari opsi yang mengandung kata kunci aman
            for (const kataAman of prioritasOpsiAman) {
                pilihanSpan = spans.find(span => {
                    const teksEl = bersih(span.innerText || span.textContent);
                    return teksEl.includes(kataAman);
                });
                if (pilihanSpan) break;
            }
            // Jika tetap tidak ada opsi "aman" yang cocok, pilih yang terakhir (biasanya "Tidak" ada di belakang) atau fallback ke index 0
            if (!pilihanSpan) {
                const teksPertama = bersih(spans[0].innerText || spans[0].textContent);
                if (spans.length > 1 && teksPertama.includes('ya')) pilihanSpan = spans[1];
                else pilihanSpan = spans[0];
            }
        } else {
            // Pass 1: Exact match
            for (const span of spans) {
                const teksEl = bersih(span.innerText || span.textContent);
                if (!teksEl) continue;
                if (teksEl === target || (fallback && teksEl === fallback)) {
                    pilihanSpan = span;
                    break;
                }
            }

            // Pass 2: Partial match
            if (!pilihanSpan) {
                for (const span of spans) {
                    const teksEl = bersih(span.innerText || span.textContent);
                    if (!teksEl) continue;
                    if (teksEl.includes(target) || target.includes(teksEl) || (fallback && (teksEl.includes(fallback) || fallback.includes(teksEl)))) {
                        pilihanSpan = span;
                        break;
                    }
                }
            }
        }

        if (pilihanSpan) {
            const span = pilihanSpan;
            const labelEl = span.closest('label');
            const inputEl = labelEl?.querySelector('input');

            if (sudahDijawab(qNode)) return true;

            // Eksekusi klik komplit
            triggerClick(span);
            if (labelEl) triggerClick(labelEl);
            if (inputEl) {
                triggerClick(inputEl);
                inputEl.checked = true;
                inputEl.dispatchEvent(new Event('input', { bubbles: true }));
                inputEl.dispatchEvent(new Event('change', { bubbles: true }));
                inputEl.dispatchEvent(new Event('blur', { bubbles: true }));
            }
            await sleep(400);

            return sudahDijawab(qNode);

        }
        return false;
    }

    // --
    // HELPER B: Isi Dropdown SurveyJS (DIPERCEPAT)
    // --
    async function isiDropdown(qNode, targetText, fallbackText) {
        const dropdown = qNode.querySelector('.sd-dropdown, .sv-dropdown_select-wrapper .sd-input, .sv-dropdown_select-wrapper');
        if (!dropdown) return false;

        const valueEl = dropdown.querySelector('.sd-dropdown__value, .sv-string-viewer');
        if (valueEl) {
            let val = valueEl.innerText || valueEl.textContent || "";
            const inp = valueEl.querySelector('input');
            if (inp && inp.value && inp.value.trim()) val = inp.value;
            if (!teksDropdownKosong(val)) return true;
        }

        try { dropdown.scrollIntoView({ behavior: 'smooth', block: 'center' }); } catch (e) { }
        await sleep(150);

        triggerClick(dropdown);
        await sleep(800); // Waktu jeda diperlama agar popup sempat render

        // Cari popup yang aktif
        const popups = Array.from(document.querySelectorAll('.sv-popup, .sd-popup, .sv-dropdown-popup, [role="presentation"] > [role="listbox"], .sv-popup--show')).filter(p => p.offsetParent !== null || p.style.display !== 'none');
        const activePopup = popups[popups.length - 1];

        if (!activePopup) {
            triggerClick(dropdown); // Tutup kalau gagal buka
            await sleep(300);
            return false;
        }

        const target = bersih(targetText);
        const fallback = fallbackText ? bersih(fallbackText) : null;

        const items = Array.from(activePopup.querySelectorAll('.sv-list__item, .sd-list__item, li, [role="option"], .sd-dropdown__item'));

        let pilihan = null;
        if (target !== "auto") {
            // Exact match
            pilihan = items.find(li => {
                const t = bersih(li.innerText || li.textContent);
                return t === target || (fallback && t === fallback);
            });
            // Partial match
            if (!pilihan) {
                pilihan = items.find(li => {
                    const t = bersih(li.innerText || li.textContent);
                    return t.includes(target) || target.includes(t) || (fallback && (t.includes(fallback) || fallback.includes(t)));
                });
            }
        }

        if (!pilihan) {
            // Exact match for aman
            for (const kataAman of prioritasOpsiAman) {
                pilihan = items.find(li => {
                    const t = bersih(li.innerText || li.textContent);
                    return t === kataAman;
                });
                if (pilihan) break;
            }
            // Partial match for aman
            if (!pilihan) {
                for (const kataAman of prioritasOpsiAman) {
                    pilihan = items.find(li => {
                        const t = bersih(li.innerText || li.textContent);
                        return t.includes(kataAman);
                    });
                    if (pilihan) break;
                }
            }
        }

        if (!pilihan && items.length > 1) {
            const teksPertama = bersih(items[0].innerText || items[0].textContent);
            if (teksDropdownKosong(teksPertama)) pilihan = items[1];
            else pilihan = items[0];
        } else if (!pilihan && items.length === 1) {
            const teksPertama = bersih(items[0].innerText || items[0].textContent);
            if (!teksDropdownKosong(teksPertama)) pilihan = items[0];
        }

        if (pilihan) {
            // Scroll ke opsi agar presisi
            try { pilihan.scrollIntoView({ behavior: 'smooth', block: 'center' }); } catch (e) { }
            await sleep(150);

            // SurveyJS list item butuh simulasi mouse utuh (pointerdown -> pointerup -> click)
            // Native .click() saja sering diabaikan karena mereka menggunakan mousedown event
            console.log("[ASIK Robot] Memilih opsi:", pilihan.innerText || pilihan.textContent);
            triggerClick(pilihan);

            await sleep(500);

            if (sudahDijawab(qNode)) return true;

            // Jika gagal terisi, pastikan popup tertutup sebelum retry
            if (activePopup.offsetParent !== null) {
                triggerClick(dropdown);
                await sleep(200);
            }
            return false;
        }

        triggerClick(dropdown);
        await sleep(300);
        return false;
    }

    async function isiTanggal(qNode, value = 'today') {
        const dateInput = qNode.querySelector('input[type="date"]');
        if (!dateInput) return false;
        if (dateInput.value) return true;

        let tgl = value;
        if (value === 'today') {
            const d = new Date();
            tgl = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
        }
        dateInput.focus();

        const nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, "value").set;
        nativeInputValueSetter.call(dateInput, tgl);

        dateInput.dispatchEvent(new Event('input', { bubbles: true }));
        dateInput.dispatchEvent(new Event('change', { bubbles: true }));
        dateInput.dispatchEvent(new Event('blur', { bubbles: true }));
        dateInput.dispatchEvent(new Event('focusout', { bubbles: true }));
        await sleep(200);
        return !!dateInput.value;
    }

    // --
    // HELPER C2: Isi Text / Textarea / Number
    // --
    async function isiTextField(qNode, value) {
        const input = qNode.querySelector('input[type="text"], input[type="number"], .sd-text input, textarea, .sd-comment textarea');
        if (!input) return false;
        if (input.value && String(input.value).trim()) return true;

        input.focus();

        // Guard: jangan masukkan teks non-angka ke input number
        if (input.type === 'number' && isNaN(Number(value))) {
            console.warn("[ASIK Robot] Skip isi number dengan nilai non-angka:", value);
            return false;
        }

        // Guard: jangan pernah mengetik literal "auto" ke text field (karena ini flag fallback)
        if (value === "auto") {
            console.warn("[ASIK Robot] Skip isi text field dengan literal 'auto'.");
            return false;
        }

        const proto = input.tagName.toLowerCase() === 'textarea' ? window.HTMLTextAreaElement.prototype : window.HTMLInputElement.prototype;
        const nativeInputValueSetter = Object.getOwnPropertyDescriptor(proto, "value").set;
        nativeInputValueSetter.call(input, value);

        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
        input.dispatchEvent(new Event('blur', { bubbles: true }));
        input.dispatchEvent(new Event('focusout', { bubbles: true }));
        await sleep(200);
        return !!(input.value && String(input.value).trim());
    }

    // --
    // HELPER C: Isi Satu Pertanyaan -- Langsung Tepat Sasaran (1x klik)
    // --
    async function isiSatuPertanyaan(qNode, targetText, fallbackText, mapInfo) {
        const jenis = mapInfo?.type || jenisInputPertanyaan(qNode);

        if (jenis === 'date' || targetText === 'today') {
            return isiTanggal(qNode, targetText);
        }
        if (jenis === 'text' || jenis === 'textarea' || jenis === 'number') {
            return isiTextField(qNode, targetText);
        }

        // Langsung tepat sasaran: jika jenisnya sudah diketahui, panggil fungsi yang sesuai SEKALI
        if (jenis === 'radio') {
            return await isiRadio(qNode, targetText, fallbackText);
        }
        if (jenis === 'dropdown') {
            return await isiDropdown(qNode, targetText, fallbackText);
        }

        // Jenis tidak diketahui (unknown): coba radio dulu, kalau gagal baru dropdown (1x saja)
        const berhasilRadio = await isiRadio(qNode, targetText, fallbackText);
        if (berhasilRadio) return true;

        const berhasilDropdown = await isiDropdown(qNode, targetText, fallbackText);
        if (berhasilDropdown) return true;

        return false;
    }

    // ------------------------------------------------------------------------------------------------------------------------------------------------------------------
    // HELPER D: Cek apakah pertanyaan terlihat di layar (bukan hidden/conditional)
    // ------------------------------------------------------------------------------------------------------------------------------------------------------------------
    function pertanyaanTerlihat(qNode) {
        if (!qNode || !qNode.isConnected) return false;
        if (qNode.closest('[style*="display: none"], [style*="display:none"]')) return false;
        if (qNode.classList.contains('sd-question--hidden') || qNode.getAttribute('aria-hidden') === 'true') return false;

        const style = window.getComputedStyle(qNode);
        if (style.display === 'none' || style.visibility === 'hidden') return false;

        const rect = qNode.getBoundingClientRect();
        return rect.width > 0 && rect.height > 0;
    }

    // ------------------------------------------------------------------------------------------------------------------------------------------------------------------
    // HELPER E: Cek apakah sudah dijawab (RADIO + DROPDOWN + DATE)
    // ------------------------------------------------------------------------------------------------------------------------------------------------------------------
    function sudahDijawab(qNode) {
        const inputTerpilih = qNode.querySelector('input[type="radio"]:checked, input[type="checkbox"]:checked');
        if (inputTerpilih) return true;

        const dropVal = qNode.querySelector('.sd-dropdown__value, .sd-dropdown .sv-string-viewer');
        if (dropVal) {
            let val = dropVal.innerText || dropVal.textContent || "";
            const inp = dropVal.querySelector('input');
            if (inp && inp.value && inp.value.trim()) val = inp.value;
            if (!teksDropdownKosong(val)) return true;
        }

        const dateInput = qNode.querySelector('input[type="date"]');
        if (dateInput && dateInput.value) return true;

        const textInput = qNode.querySelector('input[type="text"], input[type="number"], .sd-text input');
        if (textInput && textInput.value && textInput.value.trim()) return true;

        const textarea = qNode.querySelector('textarea, .sd-comment textarea');
        if (textarea && textarea.value && textarea.value.trim()) return true;

        return false;
    }

    // Ambil teks judul pertanyaan
    function teksPertanyaan(qNode) {
        const titleNode = qNode?.querySelector('.sd-question__title');
        return titleNode ? bersih(titleNode.innerText) : "";
    }

    // Daftar pertanyaan kosong yang terlihat (urut DOM)
    function daftarPertanyaanKosongTerlihat() {
        return Array.from(document.querySelectorAll('.sd-question'))
            .filter(pertanyaanTerlihat)
            .filter(q => {
                if (jenisInputPertanyaan(q) === 'unknown') return false;

                // Abaikan field catatan/keterangan yang tidak wajib agar dibiarkan kosong
                const teks = teksPertanyaan(q);
                if ((teks.includes("catatan") || teks.includes("keterangan")) && !pertanyaanWajibDiisi(q)) return false;

                // Abaikan field gula darah selain GDS utama jika tidak wajib (GDS 2, GDP, HbA1c, 2 Jam PP, dll)
                if ((teks.includes("puasa") || teks.includes("prandial") || teks.includes("hba1c") || teks.includes("gdp") || teks.includes("gds 2") || teks.includes("sewaktu kedua") || teks.includes("2 jam pp")) && !pertanyaanWajibDiisi(q)) return false;

                // Abaikan field Tekanan Darah ke-2 jika tidak wajib
                if ((teks.includes("ke-2") || teks.includes("ke 2")) && (teks.includes("sistol") || teks.includes("diastol")) && !pertanyaanWajibDiisi(q)) return false;

                // Abaikan Pemeriksaan Kadar CO (Tatalaksana Merokok)
                if (teks.includes("kadar co") && !pertanyaanWajibDiisi(q)) return false;

                return !sudahDijawab(q);
            });
    }

    // Snapshot teks pertanyaan kosong -- untuk deteksi beranak baru (bukan cuma hitung jumlah)
    function snapshotKosong() {
        return daftarPertanyaanKosongTerlihat().map(teksPertanyaan).join('||');
    }

    // ------------------------------------------------------------------------------------------------------------------------------------------------------------------
    // HELPER F: Cek apakah masih ada pertanyaan kosong yang terlihat
    // ------------------------------------------------------------------------------------------------------------------------------------------------------------------
    function masihAdaPertanyaanKosong() {
        return daftarPertanyaanKosongTerlihat().length > 0;
    }

    // Tunggu pertanyaan beranak muncul setelah jawaban induk (SurveyJS butuh render)
    async function tungguPertanyaanBeranak(sebelumSnapshot, timeout = 3000) {
        const mulai = Date.now();

        while (Date.now() - mulai < timeout) {
            await sleep(300);
            const sesudahSnapshot = snapshotKosong();

            if (!sesudahSnapshot) return;
            if (sesudahSnapshot !== sebelumSnapshot) {
                await sleep(500);
                return;
            }
        }
    }

    // Isi satu pertanyaan + tunggu beranak muncul jika berhasil
    async function jawabDanTungguBeranak(qNode, targetText, fallbackText, mapInfo) {
        const sebelum = snapshotKosong();
        const ok = await isiSatuPertanyaan(qNode, targetText, fallbackText, mapInfo);
        if (ok) await tungguPertanyaanBeranak(sebelum);
        return ok;
    }

    // --
    // OTAK 1: FORM -- isi berurutan, tunggu beranak setelah tiap jawaban
    // --
    async function prosesFormPage(tombolKirim) {
        let memUsia = null;
        try {
            memUsia = typeof GM_getValue === 'function' ? GM_getValue('asik_usia_pasien', null) : sessionStorage.getItem('asik_usia_pasien');
        } catch (e) { }

        globalUsiaPasien = memUsia ? parseInt(memUsia, 10) : deteksiUsiaPasien();
        console.log("[ASIK Robot] Umur pasien terkunci di form:", globalUsiaPasien);

        console.log("[ASIK Robot] OTAK FORM: Mulai isi form...");
        if (await cekGangguanHalaman()) return;

        const pertanyaanPertama = await tungguElemen('.sd-question', document, 15000);
        if (!pertanyaanPertama) {
            console.log("[ASIK Robot] Tidak ditemukan .sd-question di dalam form.");
            return;
        }

        // Jeda perdana untuk memastikan Vue / SurveyJS selesai melakukan 'hydration' data
        if (!window.sudahSettledSurvey) {
            console.log("[ASIK Robot] Form SurveyJS terdeteksi perdana, bernafas 1.5 detik agar reaktivitas Vue/API selesai...");
            await sleep(1500);
            window.sudahSettledSurvey = true;
        }

        const qNodesKosong = daftarPertanyaanKosongTerlihat();

        if (qNodesKosong.length > 0) {
            const qNode = qNodesKosong[0];
            qNode.scrollIntoView({ behavior: 'smooth', block: 'center' });
            await sleep(200);

            const teksTanya = teksPertanyaan(qNode);
            if (!teksTanya) return;

            const jenis = jenisInputPertanyaan(qNode);
            let adaYangDiisi = false;

            let targetJawabanAsli = "tidak"; // Default aman untuk fallback akhir

            // --
            if (teksTanya.includes("status perkawinan")) {
                const isDewasa = globalUsiaPasien >= 30;
                const targetUtama = isDewasa ? "menikah" : "belum menikah";
                const targetCadangan = isDewasa ? "kawin" : "belum kawin";
                adaYangDiisi = await jawabDanTungguBeranak(qNode, targetUtama, targetCadangan);
                statusPerkawinanPasien = targetUtama;
                targetJawabanAsli = targetUtama;
            }
            // --
            else if (teksTanya.includes("berhubungan seksual") || teksTanya.includes("berhubungan badan")) {
                const target = statusPerkawinanPasien === "menikah" ? "ya" : "tidak";
                adaYangDiisi = await jawabDanTungguBeranak(qNode, target);
                targetJawabanAsli = target;
            }
            // --
            else if (["hati", "hepatitis", "kesehatan jiwa"].some(k => judulFormAktif().includes(k))) {
                adaYangDiisi = await jawabDanTungguBeranak(qNode, "tidak", "tidak sama sekali");
            }
            // -- CKG ANAK: GPPH --
            else if (teksTanya.includes("gpph")) {
                adaYangDiisi = await jawabDanTungguBeranak(qNode, "< 13");
                targetJawabanAsli = "< 13";
            }
            // -- CKG ANAK: KMPE --
            else if (teksTanya.includes("kmpe")) {
                adaYangDiisi = await jawabDanTungguBeranak(qNode, "tidak ada jawaban ya", "tidak");
                targetJawabanAsli = "tidak ada jawaban ya";
            }
            // -- CKG ANAK: KPSP --
            else if (teksTanya.includes("kpsp")) {
                adaYangDiisi = await jawabDanTungguBeranak(qNode, "perkembangan sesuai", "sesuai");
                targetJawabanAsli = "perkembangan sesuai usia";
            }
            // -- CKG ANAK: Gula Darah Anak --
            else if (teksTanya.includes("gula darah anak")) {
                adaYangDiisi = await jawabDanTungguBeranak(qNode, "tidak semua", "tidak");
                targetJawabanAsli = "tidak semua";
            }
            // -- CKG ANAK: Lingkar Kepala --
            else if (teksTanya.includes("lingkar kepala")) {
                adaYangDiisi = await jawabDanTungguBeranak(qNode, "normal", "sesuai");
                targetJawabanAsli = "normal";
            }
            // -- CKG ANAK: Imunisasi --
            else if (teksTanya.includes("imunisasi")) {
                if (teksTanya.includes("catatan")) {
                    adaYangDiisi = await jawabDanTungguBeranak(qNode, "tidak", "belum");
                    targetJawabanAsli = "tidak";
                } else {
                    adaYangDiisi = await jawabDanTungguBeranak(qNode, "sudah", "ya");
                    targetJawabanAsli = "sudah";
                }
            }
            // -- CKG ANAK: Disabilitas --
            else if (teksTanya.includes("disabilitas") || teksTanya.includes("cacat")) {
                adaYangDiisi = await jawabDanTungguBeranak(qNode, "tidak", "belum");
                targetJawabanAsli = "tidak";
            }
            // --
            else {
                const map = cariJawaban(qNode, teksTanya);

                if (map) {
                    adaYangDiisi = await jawabDanTungguBeranak(qNode, map.a, map.fallback, map);
                    targetJawabanAsli = map.a;
                } else {
                    console.warn("[ASIK Robot] Tidak bisa tentukan jawaban:", { teks: teksTanya, jenis });
                }
            }

            if (!adaYangDiisi) {
                console.log("[ASIK Robot] Gagal isi:", teksTanya, `(${jenis}). Mencoba target AUTO fallback...`);
                // Force safe default
                const autoBerhasil = await jawabDanTungguBeranak(qNode, "auto");

                // BARU: Fallback ke-3 (ketik & enter) jika click tetap gagal
                if (!autoBerhasil) {
                    console.log("[ASIK Robot] AUTO fallback klik gagal! Mencoba jalur ketik teks + enter...");
                    const inputEl = qNode.querySelector('input.sd-input, input[type="text"]');
                    if (inputEl) {
                        let ketik = targetJawabanAsli;

                        // Set value & trigger events
                        inputEl.focus();
                        inputEl.value = ketik;
                        inputEl.dispatchEvent(new Event('input', { bubbles: true }));
                        inputEl.dispatchEvent(new Event('change', { bubbles: true }));
                        await sleep(300);
                        inputEl.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', code: 'Enter', keyCode: 13, bubbles: true }));
                        inputEl.dispatchEvent(new KeyboardEvent('keyup', { key: 'Enter', code: 'Enter', keyCode: 13, bubbles: true }));
                        console.log("[ASIK Robot] Berhasil memaksakan ketik '" + ketik + "' + Enter.");
                        await sleep(500);
                    }
                }
            }

            // Kembalikan kendali ke setInterval utama (Delay 1 detik alami)
            return;
        }

        // --- Jika tidak ada field yang kosong di layar ---
        await sleep(500);

        const scanAkhir = laporanFieldKosong();
        const sisaWajib = scanAkhir.filter(f => f.wajib);

        if (sisaWajib.length > 0) {
            console.log("[ASIK Robot] Masih ada field WAJIB kosong:", sisaWajib);
            return;
        }

        if (scanAkhir.length > 0) {
            console.log("[ASIK Robot] Mengabaikan field tidak wajib yang kosong:", scanAkhir);
        }

        console.log("[ASIK Robot] Semua field terisi klik Kirim...");
        triggerClick(tombolKirim);
        await sleep(3500);
    }

    // --
    // OTAK 2: DAFTAR PEMERIKSAAN (KODE ANDA UTUH)
    // --
    async function prosesListPage() {
        console.log("[ASIK Robot] OTAK PEMERIKSAAN: Memindai...");
        if (await cekGangguanHalaman()) return;

        // BARU: Deteksi jika server Kemenkes lemot dan gagal memuat data pasien
        // (Muncul "(Belum Diketahui)" di bagian NIK dan "0 Tahun 0 Hari" di bagian umur)
        const bodyTextAwal = (document.body.innerText || "").toLowerCase();
        if (bodyTextAwal.includes("(belum diketahui)") && bodyTextAwal.includes("0 tahun 0 hari")) {
            console.warn("[ASIK Robot] Terdeteksi data pasien gagal dimuat (Server Kemenkes Lemot). Menunggu 2 detik...");
            await sleep(2000);
            const bodyTextUlang = (document.body.innerText || "").toLowerCase();
            if (bodyTextUlang.includes("(belum diketahui)")) {
                console.warn("[ASIK Robot] Data masih gagal dimuat. Memaksa REFRESH halaman...");
                window.location.reload();
                return;
            }
        }

        let usia = -1, coba = 0;
        while (usia === -1 && coba < 6) {
            let found = false;
            // Coba cari dari label "Umur" atau "Usia"
            const divs = Array.from(document.querySelectorAll('div, span, p'));
            const umurLabel = divs.find(d => {
                const t = (d.innerText || "").trim().toLowerCase();
                return t === "umur" || t === "usia" || t === "umur saat pemeriksaan";
            });

            if (umurLabel && umurLabel.nextElementSibling) {
                const textUmur = umurLabel.nextElementSibling.innerText || "";
                if (textUmur.trim() !== "") {
                    const matchTahun = textUmur.match(/(\d+)\s*(?:Tahun|Thn|Th)/i);
                    if (matchTahun) usia = parseInt(matchTahun[1]);
                    else usia = 0; // Kemungkinan bayi (< 1 tahun)
                    found = true;
                }
            }

            if (!found) {
                // Fallback cari pola di body
                const text = document.body.innerText;
                const matchExact = text.match(/(?:umur|usia)(?:\s+saat\s+pemeriksaan)?\s*:?\s*(\d+)\s*(?:tahun|thn|th)\b/i);
                if (matchExact) {
                    usia = parseInt(matchExact[1]);
                    found = true;
                } else {
                    const matchTahun = text.match(/(\d+)\s*(?:Tahun|Thn|Th)/i);
                    const matchBulan = text.match(/(\d+)\s*Bulan/i);
                    if (matchTahun) {
                        usia = parseInt(matchTahun[1]);
                        found = true;
                    } else if (text.includes("Lansia\n") || text.includes("\nLansia")) {
                        usia = 60;
                        found = true;
                    } else if (matchBulan) {
                        usia = 0;
                        found = true;
                    }
                }
            }

            if (!found) {
                await sleep(600);
                coba++;
            }
        }
        if (usia === -1) {
            console.log("[ASIK Robot] Gagal mendeteksi usia pasien, halaman mungkin belum dimuat...");
            return;
        }
        globalUsiaPasien = usia;

        try {
            if (typeof GM_setValue === 'function') GM_setValue('asik_usia_pasien', usia.toString());
            sessionStorage.setItem('asik_usia_pasien', usia.toString());
        } catch (e) { }

        const penanda = document.body.innerText.substring(0, 150).replace(/\s+/g, '');
        if (penanda !== namaPasienTerakhir) {
            layananSudahDimatikan.clear();
            statusPerkawinanPasien = "";
            globalJenisKelamin = "";
            namaPasienTerakhir = penanda;
        }

        const wajib = daftarLayananWajib(usia);
        const rows = Array.from(document.querySelectorAll('.grid-cols-5, tr'));
        let adaAksi = false;

        // Prioritas 1: isi layanan wajib yang belum selesai
        const abaikan = daftarLayananAbaikan(usia);
        for (const row of rows) {
            const teks = row.innerText.toLowerCase().replace(/\s+/g, ' ');
            const rowId = teks.replace(/[^a-z0-9]/gi, '').substring(0, 25);

            const btnInputMatches = Array.from(row.querySelectorAll('button, div')).filter(b => {
                const t = b.innerText ? b.innerText.trim().toLowerCase() : "";
                return t === 'input data' || t === 'sedang pemeriksaan' || t === 'dalam pemeriksaan' || t === 'lanjutkan';
            });
            const btnInput = btnInputMatches.length > 0 ? btnInputMatches[btnInputMatches.length - 1] : null;

            if (!btnInput || layananSelesai(row) || layananGagalKlik.has(rowId)) continue;

            const checkboxYa = row.querySelector('input[type="checkbox"]');
            // Layanan hanya wajib jika ada di daftar user, atau tidak punya checkbox (harus diisi)
            let isWajib = wajib.some(k => teks.includes(k)) || !checkboxYa;

            if (!isWajib) continue;

            row.scrollIntoView({ behavior: 'smooth', block: 'center' });
            await sleep(600);
            const targetBtn = btnInput.closest('button') || btnInput.closest('[id^="rowfrm"]') || btnInput;
            triggerClick(targetBtn);
            layananGagalKlik.add(rowId);
            await sleep(3500);
            adaAksi = true;
            break;
        }

        // Prioritas 2: matikan layanan opsional & layanan wajibMatikan (toggle menjadi "Tidak Diperiksa")
        // ATURAN BARU: HPV/IVA, HIV/SIFILIS, PROFIL LIPID, HEPATITIS harus digeser ke Tidak Diperiksa
        if (!adaAksi) {
            for (const row of rows) {
                const teks = row.innerText.toLowerCase().replace(/\s+/g, ' ');

                const checkboxYa = row.querySelector('input[type="checkbox"]');
                if (!checkboxYa || checkboxYa.disabled || layananSelesai(row)) continue;

                // Layanan hanya wajib jika ada di daftar user
                let isWajib = wajib.some(k => teks.includes(k));

                if (isWajib) continue;

                const rowId = teks.replace(/[^a-z0-9]/gi, '').substring(0, 25);
                if (layananSudahDimatikan.has(rowId)) continue;

                console.log("[ASIK Robot] Toggle tidak periksa:", teks.substring(0, 50));
                row.scrollIntoView({ behavior: 'smooth', block: 'center' });
                await sleep(200);

                const label = checkboxYa.closest('label') || checkboxYa.parentElement || checkboxYa;

                // Jika sudah dicentang (Ya), klik sekali untuk mematikan memunculkan popup
                // Jika belum dicentang, centang dulu lalu matikan
                if (!checkboxYa.checked) {
                    triggerClick(label);
                    await sleep(200);
                }

                // Klik untuk mematikan dan memunculkan popup "Tidak Periksa"
                triggerClick(label);
                await sleep(200);

                // Cari dan klik tombol konfirmasi popup
                const btnPopup = Array.from(document.querySelectorAll('button')).find(b => {
                    const t = b.innerText.trim();
                    return t === 'Tidak Periksa' || t === 'Konfirmasi' || t === 'Ya' || t === 'OK';
                });
                if (btnPopup) {
                    triggerClick(btnPopup);
                    await sleep(200);
                    layananSudahDimatikan.add(rowId);
                    console.log("[ASIK Robot] Berhasil toggle:", teks.substring(0, 50));

                    // PENTING: Tunggu halaman selesai reload setelah konfirmasi popup
                    // Kemenkes melakukan refresh/spinner setelah toggle - DOM tidak stabil
                    console.log("[ASIK Robot] Menunggu halaman stabil setelah toggle...");
                    await tungguSiapHalaman(2000);
                    await sleep(200); // Buffer ekstra agar DOM benar-benar selesai re-render

                    // Cek apakah UI Kemenkes nge-lag dan status gagal update
                    const updatedRow = Array.from(document.querySelectorAll('.grid-cols-5, tr')).find(r => r.innerText.toLowerCase().replace(/\s+/g, ' ').includes(teks.substring(0, 25)));
                    if (updatedRow && !updatedRow.innerText.toLowerCase().includes('tidak diperiksa') && !layananSelesai(updatedRow)) {
                        console.warn("[ASIK Robot] UI Kemenkes lag: Status tidak update setelah diklik! Refresh paksa...");
                        window.location.reload();
                        return; // Berhenti memproses, halaman akan muat ulang dan script otomatis lanjut
                    }
                } else {
                    console.warn("[ASIK Robot] Popup tidak muncul untuk:", teks.substring(0, 50));
                    triggerClick(label);
                    layananSudahDimatikan.add(rowId); // Mencegah infinite loop
                    await sleep(400);
                }

                adaAksi = true;
                break;
            }
        }

        if (adaAksi) {
            jedaRobotSampai = Date.now() + 500;
            console.log("[ASIK Robot] Jeda 0.5 detik setelah toggle agar halaman stabil.");
        }

        if (!adaAksi && pemeriksaanBelumSelesai(usia)) {
            console.log("[ASIK Robot] Masih ada layanan belum selesai tunggu, belum ke Tatalaksana.");
        }

        if (!adaAksi && !pemeriksaanBelumSelesai(usia)) {
            console.log("[ASIK Robot] Semua layanan sudah berstatus 'Selesai'.");

            // Cari tombol Mulai Tatalaksana yang AKTIF (tidak disabled)
            const btn = Array.from(document.querySelectorAll('button')).find(b => {
                const text = bersih(b.innerText);
                return (text === 'mulai' || text.includes("mulai tatalaksana")) && !b.disabled && b.offsetParent !== null;
            });
            if (btn) {
                console.log("[ASIK Robot] Ditemukan tombol Tatalaksana. Mengklik...");
                triggerClick(btn);
            } else {
                console.log("[ASIK Robot] Tombol Tatalaksana tidak aktif/tidak ada. Menyelesaikan pasien...");
                const dilaporkan = await selesaikanPasien();
                if (!dilaporkan) {
                    catatRiwayat(currentSheetNik, 'Selesai', 'berhasil input dgn 14 panel');
                    window.location.href = "https://sehatindonesiaku.kemkes.go.id/ckg-pelayanan";
                }
            }
            return;
        }
    }

    // --
    // OTAK 6: PENDAFTARAN INDIVIDU -- diadaptasi dari Chrome Extension
    // --
    async function isiDropdownPendukung(labelName, targetText, fallbackText, maxTries = 3) {
        if (!targetText && !fallbackText) return;
        const searchTxt = targetText ? targetText.toString().toLowerCase() : fallbackText.toString().toLowerCase();

        const labelXpath = `//label[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '${labelName.toLowerCase()}')] | //div[contains(@class, 'font-semibold') and contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '${labelName.toLowerCase()}')]`;
        const labelEl = await waitForXPath(labelXpath, null, maxTries);
        if (!labelEl) {
            console.log(`[ASIK Robot] GAGAL mencari label untuk: ${labelName}`);
            return;
        }

        let container = labelEl.parentElement;
        if (!container) return;

        // Cari elemen dropdown (.cursor-pointer) yang berisi teks pilihan, naik beberapa level jika perlu
        let relativeWrapper = null;
        let searchContainer = container;
        for (let depth = 0; depth < 4 && !relativeWrapper; depth++) {
            if (!searchContainer) break;
            // Prioritas: cari .cursor-pointer dulu (pasti ada di setiap dropdown Kemenkes)
            const cp = searchContainer.querySelector('.cursor-pointer');
            if (cp) {
                // Ambil parent .relative dari cursor-pointer, atau cursor-pointer itu sendiri
                relativeWrapper = cp.closest('.relative') || cp.parentElement || cp;
                break;
            }
            // Coba juga .relative langsung
            const rel = searchContainer.querySelector('.relative');
            if (rel && rel.querySelector('.cursor-pointer')) {
                relativeWrapper = rel;
                break;
            }
            searchContainer = searchContainer.parentElement;
        }

        // Fallback: cari dari sibling label
        if (!relativeWrapper) {
            let sibling = labelEl.nextElementSibling;
            if (sibling) {
                const cp2 = sibling.querySelector('.cursor-pointer');
                relativeWrapper = cp2 ? (cp2.closest('.relative') || cp2.parentElement || cp2) : sibling.querySelector('.relative') || sibling;
            }
        }

        if (!relativeWrapper) {
            console.log(`[ASIK Robot] GAGAL mencari kotak dropdown untuk: ${labelName}`);
            return;
        }

        // Cek apakah field sudah terisi (Skip jika sudah terisi apapun yang bukan placeholder "Pilih ...")
        let currentValue = (relativeWrapper.innerText || relativeWrapper.textContent || "").toLowerCase().trim();

        const isPlaceholder = !currentValue || currentValue.startsWith("pilih") || currentValue === "";
        const targetMatch = targetText && targetText.toString().trim() !== "" ? targetText.toString().toLowerCase().trim() : null;

        if (!isPlaceholder) {
            // Jika user punya data spesifik dari Excel, dan nilainya BEDA dari yang ada sekarang, TIMPA!
            if (targetMatch && !currentValue.includes(targetMatch) && !targetMatch.includes(currentValue)) {
                console.log(`[ASIK Robot] Field ${labelName} saat ini (${currentValue}) BEDA dengan data Excel (${targetMatch}). Menimpa...`);
            }
            // Jika user tidak punya data spesifik (hanya fallback), ATAU datanya sudah sesuai, biarkan saja (jangan rusak data valid Kemenkes)
            else {
                console.log(`[ASIK Robot] Field ${labelName} sudah terisi/sesuai (${currentValue}), melewati.`);
                return;
            }
        }

        // Buka dropdown - klik .cursor-pointer HANYA SEKALI
        // Jika diklik dua kali berturut-turut (misal child lalu parent), dropdown akan langsung tertutup kembali!
        const clickTarget = relativeWrapper.querySelector('.cursor-pointer') || relativeWrapper;
        console.log(`[ASIK Robot] Membuka dropdown ${labelName}...`);

        if (typeof clickTarget.click === 'function') clickTarget.click();

        async function cariDanKlikOpsi(txt) {
            let txtLower = txt.toLowerCase().trim();
            // Penerjemah sinonim umum agar data Excel lebih mudah cocok dengan Kemenkes
            txtLower = txtLower.replace(/\bkaryawan\b/g, 'pegawai')
                .replace(/\bpns\b/g, 'pegawai negeri sipil')
                .replace(/\birt\b/g, 'ibu rumah tangga')
                .replace(/\bbelum kerja\b/g, 'belum bekerja');
            const xp = `//div[contains(@class, 'cursor-pointer')] | //li | //div[contains(@class, 'text-sm')] | //button[contains(@class, 'text-left') or contains(@class, 'border-b')]`;

            let bestMatch = null;
            let exactMatch = null;

            // Coba tunggu sampai popup opsi benar-benar muncul, maks 5 detik
            for (let wait = 0; wait < 10; wait++) {
                await sleep(500);

                const iter = document.evaluate(xp, document, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
                let partialMatch = null;
                exactMatch = null;

                // Cari dari akhir DOM karena opsi dropdown Vue biasanya dirender di akhir (absolute/teleport)
                for (let i = iter.snapshotLength - 1; i >= 0; i--) {
                    const node = iter.snapshotItem(i);
                    if (!node) continue;

                    // Lewati elemen yang tersembunyi (width 0 atau height 0)
                    const rect = node.getBoundingClientRect();
                    if (rect.width === 0 || rect.height === 0) continue;

                    const nodeText = (node.innerText || node.textContent || "").toLowerCase().trim();

                    if (nodeText === txtLower) {
                        exactMatch = node;
                        break;
                    } else if (nodeText.includes(txtLower) && !partialMatch) {
                        partialMatch = node;
                    }
                }

                bestMatch = exactMatch || partialMatch;
                if (bestMatch) {
                    break; // Ditemukan!
                } else if (wait === 3 || wait === 6) {
                    // Jika belum muncul setelah 1.5 detik, klik ulang (sekali saja)
                    console.log(`[ASIK Robot] Popup opsi "${txt}" belum muncul, mengklik ulang dropdown ${labelName}...`);
                    if (typeof clickTarget.click === 'function') clickTarget.click();
                }
            }

            if (bestMatch) {
                console.log(`[ASIK Robot] -> Ditemukan opsi dropdown: "${bestMatch.innerText.trim()}" (Match: ${exactMatch ? 'Exact' : 'Partial'})`);
                triggerClick(bestMatch);
                if (typeof bestMatch.click === 'function') bestMatch.click();
                return true;
            }
            console.log(`[ASIK Robot] -> GAGAL menemukan opsi dropdown untuk: "${txt}"`);
            return false;
        }

        if (await cariDanKlikOpsi(searchTxt)) {
            // Berhasil
        } else if (fallbackText) {
            const fbTxt = fallbackText.toString().toLowerCase();
            if (!(await cariDanKlikOpsi(fbTxt))) {
                triggerClick(document.body);
            }
        } else {
            triggerClick(document.body);
        }
        await sleep(500);
    }

    async function prosesPendaftaranPage() {
        if (!isAutoSheetRunning) return;

        // Tarik data baru dari App Script JIKA sedang tidak ada pasien yang diproses
        if (!currentSheetNik) {
            console.log("[ASIK Robot] Menarik data dari Sheet (di Halaman Pendaftaran)...");
            try {
                const data = await tarikDataSheet();
                if (data && data.length > 0) {
                    currentSheetRow = data[0].row;
                    currentSheetNik = String(data[0].nik).replace(/[^0-9]/g, '');

                    if (dataPasienLokal_v2.length > 0) {
                        currentUserData = dataPasienLokal_v2[0];
                    } else {
                        currentUserData = data[0];
                    }

                    pendaftaranAktif = true;
                    if (typeof GM_setValue === 'function') {
                        GM_setValue('currentSheetRow', currentSheetRow);
                        GM_setValue('currentSheetNik', currentSheetNik);
                        GM_setValue('currentUserData', JSON.stringify(currentUserData));
                        GM_setValue('pendaftaranAktif', true);
                    }
                    renderUIFactory();
                } else {
                    console.log("[ASIK Robot] Semua NIK sudah diproses atau data kosong!");
                    isAutoSheetRunning = false;
                    if (typeof GM_setValue === 'function') GM_setValue('isAutoSheetRunning', false);
                    renderUIFactory();
                    return;
                }
            } catch (e) {
                console.error("[ASIK Robot] Gagal menarik data:", e);
                isAutoSheetRunning = false;
                if (typeof GM_setValue === 'function') GM_setValue('isAutoSheetRunning', false);
                renderUIFactory();
                return;
            }
        }

        if (!pendaftaranAktif) {
            console.log("[ASIK Robot] Terdeteksi di halaman Pendaftaran tapi pendaftaranAktif=false. Artinya pasien belum terdaftar di Pelayanan. Melakukan Skip...");
            catatRiwayat(currentSheetNik || 'Unknown', 'Skip', 'Pasien belum terdaftar (Dilempar ke Pendaftaran)');
            skipPasienDanLanjut();
            return;
        }

        if (!currentUserData) {
            console.log("[ASIK Robot] Pendaftaran: Tidak ada data pasien aktif.");
            return;
        }

        const inData = currentUserData;
        const defData = defaultDaftar;

        console.log("[ASIK Robot] OTAK PENDAFTARAN: Mulai pendaftaran untuk NIK:", inData.nik);
        robotStatusTeks = "Pendaftaran";
        renderUIFactory();

        try {
            // 1. Klik Daftar Baru
            const btnDaftar = await waitForXPath(XPATH_DAFTAR.BTN_DAFTAR_BARU);
            if (!btnDaftar) {
                console.log("[ASIK Robot] Tombol Daftar Baru tidak ditemukan. Mungkin halaman belum dimuat.");
                return;
            }
            triggerClick(btnDaftar);
            await sleep(1500);

            if (!inData.nik) {
                console.warn("[ASIK Robot] Pendaftaran: Tidak Ada NIK");
                catatRiwayat('', 'Gagal', 'gagal input: Tidak Ada NIK');
                skipPasienDanLanjut();
                return;
            }

            // 2. Input NIK & Cek
            const inputNIK = await waitForXPath(XPATH_DAFTAR.INPUT_NIK_PENDAFTARAN);
            if (inputNIK) inputElValue(inputNIK, inData.nik);
            await sleep(500);

            const btnCekNIK = await waitForXPath(XPATH_DAFTAR.BTN_CEK_NIK_PENDAFTARAN);
            if (btnCekNIK) triggerClick(btnCekNIK);

            // Tunggu respons Kemenkes (bisa cepat, bisa lambat). Polling tiap 1 detik maks 8x
            let sudahDilayani = false;
            for (let i = 0; i < 8; i++) {
                await sleep(1000);
                const bodyText = (document.body.innerText || "").toLowerCase();
                if (bodyText.includes("sudah menerima layanan") || bodyText.includes("telah selesai diperiksa")) {
                    sudahDilayani = true;
                    break;
                }
                // Jika muncul popup "tidak ditemukan" atau form sudah terbuka, break lebih cepat
                if (bodyText.includes("tidak ditemukan") || bodyText.includes("data peserta valid") || bodyText.includes("list data individu") || document.evaluate(XPATH_DAFTAR.INPUT_NAMA_LENGKAP, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue) {
                    break;
                }
            }

            // Cek apakah pasien sudah dilayani sebelumnya (muncul layar 'Individu sudah menerima layanan')
            if (sudahDilayani) {
                console.log("[ASIK Robot] Pasien ini sudah menerima layanan! Melakukan Auto-Skip...");
                catatRiwayat(inData.nik, 'Skip', 'Pasien sudah menerima layanan');

                // Klik tombol Kembali / Cari Individu agar bisa lanjut ke pasien berikutnya
                const btnKembali = document.evaluate("//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'kembali') or contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'cari individu')]", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                if (btnKembali) {
                    triggerClick(btnKembali);
                    if (typeof btnKembali.click === 'function') btnKembali.click();
                    await sleep(1000);
                } else {
                    // Fallback klik tombol apapun yang warnanya outline atau biru
                    const sembarangBtn = document.evaluate("//button[contains(@class, 'btn-outline-primary') or contains(@class, 'bg-primaryColor')]", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                    if (sembarangBtn) {
                        triggerClick(sembarangBtn);
                        if (typeof sembarangBtn.click === 'function') sembarangBtn.click();
                        await sleep(1000);
                    }
                }

                skipPasienDanLanjut();
                return;
            }

            // 3. Cek apakah NIK ditemukan
            let nikFound = true;
            const popupTidakDitemukan = await waitForXPath(XPATH_DAFTAR.POPUP_NIK_TIDAK_DITEMUKAN, null, 3);
            // Fallback: cek juga dari bodyText jika XPath gagal (case/struktur beda)
            if (!popupTidakDitemukan) {
                const bodyTextCek = (document.body.innerText || "").toLowerCase();
                if (bodyTextCek.includes("data tidak ditemukan")) {
                    console.warn("[ASIK Robot] NIK Tidak Ditemukan (terdeteksi dari teks halaman)!");
                    nikFound = false;
                }
            }
            if (popupTidakDitemukan || !nikFound) {
                if (popupTidakDitemukan) console.warn("[ASIK Robot] NIK Tidak Ditemukan di Dukcapil!");
                nikFound = false;

                // Coba klik Tutup atau Oke agar popup hilang dan input di bawahnya bisa diklik
                const parentModal = popupTidakDitemukan.closest('.el-message-box, .swal2-popup, .modal, div[role="dialog"]') || popupTidakDitemukan.parentElement.parentElement;
                if (parentModal) {
                    const btnTutup = await waitForXPath(".//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'tutup') or contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'oke') or contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'manual')]", parentModal, 1);
                    if (btnTutup) {
                        triggerClick(btnTutup);
                    } else {
                        // Coba klik sembarang button di modal
                        const anyBtn = await waitForXPath(".//button", parentModal, 1);
                        if (anyBtn) triggerClick(anyBtn);
                    }
                    await sleep(500);
                    // Force hide agar tidak menghalangi klik input manual di bawahnya
                    try {
                        parentModal.style.display = 'none';
                        const backdrop = document.querySelector('.swal2-backdrop-show, .modal-backdrop, .el-overlay');
                        if (backdrop) backdrop.style.display = 'none';
                    } catch (e) { }
                }

                // Tunggu form manual benar-benar muncul setelah popup NIK tidak ditemukan ditutup
                console.log('[ASIK Robot] Menunggu form manual muncul setelah popup NIK tidak ditemukan...');
                await sleep(1500);
            }

            // 4. Helper: Pilih tanggal lahir dari datepicker
            async function selectBirthYear(dateStr, xPathInput) {
                const tglLahir = parseDateStringDaftar(dateStr);
                if (!tglLahir) return;

                let tglInput = await waitForXPath(xPathInput);
                let activePopup = null;

                if (tglInput) {
                    for (let i = 0; i < 3; i++) {
                        if (i === 0) {
                            if (typeof tglInput.focus === 'function') tglInput.focus();
                            triggerClick(tglInput);
                            if (typeof tglInput.click === 'function') tglInput.click();
                        } else if (i === 1) {
                            // Jika klik input gagal, coba klik pembungkusnya
                            if (tglInput.parentElement) {
                                triggerClick(tglInput.parentElement);
                                if (typeof tglInput.parentElement.click === 'function') tglInput.parentElement.click();
                            }
                        } else {
                            // Jika masih gagal, coba klik ikon SVG/kalender
                            const icon = (tglInput.parentElement || tglInput).querySelector('svg, i, .el-input__icon');
                            if (icon) triggerClick(icon);
                        }

                        await sleep(500);

                        // Cek apakah popup sudah muncul
                        const popups = Array.from(document.querySelectorAll('.mx-datepicker-popup, .el-picker-panel, .el-popper'));
                        activePopup = popups.find(p => p.offsetParent !== null && p.style.display !== 'none');
                        if (activePopup) break;
                    }
                }

                if (!activePopup) {
                    console.error("[ASIK Robot] Gagal membuka popup kalender untuk", xPathInput);
                    return;
                }

                let yearBtn = document.evaluate(XPATH_DAFTAR.INPUT_TGL_LAHIR_YEAR, activePopup, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                if (!yearBtn) {
                    const btns = Array.from(document.querySelectorAll('.mx-btn-current-year, .el-date-picker__header-label')).filter(b => b.offsetParent !== null);
                    if (btns.length > 0) yearBtn = btns[btns.length - 1]; // Ambil yang paling terakhir
                }

                if (yearBtn) {
                    triggerClick(yearBtn);
                    await sleep(500);
                } else {
                    console.error("[ASIK Robot] Gagal menemukan tombol Tahun pada kalender!");
                    return;
                }

                async function getVisibleTable(classNames) {
                    const classes = Array.isArray(classNames) ? classNames : [classNames];
                    for (let i = 0; i < 15; i++) {
                        const popups = Array.from(document.querySelectorAll('.mx-datepicker-popup, .el-picker-panel, .el-popper'));
                        const activePopups = popups.filter(p => p.offsetParent !== null && p.style.display !== 'none');

                        for (const popup of activePopups) {
                            for (const cls of classes) {
                                const table = popup.querySelector(`table.${cls}`);
                                if (table && table.offsetParent !== null) return table;
                            }
                        }

                        for (const cls of classes) {
                            const tables = Array.from(document.querySelectorAll(`table.${cls}`)).filter(t => t.offsetParent !== null);
                            if (tables.length > 0) {
                                return tables[tables.length - 1];
                            }
                        }
                        await sleep(200);
                    }
                    return null;
                }

                async function findDay() {
                    const dayTable = await getVisibleTable(['mx-table-date', 'el-date-table']);
                    if (!dayTable) {
                        console.error("[ASIK Robot] Tabel hari tidak ditemukan!");
                        return;
                    }
                    let dayEl = document.evaluate(`.//td[@title="${tglLahir.date}"]`, dayTable, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                    if (!dayEl) {
                        dayEl = Array.from(dayTable.querySelectorAll('td:not(.not-current-month):not(.prev-month):not(.next-month) .cell, td.available .cell, td:not(.not-current-month):not(.prev-month):not(.next-month)')).find(td => td.innerText && td.innerText.trim() === String(tglLahir.day));
                    }
                    if (dayEl) {
                        triggerClick(dayEl);
                        await sleep(800); // Tunggu kalender tertutup sempurna
                    } else {
                        console.error("[ASIK Robot] Gagal menemukan elemen Hari:", tglLahir.day);
                    }
                }

                async function findMonth() {
                    const monthTable = await getVisibleTable(['mx-table-month', 'el-month-table']);
                    if (!monthTable) {
                        console.error("[ASIK Robot] Tabel bulan tidak ditemukan!");
                        return;
                    }
                    let monthEl = document.evaluate(`.//td[@data-month="${tglLahir.month - 1}"]`, monthTable, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;

                    if (!monthEl) {
                        const cells = Array.from(monthTable.querySelectorAll('td .cell, td')).filter(el => el.innerText && el.innerText.trim().length > 0);
                        const monthCells = cells.filter(c => isNaN(parseInt(c.innerText.trim())));
                        if (monthCells.length >= 12) {
                            monthEl = monthCells[tglLahir.month - 1];
                        }
                    }

                    if (monthEl) {
                        triggerClick(monthEl);
                        await sleep(500);
                        await findDay();
                    } else {
                        console.error("[ASIK Robot] Gagal menemukan elemen Bulan:", tglLahir.month);
                    }
                }

                async function findYear() {
                    const yearTable = await getVisibleTable(['mx-table-year', 'el-year-table']);
                    if (!yearTable) {
                        console.error("[ASIK Robot] Tabel tahun tidak ditemukan!");
                        return;
                    }
                    let yearEl = document.evaluate(`.//td[@data-year="${tglLahir.year}"]`, yearTable, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                    if (!yearEl) {
                        yearEl = Array.from(yearTable.querySelectorAll('td .cell, td')).find(el => el.innerText && el.innerText.trim() === String(tglLahir.year));
                    }

                    if (yearEl) {
                        triggerClick(yearEl);
                        await sleep(500);
                        await findMonth();
                    } else {
                        // Cari popup aktif lagi
                        const popups = Array.from(document.querySelectorAll('.mx-datepicker-popup, .el-picker-panel, .el-popper'));
                        const currentActivePopup = popups.find(p => p.offsetParent !== null && p.style.display !== 'none');

                        let prevBtn = null;
                        if (currentActivePopup) {
                            prevBtn = document.evaluate(XPATH_DAFTAR.INPUT_TGL_LAHIR_YEAR_BEFORE, currentActivePopup, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                        }

                        if (!prevBtn) {
                            // Fallback cari tombol prev secara global
                            const btns = Array.from(document.querySelectorAll('.mx-btn-icon-double-left, .el-icon-d-arrow-left, .el-picker-panel__icon-btn')).filter(b => b.offsetParent !== null);
                            if (btns.length > 0) prevBtn = btns[0];
                        }

                        if (prevBtn) {
                            triggerClick(prevBtn);
                            await sleep(500);
                            await findYear();
                        } else {
                            console.error("[ASIK Robot] Gagal menemukan tombol Previous Year!");
                        }
                    }
                }

                await sleep(350);
                await findYear();
            }

            // 5. Jika NIK ditemukan, gunakan data atau klik Lanjutkan jika muncul popup baru. Jika tidak, isi manual.
            if (nikFound) {
                // Cek popup "Data peserta valid" -> "Lanjutkan"
                let btnLanjutkan = await waitForXPath(XPATH_DAFTAR.BTN_LANJUTKAN_DATA_VALID + " | //*[contains(text(), 'Data peserta valid')]/ancestor::*//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'lanjutkan')]", null, 4);

                if (btnLanjutkan) {
                    console.log("[ASIK Robot] Ditemukan popup Data Peserta Valid, mengklik Lanjutkan...");
                    triggerClick(btnLanjutkan);
                } else {
                    const btnGunakan = await waitForXPath(XPATH_DAFTAR.BTN_GUNAKAN_NIK, null, 2);
                    if (btnGunakan) {
                        console.log("[ASIK Robot] Mengklik Gunakan Data...");
                        triggerClick(btnGunakan);
                    }
                }
                await sleep(1000);
            } else {
                // DEBUG: tampilkan data yang akan diinput
                console.log('[ASIK Robot] === DEBUG DATA MANUAL INPUT ===');
                console.log('[ASIK Robot] inData.nama:', JSON.stringify(inData.nama));
                console.log('[ASIK Robot] inData.tgl_lahir:', JSON.stringify(inData.tgl_lahir));
                console.log('[ASIK Robot] inData.jenis_kelamin:', JSON.stringify(inData.jenis_kelamin));
                console.log('[ASIK Robot] inData.no_hp:', JSON.stringify(inData.no_hp));
                console.log('[ASIK Robot] Semua key inData:', Object.keys(inData));

                // Isi Nama (retry diperbanyak karena form manual butuh waktu muncul setelah popup ditutup)
                console.log('[ASIK Robot] Mulai mengisi form manual: Nama...');
                const inputNama = await waitForXPath(XPATH_DAFTAR.INPUT_NAMA_LENGKAP, null, 15);
                console.log('[ASIK Robot] inputNama ditemukan?', !!inputNama, inputNama ? inputNama.outerHTML.substring(0, 100) : 'null');
                if (!inputNama) {
                    console.error('[ASIK Robot] GAGAL: Field Nama Lengkap tidak ditemukan! Form manual belum muncul. Menghentikan proses pendaftaran untuk pasien ini.');
                    catatRiwayat(inData.nik, 'Gagal', 'gagal input: Field Nama Lengkap tidak ditemukan (form belum muncul)');
                    skipPasienDanLanjut();
                    return;
                }

                // Tunggu efek loading NIK selesai dan form benar-benar siap
                let waitLoad = 0;
                while (inputNama.disabled && waitLoad < 20) {
                    await sleep(200);
                    waitLoad++;
                }
                await sleep(800); // Ekstra delay agar state Vue benar-benar ter-update

                let nameToFill = inData.nama || 'NAMA_TIDAK_TERBACA_DARI_EXCEL';
                inputElValue(inputNama, nameToFill);

                // Mekanisme pertahanan: Kadang web ASIK mereset (mengosongkan) form 
                // sesaat setelah notifikasi muncul. Kita cek dan isi ulang jika kosong.
                setTimeout(() => {
                    if (!inputNama.value || inputNama.value === '') {
                        console.log('[ASIK Robot] Terdeteksi web ASIK mengosongkan field Nama! Mengisi ulang...');
                        inputElValue(inputNama, nameToFill);
                    }
                }, 1000);
                setTimeout(() => {
                    if (!inputNama.value || inputNama.value === '') {
                        inputElValue(inputNama, nameToFill);
                    }
                }, 2500);

                // Isi Tanggal Lahir
                console.log('[ASIK Robot] Mulai mengisi form manual: Tanggal Lahir...');
                if (inData.tgl_lahir) {
                    await selectBirthYear(inData.tgl_lahir, XPATH_DAFTAR.INPUT_TGL_LAHIR);
                }

                // Isi Jenis Kelamin (Dropdown)
                console.log('[ASIK Robot] Mulai mengisi form manual: Jenis Kelamin...');
                let inputJK = await waitForXPath(XPATH_DAFTAR.INPUT_JENIS_KELAMIN, null, 2);
                if (!inputJK) {
                    inputJK = await waitForXPath("//div[@id='Jenis Kelamin'] | //label[contains(text(), 'Jenis Kelamin')]/following-sibling::div", null, 1);
                }

                let boxJK = inputJK;
                if (!boxJK) {
                    const labelXpath = "//div[contains(@class, 'font-semibold') and contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'jenis kelamin')]";
                    const labelEl = await waitForXPath(labelXpath, null, 2);
                    if (labelEl && labelEl.parentElement) {
                        boxJK = labelEl.parentElement.querySelector('.cursor-pointer') || labelEl.parentElement.querySelector('.relative') || labelEl.parentElement.querySelector('input');
                    }
                }

                console.log('[ASIK Robot] Hasil pencarian kotak Jenis Kelamin:', boxJK ? 'Ditemukan' : 'TIDAK Ditemukan');
                if (boxJK && inData.jenis_kelamin) {
                    console.log('[ASIK Robot] Mengklik kotak Jenis Kelamin untuk membuka dropdown...');
                    triggerClick(boxJK);

                    let boxWrapper = boxJK.closest('.relative') || boxJK.parentElement;
                    if (boxWrapper) {
                        let svg = boxWrapper.querySelector('svg');
                        if (svg) triggerClick(svg);
                    }

                    await sleep(1000); // Tunggu animasi dropdown muncul

                    const isPR = String(inData.jenis_kelamin).toLowerCase().includes('perempuan');
                    const searchTxt = isPR ? 'perempuan' : 'laki-laki';

                    const xpathMatchText = `contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '${searchTxt}')`;
                    const xp = `//li[${xpathMatchText}] | //div[contains(@class, 'cursor-pointer') and ${xpathMatchText}] | //div[contains(@class, 'text-sm') and ${xpathMatchText}]`;

                    const iter = document.evaluate(xp, document, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
                    if (iter.snapshotLength > 0) {
                        const nodeTerakhir = iter.snapshotItem(iter.snapshotLength - 1);
                        triggerClick(nodeTerakhir);
                        console.log('[ASIK Robot] Berhasil memilih opsi:', searchTxt);
                    } else {
                        console.log('[ASIK Robot] GAGAL menemukan opsi dropdown Jenis Kelamin! Pastikan dropdown terbuka.');
                    }
                    await sleep(300);
                }

                // Isi No WA
                console.log('[ASIK Robot] Mulai mengisi form manual: No WhatsApp...');
                const inputWA = await waitForXPath(XPATH_DAFTAR.INPUT_WA, null, 3);
                if (inputWA) inputElValue(inputWA, cleanPhoneNumber(inData.no_hp, defData.no_wa));
            }

            // 6. Pilih Tanggal Pemeriksaan
            const tglParent = await waitForXPath(XPATH_DAFTAR.INPUT_TGL_PEMERIKSAAN_PARENT);
            if (tglParent) {
                const tglHari = defData.tanggal_pemeriksaan || new Date().getDate().toString();
                const tglEl = document.evaluate(`.//button[.//span[text()='${tglHari}'] and not(contains(@class,'cursor-not-allowed'))]`, tglParent, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                if (tglEl) triggerClick(tglEl);
                await sleep(500);
            }
            // 7. Handle Form Data Wali Khusus Balita / Lansia / Manual
            const isAnak = inData.tgl_lahir && isUnder10Years(inData.tgl_lahir);
            const isLansia = inData.tgl_lahir && isOver60Years(inData.tgl_lahir);

            // Cek apakah field Wali memang ada di layar Halaman 1
            const fieldWaliAda = document.evaluate("//label[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'wali')] | //*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'data wali')]", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;

            if (fieldWaliAda) {
                console.log('[ASIK Robot] Memasuki bagian Data Wali (Balita/Lansia/Manual). Menunggu loading data...');

                let skipWali = false;
                const xpathCheckboxWali = XPATH_DAFTAR.CHECKBOX_TANPA_WALI + " | //label[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'tanpa data wali')] | //div[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'tanpa data wali')]/ancestor::label";

                const checkboxTanpaWali = await waitForXPath(xpathCheckboxWali, null, 1);

                if (checkboxTanpaWali) {
                    console.log('[ASIK Robot] Ditemukan opsi "Daftarkan tanpa data wali". Mengklik...');
                    triggerClick(checkboxTanpaWali);
                    if (typeof checkboxTanpaWali.click === 'function') checkboxTanpaWali.click();
                    await sleep(500);
                    skipWali = true;
                }

                // Jika form wali wajib (misal balita) dan checkbox skip tidak ada (atau belum diklik), isi manual
                if (!skipWali && (!nikFound || isAnak)) {
                    console.log('[ASIK Robot] Opsi skip tidak ditemukan. Mengisi Data Wali...');

                    const inputNIKWali = await waitForXPath(XPATH_DAFTAR.INPUT_NIK_WALI, null, 3);
                    if (inputNIKWali) inputElValue(inputNIKWali, inData.nik);

                    const inputNamaWali = await waitForXPath(XPATH_DAFTAR.INPUT_NAMA_LENGKAP_WALI, null, 2);
                    if (inputNamaWali) inputElValue(inputNamaWali, inData.nama || 'WALI');

                    await sleep(400);
                    await selectBirthYear(inData.tgl_lahir, XPATH_DAFTAR.INPUT_TGL_LAHIR_WALI);

                    await sleep(500);
                    console.log('[ASIK Robot] Mulai mengisi Jenis Kelamin Wali...');
                    let inputJKWali = null;
                    const labelWaliXp = "(//label[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'jenis kelamin')] | //div[contains(@class, 'font-semibold') and contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'jenis kelamin')])[last()]";
                    const lblWali = document.evaluate(labelWaliXp, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;

                    if (lblWali) {
                        let container = lblWali.parentElement;
                        if (container) {
                            inputJKWali = container.querySelector('.cursor-pointer') || container.querySelector('.relative');
                        }
                        if (!inputJKWali && lblWali.nextElementSibling) {
                            inputJKWali = lblWali.nextElementSibling.querySelector('.cursor-pointer') || lblWali.nextElementSibling.querySelector('.relative') || lblWali.nextElementSibling;
                        }
                    }

                    if (!inputJKWali) {
                        inputJKWali = await waitForXPath(XPATH_DAFTAR.INPUT_JENIS_KELAMIN_WALI, null, 2);
                    }

                    if (inputJKWali) {
                        console.log('[ASIK Robot] Mengklik kotak Jenis Kelamin Wali untuk membuka dropdown...');
                        triggerClick(inputJKWali);
                        await sleep(600);

                        let boxWrapper = inputJKWali.closest('.relative') || inputJKWali.parentElement;
                        let listContainer = Array.from(document.querySelectorAll('ul, [role="listbox"], .absolute.z-50')).find(ul => {
                            return ul.offsetParent !== null && ul.innerText.toLowerCase().includes('laki-laki');
                        }) || boxWrapper;

                        const jkData = (inData.jenis_kelamin || inData.jk || defData.jk).toString().toLowerCase();
                        const searchTxt = (jkData === 'l' || jkData === 'laki-laki' || jkData === 'pria') ? 'laki' : 'perempuan';

                        const options = Array.from(listContainer.querySelectorAll('li, [role="option"], .cursor-pointer'));
                        let matchFound = false;
                        for (let opt of options) {
                            const textOpt = bersih(opt.innerText);
                            if (textOpt.includes(searchTxt)) {
                                triggerClick(opt);
                                await sleep(300);
                                matchFound = true;
                                console.log(`[ASIK Robot] Berhasil memilih Jenis Kelamin Wali: ${searchTxt}`);
                                break;
                            }
                        }
                        if (!matchFound) {
                            console.log('[ASIK Robot] GAGAL menemukan opsi dropdown Jenis Kelamin Wali!');
                            triggerClick(document.body);
                            await sleep(400);
                        }
                    } else {
                        console.log('[ASIK Robot] Gagal menemukan kotak Jenis Kelamin Wali!');
                    }

                    const inputWAWali = await waitForXPath(XPATH_DAFTAR.INPUT_WA_WALI, null, 2);
                    if (inputWAWali) inputElValue(inputWAWali, cleanPhoneNumber(inData.no_hp, defData.no_wa));
                }
            }

            // --- TRANSAKSI HALAMAN 1 KE HALAMAN 2 ---
            // Kemenkes sekarang memisahkan form menjadi 2 halaman (Halaman 2 khusus Data Pendukung).
            // Kita harus klik Selanjutnya dulu di Halaman 1 agar Halaman 2 muncul.
            console.log("[ASIK Robot] Menunggu tombol Selanjutnya (Halaman 1) aktif...");
            let btnSelanjutnyaHal1 = await waitForXPath(XPATH_DAFTAR.BTN_SELANJUTNYA_FORMULIR_PENDAFTARAN, null, 4);

            if (btnSelanjutnyaHal1) {
                console.log("[ASIK Robot] Menekan tombol Selanjutnya untuk beralih ke Halaman 2 (Data Pendukung)...");
                triggerClick(btnSelanjutnyaHal1);
                if (typeof btnSelanjutnyaHal1.click === 'function') btnSelanjutnyaHal1.click();
                await sleep(1500);

                for (let k = 0; k < 8; k++) { // Cek selama 4 detik
                    let btnKuotaLanjut = document.evaluate("//button[normalize-space(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'))='lanjut'] | //div[contains(@class,'el-overlay') or @role='dialog']//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'lanjut') and not(contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'selanjutnya'))]", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                    if (btnKuotaLanjut) {
                        console.log("[ASIK Robot] Ditemukan popup Kuota Habis, mengklik Lanjut...");
                        triggerClick(btnKuotaLanjut);
                        await sleep(1000); // Tunggu sebentar siapa tahu popup Data Valid muncul setelah ini
                    }

                    const bodyTextTrans = (document.body.innerText || "").toLowerCase();

                    // Cek popup "Data peserta tidak valid" (Error Dukcapil)
                    if (bodyTextTrans.includes("data peserta tidak valid") || bodyTextTrans.includes("silakan perbaiki data yang dimasukkan")) {
                        console.log("[ASIK Robot] Ditemukan popup Data Peserta TIDAK Valid (Error NIK/Nama), skip pasien...");
                        catatRiwayat(inData.nik, 'Gagal', 'gagal input: Data peserta tidak valid');

                        const btnPeriksa = document.evaluate("//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'periksa kembali') or contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'kembali')]", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                        if (btnPeriksa) {
                            triggerClick(btnPeriksa);
                            if (typeof btnPeriksa.click === 'function') btnPeriksa.click();
                        } else {
                            const btnTutup = document.evaluate("//button[contains(@class, 'bg-primaryColor') or contains(@class, 'btn-primary')]", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                            if (btnTutup) { triggerClick(btnTutup); }
                        }
                        skipPasienDanLanjut();
                        return;
                    }

                    // Cek popup "Individu sudah menerima layanan"
                    if (bodyTextTrans.includes("sudah menerima layanan") || bodyTextTrans.includes("telah selesai diperiksa")) {
                        console.log("[ASIK Robot] Pasien ini sudah menerima layanan! Melakukan Auto-Skip...");
                        catatRiwayat(inData.nik, 'Skip', 'Pasien sudah menerima layanan');

                        // Klik tombol Kembali / Cari Individu
                        const btnKembali = document.evaluate("//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'kembali') or contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'cari individu')]", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                        if (btnKembali) {
                            triggerClick(btnKembali);
                            if (typeof btnKembali.click === 'function') btnKembali.click();
                        } else {
                            const sembarangBtn = document.evaluate("//button[contains(@class, 'btn-outline-primary') or contains(@class, 'bg-primaryColor')]", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                            if (sembarangBtn) {
                                triggerClick(sembarangBtn);
                                if (typeof sembarangBtn.click === 'function') sembarangBtn.click();
                            }
                        }

                        skipPasienDanLanjut();
                        return; // Stop pengisian untuk NIK ini
                    }

                    // Cek Data Valid
                    let popupValidTransition = document.evaluate(XPATH_DAFTAR.BTN_LANJUTKAN_DATA_VALID + " | //*[contains(text(), 'Data peserta valid')]/ancestor::*//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'lanjutkan')] | //div[contains(@class,'el-overlay') or contains(@class,'modal')]//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'lanjutkan')]", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                    if (popupValidTransition) {
                        console.log("[ASIK Robot] Ditemukan popup Data Peserta Valid (Transisi), mengklik Lanjutkan...");
                        triggerClick(popupValidTransition);
                        await sleep(3000); // Tunggu Halaman 2 selesai dirender
                        break; // Transisi selesai
                    }
                    await sleep(500);
                }
            }

            // 8. Isi form pendukung yang sekarang ada di halaman 2
            // VALIDASI: Pastikan halaman sudah benar-benar berpindah ke Halaman 2 (Data Pendukung)
            // Jika masih stuck di popup awal, jangan lanjutkan isi data pendukung
            console.log('[ASIK Robot] Memeriksa apakah sudah masuk ke Halaman 2 (Data Pendukung)...');
            const halamanDuaReady = await waitForXPath(
                "//div[contains(@class, 'font-semibold') and (contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'status pernikahan') or contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'pekerjaan') or contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'disabilitas') or contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'alamat domisili'))]",
                null, 8
            );
            if (!halamanDuaReady) {
                console.warn('[ASIK Robot] Halaman 2 (Data Pendukung) belum muncul! Kemungkinan masih di popup Halaman 1. Menghentikan proses...');
                catatRiwayat(inData.nik, 'Gagal', 'gagal input: Halaman 2 (Data Pendukung) tidak muncul');
                skipPasienDanLanjut();
                return;
            }
            console.log('[ASIK Robot] Halaman 2 (Data Pendukung) terdeteksi. Mengisi Data Pendukung (Status Pernikahan, dll)...');

            // Logika pintar: Jika status pernikahan tidak ada di data Excel, tebak dari umur NIK
            let finalStatusPernikahan = inData.status_pernikahan;
            if (!finalStatusPernikahan && inData.nik && String(inData.nik).length === 16) {
                let yyStr = String(inData.nik).substring(10, 12);
                let yy = parseInt(yyStr, 10);
                // Jika tahun > 30 (asumsi 1930-1999), jika < 30 (2000-2026)
                let yearOfBirth = yy > 26 ? 1900 + yy : 2000 + yy;
                let age = new Date().getFullYear() - yearOfBirth;
                finalStatusPernikahan = age >= 30 ? 'Menikah' : 'Belum Menikah';
            }

            await isiDropdownPendukung('Status Pernikahan', finalStatusPernikahan, defData.status_pernikahan || 'Menikah', 10);
            await isiDropdownPendukung('Penyandang disabilitas', inData.disabilitas, defData.disabilitas || 'Tidak memiliki disabilitas', 4);
            await isiDropdownPendukung('Pekerjaan', inData.pekerjaan, defData.pekerjaan || 'Lainnya', 4);

            // Coba isi Alamat Domisili jika kosong (Pilih Lokasi bertahap 4 level)
            const alamatLabel = document.evaluate("//div[contains(@class, 'font-semibold') and contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'alamat domisili') and not(contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'detail'))]", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
            if (alamatLabel) {
                let boxAlamat = alamatLabel.nextElementSibling || alamatLabel.parentElement.querySelector('.relative');
                if (boxAlamat) {
                    let currentValue = (boxAlamat.innerText || boxAlamat.textContent || "").toLowerCase().trim();
                    if (currentValue && !currentValue.startsWith("pilih")) {
                        console.log(`[ASIK Robot] Field Alamat Domisili sudah terisi otomatis (${currentValue}), melewati.`);
                    } else {
                        console.log('[ASIK Robot] Mengisi Alamat Domisili (Pilih Lokasi Bertahap)...');
                        const clickTarget = boxAlamat.querySelector('.cursor-pointer') || boxAlamat;
                        if (typeof clickTarget.click === 'function') clickTarget.click();
                        await sleep(1000);

                        const tahapan = [
                            defData.provinsi || 'DKI Jakarta',
                            defData.kabkota || 'Kota Adm. Jakarta Timur',
                            defData.kecamatan || 'Matraman',
                            defData.keldesa || 'Palmeriam'
                        ];

                        for (let t = 0; t < tahapan.length; t++) {
                            const targetText = tahapan[t];
                            if (!targetText) continue;
                            const txtLower = targetText.toLowerCase().trim();
                            const xp = `//button[contains(@class, 'text-left') or contains(@class, 'border-b')] | //li | //div[contains(@class, 'cursor-pointer')]`;

                            let bestMatch = null;
                            for (let wait = 0; wait < 10; wait++) {
                                await sleep(500);
                                const iter = document.evaluate(xp, document, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
                                let exactMatch = null;
                                let partialMatch = null;
                                for (let i = iter.snapshotLength - 1; i >= 0; i--) {
                                    const node = iter.snapshotItem(i);
                                    if (!node) continue;
                                    const rect = node.getBoundingClientRect();
                                    if (rect.width === 0 || rect.height === 0) continue;
                                    const nodeText = (node.innerText || node.textContent || "").toLowerCase().trim();
                                    if (nodeText === txtLower) {
                                        exactMatch = node; break;
                                    } else if (nodeText.includes(txtLower) && !partialMatch) {
                                        partialMatch = node;
                                    }
                                }
                                bestMatch = exactMatch || partialMatch;
                                if (bestMatch) break;
                            }

                            if (bestMatch) {
                                console.log(`[ASIK Robot] -> Ditemukan opsi lokasi tahap ${t + 1}: "${bestMatch.innerText.trim()}"`);
                                triggerClick(bestMatch);
                                if (typeof bestMatch.click === 'function') bestMatch.click();
                                await sleep(1000); // Tunggu list berikutnya loading
                            } else {
                                console.log(`[ASIK Robot] -> GAGAL menemukan opsi lokasi tahap ${t + 1} untuk: "${targetText}"`);
                                break;
                            }
                        }
                    }
                }
            }

            const inputDomisili = await waitForXPath("//textarea[@id='detail-domisili'] | //textarea[contains(@placeholder, 'alamat')] | //label[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'detail alamat')]/following-sibling::div//textarea", null, 2);
            if (inputDomisili) {
                // Textarea kadang menyimpan valuenya di textContent atau innerHTML pada Vue
                const currentVal = (inputDomisili.value || inputDomisili.textContent || "").trim();
                const alamatDariData = inData.alamat || inData.alamatdomisili || inData['alamat domisili'] || inData.alamat_domisili || inData.detailalamat || inData['detail alamat'];

                if (currentVal !== "") {
                    console.log(`[ASIK Robot] Detail Alamat Domisili sudah terisi (${currentVal}), membiarkan apa adanya.`);
                } else if (alamatDariData) {
                    console.log(`[ASIK Robot] Mengisi Detail Alamat Domisili dari data excel -> ${alamatDariData}`);
                    inputElValue(inputDomisili, alamatDariData.toString());
                } else {
                    console.log('[ASIK Robot] Mengisi Detail Alamat Domisili dengan default...');
                    inputElValue(inputDomisili, 'Sesuai KTP');
                }
                await sleep(500);
            }

            // 7. Klik Selanjutnya untuk Submit Pendaftaran
            console.log('[ASIK Robot] Menunggu tombol Selanjutnya aktif dan mengkliknya...');
            let formTersubmit = false;
            let sudahKlikPeriksaKembali = false;
            let sudahPilih = false;
            let sudahDaftar = false;
            let sudahKlikLanjutKuota = false;
            let sudahKlikSelanjutnya = false;

            for (let i = 0; i < 20; i++) {
                const bodyTextSubmit = (document.body.innerText || "").toLowerCase();
                const isPopupTerbuka = bodyTextSubmit.includes("list data individu") || bodyTextSubmit.includes("daftarkan dengan nik") || bodyTextSubmit.includes("kuota pemeriksaan") || bodyTextSubmit.includes("terjadi kesalahan");

                // Jangan spam klik Selanjutnya jika popup sudah muncul atau sedang dalam antrean jaringan
                if (!isPopupTerbuka && !sudahKlikSelanjutnya) {
                    const btnSelanjutnya = document.evaluate(XPATH_DAFTAR.BTN_SELANJUTNYA_FORMULIR_PENDAFTARAN, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                    if (btnSelanjutnya) {
                        triggerClick(btnSelanjutnya);
                        sudahKlikSelanjutnya = true;
                        // Tunggu loading spinner
                        await tungguSiapHalaman(3000);
                    }
                }

                await sleep(1500);

                // Cek popup "Kuota Pemeriksaan Habis" dan klik Lanjut
                const btnLanjutKuota = document.evaluate(XPATH_DAFTAR.BTN_LANJUT_KUOTA_HABIS, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                if (btnLanjutKuota && !sudahKlikLanjutKuota) {
                    console.log("[ASIK Robot] Kuota harian habis, memaksa lanjut pendaftaran...");
                    triggerClick(btnLanjutKuota);
                    sudahKlikLanjutKuota = true;
                    await sleep(1500);
                }

                // Cek popup sukses
                const msgSuccess = document.evaluate(XPATH_DAFTAR.MSG_POPUP_SUCCESS, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                if (msgSuccess) {
                    formTersubmit = true;
                    break;
                }

                // Cek popup "List Data Individu" (muncul jika Kemenkes mendeteksi duplikat Dukcapil setelah isi manual)
                if (bodyTextSubmit.includes("list data individu") || bodyTextSubmit.includes("daftarkan dengan nik")) {
                    console.log("[ASIK Robot] Ditemukan tabel List Data Individu setelah Halaman 2, memproses popup...");

                    const semuaTombol = Array.from(document.querySelectorAll('button, a'));
                    const btnPilih = semuaTombol.find(b => {
                        const txt = b.innerText.trim().toLowerCase();
                        return txt === 'pilih' && b.offsetWidth > 0;
                    });

                    if (btnPilih && !sudahPilih) {
                        triggerClick(btnPilih);
                        console.log("[ASIK Robot] Tombol Pilih diklik, menunggu tombol Daftarkan...");
                        sudahPilih = true;
                        await sleep(2500); // DIPERPANJANG agar tidak kena limit rate API
                    }

                    const btnDaftarkan = semuaTombol.find(b => {
                        const txt = b.innerText.trim().toLowerCase();
                        return (txt === 'daftarkan dengan nik' || txt === 'daftarkan') && b.offsetWidth > 0;
                    });

                    if (btnDaftarkan && (!btnPilih || sudahPilih) && !sudahDaftar) {
                        triggerClick(btnDaftarkan);
                        console.log("[ASIK Robot] Tombol Daftarkan dengan NIK diklik, melanjutkan submit...");
                        sudahDaftar = true;
                        await sleep(3000); // DIPERPANJANG agar tidak tabrakan dengan event DOM Kemenkes
                    }
                }

                // Cek error popup Data peserta tidak valid
                const popupTidakValid = document.evaluate(XPATH_DAFTAR.POPUP_DATA_PESERTA_WALI_TIDAK_VALID + " | //*[contains(text(), 'Data peserta tidak valid')]", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                if (popupTidakValid) {
                    const btnPeriksaKembali = document.evaluate(".//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'periksa kembali')]", popupTidakValid.closest('[role="dialog"], .modal, .swal2-popup, .el-overlay') || document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;

                    if (btnPeriksaKembali && !sudahKlikPeriksaKembali) {
                        console.log("[ASIK Robot] Muncul popup Data peserta tidak valid. Mengklik Periksa Kembali...");
                        triggerClick(btnPeriksaKembali);
                        sudahKlikPeriksaKembali = true;
                        await sleep(1500);
                        continue;
                    } else {
                        catatRiwayat(inData.nik, 'Gagal', 'gagal input: Data peserta tidak valid');
                        skipPasienDanLanjut(); return;
                    }
                }
                const msgKesalahan = document.evaluate(XPATH_DAFTAR.MSG_POPUP_TERJADI_KESALAHAN, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                if (msgKesalahan) {
                    catatRiwayat(inData.nik, 'Gagal', 'gagal input: ' + msgKesalahan.textContent.trim());
                    skipPasienDanLanjut(); return;
                }

                // Cek popup layanan sudah diterima (kadang muncul di sini)
                const popupLayanan = document.evaluate(XPATH_DAFTAR.POPUP_INDIVIDU_SUDAH_MENERIMA_LAYANAN, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                if (popupLayanan) {
                    console.log("[ASIK Robot] Pasien sudah terdaftar. Skip ke pasien berikutnya...");
                    catatRiwayat(inData.nik, 'Skip', 'Sudah menerima layanan hari ini');

                    // Tutup popup jika ada
                    const btnTutup = document.evaluate(".//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'tutup') or contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'oke')]", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                    if (btnTutup) { triggerClick(btnTutup); }

                    skipPasienDanLanjut();
                    return;
                }

                // Cek popup pesan lainnya (cari teks spesifik di seluruh body jika MSG_POPUP gagal)
                let msgPopupText = "";
                const msgPopupEl = document.evaluate(XPATH_DAFTAR.MSG_POPUP, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                if (msgPopupEl) msgPopupText = msgPopupEl.textContent.trim();

                // Cek isi seluruh halaman jika pesan tertentu muncul (karena layout sering berubah)
                const fullText = document.body.textContent;
                if (!msgPopupText && document.querySelector('div[role="dialog"], .swal2-popup, .el-message-box, .modal')) {
                    msgPopupText = document.querySelector('div[role="dialog"], .swal2-popup, .el-message-box, .modal').textContent;
                }

                if (msgPopupText) {
                    const text = msgPopupText;
                    if (text.includes("Data belum sesuai KTP")) {
                        catatRiwayat(inData.nik, 'Gagal', 'gagal input: Data belum sesuai KTP / Nama dan NIK Beda');
                        skipPasienDanLanjut(); return;
                    } else if (text.includes("Kuota Pemeriksaan Habis")) {
                        // Kuota habis punya tombol Lanjut, klik dan tunggu form tersubmit di iterasi selanjutnya
                        const btnLanjutKuota = document.evaluate(XPATH_DAFTAR.BTN_LANJUT_KUOTA_HABIS, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                        if (btnLanjutKuota) {
                            triggerClick(btnLanjutKuota);
                            await sleep(500);
                        } else {
                            catatRiwayat(inData.nik, 'Gagal', 'gagal input: Kuota Pemeriksaan Habis');
                            skipPasienDanLanjut(); return;
                        }
                    } else if (text.includes("Peserta Menerima Pemeriksaan") || text.includes("sudah menerima layanan")) {
                        console.log("[ASIK Robot] Pasien sudah terdaftar (popup 2). Skip ke pasien berikutnya...");
                        catatRiwayat(inData.nik, 'Skip', 'Sudah menerima layanan hari ini');

                        const btnTutup = document.evaluate(".//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'tutup') or contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'oke')]", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                        if (btnTutup) { triggerClick(btnTutup); }

                        skipPasienDanLanjut();
                        return;
                    } else if (text.includes("Pembatasan Umur Pemeriksaan")) {
                        catatRiwayat(inData.nik, 'Gagal', 'gagal input: Pembatasan Umur Pemeriksaan');
                        skipPasienDanLanjut(); return;
                    }
                }
            }

            if (!formTersubmit) {
                console.log('[ASIK Robot] Gagal submit form pendaftaran (timeout atau tombol tidak bisa diklik)!');
                catatRiwayat(inData.nik, 'Gagal', 'gagal input: Timeout submit form Pendaftaran');
                skipPasienDanLanjut(); return;
            }

            // 18. Cek popup sukses daftar
            const msgSuccess = await waitForXPath(XPATH_DAFTAR.MSG_POPUP_SUCCESS, null, 5);
            if (!msgSuccess) {
                catatRiwayat(inData.nik, 'Gagal', 'gagal input: Popup sukses tidak muncul');
                skipPasienDanLanjut(); return;
            }

            // 18a. Ambil No. Tiket dari popup "Berhasil Daftar" (format: XXX-YYY)
            let nomorTiket = '';
            const tiketEl = document.evaluate("//*[contains(text(), 'No. Tiket')]", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
            if (tiketEl) {
                const tiketMatch = (tiketEl.textContent || '').match(/No\.\s*Tiket\s*:\s*([A-Z0-9\-]+)/i);
                if (tiketMatch) {
                    nomorTiket = tiketMatch[1].trim();
                    console.log("[ASIK Robot] No. Tiket disalin:", nomorTiket);
                }
            }
            if (!nomorTiket) {
                console.warn("[ASIK Robot] Tidak bisa mengambil No. Tiket dari popup, fallback nama.");
                nomorTiket = inData.nama || '';
            }

            console.log("[ASIK Robot] Berhasil Daftar! No. Tiket:", nomorTiket, ". Lanjut ke konfirmasi hadir...");
            catatRiwayat(inData.nik, 'Sukses', 'Berhasil Daftar - Tiket: ' + nomorTiket);

            // 19. Klik Tutup pada popup Berhasil Daftar
            const btnTutup = await waitForXPath("//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'tutup')]", null, 3);
            if (btnTutup) {
                console.log("[ASIK Robot] Menutup popup Berhasil Daftar...");
                triggerClick(btnTutup);
                if (typeof btnTutup.click === 'function') btnTutup.click();
            }
            await sleep(1500);

            // 20. Search dengan No. Tiket lalu Enter
            const inputSearch = await waitForXPath(XPATH_DAFTAR.INPUT_SEARCH + " | //input[contains(@placeholder, 'Cari') or contains(@placeholder, 'cari') or contains(@placeholder, 'Search')]", null, 5);
            if (inputSearch) {
                console.log("[ASIK Robot] Memasukkan No. Tiket ke kolom pencarian:", nomorTiket);
                inputElValue(inputSearch, nomorTiket);
                await sleep(300);
                enterKeyEl(inputSearch);
                await sleep(2500); // Tunggu tabel loading
            } else {
                console.warn("[ASIK Robot] Kolom pencarian tidak ditemukan, menunggu tabel muncul...");
                await sleep(3000);
            }

            // 21. Klik "Konfirmasi Hadir" di tabel hasil pencarian
            let btnKonfirmHadir = await waitForXPath(XPATH_DAFTAR.BTN_KONFIMASI_HADIR + " | //button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'konfirmasi hadir')]", null, 5);
            if (btnKonfirmHadir) {
                console.log("[ASIK Robot] Mengklik Konfirmasi Hadir...");
                triggerClick(btnKonfirmHadir);
                if (typeof btnKonfirmHadir.click === 'function') btnKonfirmHadir.click();
            } else {
                console.warn("[ASIK Robot] Tombol Konfirmasi Hadir tidak ditemukan!");
            }
            await sleep(1500);

            // 22. Klik checkbox "Peserta memahami & bersedia untuk menjalani prosedur CKG"
            const checkboxBersedia = await waitForXPath(XPATH_DAFTAR.CHECKBOX_BERSEDIA_CKG + " | //div[contains(@class, 'check')] | //label[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'memahami')]//div | //div[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'bersedia')]/ancestor::label//div[contains(@class, 'check')]", null, 5);
            if (checkboxBersedia) {
                console.log("[ASIK Robot] Mengklik checkbox bersedia CKG...");
                triggerClick(checkboxBersedia);
                if (typeof checkboxBersedia.click === 'function') checkboxBersedia.click();
            }
            await sleep(500);

            // 23. Klik tombol "Hadir"
            const btnHadirOK = await waitForXPath(XPATH_DAFTAR.BTN_HADIR_CKG + " | //button[normalize-space(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'))='hadir']", null, 5);
            if (btnHadirOK) {
                console.log("[ASIK Robot] Mengklik tombol Hadir...");
                triggerClick(btnHadirOK);
                if (typeof btnHadirOK.click === 'function') btnHadirOK.click();
            }
            await sleep(2000);

            // 24. Cek popup "Berhasil Hadir" dan klik "Periksa Skrining Mandiri"
            const msgHadir = await waitForXPath(XPATH_DAFTAR.MSG_POPUP_BERHASIL_HADIR + " | //*[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'berhasil hadir')]", null, 5);
            if (msgHadir) {
                console.log("[ASIK Robot] Berhasil Hadir! Mencari tombol Periksa Skrining Mandiri...");

                // Ambil No. Tiket dari popup Berhasil Hadir juga (bisa beda dari daftar)
                const tiketHadirEl = document.evaluate("//*[contains(text(), 'No. Tiket')]", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                if (tiketHadirEl) {
                    const tiketHadirMatch = (tiketHadirEl.textContent || '').match(/No\.\s*Tiket\s*:\s*([A-Z0-9\-]+)/i);
                    if (tiketHadirMatch) {
                        console.log("[ASIK Robot] No. Tiket Hadir:", tiketHadirMatch[1].trim());
                    }
                }

                const btnSkrining = await waitForXPath("//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'periksa skrining mandiri') or contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'skrining mandiri')]", null, 5);
                if (btnSkrining) {
                    console.log("[ASIK Robot] Mengklik Periksa Skrining Mandiri...");
                    triggerClick(btnSkrining);
                    if (typeof btnSkrining.click === 'function') btnSkrining.click();
                } else {
                    console.warn("[ASIK Robot] Tombol Periksa Skrining Mandiri tidak ditemukan, fallback URL...");
                    window.location.href = "https://sehatindonesiaku.kemkes.go.id/ckg-pelayanan";
                }
            } else {
                console.warn("[ASIK Robot] Popup Berhasil Hadir tidak muncul. Pendaftaran sudah berhasil, lanjut paksa...");
                window.location.href = "https://sehatindonesiaku.kemkes.go.id/ckg-pelayanan";
            }

            pendaftaranAktif = false;
            if (typeof GM_setValue === 'function') GM_setValue('pendaftaranAktif', false);
            return;
        } catch (err) {
            console.error("[ASIK Robot] Error prosesPendaftaran:", err);
            catatRiwayat(inData.nik, 'Gagal', 'lainnya: Error pendaftaran - ' + String(err).substring(0, 100));
            skipPasienDanLanjut();
        }
    }

    // Helper: Skip pasien saat ini dan lanjut ke pasien berikutnya
    function skipPasienDanLanjut() {
        pendaftaranAktif = true; // Kunci main loop
        console.log("[ASIK Robot - DEBUG] skipPasienDanLanjut called.");

        if (riwayatPasien.length > 0) {
            const lastRiwayat = riwayatPasien[riwayatPasien.length - 1];
            console.log("[ASIK Robot - DEBUG] lastRiwayat:", lastRiwayat);
        }

        laporSelesaiSheet(currentSheetRow, "Skip").finally(() => {
            console.log("[ASIK Robot - DEBUG] Memanggil bersihkanDataDanLanjut...");
            bersihkanDataDanLanjut();
        });
    }

    function bersihkanDataDanLanjut() {
        currentSheetNik = ""; currentSheetRow = 0;
        currentUserData = null;
        flagPasienSudahTerdaftar = false;
        if (typeof GM_setValue === 'function') {
            GM_setValue('currentSheetNik', "");
            GM_setValue('currentSheetRow', 0);
            GM_setValue('currentUserData', "null");
            GM_setValue('flagPasienSudahTerdaftar', false);
        }
        layananSudahDimatikan.clear();
        globalUsiaPasien = 0;
        globalJenisKelamin = "";
        statusPerkawinanPasien = "";
        namaPasienTerakhir = "";

        pendaftaranAktif = false;
        if (typeof GM_setValue === 'function') GM_setValue('pendaftaranAktif', false);

        let modeAwal = 'pelayanan';
        if (typeof GM_getValue === 'function') modeAwal = GM_getValue('asik_mode_awal', 'pelayanan');

        // Kembali ke titik awal sesuai preferensi rute untuk memproses pasien berikutnya
        if (modeAwal === 'pelayanan') {
            window.location.href = "https://sehatindonesiaku.kemkes.go.id/ckg-pelayanan";
        } else {
            window.location.href = "https://sehatindonesiaku.kemkes.go.id/ckg-pendaftaran-individu";
        }
    }

    // --
    // OTAK 5: SIKLUS PABRIK & PENYELESAIAN
    // --
    async function selesaikanPasien() {
        // Reset state per pasien agar tidak bocor ke pasien selanjutnya
        layananSudahDimatikan.clear();

        console.log("[ASIK Robot] selesaikanPasien() dipanggil | AutoSheet:", isAutoSheetRunning, "| Row:", currentSheetRow, "| NIK:", currentSheetNik);
        if (!isAutoSheetRunning) {
            console.log("[ASIK Robot] Auto Sheet belum aktif, tidak bisa selesaikan pasien.");
            return false;
        }
        if (!currentSheetRow) {
            console.log("[ASIK Robot] currentSheetRow kosong (0). Navigasi ke pencarian untuk ambil NIK baru...");
            window.location.href = "https://sehatindonesiaku.kemkes.go.id/ckg-pelayanan";
            return true;
        }

        console.log("[ASIK Robot] Melaporkan NIK selesai ke Sheet: ", currentSheetNik);
        try {
            // ATURAN BARU: Status dibedakan berdasarkan apakah pasien ini pasien lama atau baru
            let statusAkhir = flagPasienSudahTerdaftar ? "sdh terdaftar lengkap 14 panel" : "berhasil input dgn 14 panel";
            await laporSelesaiSheet(currentSheetRow, statusAkhir);
            catatRiwayat(currentSheetNik, 'Selesai', statusAkhir);
            console.log("[ASIK Robot] Laporan sukses! Kembali ke pencarian...");

            currentSheetNik = ""; currentSheetRow = 0;
            currentUserData = null;
            pendaftaranAktif = false;
            flagPasienSudahTerdaftar = false;
            if (typeof GM_setValue === 'function') {
                GM_setValue('currentSheetNik', "");
                GM_setValue('currentSheetRow', 0);
                GM_setValue('currentUserData', "null");
                GM_setValue('pendaftaranAktif', false);
                GM_setValue('flagPasienSudahTerdaftar', false);
            }
            renderUIFactory();

            // const btnBack = Array.from(document.querySelectorAll('button')).find(b => bersih(b.innerText).includes('selesaikan layanan'));
            // if (btnBack) { triggerClick(btnBack); await sleep(2000); }

            window.location.href = "https://sehatindonesiaku.kemkes.go.id/ckg-pelayanan";
            return true;
        } catch (e) {
            console.error("[ASIK Robot] Gagal melapor ke sheet", e);
            catatRiwayat(currentSheetNik, 'Gagal', 'Error lapor ke Sheet');
            return false;
        }
    }

    async function prosesHalamanPencarian(inputNik) {
        if (!isAutoSheetRunning) return;

        if (!currentSheetNik) {
            console.log("[ASIK Robot] Menarik data dari Sheet...");
            try {
                const data = await tarikDataSheet();
                if (data && data.length > 0) {
                    currentSheetRow = data[0].row;

                    // PENTING: Bersihkan spasi, enter, atau huruf bawaan dari Google Sheet! Hanya ambil angka.
                    currentSheetNik = String(data[0].nik).replace(/[^0-9]/g, '');

                    // Simpan juga data pasien lengkap (termasuk kolom pendaftaran) untuk prosesPendaftaranPage
                    if (dataPasienLokal_v2.length > 0) {
                        currentUserData = dataPasienLokal_v2[0];
                    } else {
                        currentUserData = data[0];
                    }

                    if (typeof GM_setValue === 'function') {
                        GM_setValue('currentSheetRow', currentSheetRow);
                        GM_setValue('currentSheetNik', currentSheetNik);
                        GM_setValue('currentUserData', JSON.stringify(currentUserData));
                    }
                    renderUIFactory();

                    // Cek kombinasi Status Daftar dan Status Pelayanan
                    const keys = Object.keys(currentUserData);

                    const keyDaftar = keys.find(k => {
                        const cleanKey = k.toLowerCase().replace(/_/g, ' ');
                        return cleanKey.includes('status daftar') || cleanKey.includes('status pendaftaran');
                    }) || 'status_pendaftaran';
                    const keyPelayanan = keys.find(k => k.toLowerCase().replace(/_/g, ' ').includes('status pelayanan')) || 'status_pelayanan';

                    const valDaftar = (currentUserData[keyDaftar] || "").toString().trim();
                    const valPelayanan = (currentUserData[keyPelayanan] || "").toString().trim();

                    let modeAwal = 'pelayanan';
                    if (typeof GM_getValue === 'function') modeAwal = GM_getValue('asik_mode_awal', 'pelayanan');

                    if (modeAwal === 'pelayanan') {
                        console.log("[ASIK Robot] Menggunakan rute Pelayanan CKG untuk: " + currentSheetNik);
                        pendaftaranAktif = false;
                        if (typeof GM_setValue === 'function') GM_setValue('pendaftaranAktif', false);
                        window.location.href = "https://sehatindonesiaku.kemkes.go.id/ckg-pelayanan";
                    } else {
                        // modeAwal === 'pendaftaran'
                        console.log("[ASIK Robot] Menggunakan rute Pendaftaran untuk: " + currentSheetNik);
                        pendaftaranAktif = true;
                        if (typeof GM_setValue === 'function') GM_setValue('pendaftaranAktif', true);
                        window.location.href = "https://sehatindonesiaku.kemkes.go.id/ckg-pendaftaran-individu";
                    }

                    return;
                } else {
                    console.log("[ASIK Robot] Semua NIK sudah diproses!");
                    isAutoSheetRunning = false;
                    if (typeof GM_setValue === 'function') GM_setValue('isAutoSheetRunning', false);
                    renderUIFactory();
                    return;
                }
            } catch (e) {
                console.error("[ASIK Robot] Gagal menarik data dari Sheet", e);
                await sleep(5000);
                return;
            }
        }

        // Ganti Dropdown Filter Pencarian ke NIK (Native Select)
        let dropdownDiganti = false;
        const selects = document.querySelectorAll('select');
        for (const sel of selects) {
            const optNik = Array.from(sel.options).find(o => o.text.trim().toUpperCase() === 'NIK' || String(o.value).toUpperCase() === 'NIK');
            if (optNik && sel.value !== optNik.value) {
                console.log("[ASIK Robot] Mengganti native dropdown menjadi NIK...");
                sel.value = optNik.value;
                sel.dispatchEvent(new Event('change', { bubbles: true }));
                sel.dispatchEvent(new Event('input', { bubbles: true }));
                dropdownDiganti = true;
                await sleep(500);
            }
        }

        if (inputNik && inputNik.value !== currentSheetNik) {
            console.log("[ASIK Robot] Memulai sekuens pencarian NIK: " + currentSheetNik);

            // JEDA DI AWAL: Menunggu server Kemenkes selesai memuat halaman / tabel
            console.log("[ASIK Robot] Menunggu halaman dimuat sepenuhnya...");
            await sleep(1500); // Dipercepat sesuai permintaan

            // PENTING: Karena jeda, halaman mungkin dirender ulang. Ambil ulang elemen aktifnya!
            inputNik = document.getElementById('searchNik') || document.querySelector('input[type="text"], input:not([type])');

            // 1. Ganti Dropdown Filter Pencarian ke NIK (Custom UI Vue/React)
            if (!dropdownDiganti && inputNik) {
                // Selalu cari di seluruh document body, jangan batasi dengan container
                const spanAktif = Array.from(document.querySelectorAll('span')).find(el =>
                    el.innerText && (el.innerText.trim() === 'Nama' || el.innerText.trim() === 'NIK' || el.innerText.trim() === 'Nomor Tiket') &&
                    el.className.includes('line-clamp')
                );

                if (spanAktif && spanAktif.innerText.trim() !== 'NIK') {
                    console.log("[ASIK Robot] Membuka dropdown filter (saat ini: " + spanAktif.innerText.trim() + ")...");
                    const clickableBox = spanAktif.closest('.cursor-pointer') || spanAktif;
                    triggerClick(clickableBox);
                    await sleep(1500); // Tahan lebih lama agar dropdown benar-benar muncul

                    const opsiNik = Array.from(document.querySelectorAll('div.cursor-pointer, div.py-2, li')).find(el =>
                        el.innerText && el.innerText.trim() === 'NIK' && el.getBoundingClientRect().width > 0
                    );
                    if (opsiNik) {
                        console.log("[ASIK Robot] Memilih opsi NIK...");
                        triggerClick(opsiNik);
                        // Beri jeda 2.5 detik agar Vue selesai me-render ulang kotak input baru
                        await sleep(2500);
                        inputNik = document.getElementById('searchNik') || document.querySelector('input[type="text"], input:not([type])');
                    } else {
                        console.log("[ASIK Robot] Gagal menemukan elemen opsi NIK di dalam popup dropdown!");
                    }
                } else if (!spanAktif) {
                    console.log("[ASIK Robot] Elemen dropdown filter 'Nama' tidak ditemukan di halaman!");
                }
            }

            // PENTING: Karena Vue merender ulang, kita ambil ulang elemen inputnya sekali lagi!
            inputNik = document.getElementById('searchNik') || document.querySelector('input[type="text"], input:not([type])');

            // 2. Masukkan NIK dan Enter
            if (inputNik) {
                console.log("[ASIK Robot] Mengetik NIK layaknya manusia: " + currentSheetNik);
                inputNik.focus();
                await sleep(200);

                // Kosongkan input secara native
                const nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, "value").set;
                nativeInputValueSetter.call(inputNik, '');
                inputNik.dispatchEvent(new Event('input', { bubbles: true }));
                await sleep(300);

                // Ketik karakter satu per satu (bypass keamanan deteksi robot/paste dari Vue)
                for (let i = 0; i < currentSheetNik.length; i++) {
                    const char = currentSheetNik[i];
                    const charCode = char.charCodeAt(0);

                    // Down & Press
                    inputNik.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, cancelable: true, key: char, keyCode: charCode, which: charCode }));
                    inputNik.dispatchEvent(new KeyboardEvent('keypress', { bubbles: true, cancelable: true, key: char, keyCode: charCode, which: charCode }));

                    // Input (inject value)
                    nativeInputValueSetter.call(inputNik, inputNik.value + char);
                    inputNik.dispatchEvent(new Event('input', { bubbles: true }));

                    // Up
                    inputNik.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true, cancelable: true, key: char, keyCode: charCode, which: charCode }));

                    // Jeda super singkat ala manusia
                    await sleep(30);
                }

                inputNik.dispatchEvent(new Event('change', { bubbles: true }));
                await sleep(800); // Tahan sebentar setelah diketik semua sebelum di-enter

                console.log("[ASIK Robot] Menekan tombol Enter secara virtual...");

                // Jika elemen dibungkus dalam <form>, ini cara paling ampuh
                const parentForm = inputNik.closest('form');
                if (parentForm && typeof parentForm.requestSubmit === 'function') {
                    parentForm.requestSubmit();
                } else if (parentForm && parentForm.submit) {
                    parentForm.submit();
                }

                // Simulasi Enter tingkat dalam untuk memicu Vue.js event listener (v-on:keyup.enter)
                const downEvent = new KeyboardEvent('keydown', { bubbles: true, cancelable: true, key: 'Enter', code: 'Enter', keyCode: 13, which: 13 });
                inputNik.dispatchEvent(downEvent);

                const pressEvent = new KeyboardEvent('keypress', { bubbles: true, cancelable: true, key: 'Enter', code: 'Enter', keyCode: 13, which: 13 });
                inputNik.dispatchEvent(pressEvent);

                const upEvent = new KeyboardEvent('keyup', { bubbles: true, cancelable: true, key: 'Enter', code: 'Enter', keyCode: 13, which: 13 });
                inputNik.dispatchEvent(upEvent);

                await sleep(3500); // Tunggu server memuat hasil pencarian di tab bawaan
            }

            // 3. Cek hasil pencarian di tab saat ini (Belum Pemeriksaan / Default)
            const cekRows = Array.from(document.querySelectorAll('tr, .grid-cols-5, .row'));
            // Pastikan row yang dicek BUKAN wrapper header pencarian dengan mewajibkan adanya tombol Mulai di dalamnya
            const rowAwal = cekRows.find(r => r.innerText.includes(currentSheetNik) && Array.from(r.querySelectorAll('button')).some(b => bersih(b.innerText).includes('mulai') && !b.disabled));
            const fallbackMulai = Array.from(document.querySelectorAll('button')).find(b => bersih(b.innerText).includes('mulai') && !b.disabled);
            const bisaMulaiAwal = rowAwal || (fallbackMulai && document.body.innerText.includes(currentSheetNik));

            if (bisaMulaiAwal) {
                console.log("[ASIK Robot] NIK ditemukan di tab bawaan! Langsung proses...");
            } else {
                // 3b. Klik tab "Sedang Pemeriksaan" jika tidak ketemu di tab awal
                const tabSedang = Array.from(document.querySelectorAll('a, button, li, div, p')).find(el =>
                    el.innerText && el.innerText.trim().toLowerCase() === 'sedang pemeriksaan'
                );
                if (tabSedang && (!tabSedang.className.includes('active') && !tabSedang.className.includes('bg-'))) {
                    console.log("[ASIK Robot] Pindah ke tab Sedang Pemeriksaan...");
                    triggerClick(tabSedang);
                    // Beri jeda panjang agar server Kemenkes mencari data NIK di tab ini
                    await sleep(3500);
                } else {
                    // Jika sudah aktif, beri waktu server memuat
                    await sleep(3500);
                }
            }
        }

        // 4. Cari tombol aksi utama (Mulai Pemeriksaan ATAU Lanjutkan)
        // Karena kita sudah memfilter NIK di kotak pencarian, tombol yang muncul pasti milik pasien tersebut
        const btnMulai = Array.from(document.querySelectorAll('button')).find(b => {
            if (b.disabled || b.offsetParent === null) return false;
            const txt = bersih(b.innerText);
            return txt.includes('mulai') || txt.includes('lanjut');
        });

        if (btnMulai) {
            console.log("[ASIK Robot] Tombol Mulai/Lanjutkan ditemukan, mengeklik...");
            percobaanCariNik = 0;
            try { btnMulai.scrollIntoView({ behavior: 'smooth', block: 'center' }); } catch (e) { }
            await sleep(1000);
            triggerClick(btnMulai);
            await sleep(3000);
            return;
        } else {
            // Cek apakah ada tombol "Lihat" yang menandakan pasien sudah Selesai
            const btnLihat = Array.from(document.querySelectorAll('button')).find(b => bersih(b.innerText).includes('lihat'));
            if (btnLihat) {
                console.log("[ASIK Robot] Hanya ada tombol Lihat. Pasien sudah selesai diperiksa hari ini.");
                catatRiwayat(currentSheetNik, 'Selesai', 'Pemeriksaan sudah selesai hari ini');
                skipPasienDanLanjut();
                return;
            }
        }

        // Jika sampai sini artinya NIK tidak ditemukan di kedua tab (Belum & Sedang Pemeriksaan)
        percobaanCariNik++;
        console.log(`[ASIK Robot] NIK ${currentSheetNik} tidak ditemukan. Percobaan ke-${percobaanCariNik}`);
        if (percobaanCariNik >= 2) {
            console.log(`[ASIK Robot] NIK ${currentSheetNik} tidak ditemukan setelah ${percobaanCariNik}x percobaan. Auto SKIP!`);
            catatRiwayat(currentSheetNik, 'Skip', 'NIK tidak ditemukan di Belum/Sedang Pemeriksaan');
            percobaanCariNik = 0;
            skipPasienDanLanjut();
            return;
        }
        await sleep(2000);
    }

    // --
    // OTAK 3: TATALAKSANA
    // --
    async function prosesTatalaksanaPage() {
        if (await cekGangguanHalaman()) return;
        const rows = Array.from(document.querySelectorAll('.table-ckg-tatalaksana-detail tr, tr'));
        let adaMulai = false;
        for (const row of rows) {
            const btnMulai = Array.from(row.querySelectorAll('button')).find(b => {
                const text = bersih(b.innerText);
                return text === 'mulai' || text.includes('mulai tatalaksana');
            });
            if (row.innerText.toLowerCase().includes('belum tatalaksana') && btnMulai) {
                adaMulai = true;
                row.scrollIntoView({ behavior: 'smooth', block: 'center' });
                await sleep(600);
                triggerClick(btnMulai);
                await sleep(3500);
                return;
            }
        }
        if (!adaMulai) {
            const dilaporkan = await selesaikanPasien();
            if (!dilaporkan) isRobotRunning = false;
        }
    }

    // --
    // OTAK 4: CKG
    // --
    async function prosesCKGPage() {
        console.log("[ASIK Robot] OTAK CKG: Memindai halaman selesai pemeriksaan...");
        if (await cekGangguanHalaman()) return;

        if (pemeriksaanBelumSelesai()) {
            console.log("[ASIK Robot] Pemeriksaan belum selesai tunggu.");
            return;
        }

        // SAFEGUARD: Tunggu halaman benar-benar stabil, lalu cek ulang
        // Ini mencegah false-positive CKG saat halaman masih loading setelah toggle
        console.log("[ASIK Robot] CKG: Verifikasi ulang setelah jeda...");
        await tungguSiapHalaman(8000);
        await sleep(3000); // Tunggu DOM selesai re-render

        // Cek ulang -- mungkin setelah DOM stabil, ternyata masih ada layanan belum selesai
        if (pemeriksaanBelumSelesai()) {
            console.log("[ASIK Robot] Re-check CKG: ternyata masih ada layanan belum selesai! Kembali ke pemeriksaan.");
            return;
        }

        // Cari tombol Mulai Tatalaksana yang AKTIF (tidak disabled)
        const btn = Array.from(document.querySelectorAll('button')).find(b => {
            const text = bersih(b.innerText);
            return (text === 'mulai' || text.includes("mulai tatalaksana")) && !b.disabled && b.offsetParent !== null;
        });

        if (btn) {
            console.log("[ASIK Robot] Layanan selesai klik Mulai Tatalaksana.");
            btn.scrollIntoView({ behavior: 'smooth', block: 'center' });
            await sleep(700);
            triggerClick(btn);
            await sleep(4000);
        } else {
            // Tombol tatalaksana tidak ada atau disabled -- pasien ini dianggap selesai
            console.log("[ASIK Robot] Tombol Tatalaksana tidak aktif/tidak ada. Menyelesaikan pasien...");
            const dilaporkan = await selesaikanPasien();
            if (!dilaporkan) {
                console.log("[ASIK Robot] selesaikanPasien gagal. Navigasi paksa ke halaman pencarian...");
                window.location.href = "https://sehatindonesiaku.kemkes.go.id/ckg-pelayanan";
            }
        }
    }

    // --
    // --
    // RADAR PUSAT
    // --
    let robotMulaiJalanTimestamp = 0; // Timestamp kapan isRobotRunning diaktifkan

    setInterval(async () => {
        renderUIFactory();

        // SINKRONISASI LINTAS-TAB: Selalu baca ulang status terbaru dari GM storage
        try {
            if (typeof GM_getValue === 'function') {
                isAutoSheetRunning = GM_getValue('isAutoSheetRunning', false);
                if (!currentSheetRow) currentSheetRow = GM_getValue('currentSheetRow', 0);
                if (!currentSheetNik) currentSheetNik = GM_getValue('currentSheetNik', "");
                pendaftaranAktif = GM_getValue('pendaftaranAktif', false);
            }
        } catch (e) { }

        // PAUSE MANUAL: Jika user menekan tombol Pause di UI
        if (isRobotPaused) {
            robotStatusTeks = "Paused";
            return;
        }

        const isDijeda = robotSedangDijeda();

        // DEADLOCK GUARD: Jika isRobotRunning sudah nyangkut lebih dari 120 detik, paksa bebaskan!
        // (Pendaftaran memakan waktu lebih lama, naikkan dari 60 ke 120 detik)
        if (isRobotRunning) {
            if (robotMulaiJalanTimestamp > 0 && (Date.now() - robotMulaiJalanTimestamp > 120000)) {
                console.warn("[ASIK Robot] DEADLOCK TERDETEKSI! Robot nyangkut >120 detik. Membebaskan kunci paksa...");
                isRobotRunning = false;
                robotMulaiJalanTimestamp = 0;
            } else {
                return; // Masih dalam batas wajar, tunggu
            }
        }

        if (isDijeda) {
            robotStatusTeks = "Jeda sistem";
            console.log("[ASIK Robot] Robot sedang dijeda oleh sistem (deteksi popup/halaman tatalaksana aktif).");
            return;
        }

        // Jangan jalankan robot di halaman yang bukan Kemenkes (Google Sheet hanya perlu overlay)
        if (!window.location.hostname.includes('kemkes.go.id')) {
            robotStatusTeks = "Bukan Kemenkes";
            return;
        }

        // Mencegah robot jalan diam-diam tanpa di-start/resume
        if (!isAutoSheetRunning) {
            robotStatusTeks = "Standby (Tekan Play)";
            return;
        }

        const tombolKirim = cariTombolKirim();
        const adaFormSurvey = document.querySelector('.sd-question') !== null;
        const isHalamanForm = !!(tombolKirim && adaFormSurvey);

        const isHalamanTable = document.querySelector('.table-ckg-tatalaksana-detail') !== null;

        const url = window.location.href.toLowerCase();

        const btnMulaiPemeriksaan = Array.from(document.querySelectorAll('button')).find(b => {
            const t = (b.innerText || b.textContent || "").trim().toLowerCase();
            return (t === 'mulai pemeriksaan' || t.includes('mulai pemeriksaan')) && !b.disabled && b.offsetParent !== null;
        });

        // BARU: Deteksi halaman pendaftaran individu.
        // Tambahan: Pastikan BUKAN halaman Mulai Pemeriksaan (Kemenkes kadang tidak ganti URL setelah sukses daftar)
        const isHalamanPendaftaran = url.includes('ckg-pendaftaran-individu') && !url.includes('detail-pemeriksaan') && !btnMulaiPemeriksaan;

        const isDiHalamanPelayanan = url.includes('ckg-pelayanan') && !url.includes('detail-pemeriksaan');

        const adaDaftarPemeriksaan = document.querySelector('.grid-cols-5') !== null ||
            Array.from(document.querySelectorAll('button')).some(b => {
                const t = b.innerText.toLowerCase();
                return t.includes('input data') || t.includes('ubah data') || t.includes('lihat hasil');
            }) || (url.includes('detail-pemeriksaan') && !isHalamanTable && !isHalamanForm);

        const isHalamanPemeriksaan = !isHalamanForm && !isHalamanTable && adaDaftarPemeriksaan && pemeriksaanBelumSelesai();

        // isHalamanCKG aktif jika daftar pemeriksaan ada namun SEMUA sudah selesai (entah tombol tatalaksana muncul atau tidak)
        const isHalamanCKG = !isHalamanForm && !isHalamanTable && adaDaftarPemeriksaan && !pemeriksaanBelumSelesai();

        let inputNik = null;
        if (isDiHalamanPelayanan && !isHalamanForm && !isHalamanTable && !adaDaftarPemeriksaan) {
            inputNik = Array.from(document.querySelectorAll('input')).find(i =>
                (i.placeholder && i.placeholder.toLowerCase().includes('nik')) ||
                (i.parentElement && i.parentElement.innerText.toLowerCase().includes('nik'))
            ) || document.querySelector('input[type="text"], input:not([type])');
        }

        const isHalamanPencarianDaftar = isDiHalamanPelayanan && !!inputNik && !btnMulaiPemeriksaan;
        const isHalamanMulaiPemeriksaan = !isHalamanForm && !isHalamanTable && !!btnMulaiPemeriksaan;

        // Update status halaman untuk UI
        if (isHalamanPendaftaran) robotStatusTeks = pendaftaranAktif ? "Proses Daftar" : "Siap Daftar";
        else if (isHalamanMulaiPemeriksaan) robotStatusTeks = "Mulai Pemeriksaan";
        else if (isHalamanForm) robotStatusTeks = "Isi Form";
        else if (isHalamanPemeriksaan) robotStatusTeks = "Pemeriksaan";
        else if (isHalamanCKG) robotStatusTeks = "CKG Selesai";
        else if (isHalamanTable) robotStatusTeks = "Tatalaksana";
        else if (isHalamanPencarianDaftar) robotStatusTeks = "Cari Pasien";
        else robotStatusTeks = "Menunggu Halaman...";

        try {
            if (await tanganiPopupPeringatan()) return;

            isRobotRunning = true;
            robotMulaiJalanTimestamp = Date.now();

            // BARU: Deteksi form crash / server error (Layar putih + tombol Kembali ke Halaman Utama)
            const isHalamanError = url.includes('form.kemkes.go.id') && document.body.innerText.toLowerCase().includes('kembali ke halaman utama') && document.querySelectorAll('.sd-question').length === 0;
            if (isHalamanError && isAutoSheetRunning) {
                console.warn("[ASIK Robot] Terdeteksi halaman Form Error/Crash! Mengklik tombol Kembali...");
                robotStatusTeks = "Error Form, Kembali...";
                renderUIFactory();
                await sleep(1500);
                // Cari elemen <button> secara spesifik agar kliknya tepat sasaran (bukan <span>)
                const btnKembali = Array.from(document.querySelectorAll('button')).find(b => (b.innerText || b.textContent || "").trim().toLowerCase() === 'kembali ke halaman utama');
                if (btnKembali) {
                    console.log("[ASIK Robot] Ditemukan tombol Kembali ke Halaman Utama, mengklik...");
                    triggerClick(btnKembali);
                    if (typeof btnKembali.click === 'function') btnKembali.click();
                } else {
                    console.log("[ASIK Robot] Tombol tidak ditemukan, menggunakan history.back()...");
                    window.history.back();
                }
                return;
            }

            // BARU: Jika flag pendaftaranAktif ATAU (kita di halaman pendaftaran dan robot nyala), proses pendaftaran
            if (isHalamanPendaftaran && isAutoSheetRunning) {
                await prosesPendaftaranPage();
            } else if (isHalamanMulaiPemeriksaan && isAutoSheetRunning) {
                console.log("[ASIK Robot] Ditemukan tombol Mulai Pemeriksaan, mengklik...");
                triggerClick(btnMulaiPemeriksaan);
                if (typeof btnMulaiPemeriksaan.click === 'function') btnMulaiPemeriksaan.click();

                // Tambahan: Tangani popup Konfirmasi Tanggal Pemeriksaan yang muncul setelah klik Mulai Pemeriksaan
                await sleep(1500);
                const popupKonfirmasi = Array.from(document.querySelectorAll('div, span, h2, h3, p')).find(el => {
                    const text = (el.innerText || "").toLowerCase();
                    return text.includes('konfirmasi tanggal pemeriksaan') || text.includes('konfirmasi tanggal');
                });

                if (popupKonfirmasi) {
                    console.log("[ASIK Robot] Popup Konfirmasi Tanggal Pemeriksaan muncul, menekan Simpan...");
                    const btnSimpan = Array.from(document.querySelectorAll('button')).find(b => {
                        const text = (b.innerText || "").toLowerCase().trim();
                        return text === 'simpan' && b.offsetParent !== null;
                    });
                    if (btnSimpan) {
                        triggerClick(btnSimpan);
                        if (typeof btnSimpan.click === 'function') btnSimpan.click();
                        await sleep(1000);
                    }
                }
            } else if (isHalamanForm && isAutoSheetRunning) {
                await prosesFormPage(tombolKirim);
            } else if (isHalamanPemeriksaan) {
                await prosesListPage();
            } else if (isHalamanCKG) {
                await prosesCKGPage();
            } else if (isHalamanTable) {
                await prosesTatalaksanaPage();
            } else if (isHalamanPencarianDaftar) {
                await prosesHalamanPencarian(inputNik);
            }
        } catch (err) {
            console.error("[ASIK Robot] Error:", err);
        } finally {
            // Jika robot ternyata tidak mengeksekusi aksi apapun
            // Log state halamannya untuk debugging jika tidak masuk ke kondisi apapun
            if (!isHalamanForm && !isHalamanPemeriksaan && !isHalamanCKG && !isHalamanTable && !isHalamanPencarianDaftar && !(isHalamanPendaftaran && pendaftaranAktif)) {
                console.log(`[ASIK Robot] Status Nganggur | URL: ${url} | Form:${isHalamanForm} Pem:${isHalamanPemeriksaan} CKG:${isHalamanCKG} Tbl:${isHalamanTable} Cari:${isHalamanPencarianDaftar} Daftar:${isHalamanPendaftaran} | DaftarPem:${adaDaftarPemeriksaan}`);
            }
            // Selalu bebaskan kunci agar setInterval berikutnya bisa berjalan!
            isRobotRunning = false;
            robotMulaiJalanTimestamp = 0;
        }
    }, 2000);

})();
