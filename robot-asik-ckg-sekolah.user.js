// ==UserScript==
// @name         Robot ASIK - Full Auto CKG Sekolah (Excel)
// @namespace    http://tampermonkey.net/
// @version      5.0
// @description  Automasi form pelayanan CKG Sekolah Sehat Indonesiaku dengan Input Excel (Update Layout 2026)
// @author       Faris / ASIK Automation
// @match        *://*/*
// @grant        GM_setValue
// @grant        GM_getValue
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// ==/UserScript==

(function () {
    'use strict';
    console.log('[ASIK Robot CKG Sekolah] === AKTIF (Versi 5.0 Full Auto Excel) ===');

    let loopTimerId = null; // Timer ID untuk membatalkan loop saat STOP

    // --
    // UTILITY FUNCTIONS
    // --
    const sleep = (ms) => new Promise(r => setTimeout(r, ms));

    // isElementVisible: Mendeteksi visibilitas elemen secara akurat, termasuk elemen dengan position: fixed (seperti modal/popup) dan tab di latar belakang
    const isElementVisible = (el) => {
        if (!el) return false;
        if (el.offsetParent !== null) return true;
        try {
            const style = window.getComputedStyle ? window.getComputedStyle(el) : null;
            if (style && (style.display === 'none' || style.visibility === 'hidden' || style.opacity === '0')) {
                return false;
            }
            const rect = el.getBoundingClientRect();
            if (rect.width > 0 || rect.height > 0) return true;
            if (style && (style.position === 'fixed' || style.position === 'sticky' || style.position === 'absolute')) {
                return true;
            }
        } catch (e) { }
        return false;
    };

    // triggerClick: Sesuai arsitektur robot-asik.user.js (MouseEvent + native click + Keyboard simulation)
    const triggerClick = (el) => {
        if (!el) return;

        const isSurveyJs = el.closest('.sv-root-modern, .sd-root-modern, .sd-question, .sv-question, .sd-dropdown');

        ['mouseover', 'pointerdown', 'mousedown', 'pointerup', 'mouseup'].forEach(ev => {
            try {
                el.dispatchEvent(new MouseEvent(ev, { bubbles: true, cancelable: true, view: document.defaultView, detail: 1 }));
            } catch (e) { }
        });

        if (isSurveyJs) {
            try {
                el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: document.defaultView, detail: 1 }));
            } catch (e) { }
        } else {
            if (typeof el.click === 'function') {
                try { el.click(); } catch (e) { }
            }
        }

        // Simulasi klik dengan keyboard (Enter & Space) untuk trigger event handler Vue/Nuxt (hindari input/textarea dan dropdown SurveyJS agar tidak tertutup kembali)
        if (!isSurveyJs && el.tagName !== 'INPUT' && el.tagName !== 'TEXTAREA') {
            try {
                el.focus();
                el.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', code: 'Enter', keyCode: 13, bubbles: true }));
                el.dispatchEvent(new KeyboardEvent('keyup', { key: 'Enter', code: 'Enter', keyCode: 13, bubbles: true }));
            } catch (e) { }
        }
    };

    const bersih = (str) =>
        str ? str.replace(/[\u00A0\u200B\u200C\u200D\uFEFF]+/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase() : "";

    const tungguElemen = (selector, root = document, timeout = 10000) =>
        new Promise(resolve => {
            const el = root.querySelector(selector);
            if (el) return resolve(el);
            const ob = new MutationObserver(() => {
                const found = root.querySelector(selector);
                if (found) { ob.disconnect(); resolve(found); }
            });
            ob.observe(root, { childList: true, subtree: true });
            setTimeout(() => { ob.disconnect(); resolve(null); }, timeout);
        });

    const cariElemenTeks = (selector, teks, root = document) => {
        const elements = Array.from(root.querySelectorAll(selector));
        return elements.find(el => bersih(el.textContent) === bersih(teks));
    };

    const cariInputSampingLabel = (labelTeks) => {
        const labels = Array.from(document.querySelectorAll('label, div, span'));
        const label = labels.find(el => bersih(el.textContent).includes(bersih(labelTeks)));
        if (label) {
            return label.querySelector('input') || (label.parentElement && label.parentElement.querySelector('input'));
        }
        return null;
    };

    const isiInput = (el, nilai, tekanEnter = false) => {
        if (!el || nilai === undefined || nilai === null) return;
        const strVal = String(nilai);
        el.focus();
        el.value = strVal;
        el.setAttribute('value', strVal);
        if (el._value !== undefined) el._value = strVal;

        const proto = el.tagName === 'TEXTAREA' ? window.HTMLTextAreaElement.prototype : window.HTMLInputElement.prototype;
        const setter = Object.getOwnPropertyDescriptor(proto, 'value')?.set;
        if (setter) setter.call(el, strVal);

        el.dispatchEvent(new Event('input', { bubbles: true, composed: true }));
        el.dispatchEvent(new InputEvent('input', { bubbles: true, composed: true, inputType: 'insertText', data: strVal }));
        ['keydown', 'keypress', 'change', 'keyup'].forEach(e => {
            try { el.dispatchEvent(new Event(e, { bubbles: true })); } catch (err) { }
        });

        // PENTING: Panggil blur native dan dispatch focusout agar SurveyJS memvalidasi dan mencatat nilai input
        try { el.blur(); } catch (err) { }
        try { el.dispatchEvent(new FocusEvent('blur', { bubbles: true })); } catch (err) { }
        try { el.dispatchEvent(new FocusEvent('focusout', { bubbles: true })); } catch (err) { }

        if (tekanEnter) {
            try {
                el.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', keyCode: 13, bubbles: true }));
                el.dispatchEvent(new KeyboardEvent('keyup', { key: 'Enter', keyCode: 13, bubbles: true }));
            } catch (err) { }
        }
    };

    // --
    // HELPER: KONFIRMASI POP-UP MODAL "DATA TIDAK DIPERIKSA?"
    // --
    const tanganiModalKonfirmasiTidakDiperiksa = async () => {
        for (let attempt = 0; attempt < 8; attempt++) {
            // Modal popup konfirmasi "Data Tidak Diperiksa?" selalu berupa dialog modal fixed
            const allModals = Array.from(document.querySelectorAll('div[role="dialog"], .modal, div.fixed[class*="z-"], .swal2-container'))
                .filter(el => !el.closest('#asik-ckg-panel, #asik-dashboard-overlay'));

            const modalContainer = allModals.find(el => {
                const t = bersih(el.textContent);
                return (t.includes('data tidak diperiksa') || t.includes('tidak diperiksa?')) &&
                    (t.includes('tidak periksa') || t.includes('yakin'));
            });

            // WAJIB: Jika modalContainer tidak ditemukan, JANGAN pernah mencari di luar modal!
            if (!modalContainer) {
                if (attempt < 7) await sleep(250);
                continue;
            }

            // Cari tombol "Tidak Periksa" khusus di dalam modalContainer saja
            const candidateButtons = Array.from(modalContainer.querySelectorAll('button, div.cursor-pointer, a, span'))
                .filter(b => {
                    if (!isElementVisible(b) || b.closest('#asik-ckg-panel, #asik-dashboard-overlay')) return false;
                    const t = bersih(b.textContent);
                    // PENTING: Hanya tombol konfirmasi "Tidak Periksa", JANGAN tombol "Batal" / "Kembali" / opsi form
                    if (t.includes('batal') || t.includes('kembali') || t.includes('ulang')) return false;
                    return t === 'tidak periksa' || t === '"tidak periksa"' || (t.includes('tidak periksa') && !t.includes('data'));
                });

            const btnKonfirmasi = candidateButtons.find(b => b.tagName === 'BUTTON') ||
                candidateButtons.find(b => b.closest('button')) ||
                candidateButtons[0];

            if (btnKonfirmasi) {
                const targetBtn = btnKonfirmasi.closest('button') || btnKonfirmasi;
                console.log("[CKG] ⚠️ Terdeteksi Popup Modal 'Data Tidak Diperiksa?'. Mengklik tombol konfirmasi:", targetBtn.textContent.trim());

                ['pointerdown', 'mousedown', 'pointerup', 'mouseup', 'click'].forEach(evt => {
                    try {
                        targetBtn.dispatchEvent(new MouseEvent(evt, { bubbles: true, cancelable: true, view: window, detail: 1 }));
                    } catch (e) { }
                });

                if (typeof targetBtn.click === 'function') {
                    try { targetBtn.click(); } catch (e) { }
                }

                const child = targetBtn.querySelector('div, span');
                if (child) {
                    try { child.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window })); } catch (e) { }
                }

                await sleep(1000);
                return true;
            }

            await sleep(250);
        }
        return false;
    };

    // --
    // HELPER: KONFIRMASI POP-UP MODAL "SELESAIKAN LAYANAN" ("Data Pemeriksaan akan Dikunci" -> Klik "Konfirmasi")
    // --
    const tanganiModalSelesaikanLayanan = async () => {
        for (let attempt = 0; attempt < 8; attempt++) {
            const allModals = Array.from(document.querySelectorAll('div[role="dialog"], .modal, div.fixed, div[class*="fixed"][class*="z-"], div[class*="shadow-standard"], div[class*="rounded-xl"], div[class*="swal2-container"], div[class*="popup"]'))
                .filter(el => !el.closest('#asik-ckg-panel, #asik-dashboard-overlay'));

            const modalContainer = allModals.find(el => {
                const t = bersih(el.textContent);
                return t.includes('data pemeriksaan akan dikunci') || t.includes('akan dikunci') || t.includes('tidak dapat menambahkan data') || t.includes('rapor kesehatan peserta') || t.includes('selesaikan layanan') || t.includes('menyelesaikan') || t.includes('apakah anda yakin');
            });

            if (modalContainer || allModals.length > 0) {
                const container = modalContainer || allModals[allModals.length - 1];
                const candidateButtons = Array.from(container.querySelectorAll('button, div[class*="cursor-pointer"], .btn-fill-primary, div.btn-fill-primary, div.btn-outline-error, a, span'))
                    .filter(b => {
                        if (!isElementVisible(b) || b.closest('#asik-ckg-panel, #asik-dashboard-overlay')) return false;
                        const t = bersih(b.textContent);
                        if (t === 'periksa kembali' || t.includes('periksa kembali') || t === 'batal' || t === 'kembali') return false;
                        return t === 'konfirmasi' || t.includes('konfirmasi') || t === 'ya' || t === 'ya, selesaikan' || t === 'selesaikan' || t === 'simpan';
                    });

                const btnKonfirmasi = candidateButtons.find(b => bersih(b.textContent) === 'konfirmasi') ||
                    candidateButtons.find(b => bersih(b.textContent).includes('konfirmasi')) ||
                    candidateButtons.find(b => b.tagName === 'BUTTON') ||
                    candidateButtons.find(b => b.closest('button')) ||
                    candidateButtons[0];

                if (btnKonfirmasi) {
                    const targetBtn = btnKonfirmasi.closest('button') || btnKonfirmasi;
                    console.log("[CKG] ⚠️ Terdeteksi Popup Modal 'Data Pemeriksaan akan Dikunci'. Mengklik tombol Konfirmasi:", targetBtn.textContent.trim());
                    targetBtn.scrollIntoView({ behavior: 'smooth', block: 'center' });
                    await sleep(300);
                    triggerClick(targetBtn);
                    if (typeof targetBtn.click === 'function') {
                        try { targetBtn.click(); } catch (e) { }
                    }
                    await sleep(1500);
                    return true;
                }
            }
            await sleep(400);
        }
        return false;
    };

    // --
    // HELPER: PENANGANAN POPUP "Data pemeriksaan sedang diproses / Silakan coba secara berkala"
    // --
    const tanganiModalSedangDiproses = async () => {
        const bodyText = bersih(document.body ? document.body.textContent || '' : '');
        const hasText = (bodyText.includes('sedang diproses') || bodyText.includes('coba secara berkala')) &&
            (bodyText.includes('data pemeriksaan') || bodyText.includes('secara berkala') || bodyText.includes('sedang diproses'));

        if (!hasText) return false;

        console.log("[CKG] ℹ️ Terdeteksi pesan 'Data pemeriksaan sedang diproses / Silakan coba secara berkala'. Mencari tombol Tutup...");

        // 1. Cari container modal terdekat / dialog card
        const allModals = Array.from(document.querySelectorAll('div[role="dialog"], .modal, div.fixed, div[class*="fixed"], div[class*="z-"], div[class*="shadow-"], div[class*="rounded-"], div[class*="swal2"], div[class*="popup"], div[class*="backdrop"], div[class*="dialog"], div.bg-white, section, main, body'))
            .filter(el => !el.closest('#asik-ckg-panel, #asik-dashboard-overlay') && isElementVisible(el));

        const matchedModals = allModals.filter(container => {
            const t = bersih(container.textContent || '');
            return (t.includes('sedang diproses') || t.includes('coba secara berkala'));
        });

        // Urutkan dari elemen paling kecil/spesifik (kartu dialog terdalam)
        matchedModals.sort((a, b) => (a.textContent || '').length - (b.textContent || '').length);

        const targetContainer = matchedModals[0] || document.body;

        // Cari tombol "Tutup" di dalam modal
        const candidateButtons = Array.from(targetContainer.querySelectorAll('button, div[role="button"], div[class*="cursor-pointer"], .btn-fill-primary, div.btn-fill-primary, div.btn-primary, a, span, input[type="button"]'))
            .filter(b => {
                if (!isElementVisible(b) || b.closest('#asik-ckg-panel, #asik-dashboard-overlay')) return false;
                const bt = bersih(b.textContent || b.value || b.getAttribute('aria-label') || '');
                return bt === 'tutup' || bt.startsWith('tutup') || bt === 'ok' || bt === 'close' || bt.includes('tutup');
            });

        let btnTutup = candidateButtons.find(b => b.tagName === 'BUTTON') ||
            candidateButtons.find(b => b.closest('button')) ||
            candidateButtons[0];

        // Fallback global jika tombol Tutup ada di luar targetContainer
        if (!btnTutup) {
            const globalButtons = Array.from(document.querySelectorAll('button, div[role="button"], div[class*="cursor-pointer"], .btn-fill-primary, a, span'))
                .filter(b => {
                    if (!isElementVisible(b) || b.closest('#asik-ckg-panel, #asik-dashboard-overlay')) return false;
                    const bt = bersih(b.textContent || b.value || '');
                    return bt === 'tutup';
                });
            btnTutup = globalButtons.find(b => b.tagName === 'BUTTON') ||
                globalButtons.find(b => b.closest('button')) ||
                globalButtons[0];
        }

        if (btnTutup) {
            const targetBtn = btnTutup.closest('button') || btnTutup;
            console.log("[CKG] ✅ Mengklik tombol Tutup pada popup 'Data pemeriksaan sedang diproses':", targetBtn.textContent.trim());
            targetBtn.scrollIntoView({ behavior: 'smooth', block: 'center' });
            await sleep(200);
            triggerClick(targetBtn);
            if (typeof targetBtn.click === 'function') {
                try { targetBtn.click(); } catch (e) { }
            }
            await sleep(1000);
            return true;
        }
        return false;
    };

    // --
    // HELPER: KONFIRMASI POP-UP MODAL "KIRIM RAPOR"
    // --
    const tanganiModalKirimRapor = async () => {
        for (let attempt = 0; attempt < 8; attempt++) {
            // Cek jika sudah muncul popup "Data pemeriksaan sedang diproses"
            const handledDiproses = await tanganiModalSedangDiproses();
            if (handledDiproses) return true;

            const allModals = Array.from(document.querySelectorAll('div[role="dialog"], .modal, div.fixed, div[class*="fixed"][class*="z-"], div[class*="shadow-standard"], div[class*="rounded-xl"], div[class*="swal2-container"], div[class*="popup"]'))
                .filter(el => !el.closest('#asik-ckg-panel, #asik-dashboard-overlay'));

            const modalContainer = allModals.find(el => {
                const t = bersih(el.textContent);
                return t.includes('kirim') || t.includes('rapor') || t.includes('rapot') || t.includes('pesan') || t.includes('yakin') || t.includes('notifikasi') || t.includes('sedang diproses') || t.includes('coba secara berkala');
            });

            if (modalContainer || allModals.length > 0) {
                const container = modalContainer || allModals[allModals.length - 1];
                const candidateButtons = Array.from(container.querySelectorAll('button, div[class*="cursor-pointer"], .btn-fill-primary, div.btn-fill-primary, a, span'))
                    .filter(b => {
                        if (!isElementVisible(b) || b.closest('#asik-ckg-panel, #asik-dashboard-overlay')) return false;
                        const t = bersih(b.textContent);
                        return t === 'kirim' || t === 'ya, kirim' || t === 'ya' || t === 'kirim rapor' || t === 'ok' || t === 'konfirmasi' || t === 'tutup' || t.includes('kirim') || t.includes('tutup');
                    });

                const btnKonfirmasi = candidateButtons.find(b => b.tagName === 'BUTTON') ||
                    candidateButtons.find(b => b.closest('button')) ||
                    candidateButtons[0];

                if (btnKonfirmasi) {
                    const targetBtn = btnKonfirmasi.closest('button') || btnKonfirmasi;
                    console.log("[CKG] 📄 Terdeteksi Popup Modal Kirim Rapor. Mengklik konfirmasi:", targetBtn.textContent.trim());
                    triggerClick(targetBtn);
                    if (typeof targetBtn.click === 'function') {
                        try { targetBtn.click(); } catch (e) { }
                    }
                    await sleep(1500);

                    // Polling tunggu popup kedua 'Data pemeriksaan sedang diproses'
                    for (let p = 0; p < 8; p++) {
                        const closed = await tanganiModalSedangDiproses();
                        if (closed) break;
                        await sleep(600);
                    }
                    return true;
                }
            }
            await sleep(400);
        }
        return false;
    };

    // --
    // OBSERVER OTOMATIS: POPUP "Data pemeriksaan sedang diproses" -> KLIK TUTUP
    // --
    let lastAutoCloseDiprosesTime = 0;
    const initSedangDiprosesObserver = () => {
        if (!document.body) {
            setTimeout(initSedangDiprosesObserver, 500);
            return;
        }
        const observer = new MutationObserver(() => {
            const now = Date.now();
            if (now - lastAutoCloseDiprosesTime < 1500) return;
            const bodyText = bersih(document.body ? document.body.textContent || '' : '');
            if ((bodyText.includes('sedang diproses') || bodyText.includes('coba secara berkala')) &&
                (bodyText.includes('data pemeriksaan') || bodyText.includes('secara berkala') || bodyText.includes('sedang diproses'))) {
                const buttons = Array.from(document.querySelectorAll('button, div[role="button"], div[class*="cursor-pointer"], .btn-fill-primary, a, span'))
                    .filter(b => {
                        if (!isElementVisible(b) || b.closest('#asik-ckg-panel, #asik-dashboard-overlay')) return false;
                        const bt = bersih(b.textContent || b.value || '');
                        return bt === 'tutup';
                    });
                const btn = buttons.find(b => b.tagName === 'BUTTON') || buttons[0];
                if (btn) {
                    lastAutoCloseDiprosesTime = now;
                    console.log("[CKG] ⚡ MutationObserver: Menutup otomatis popup 'Data pemeriksaan sedang diproses'!");
                    const targetBtn = btn.closest('button') || btn;
                    triggerClick(targetBtn);
                    if (typeof targetBtn.click === 'function') {
                        try { targetBtn.click(); } catch (e) { }
                    }
                }
            }
        });
        observer.observe(document.body, { childList: true, subtree: true });
    };
    initSedangDiprosesObserver();

    // --
    // HELPER: CARI TOMBOL AKSI / EDIT / LIHAT PADA BARIS LAYANAN (WALAU SUDAH TERISI)
    // --
    const cariTombolEditLayanan = (row) => {
        if (!row) return null;
        const allClickables = Array.from(row.querySelectorAll('button, div.cursor-pointer, a, div[role="button"], span.cursor-pointer, [class*="cursor-pointer"], svg, img')).filter(el => {
            if (!isElementVisible(el) || el.closest('#asik-ckg-panel, #asik-dashboard-overlay')) return false;
            const t = bersih(el.textContent || '');
            if (t.includes('tidak diperiksa') || t.includes('konfirmasi')) return false;
            return true;
        });

        // 1. Prioritas tombol dengan teks aksi ubah / edit / lihat / input / mulai
        const priority = allClickables.find(el => {
            const t = bersih(el.textContent || el.getAttribute('title') || el.getAttribute('aria-label') || '');
            return t === 'ubah data' || t === 'ubah' || t === 'edit' || t === 'lihat data' || t === 'lihat hasil' || t === 'lihat' || t === 'input data' || t === 'lanjutkan' || t === 'mulai' || t.includes('ubah') || t.includes('edit');
        });
        if (priority) return priority.closest('button') || priority;

        // 2. Cari button apa pun di baris
        const btn = allClickables.find(el => el.tagName === 'BUTTON');
        if (btn) return btn;

        // 3. Cari elemen dengan class cursor-pointer
        const cp = allClickables.find(el => (el.className || '').includes('cursor-pointer'));
        if (cp) return cp;

        return row.querySelector('button, a, div') || row;
    };

    // --
    // HELPER: DETEKSI KONDISI DATA KHUSUS SISWA DARI EXCEL (MATA / TELINGA / GIGI)
    // --
    const isPositiveIssue = (val) => {
        if (val === undefined || val === null) return false;
        const v = bersih(String(val));
        if (!v) return false;
        // Nilai eksplisit negatif / tidak ada / normal / kosong / N (termasuk N/N, 6/6) -> false (Normal)
        if (['n', 'n.', 'n/n', 'n / n', 'n,n', 'N', 'nn', '6/6', '6 / 6', 'tdk', 'tidak', 'tidak ada', 'tak ada', 'normal', '0', '-', '--', 'false', 'nihil', 'sehat', 'no', 't'].includes(v)) {
            return false;
        }
        if (/^[n\s\/\.,\-_]+$/i.test(v)) {
            return false;
        }
        // Selain N/n, kosong, dan normal -> dianggap ADA masalah (angka, silinder, plus, minus, dll.)
        return true;
    };

    const hasMataIssue = (student) => student ? isPositiveIssue(student.MATA) : false;
    const hasTelingaIssue = (student) => student ? isPositiveIssue(student.TELINGA) : false;
    const hasGigiIssue = (student) => student ? isPositiveIssue(student.GIGI) : false;

    const hasAnyIssue = (student) => {
        return hasMataIssue(student) || hasTelingaIssue(student) || hasGigiIssue(student);
    };

    // --
    // STATE MANAGEMENT LOKAL & GM STORAGE (Sinkronisasi Lintas Halaman / Tab)
    // --
    const getState = () => {
        try {
            if (typeof GM_getValue === 'function') {
                const gm = GM_getValue('asik_ckg_state', null);
                if (gm) return typeof gm === 'string' ? JSON.parse(gm) : gm;
            }
            const ls = localStorage.getItem('asik_ckg_state');
            if (ls) return JSON.parse(ls);
        } catch (e) { }
        return { data: [], index: 0, running: false, phase: 'IDLE', sekolahTarget: '', sweepMode: false, sweepSheetIndex: 0, sweepCurrentStudent: null, sweepEmptyCheckCount: 0 };
    };

    const setState = (stateObj) => {
        try {
            if (typeof GM_setValue === 'function') {
                GM_setValue('asik_ckg_state', JSON.stringify(stateObj));
            }
        } catch (e) { }
        try {
            localStorage.setItem('asik_ckg_state', JSON.stringify(stateObj));
        } catch (e) { }
        updateUI();
    };

    // --
    // HELPER: PENCOCOKAN NAMA FLEKSIBEL (TOLERAN PERBEDAAN SPASI, TITIK, & VARIASI KATA)
    // Contoh: "AHMAD ALZAM ALBASIT" vs "AHMAD ALZAM AL BASIT"
    // --
    function stripName(n) {
        if (!n) return '';
        return n.toLowerCase().replace(/[^a-z0-9]/gi, '');
    }

    function normalizeName(n) {
        if (!n) return '';
        return n.toLowerCase()
            .replace(/[^a-z0-9\s]/gi, ' ')
            .replace(/\s+/g, ' ')
            .trim();
    }

    function isNamaCocok(namaA, namaB) {
        if (!namaA || !namaB) return false;

        const normA = normalizeName(namaA);
        const normB = normalizeName(namaB);
        if (!normA || !normB) return false;

        // 1. Exact string match
        if (normA === normB) return true;

        // 2. Space-stripped match (e.g. "AHMAD ALZAM ALBASIT" vs "AHMAD ALZAM AL BASIT")
        const stripA = stripName(normA);
        const stripB = stripName(normB);
        if (stripA === stripB) return true;

        // 3. Substring without spaces (jika panjang cukup signifikan)
        if (stripA.length >= 6 && stripB.length >= 6) {
            const minLen = Math.min(stripA.length, stripB.length);
            const maxLen = Math.max(stripA.length, stripB.length);
            if ((stripA.includes(stripB) || stripB.includes(stripA)) && (minLen / maxLen >= 0.75)) {
                return true;
            }
        }

        // 4. Token / Word level matching
        const wordsA = normA.split(' ').filter(w => w.length >= 2);
        const wordsB = normB.split(' ').filter(w => w.length >= 2);

        const exactCommon = wordsA.filter(w => wordsB.includes(w));
        if (exactCommon.length >= 2) return true;
        if (wordsA.length === 1 && wordsB.length === 1 && wordsA[0] === wordsB[0] && wordsA[0].length >= 4) return true;

        // 5. Cross-token concatenated match (e.g. "al"+"basit" vs "albasit")
        if (wordsA.length >= 2 && wordsA.every(w => stripB.includes(w))) return true;
        if (wordsB.length >= 2 && wordsB.every(w => stripA.includes(w))) return true;

        return false;
    }

    function getSearchQuery(fullName, retryCount = 0) {
        if (!fullName) return '';
        const clean = bersih(fullName);
        const words = clean.split(/\s+/).filter(w => w.length > 0);

        // Jika retry >= 1, gunakan kata pertama yang panjang
        if (retryCount >= 1) {
            const longWord = words.find(w => w.length >= 4) || words[0];
            return longWord || clean;
        }

        // Jika kata pertama sangat pendek (misal "M.", "M", "AL"), ambil 2 kata pertama
        if (words.length >= 2 && words[0].replace(/[^a-z0-9]/gi, '').length <= 2) {
            return words[0] + ' ' + words[1];
        }

        // Jika ada 2 kata atau lebih, gunakan 2 kata pertama untuk menghindari perbedaan spasi pada nama belakang
        // Contoh: "AHMAD ALZAM AL BASIT" -> "AHMAD ALZAM" (menemukan "AHMAD ALZAM ALBASIT" di web)
        if (words.length >= 2) {
            return words[0] + ' ' + words[1];
        }

        return words[0] || clean;
    }

    function findStudentInWorkbook(rowNama, wbData, fallbackKelas) {
        const defaultObj = {
            NAMA: rowNama || 'Siswa',
            KELAS: fallbackKelas || '1',
            TB: '160',
            BB: '46',
            GDS: '90',
            HB: '12',
            SISTOLE: '110',
            DIASTOLE: '70',
            MATA: 'normal',
            TELINGA: 'normal',
            GIGI: 'normal'
        };

        if (!rowNama || !wbData || !wbData.sheets) {
            console.log(`[CKG Sweep] ℹ️ Siswa "${rowNama}" menggunakan data default (Excel belum dimuat).`);
            return defaultObj;
        }

        let allStudents = [];
        Object.keys(wbData.sheets).forEach(k => {
            const list = wbData.sheets[k] || [];
            allStudents = allStudents.concat(list);
        });

        // 1. Pencocokan cerdas toleran spasi (isNamaCocok)
        let match = allStudents.find(s => isNamaCocok(s.NAMA, rowNama));

        // 2. Jika belum cocok, cari khusus di sheet kelas yang bersangkutan
        if (!match && fallbackKelas) {
            const classStudents = wbData.sheets[fallbackKelas] || wbData.sheets[String(fallbackKelas)] || [];
            match = classStudents.find(s => isNamaCocok(s.NAMA, rowNama));
        }

        if (match) {
            const displayGds = match.GDS ? `${match.GDS} (dari Excel)` : '90 (default)';
            console.log(`[CKG Sweep] 🎯 Siswa "${rowNama}" DITEMUKAN di Excel! Menggunakan data Excel: "${match.NAMA}" (Kelas ${match.KELAS}, TB:${match.TB || '160'}, BB:${match.BB || '46'}, GDS:${displayGds}, HB:${match.HB || '12'}, Mata:${match.MATA || 'Normal'}, Telinga:${match.TELINGA || 'Normal'}, Gigi:${match.GIGI || 'Normal'})`);
            return {
                ...match,
                NAMA: rowNama // Tetap gunakan nama dari baris web agar sinkron
            };
        }

        console.log(`[CKG Sweep] ℹ️ Siswa "${rowNama}" TIDAK DITEMUKAN di Excel -> Menggunakan nilai default normal (TB: 160, BB: 46, GDS: 90, HB: 12, Normal).`);
        return defaultObj;
    }

    // --
    // KUNCI JAWABAN & ATURAN CKG SEKOLAH (HASIL ANALISIS FORM HTML)
    // --
    const pertanyaanKhususCKG = [
        // Imunisasi
        { q: "imunisasi polio tetes", a: "ya", fallback: "sudah" },
        { q: "imunisasi", a: "ya", fallback: "sudah" },
        // Faktor Risiko Hepatitis SMP dan SMA (Semua dijawab 'Tidak' sesuai file HTML acuan)
        { q: "tes untuk hepatitis b", a: "tidak" },
        { q: "ibu kandung/saudara sekandung yang menderita hepatitis", a: "tidak" },
        { q: "ibu kandung", a: "tidak" },
        { q: "berhubungan seksual berisiko", a: "tidak" },
        { q: "tanpa pengaman", a: "tidak" },
        { q: "menerima transfusi darah", a: "tidak" },
        { q: "transfusi darah", a: "tidak" },
        { q: "cuci darah atau hemodialisis", a: "tidak" },
        { q: "cuci darah", a: "tidak" },
        { q: "hemodialisis", a: "tidak" },
        { q: "menggunakan narkoba", a: "tidak" },
        { q: "obat terlarang", a: "tidak" },
        { q: "disuntik", a: "tidak" },
        { q: "orang dengan hiv", a: "tidak" },
        { q: "odhiv", a: "tidak" },
        { q: "pengobatan hepatitis c", a: "tidak" },
        { q: "hepatitis c", a: "tidak" },
        { q: "hepatitis b", a: "tidak" },
        { q: "hepatitis", a: "tidak" },
        // TB
        { q: "batuk yang tidak sembuh", a: "tidak batuk", fallback: "tidak" },
        { q: "penurunan berat badan", a: "tidak" },
        { q: "berat badan anak anda turun", a: "tidak" },
        { q: "demam hilang timbul", a: "tidak" },
        { q: "lesu atau malaise", a: "tidak" },
        // Kebugaran Jasmani Anak & Kelayakan Tes Kebugaran
        { q: "kebugaran jasmani", a: "baik", fallback: "cukup" },
        { q: "tulang dan sendi", a: "tidak" },
        { q: "arthritis", a: "tidak" },
        { q: "masalah pada jantung", a: "tidak" },
        { q: "asma", a: "tidak" },
        { q: "kehilangan kesadaran", a: "tidak" },
        { q: "pingsan", a: "tidak" },
        { q: "sakit kepala parah", a: "tidak" },
        { q: "kelayakan tes kebugaran", a: "tidak", fallback: "ya" },
        { q: "kelayakan kebugaran", a: "tidak", fallback: "ya" },
        { q: "kebugaran", a: "baik", fallback: "cukup" },
        // Tingkat Aktivitas Fisik
        { q: "aktif secara fisik", a: "1", type: "number" },
        { q: "aktif  secara  fisik", a: "1", type: "number" },
        { q: "aktivitas fisik", a: "1", fallback: "tidak" },
        // Perilaku Merokok
        { q: "merokok dalam setahun", a: "tidak" },
        { q: "merokok", a: "tidak", fallback: "tidak pernah" },
        { q: "terpapar asap rokok", a: "tidak" },
        { q: "asap rokok", a: "tidak", fallback: "tidak pernah" },
        // Kecemasan & Kesehatan Mental Remaja
        { q: "cemas", a: "tidak", fallback: "tidak sama sekali" },
        { q: "gugup", a: "tidak", fallback: "tidak sama sekali" },
        { q: "gelisah", a: "tidak", fallback: "tidak sama sekali" },
        { q: "merasa sedih atau tertekan", a: "tidak" },
        { q: "tidak tertarik lagi", a: "tidak" },
        { q: "sering capek, sulit tidur", a: "tidak" },
        // Kesehatan Reproduksi
        { q: "reproduksi", a: "tidak", fallback: "normal" },
        // Disabilitas & Demografi
        { q: "disabilitas", a: "tidak" },
        // Gizi
        { q: "berat badan", a: "bb_excel", type: "text" },
        { q: "tinggi badan", a: "tb_excel", type: "text" },
        // Tensi
        { q: "sistol", a: "sistole_excel", type: "text" },
        { q: "diastol", a: "diastole_excel", type: "text" },
        // Penyakit Tropis Terabaikan (Frambusia, Kusta, Skabies)
        { q: "frambusia", a: "tidak ada" },
        { q: "papul", a: "tidak ada" },
        { q: "nodul", a: "tidak ada" },
        { q: "ulkus", a: "tidak ada" },
        { q: "krusta", a: "tidak ada" },
        { q: "papiloma", a: "tidak ada" },
        { q: "kusta", a: "tidak ada", fallback: "tidak" },
        { q: "bercak kulit", a: "tidak ada", fallback: "tidak" },
        { q: "skabies", a: "tidak ada", fallback: "tidak" },
        { q: "koreng/ruam", a: "tidak ada", fallback: "tidak" },
        // Gigi
        { q: "gigi karies", a: "tidak ada", fallback: "0", type: "text" },
        { q: "jumlah gigi karies", a: "tidak ada", fallback: "0" },
        // Telinga & Mata
        { q: "telinga kanan", a: "tidak ada serumen impaksi", fallback: "normal" },
        { q: "telinga kiri", a: "tidak ada serumen impaksi", fallback: "normal" },
        { q: "gangguan pendengaran", a: "normal", fallback: "tidak" },
        { q: "serumen impaksi", a: "tidak ada serumen impaksi", fallback: "tidak ada" },
        { q: "infeksi", a: "tidak ada infeksi telinga", fallback: "tidak ada" },
        { q: "selaput mata merah", a: "tidak", fallback: "normal" },
        { q: "tajam penglihatan", a: "normal (visus 6/6 - 6/9)", fallback: "normal" },
        { q: "menggunakan kacamata", a: "tidak", fallback: "tidak" },
        { q: "daya dengar", a: "normal", fallback: "baik" },
        { q: "daya lihat", a: "normal", fallback: "baik" },
        // Gula Darah & Diabetes
        { q: "diabetes", a: "tidak" },
        { q: "kencing manis", a: "tidak" },
        { q: "sering merasa haus", a: "tidak" },
        { q: "sering merasa sangat lapar", a: "tidak" },
        // Nilai GDS & Hb diisi secara dinamis sesuai data tabel masing-masing siswa (valGDS / valHB)
        // Faktor Risiko Malaria (Semua dijawab 'Tidak' sesuai file HTML acuan)
        { q: "gejala seperti; demam", a: "tidak" },
        { q: "demam, sakit kepala, dan menggigil", a: "tidak" },
        { q: "sakit malaria dan obat tidak habis diminum", a: "tidak" },
        { q: "orang sakit malaria di wilayah tempat tinggal", a: "tidak" },
        { q: "daerah berisiko tinggi malaria", a: "tidak" },
        { q: "berisiko tinggi malaria", a: "tidak" },
        { q: "obat tidak habis diminum", a: "tidak" },
        { q: "riwayat kedatangan dari daerah berisiko", a: "tidak" },
        { q: "sakit malaria", a: "tidak" },
        { q: "malaria", a: "tidak" },
        // Tatalaksana & Rujukan
        { q: "alasan tidak diberikan", a: "pengobatan tuntas", fallback: "lainnya" },
        { q: "terdaftar di puskes", a: "lainnya" },
        { q: "faskes lain", a: "lainnya" },
        { q: "edukasi", a: "ya" },
        { q: "diberikan konseling", a: "ya" },
        { q: "konseling", a: "ya" },
        { q: "rujukan", a: "tidak", fallback: "tidak dirujuk" },
        { q: "dirujuk", a: "tidak", fallback: "tidak dirujuk" },
        { q: "tanggal pelaksanaan", a: "today", type: "date" }
    ];

    const prioritasOpsiAmanCKG = [
        "tidak ada",
        "tidak pernah",
        "tidak sama sekali",
        "tidak batuk",
        "tidak",
        "normal (visus 6/6 - 6/9)",
        "normal",
        "tidak ada serumen impaksi",
        "tidak ada infeksi telinga",
        "baik",
        "cukup",
        "ya",
        "sudah",
        "0"
    ];

    // SurveyJS Helper Functions
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

    function sudahDijawab(qNode) {
        const inputTerpilih = qNode.querySelector('input[type="radio"]:checked, input[type="checkbox"]:checked');
        if (inputTerpilih) return true;

        const dropVal = qNode.querySelector('.sd-dropdown__value, .sd-dropdown .sv-string-viewer');
        if (dropVal) {
            const val = bersih(dropVal.innerText || dropVal.textContent || "");
            if (val && val !== 'pilih' && !val.includes('select')) return true;
        }

        const dateInput = qNode.querySelector('input[type="date"]');
        if (dateInput && dateInput.value) return true;

        const textInput = qNode.querySelector('input[type="text"], input[type="number"], .sd-text input');
        if (textInput && textInput.value && String(textInput.value).trim()) return true;

        return false;
    }

    function isExactMatch(candidate, target) {
        if (!candidate || !target) return false;
        return bersih(candidate) === bersih(target);
    }

    function isSafePartialMatch(candidate, target) {
        if (!candidate || !target) return false;
        const c = bersih(candidate);
        const t = bersih(target);
        if (c === t) return true;

        // Cek polaritas: Jangan cocokkan teks positif dengan teks negatif
        const cHasNegative = /\b(tidak|tak|bukan|tanpa)\b/.test(c);
        const tHasNegative = /\b(tidak|tak|bukan|tanpa)\b/.test(t);
        if (cHasNegative !== tHasNegative) return false;

        return c.includes(t) || t.includes(c);
    }

    function radioSudahSesuai(qNode, targetText, fallbackText) {
        const checkedInput = qNode.querySelector('input[type="radio"]:checked, input[type="checkbox"]:checked');
        if (!checkedInput) return false;
        const label = checkedInput.closest('label') || checkedInput.parentElement;
        const curText = label?.textContent || "";
        const target = targetText;
        const fallback = fallbackText || null;

        if (isExactMatch(curText, target) || (fallback && isExactMatch(curText, fallback))) return true;
        if (isSafePartialMatch(curText, target) || (fallback && isSafePartialMatch(curText, fallback))) return true;
        return false;
    }

    function dropdownSudahSesuai(qNode, targetText, fallbackText) {
        const dropVal = qNode.querySelector('.sd-dropdown__value, .sd-dropdown .sv-string-viewer');
        if (!dropVal) return false;
        const curText = dropVal.innerText || dropVal.textContent || "";
        if (!curText || bersih(curText) === 'pilih' || bersih(curText).includes('select')) return false;
        const target = targetText;
        const fallback = fallbackText || null;

        if (isExactMatch(curText, target) || (fallback && isExactMatch(curText, fallback))) return true;
        if (isSafePartialMatch(curText, target) || (fallback && isSafePartialMatch(curText, fallback))) return true;
        return false;
    }

    function teksPertanyaan(qNode) {
        if (!qNode) return "";
        const titleNode = qNode.querySelector('.sd-question__title, .sd-title, [class*="title"]');
        const ariaTitle = titleNode?.getAttribute('aria-label') || qNode.getAttribute('aria-label') || "";
        const inner = titleNode ? (titleNode.innerText || titleNode.textContent || "") : "";
        return bersih(inner + " " + ariaTitle);
    }

    async function isiRadioSurvey(qNode, targetText, fallbackText) {
        const target = targetText;
        const fallback = fallbackText || null;
        const spans = Array.from(qNode.querySelectorAll('label .sv-string-viewer, .sd-item__control-label .sv-string-viewer, .sd-selectbase__item-text, label span, label'));
        if (spans.length === 0) return false;

        // 1. Exact match untuk target utama
        let pilihanSpan = spans.find(span => isExactMatch(span.innerText || span.textContent, target));

        // 2. Exact match untuk fallback
        if (!pilihanSpan && fallback) {
            pilihanSpan = spans.find(span => isExactMatch(span.innerText || span.textContent, fallback));
        }

        // 3. Safe partial match untuk target utama
        if (!pilihanSpan) {
            pilihanSpan = spans.find(span => isSafePartialMatch(span.innerText || span.textContent, target));
        }

        // 4. Safe partial match untuk fallback
        if (!pilihanSpan && fallback) {
            pilihanSpan = spans.find(span => isSafePartialMatch(span.innerText || span.textContent, fallback));
        }

        // 5. Prioritas opsi aman CKG (hanya jika target belum ditemukan)
        if (!pilihanSpan) {
            for (const kataAman of prioritasOpsiAmanCKG) {
                pilihanSpan = spans.find(span => isExactMatch(span.innerText || span.textContent, kataAman) || isSafePartialMatch(span.innerText || span.textContent, kataAman));
                if (pilihanSpan) break;
            }
        }

        if (pilihanSpan) {
            const labelEl = pilihanSpan.closest('label') || pilihanSpan;
            const inputEl = labelEl?.querySelector('input') || qNode.querySelector(`input[value="${pilihanSpan.textContent?.trim()}"]`);
            triggerClick(pilihanSpan);
            if (labelEl) triggerClick(labelEl);
            if (inputEl) {
                triggerClick(inputEl);
                inputEl.checked = true;
                inputEl.dispatchEvent(new Event('input', { bubbles: true }));
                inputEl.dispatchEvent(new Event('change', { bubbles: true }));
            }
            await sleep(300);
            return sudahDijawab(qNode);
        }
        return false;
    }

    async function isiDropdownSurvey(qNode, targetText, fallbackText) {
        if (!qNode) return false;
        if (dropdownSudahSesuai(qNode, targetText, fallbackText)) return true;

        const dropdown = qNode.querySelector('.sd-dropdown, .sv-dropdown_select-wrapper .sd-input, .sv-dropdown_select-wrapper, [role="combobox"]');
        if (!dropdown) return false;

        const filterInput = dropdown.querySelector('input.sd-dropdown__filter-string-input, input[role="combobox"], input');
        const listId = dropdown.getAttribute('aria-controls') || (filterInput && filterInput.getAttribute('aria-controls'));

        // 1. KLIK UNTUK MEMBUKA DROPDOWN
        triggerClick(dropdown);
        if (filterInput) {
            try {
                filterInput.focus();
                filterInput.dispatchEvent(new Event('focus', { bubbles: true }));
                filterInput.dispatchEvent(new Event('focusin', { bubbles: true }));
            } catch (e) { }
        }
        await sleep(350);

        // 2. CARI POPUP DROPDOWN (TIDAK BOLEH FILTER DENGAN offsetParent KARENA .sv-popup MEMILIKI position: fixed!)
        let activePopup = null;
        if (listId) {
            activePopup = document.getElementById(listId);
        }
        if (!activePopup) {
            const popups = Array.from(document.querySelectorAll('.sv-popup, .sd-popup, .sv-dropdown-popup, [role="listbox"]')).filter(p => {
                const s = p.style || {};
                return s.display !== 'none' && s.visibility !== 'hidden';
            });
            activePopup = popups[popups.length - 1];
        }

        // 3. CARI ELEMEN PILIHAN / ITEMS
        let items = [];
        if (activePopup) {
            items = Array.from(activePopup.querySelectorAll('.sv-list__item, .sd-list__item, li, [role="option"], .sd-dropdown-item'));
        }
        if (items.length === 0) {
            items = Array.from(document.querySelectorAll('.sv-list__item, .sd-list__item, [role="option"], .sv-popup li, .sd-popup li')).filter(el => {
                const p = el.closest('.sv-popup, .sd-popup, .sv-dropdown-popup');
                if (p && (p.style.display === 'none' || p.style.visibility === 'hidden')) return false;
                return true;
            });
        }

        // 4. PENCOCOKAN TARGET
        const target = targetText;
        const fallback = fallbackText || null;
        let pilihan = null;

        if (items.length > 0) {
            // 4a. Exact match target
            pilihan = items.find(li => isExactMatch(li.innerText || li.textContent, target));
            // 4b. Exact match fallback
            if (!pilihan && fallback) {
                pilihan = items.find(li => isExactMatch(li.innerText || li.textContent, fallback));
            }
            // 4c. Safe partial match target
            if (!pilihan) {
                pilihan = items.find(li => isSafePartialMatch(li.innerText || li.textContent, target));
            }
            // 4d. Safe partial match fallback
            if (!pilihan && fallback) {
                pilihan = items.find(li => isSafePartialMatch(li.innerText || li.textContent, fallback));
            }
            // 4e. Prioritas opsi aman CKG
            if (!pilihan) {
                for (const kataAman of prioritasOpsiAmanCKG) {
                    pilihan = items.find(li => isExactMatch(li.innerText || li.textContent, kataAman) || isSafePartialMatch(li.innerText || li.textContent, kataAman));
                    if (pilihan) break;
                }
            }
        }

        // 5. KLIK OPSI JIKA DITEMUKAN
        if (pilihan) {
            triggerClick(pilihan);
            ['pointerdown', 'mousedown', 'pointerup', 'mouseup', 'click'].forEach(evt => {
                try { pilihan.dispatchEvent(new MouseEvent(evt, { bubbles: true, cancelable: true, view: window, detail: 1 })); } catch (e) { }
            });
            await sleep(400);
            if (sudahDijawab(qNode)) return true;
        }

        // 6. JALUR ALTERNATIF BACKGROUND: SIMULASI KEYBOARD PADA FILTER INPUT (ArrowDown + Enter)
        if (filterInput && !sudahDijawab(qNode)) {
            console.log(`[CKG] ⌨️ Mencoba fallback keyboard background untuk dropdown: "${targetText}"`);
            try {
                filterInput.focus();
                isiInput(filterInput, targetText);
                await sleep(250);
                ['ArrowDown', 'Enter'].forEach(k => {
                    const code = k === 'Enter' ? 13 : 40;
                    filterInput.dispatchEvent(new KeyboardEvent('keydown', { key: k, code: k, keyCode: code, bubbles: true }));
                    filterInput.dispatchEvent(new KeyboardEvent('keyup', { key: k, code: k, keyCode: code, bubbles: true }));
                });
                await sleep(400);
                if (sudahDijawab(qNode)) return true;
            } catch (e) { }
        }

        // 7. JALUR ALTERNATIF TERAKHIR: INJEKSI MODEL SURVEYJS / KNOCKOUT LANGSUNG (ANTI-STUCK BACKGROUND)
        try {
            const ko = window.ko || (typeof unsafeWindow !== 'undefined' ? unsafeWindow.ko : null);
            if (ko && ko.dataFor) {
                const context = ko.dataFor(dropdown);
                if (context && context.question) {
                    const qModel = context.question;
                    const choices = qModel.visibleChoices || qModel.choices || [];
                    let matchedChoice = choices.find(c => isExactMatch(c.text || c.value, target));
                    if (!matchedChoice && fallback) matchedChoice = choices.find(c => isExactMatch(c.text || c.value, fallback));
                    if (!matchedChoice) matchedChoice = choices.find(c => isSafePartialMatch(c.text || c.value, target));
                    if (!matchedChoice && choices.length > 0) matchedChoice = choices[0];

                    if (matchedChoice) {
                        console.log(`[CKG] ⚡ Injeksi model SurveyJS langsung (Background safe): "${matchedChoice.value}"`);
                        qModel.value = matchedChoice.value;
                        await sleep(300);
                        return true;
                    }
                }
            }
        } catch (e) { }

        // Tutup kembali dropdown jika gagal agar tidak menghalangi elemen lain
        if (!sudahDijawab(qNode)) {
            triggerClick(dropdown);
        }
        return sudahDijawab(qNode);
    }

    async function isiTextFieldSurvey(qNode, value) {
        const input = qNode.querySelector('input.sd-input, input[type="number"], input[type="text"], input, textarea');
        if (!input) return false;
        isiInput(input, value);

        // Direct model update ke SurveyJS question jika tersedia
        const qModel = qNode.question || input.question ||
            qNode.__vueParentComponent?.ctx?.question ||
            qNode.__vueParentComponent?.props?.question;
        if (qModel) {
            if (typeof qModel.setValue === 'function') {
                try { qModel.setValue(value); } catch (e) { }
            }
            if ('value' in qModel) {
                try { qModel.value = value; } catch (e) { }
            }
        }

        // Defocus input & trigger klik area kosong otomatis agar SurveyJS memvalidasi pertanyaan terakhir
        try { input.blur(); } catch (e) { }
        if (document.activeElement === input) {
            try { document.activeElement.blur(); } catch (e) { }
        }
        try {
            document.body.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
        } catch (e) { }

        await sleep(300);
        return !!(input.value && String(input.value).trim());
    }

    // --
    // FUNGSI UTAMA PENGISIAN FORM SURVEYJS (MULTI-PASS UNTUK PERTANYAAN BERCABANG)
    // --
    async function isiFormPelayanan(dataPasien) {
        // 0. Sinkronisasi Data Siswa Aktif: Pastikan mengambil data siswa yang benar dari Excel
        let wbData = null;
        try {
            const raw = GM_getValue('asik_ckg_workbook_data') || localStorage.getItem('asik_ckg_workbook_data');
            if (raw) wbData = typeof raw === 'string' ? JSON.parse(raw) : raw;
        } catch (e) { }

        // Coba deteksi nama siswa dari halaman aktif jika ada (misal di Detail / header ASIK)
        let pageStudentName = '';
        const nameEl = document.querySelector('.text-\\[18px\\].font-bold, div[class*="text-[18px]"].font-bold');
        if (nameEl) pageStudentName = nameEl.textContent.trim();
        if (!pageStudentName) {
            try { pageStudentName = sessionStorage.getItem('asik_ckg_active_student_name') || ''; } catch (e) { }
        }

        const candidateName = (dataPasien && dataPasien.NAMA && dataPasien.NAMA !== 'Siswa' && dataPasien.NAMA !== 'Siswa Sedang Pemeriksaan')
            ? dataPasien.NAMA
            : pageStudentName;

        if (candidateName && wbData && wbData.sheets) {
            const matched = findStudentInWorkbook(candidateName, wbData);
            if (matched && matched.NAMA !== 'Siswa') {
                dataPasien = { ...dataPasien, ...matched };
                console.log(`[CKG] 🎯 Sinkronisasi Siswa: "${matched.NAMA}" terhubung dengan data Excel! (GDS: ${matched.GDS || '(kosong)'}, HB: ${matched.HB || '(kosong)'})`);
            }
        }

        console.log('[CKG] Memeriksa & mengisi pertanyaan form untuk:', dataPasien.NAMA);

        // Parsing Tinggi Badan (TB) & Berat Badan (BB) dengan pembulatan Math.round agar angka desimal (seperti 159.9) dibulatkan menjadi 160
        function parseTB(raw) {
            if (!raw) return '160';
            const s = String(raw).trim().replace(',', '.');
            if (s === '-' || s === '--' || s === '0') return '160';
            let num = parseFloat(s);
            if (isNaN(num) || num <= 0) return '160';
            // Jika ada nilai 4 digit > 300 (misal 1599 karena titik hilang atau mm), koreksi bagi 10
            if (num > 300 && num < 3000) num = num / 10;
            return String(Math.round(num));
        }

        function parseBB(raw) {
            if (!raw) return '46';
            const s = String(raw).trim().replace(',', '.');
            if (s === '-' || s === '--' || s === '0') return '46';
            let num = parseFloat(s);
            if (isNaN(num) || num <= 0) return '46';
            // Jika ada nilai > 250 (misal 458 karena 45.8 titiknya hilang), koreksi bagi 10
            if (num > 250 && num < 2500) num = num / 10;
            return String(Math.round(num));
        }

        const valTB = parseTB(dataPasien.TB);
        const valBB = parseBB(dataPasien.BB);

        // Gula Darah Sewaktu (GDS) dari tabel: utamakan nilai dari tabel, jika kosong/tidak ada gunakan default 90
        let rawGdsStr = String(
            (dataPasien.GDS !== undefined && dataPasien.GDS !== null && String(dataPasien.GDS).trim() !== '') ? dataPasien.GDS :
                ((dataPasien.gds !== undefined && dataPasien.gds !== null && String(dataPasien.gds).trim() !== '') ? dataPasien.gds :
                    (dataPasien.GULA || dataPasien.GLUKOSA || dataPasien.GULA_DARAH || ''))
        ).trim();

        let cleanGDS = '';
        if (rawGdsStr && rawGdsStr !== '-' && rawGdsStr !== '--' && rawGdsStr !== '0') {
            if (rawGdsStr.includes('.') || rawGdsStr.includes(',')) {
                const num = parseFloat(rawGdsStr.replace(',', '.'));
                if (!isNaN(num) && num > 0) {
                    cleanGDS = String(Math.round(num));
                }
            }
            if (!cleanGDS) {
                const digits = rawGdsStr.replace(/[^0-9]/g, '');
                if (digits && digits !== '0') cleanGDS = digits;
            }
        }

        const valGDS = cleanGDS || '90';
        if (cleanGDS) {
            console.log(`[CKG] 🩸 GDS Siswa "${dataPasien.NAMA}": MENGGUNAKAN NILAI DARI TABEL EXCEL = "${cleanGDS}" (Bukan default 90)`);
        } else {
            console.log(`[CKG] 🩸 GDS Siswa "${dataPasien.NAMA}": Di tabel kosong/tidak ada -> Menggunakan default = "90"`);
        }

        // Pemeriksaan Hemoglobin / Anemia Anak dari tabel, jika kosong atau '-' default 12
        const rawHB = (dataPasien.HB || dataPasien.HEMOGLOBIN || dataPasien.ANEMIA)
            ? String(dataPasien.HB || dataPasien.HEMOGLOBIN || dataPasien.ANEMIA).trim()
            : '';
        const cleanHB = (rawHB === '-' || rawHB === '0' || rawHB === '--') ? '' : rawHB.replace(/[^0-9.,]/g, '').replace(',', '.').trim();
        const valHB = (cleanHB && cleanHB !== '0') ? cleanHB : '12';

        // Parsing Tekanan Darah (Sistol & Diastol) dari tabel (misal: "105/79" atau kolom terpisah)
        let rawSistole = dataPasien.SISTOLE ? String(dataPasien.SISTOLE).trim() : '';
        let rawDiastole = dataPasien.DIASTOLE ? String(dataPasien.DIASTOLE).trim() : '';

        if (rawSistole.includes('/')) {
            const parts = rawSistole.split('/');
            rawSistole = parts[0] || '';
            if (!rawDiastole && parts[1]) {
                rawDiastole = parts[1] || '';
            }
        } else if (rawDiastole.includes('/')) {
            const parts = rawDiastole.split('/');
            if (!rawSistole && parts[0]) rawSistole = parts[0] || '';
            rawDiastole = parts[1] || '';
        }

        const cleanSistole = rawSistole.replace(/[^0-9]/g, '');
        const cleanDiastole = rawDiastole.replace(/[^0-9]/g, '');

        // Gunakan nilai dari tabel, jika kosong gunakan default (110 / 70)
        const valSistole = cleanSistole || '110';
        const valDiastole = cleanDiastole || '70';

        const hasGigi = hasGigiIssue(dataPasien);
        const hasTelinga = hasTelingaIssue(dataPasien);
        const hasMata = hasMataIssue(dataPasien);

        // Multi-pass filling: Ulangi pengisian sampai seluruh pertanyaan bercabang (visibleIf) yang baru muncul terisi
        let pass = 0;
        const maxPasses = 4;

        while (pass < maxPasses) {
            pass++;
            const visibleQuestions = Array.from(document.querySelectorAll('.sd-question, .sv-question')).filter(isElementVisible);
            if (visibleQuestions.length === 0) break;

            let filledSomething = false;

            for (const qNode of visibleQuestions) {
                const teksTanya = teksPertanyaan(qNode);
                const jenis = jenisInputPertanyaan(qNode);

                // Pengecualian Khusus: GDS 2 (Gula Darah Sewaktu Kedua) -> TIDAK USAH DIISI
                if (teksTanya.includes("gds 2") || teksTanya.includes("gds2") || teksTanya.includes("gula darah sewaktu kedua") || (teksTanya.includes("gds") && teksTanya.includes("kedua")) || teksTanya.includes("gds 1 prediabetes")) {
                    console.log(`[CKG] ⏭️ Melewati pertanyaan GDS 2 (sesuai instruksi tidak diisi)`);
                    continue;
                }

                // 1. Prioritas Field Angka/Teks: Sistol, Diastol, TB, BB, GDS, HB
                if (jenis === 'text' || jenis === 'number') {
                    if (teksTanya.includes("sistol")) {
                        const curVal = qNode.querySelector('input')?.value;
                        if (!curVal || curVal !== valSistole) {
                            console.log(`[CKG] Mengisi Tekanan Darah Sistol: ${valSistole}`);
                            await isiTextFieldSurvey(qNode, valSistole);
                            filledSomething = true;
                        }
                        continue;
                    }
                    if (teksTanya.includes("diastol")) {
                        const curVal = qNode.querySelector('input')?.value;
                        if (!curVal || curVal !== valDiastole) {
                            console.log(`[CKG] Mengisi Tekanan Darah Diastol: ${valDiastole}`);
                            await isiTextFieldSurvey(qNode, valDiastole);
                            filledSomething = true;
                        }
                        continue;
                    }
                    if (teksTanya.includes("tinggi badan") || teksTanya.includes(" tb") || teksTanya === "tb") {
                        const curVal = qNode.querySelector('input')?.value;
                        if (!curVal || curVal !== valTB) {
                            console.log(`[CKG] Mengisi Tinggi Badan: ${valTB}`);
                            await isiTextFieldSurvey(qNode, valTB);
                            filledSomething = true;
                        }
                        continue;
                    }
                    if (teksTanya.includes("berat badan") || teksTanya.includes(" bb") || teksTanya === "bb") {
                        const curVal = qNode.querySelector('input')?.value;
                        if (!curVal || curVal !== valBB) {
                            console.log(`[CKG] Mengisi Berat Badan: ${valBB}`);
                            await isiTextFieldSurvey(qNode, valBB);
                            filledSomething = true;
                        }
                        continue;
                    }
                    if ((teksTanya.includes("gds") || teksTanya.includes("gula darah") || teksTanya.includes("glukosa")) &&
                        !teksTanya.includes("gds 2") && !teksTanya.includes("gds2") && !teksTanya.includes("kedua")) {
                        const curVal = qNode.querySelector('input')?.value;
                        if (!curVal || curVal !== valGDS) {
                            console.log(`[CKG] 🩸 Mengisi GDS (Gula Darah): ${valGDS} ${cleanGDS ? '(dari Tabel Excel)' : '(default 90)'}`);
                            await isiTextFieldSurvey(qNode, valGDS);
                            filledSomething = true;
                        }
                        continue;
                    }
                    if (teksTanya.includes("hemoglobin") || teksTanya.includes("anemia") || teksTanya.includes("kadar hb") || teksTanya.includes("hasil hb") || teksTanya.includes(" hb") || teksTanya === "hb") {
                        const curVal = qNode.querySelector('input')?.value;
                        if (!curVal || curVal !== valHB) {
                            console.log(`[CKG] 🩸 Mengisi Kadar Hemoglobin (Hb) Anemia: ${valHB}`);
                            await isiTextFieldSurvey(qNode, valHB);
                            filledSomething = true;
                        }
                        continue;
                    }
                }

                // 2. Pemeriksaan Gigi (Karies) - Dukungan Dinamis Masalah Gigi
                // Pertanyaan: "1. Berapa jumlah gigi karies? *"
                // Opsi: "Tidak ada", "1", "2", "3", ">3"
                // Aturan: Jika di data ADA pada kolom GIGI -> pilih "2"
                if (teksTanya.includes("karies") || teksTanya.includes("gigi")) {
                    const rawGigi = String(dataPasien.GIGI || '').trim();
                    let targetGigi = "tidak ada";
                    let fallbackGigi = "0";

                    if (hasGigi) {
                        const rawLower = rawGigi.toLowerCase();
                        if (rawGigi === '>3' || rawLower === '> 3') {
                            targetGigi = '>3';
                            fallbackGigi = '3';
                        } else if (/^[1-9]\d*$/.test(rawGigi)) {
                            targetGigi = rawGigi;
                            fallbackGigi = rawGigi;
                        } else {
                            // Sesuai instruksi: Jika data ADA pada kolom GIGI -> jawab "2"
                            targetGigi = '2';
                            fallbackGigi = '2';
                        }
                    }

                    if (jenis === 'text' || jenis === 'number') {
                        const curVal = qNode.querySelector('input')?.value;
                        const valTarget = hasGigi ? targetGigi : "0";
                        if (!curVal || curVal !== valTarget) {
                            console.log(`[CKG] 🦷 Mengisi Gigi Karies: ${valTarget}`);
                            await isiTextFieldSurvey(qNode, valTarget);
                            filledSomething = true;
                        }
                    } else {
                        if (!radioSudahSesuai(qNode, targetGigi, fallbackGigi)) {
                            console.log(`[CKG] 🦷 Memilih Gigi Karies: ${targetGigi}`);
                            await isiRadioSurvey(qNode, targetGigi, fallbackGigi);
                            filledSomething = true;
                        }
                    }
                    continue;
                }

                // 3. Penentuan Jawaban Target Berdasarkan Kondisi Khusus (Telinga / Mata / Lainnya)
                let target = "normal";
                let fallback = "tidak";

                // A. Kondisi Telinga (Skrining Telinga dan Mata - Anak Sekolah)
                // Pertanyaan 1 & 2: Gangguan pendengaran telinga kanan/kiri -> SELALU "Normal"
                // Pertanyaan 3 & 4: Serumen impaksi telinga kanan/kiri -> "Ada serumen impaksi" jika ADA di kolom TELINGA
                // Pertanyaan 5 & 6: Infeksi telinga kanan/kiri -> SELALU "Tidak ada infeksi telinga"
                if (teksTanya.includes("serumen") || (teksTanya.includes("telinga") && (teksTanya.includes("kanan") || teksTanya.includes("kiri")) && !teksTanya.includes("infeksi") && !teksTanya.includes("pendengaran") && !teksTanya.includes("gangguan"))) {
                    target = hasTelinga ? "ada serumen impaksi" : "tidak ada serumen impaksi";
                    fallback = hasTelinga ? "ada" : "tidak ada";
                } else if (teksTanya.includes("gangguan pendengaran") || teksTanya.includes("pendengaran") || teksTanya.includes("daya dengar")) {
                    target = "normal";
                    fallback = "normal";
                } else if (teksTanya.includes("infeksi") && (teksTanya.includes("telinga") || teksTanya.includes("kanan") || teksTanya.includes("kiri"))) {
                    target = "tidak ada infeksi telinga";
                    fallback = "tidak ada";
                }
                // B. Kondisi Mata (Skrining Telinga dan Mata / Pemeriksaan Mata)
                // Pertanyaan 7 & 8: Selaput mata merah / kornea keruh / kelopak benjolan / juling / pupil putih -> SELALU "Normal"
                // Pertanyaan 9 & 10: Hasil skrining tajam penglihatan mata kanan & kiri -> "Ada indikasi gangguan penglihatan (visus <6/9)" jika ADA di kolom MATA
                else if (teksTanya.includes("selaput mata") || teksTanya.includes("kornea") || teksTanya.includes("juling") || teksTanya.includes("kelopak") || teksTanya.includes("pupil")) {
                    target = "normal";
                    fallback = "normal";
                } else if (teksTanya.includes("tajam penglihatan") || teksTanya.includes("visus") ||
                    ((teksTanya.includes("penglihatan") || teksTanya.includes("daya lihat") || (teksTanya.includes("mata") && (teksTanya.includes("kanan") || teksTanya.includes("kiri")))) && !teksTanya.includes("selaput") && !teksTanya.includes("kacamata") && !teksTanya.includes("kornea") && !teksTanya.includes("juling") && !teksTanya.includes("kelopak"))) {
                    target = hasMata ? "ada indikasi gangguan penglihatan (visus <6/9)" : "normal (visus 6/6 - 6/9)";
                    fallback = hasMata ? "ada indikasi gangguan penglihatan" : "normal";
                } else if (teksTanya.includes("kacamata")) {
                    target = hasMata ? "ya" : "tidak";
                    fallback = hasMata ? "ya" : "tidak";
                }
                // C. Dinamis GDS & Hb pada pertanyaan standar
                else if ((teksTanya.includes("gds") || teksTanya.includes("gula darah") || teksTanya.includes("glukosa")) &&
                    !teksTanya.includes("gds 2") && !teksTanya.includes("gds2") && !teksTanya.includes("kedua")) {
                    target = valGDS;
                    fallback = valGDS;
                } else if (teksTanya.includes("hemoglobin") || teksTanya.includes("anemia") || teksTanya.includes("kadar hb") || teksTanya.includes("hasil hb") || teksTanya.includes(" hb") || teksTanya === "hb") {
                    target = valHB;
                    fallback = valHB;
                }
                // D. Pertanyaan Standar Lainnya
                else {
                    for (const map of pertanyaanKhususCKG) {
                        if (teksTanya.includes(bersih(map.q))) {
                            target = map.a;
                            fallback = map.fallback || null;
                            break;
                        }
                    }
                }

                // Cek apakah pertanyaan sudah terjawab sesuai target
                if (jenis === 'radio') {
                    if (radioSudahSesuai(qNode, target, fallback)) continue;
                    console.log(`[CKG] Mengisi pertanyaan [Pass ${pass}]: "${teksTanya.substring(0, 45)}..." -> Target: "${target}"`);
                    await isiRadioSurvey(qNode, target, fallback);
                    filledSomething = true;
                } else if (jenis === 'dropdown') {
                    if (dropdownSudahSesuai(qNode, target, fallback)) continue;
                    console.log(`[CKG] Mengisi dropdown [Pass ${pass}]: "${teksTanya.substring(0, 45)}..." -> Target: "${target}"`);
                    await isiDropdownSurvey(qNode, target, fallback);
                    filledSomething = true;
                } else if (jenis === 'text' || jenis === 'number') {
                    const textInput = qNode.querySelector('input, textarea');
                    if (textInput && textInput.value && String(textInput.value).trim()) continue;
                    console.log(`[CKG] Mengisi text input [Pass ${pass}]: "${teksTanya.substring(0, 45)}..." -> Target: "${target}"`);
                    await isiTextFieldSurvey(qNode, target);
                    filledSomething = true;
                }
                await sleep(300);
            }

            await sleep(400);

            if (!filledSomething) {
                console.log(`[CKG] Seluruh pertanyaan form SurveyJS telah terjawab tuntas pada pass ke-${pass}.`);
                break;
            }
        }
        await sleep(500);
    }

    // --
    // EXACT DROPDOWN SELECTION FOR ASIK
    // Struktur HTML dropdown Kemenkes:
    //   <div data-v-cf5614bc class="h-[2.9rem] ... cursor-pointer ...">
    //     <span data-v-cf5614bc class="line-clamp-2 text-gray-4">Pilih sekolah</span>
    //   </div>
    // Opsi popup:
    //   <div data-v-cf5614bc class="py-2 px-4 cursor-pointer text-sm hover:bg-gray-1">
    //     <div data-v-cf5614bc><div>SD NEGERI PALMERIAM 01</div></div>
    //   </div>
    // --
    const romanToNumMap = {
        'I': '1', 'II': '2', 'III': '3', 'IV': '4', 'V': '5', 'VI': '6',
        'VII': '7', 'VIII': '8', 'IX': '9',
        'X': '10', 'XI': '11', 'XII': '12'
    };
    const numToRomanMap = {
        '1': 'I', '2': 'II', '3': 'III', '4': 'IV', '5': 'V', '6': 'VI',
        '7': 'VII', '8': 'VIII', '9': 'IX',
        '10': 'X', '11': 'XI', '12': 'XII'
    };

    function extractClassNum(str) {
        if (!str) return '';
        const s = String(str).trim();
        if (!s || s.toLowerCase().startsWith('pilih')) return '';

        // Abaikan teks layanan/kuesioner (misal "Faktor Risiko TB - Anak Kelas 4-12")
        if (/faktor risiko|kuesioner|imunisasi|screening|pemeriksaan/i.test(s)) return '';

        // Bersihkan prefix umum (misal "SMA KELAS 10", "SD Kelas 1", "Tab: KELAS 10", "Rombel X.1")
        const cleanS = s.replace(/^tab:\s*/i, '')
            .replace(/^(?:sdn|sd|smp|smpn|sma|sman|smk|smkn|mas|mts|mis)\s+/i, '')
            .replace(/^(?:kelas|tingkat|rombel)\s+/i, '')
            .trim();

        // 1. Cek angka Romawi DULU untuk SMA/SMP/SD (misal "X.1", "X.5", "X-1", "X 1", "X", "XI.1", "XII.1", "XII.`1", "XII1", dst.)
        // Di SMA, X.1 artinya Kelas 10 Rombel 1; XI.2 artinya Kelas 11 Rombel 2; XII.1 artinya Kelas 12 Rombel 1!
        const romanMatch = cleanS.match(/^(XII|XI|X|IX|VIII|VII|VI|IV|V|III|II|I)(?:[\s\.\-_/`0-9a-zA-Z]|$)/i) ||
            s.match(/\b(XII|XI|X|IX|VIII|VII|VI|IV|V|III|II|I)\b/i);
        if (romanMatch) {
            const rUpper = romanMatch[1].toUpperCase();
            if (romanToNumMap[rUpper]) return romanToNumMap[rUpper];
        }

        // 2. Cek pola angka eksplisit 1-12 (misal "10.1", "10A", "10-1", "Kelas 10", "12", "1")
        // Catatan: Pada "10.1", angka kelas utamanya adalah 10 (bukan .1)
        const numMatch = cleanS.match(/^(1[0-2]|[1-9])(?:[\s\.\-_/a-zA-Z].*|$)/) ||
            s.match(/(?:kelas|tingkat|rombel|sd|smp|sma|smk)\s*(1[0-2]|[1-9])\b/i) ||
            s.match(/\b(1[0-2]|[1-9])\b/);
        if (numMatch) return numMatch[1];

        // 3. Fallback: ambil digit angka jika valid 1-12
        const digitsOnly = cleanS.replace(/[^0-9]/g, '');
        if (digitsOnly) {
            const n = parseInt(digitsOnly, 10);
            if (n >= 1 && n <= 12) return String(n);
        }

        return '';
    }

    function detectActiveClassOnAsikPage() {
        const mainArea = document.querySelector('main, section, div[class*="min-h-screen"]') || document.body;
        const allMainTriggers = Array.from(mainArea.querySelectorAll('div[class*="cursor-pointer"], div[class*="h-[2.9rem]"]'))
            .filter(el => {
                if (el.closest('#asik-ckg-panel, #asik-dashboard-overlay, aside, nav, #sidebar, [class*="sidebar"]')) return false;
                if (!isElementVisible(el)) return false;
                const span = el.querySelector('span');
                if (!span) return false;
                const t = bersih(span.textContent);
                if (!t || t === 'ckg umum' || t === 'ckg sekolah' || t === 'tatalaksana' || t === 'pelayanan') return false;
                return true;
            });

        let triggerSpan = null;
        if (allMainTriggers.length >= 2) {
            triggerSpan = allMainTriggers[1].querySelector('span') || allMainTriggers[1];
        }

        // Jika belum terdeteksi dari dropdown filter, cari dari elemen teks badge siswa di halaman detail
        if (!triggerSpan || !triggerSpan.textContent || triggerSpan.textContent.includes('pilih')) {
            const allElements = Array.from(document.querySelectorAll('div, span'))
                .filter(el => !el.closest('#asik-ckg-panel, #asik-dashboard-overlay, aside, nav, #sidebar, [class*="sidebar"]') && isElementVisible(el));
            const found = allElements.find(el => {
                const t = bersih(el.textContent);
                if (!t || t.includes('pilih kelas') || t.includes('faktor risiko') || t.includes('layanan') || t.includes('status') || t.includes('pemeriksaan')) return false;
                return /^(?:sd|smp|sma|smk)?\s*kelas\s*(1[0-2]|[1-9])\b/i.test(t) || /^(?:sd|smp|sma|smk)\s+kelas/i.test(t);
            });
            if (found) triggerSpan = found;
        }

        if (!triggerSpan) return { text: '', num: '' };
        const rawText = bersih(triggerSpan.textContent || '');
        if (!rawText || rawText === 'pilih kelas') return { text: '', num: '' };
        return { text: rawText, num: extractClassNum(rawText) };
    }

    async function pilihDropdownASIK(tipeDropdown, targetNama) {
        console.log(`[CKG] === Memilih Dropdown ${tipeDropdown}: "${targetNama}" ===`);

        const isSekolah = tipeDropdown.toLowerCase().includes('sekolah');
        const cleanTarget = bersih(targetNama);

        // 1. CARI SELURUH KOTAK DROPDOWN DI AREA KONTEN UTAMA (KECUALIKAN SIDEBAR, PANEL, & NAV)
        const mainArea = document.querySelector('main, section, div[class*="min-h-screen"]') || document.body;
        const allMainTriggers = Array.from(mainArea.querySelectorAll('div[class*="cursor-pointer"], div[class*="h-[2.9rem]"]'))
            .filter(el => {
                if (el.closest('#asik-ckg-panel, #asik-dashboard-overlay, aside, nav, #sidebar, [class*="sidebar"]')) return false;
                if (!isElementVisible(el)) return false;
                const span = el.querySelector('span');
                if (!span) return false;
                const t = bersih(span.textContent);
                if (!t || t === 'ckg umum' || t === 'ckg sekolah' || t === 'tatalaksana' || t === 'pelayanan') return false;
                return true;
            });

        // Dropdown 1 = Sekolah, Dropdown 2 = Kelas
        const triggerIndex = isSekolah ? 0 : 1;
        let clickBox = allMainTriggers[triggerIndex] || null;
        let triggerSpan = clickBox ? clickBox.querySelector('span') : null;

        // Fallback pencarian teks trigger jika tidak ditemukan berdasarkan indeks
        if (!clickBox) {
            const allSpans = Array.from(document.querySelectorAll('span, div.cursor-pointer'))
                .filter(el => !el.closest('#asik-ckg-panel, #asik-dashboard-overlay, aside, nav, #sidebar, [class*="sidebar"]') && isElementVisible(el));
            triggerSpan = allSpans.find(el => {
                const t = bersih(el.textContent);
                return isSekolah
                    ? (t === 'pilih sekolah' || (t.includes('sd') || t.includes('smp') || t.includes('sma') || t.includes('smk')))
                    : (t === 'pilih kelas' || t.includes('kelas') || /^(sd|smp|sma|smk)\s+kelas/i.test(t));
            });
            if (triggerSpan) {
                clickBox = triggerSpan.closest('div.cursor-pointer, div[class*="cursor-pointer"]') || triggerSpan.parentElement;
            }
        }

        if (!clickBox) {
            console.log(`[CKG] Kotak dropdown "${tipeDropdown}" tidak ditemukan di layar.`);
            return false;
        }

        const currentText = triggerSpan ? bersih(triggerSpan.textContent) : bersih(clickBox.textContent);
        console.log(`[CKG] Status kotak dropdown ${tipeDropdown} saat ini: "${currentText}"`);

        // 2. CEK APAKAH SUDAH TERPILIH SESUAI TARGET
        if (isSekolah) {
            const words = cleanTarget.replace(/^(sd|sdn|smp|smpn|sma|sman|smk|smkn|mas|mts|mis)\s+/i, '').trim().split(/\s+/).filter(w => w.length >= 2);
            if (currentText.includes(cleanTarget) || (words.length > 0 && words.every(w => currentText.includes(w)))) {
                console.log(`[CKG] ✅ Sekolah "${targetNama}" SUDAH TERPILIH di formulir ASIK ("${currentText}").`);
                return true;
            }
        } else {
            const targetClassNum = extractClassNum(cleanTarget);
            const currentClassNum = extractClassNum(currentText);

            if (targetClassNum && currentClassNum && targetClassNum === currentClassNum) {
                console.log(`[CKG] ✅ Kelas "${targetNama}" (Kelas ${targetClassNum}) SUDAH TERPILIH di formulir ASIK ("${currentText}").`);
                return true;
            }
            if (currentText === cleanTarget || currentText === `kelas ${cleanTarget}`) {
                console.log(`[CKG] ✅ Kelas "${targetNama}" SUDAH TERPILIH di formulir ASIK ("${currentText}").`);
                return true;
            }
        }

        // 3. KLIK KOTAK DROPDOWN UNTUK MEMBUKA POPUP OPSI
        console.log(`[CKG] Mengklik kotak dropdown "${tipeDropdown}" untuk membuka daftar opsi...`);
        triggerClick(clickBox);
        if (triggerSpan && triggerSpan !== clickBox) triggerClick(triggerSpan);
        ['pointerdown', 'mousedown', 'pointerup', 'mouseup', 'click'].forEach(evtType => {
            try {
                clickBox.dispatchEvent(new MouseEvent(evtType, { bubbles: true, cancelable: true, view: window, detail: 1 }));
                if (triggerSpan) triggerSpan.dispatchEvent(new MouseEvent(evtType, { bubbles: true, cancelable: true, view: window, detail: 1 }));
            } catch (e) { }
        });
        if (typeof clickBox.click === 'function') {
            try { clickBox.click(); } catch (e) { }
        }
        await sleep(1500); // Tempo 1.5 detik menunggu popup terbuka

        // 4. FUNGSI PENCARIAN OPSI POPUP
        const cariOpsiDiLayar = () => {
            const allCandidates = Array.from(document.querySelectorAll('div.py-2, div.px-4, div[class*="hover:bg-gray"], div.cursor-pointer.text-sm, [role="option"]'))
                .filter(el => !el.closest('#asik-ckg-panel, #asik-dashboard-overlay, aside, nav, #sidebar') && isElementVisible(el))
                .filter(el => {
                    if (el === clickBox || el === triggerSpan) return false;

                    const t = bersih(el.textContent);
                    if (!t) return false;
                    if (t.startsWith('pilih ') || t === 'ckg umum' || t === 'ckg sekolah' || t === 'tatalaksana' || t === 'pelayanan') return false;
                    if (t.includes('silakan lakukan filter') || t.includes('tampilkan pencarian')) return false;

                    // PENTING: Harus elemen baris individual (bukan pembungkus daftar banyak sekolah)
                    if (t.length > 80) return false;
                    if (el.querySelector('div.py-2, div.px-4, div[class*="hover:bg-gray"]')) return false;

                    return true;
                });

            if (isSekolah) {
                const words = cleanTarget.replace(/^(sd|sdn|smp|smpn|sma|sman|smk|smkn|mas|mts|mis)\s+/i, '').trim().split(/\s+/).filter(w => w.length >= 2);
                const mainWord = words[0] || cleanTarget;

                // 1. Match full target
                let opt = allCandidates.find(el => bersih(el.textContent).includes(cleanTarget));
                // 2. Match all keyword tokens (e.g. "palmeriam" dan "01")
                if (!opt && words.length > 1) {
                    opt = allCandidates.find(el => {
                        const t = bersih(el.textContent);
                        return words.every(w => t.includes(w));
                    });
                }
                // 3. Match main keyword
                if (!opt && mainWord.length >= 3) {
                    opt = allCandidates.find(el => bersih(el.textContent).includes(mainWord));
                }
                return opt;
            } else {
                const targetClassNum = extractClassNum(cleanTarget);
                const targetRoman = targetClassNum ? (numToRomanMap[targetClassNum] || '') : '';

                // 1. Exact text match (mendukung format SMA/SMP/SD Kelas X/10)
                let opt = allCandidates.find(el => {
                    const t = bersih(el.textContent);
                    return t === cleanTarget ||
                        t === `kelas ${cleanTarget}` ||
                        t === `sma kelas ${cleanTarget}` ||
                        t === `smp kelas ${cleanTarget}` ||
                        t === `sd kelas ${cleanTarget}` ||
                        (targetRoman && (
                            t === `kelas ${targetRoman.toLowerCase()}` ||
                            t === targetRoman.toLowerCase() ||
                            t === `sma kelas ${targetRoman.toLowerCase()}` ||
                            t === `smp kelas ${targetRoman.toLowerCase()}` ||
                            t === `sd kelas ${targetRoman.toLowerCase()}`
                        ));
                });

                // 2. Exact class number match (SANGAT PENTING: Kelas 10 TIDAK AKAN PERNAH cocok dengan 12 atau 1!)
                if (!opt && targetClassNum) {
                    opt = allCandidates.find(el => {
                        const optText = bersih(el.textContent);
                        const optNum = extractClassNum(optText);
                        return optNum === targetClassNum;
                    });
                }

                // 3. Regex boundary match (mencegah substring 1 cocok ke 10 atau 12)
                if (!opt && targetClassNum) {
                    opt = allCandidates.find(el => {
                        const t = bersih(el.textContent);
                        const reg = new RegExp(`(?:^|\\s|kelas\\s+|tingkat\\s+)${targetClassNum}(?:\\s|$|[a-z]|-)`, 'i');
                        return reg.test(t);
                    });
                }
                return opt;
            }
        };

        // 5. CARI OPSI LANGSUNG DARI POPUP
        let matchedOption = cariOpsiDiLayar();

        // 6. JIKA SEKOLAH BELUM KETEMU DI DAFTAR AWAL & ADA INPUT PENCARIAN #sekolah, KETIK KATA KUNCI
        if (isSekolah && !matchedOption) {
            const inputSekolah = document.querySelector('#sekolah, input[placeholder*="sekolah" i]');
            if (inputSekolah && isElementVisible(inputSekolah)) {
                const words = cleanTarget.replace(/^(sd|sdn|smp|smpn|sma|sman|smk|smkn|mas|mts|mis)\s+/i, '').trim().split(/\s+/);
                const keyword = words[0] || cleanTarget;
                console.log(`[CKG] Mengetik pencarian sekolah: "${keyword}"`);
                isiInput(inputSekolah, keyword, false);
                await sleep(1500);
                matchedOption = cariOpsiDiLayar();
            }
        }

        // 7. KLIK OPSI YANG DITEMUKAN
        if (matchedOption) {
            const optTarget = matchedOption.closest('div.cursor-pointer, [class*="cursor-pointer"]') || matchedOption;
            console.log(`[CKG] -> OPSI DITEMUKAN: "${optTarget.textContent.trim()}". Mengklik...`);

            try { optTarget.scrollIntoView({ block: 'nearest' }); } catch (e) { }

            ['pointerdown', 'mousedown', 'pointerup', 'mouseup', 'click'].forEach(evtType => {
                try {
                    optTarget.dispatchEvent(new MouseEvent(evtType, { bubbles: true, cancelable: true, view: window, detail: 1 }));
                } catch (e) { }
            });

            if (typeof optTarget.click === 'function') {
                try { optTarget.click(); } catch (e) { }
            }

            const childDiv = optTarget.querySelector('div');
            if (childDiv) {
                try {
                    childDiv.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
                } catch (e) { }
            }

            await sleep(2000); // Tempo 2 detik setelah memilih
            return true;
        }

        console.log(`[CKG] -> Gagal menemukan opsi untuk ${tipeDropdown} ("${targetNama}").`);
        return false;
    }

    // --
    // HELPER: KLIK TOMBOL "TAMPILKAN PENCARIAN" DENGAN RETRY & BERSIHKAN INPUT NAMA
    // --
    async function klikTombolTampilkanPencarian() {
        console.log("[CKG] === Memproses Klik Tombol 'Tampilkan Pencarian' ===");

        for (let attempt = 1; attempt <= 6; attempt++) {
            const allCandidates = Array.from(document.querySelectorAll('button, div, a, span')).filter(el => {
                if (!isElementVisible(el) || el.closest('#asik-ckg-panel, #asik-dashboard-overlay, [id*="asik-ckg"]')) return false;
                const t = bersih(el.textContent);
                return t === 'tampilkan pencarian' || t.includes('tampilkan pencarian');
            });

            let btn = allCandidates.find(el => el.tagName === 'BUTTON' && !el.disabled && !el.classList.contains('cursor-not-allowed')) ||
                allCandidates.find(el => el.closest('button') && !el.closest('button').disabled && !el.closest('button').classList.contains('cursor-not-allowed')) ||
                allCandidates.find(el => !el.classList.contains('cursor-not-allowed') && !el.classList.contains('bg-disabled')) ||
                allCandidates[0];

            if (btn) {
                const targetBtn = btn.closest('button') || btn;
                console.log(`[CKG] Tombol 'Tampilkan Pencarian' ditemukan (Percobaan ${attempt}):`, targetBtn);

                // Scroll into view
                try { targetBtn.scrollIntoView({ behavior: 'smooth', block: 'center' }); } catch (e) { }
                await sleep(200);

                // Bersihkan input pencarian #searchNik terlebih dahulu agar tidak menyaring hasil baru dengan kata kunci lama
                const searchInp = document.querySelector('input#searchNik') || document.querySelector('input[placeholder*="cari"], input[placeholder*="nik"]');
                if (searchInp) {
                    isiInput(searchInp, '', false);
                    await sleep(100);
                }

                // Trigger full click events on targetBtn and its children
                ['pointerdown', 'mousedown', 'pointerup', 'mouseup', 'click'].forEach(evt => {
                    try {
                        targetBtn.dispatchEvent(new MouseEvent(evt, { bubbles: true, cancelable: true, view: window, detail: 1 }));
                    } catch (e) { }
                });

                if (typeof targetBtn.click === 'function') {
                    try { targetBtn.click(); } catch (e) { }
                }

                const childSpan = targetBtn.querySelector('span, div');
                if (childSpan) {
                    try {
                        childSpan.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
                    } catch (e) { }
                }

                console.log("[CKG] Berhasil mengklik 'Tampilkan Pencarian'. Menunggu server memuat daftar siswa...");
                await sleep(2500); // Waktu jeda agar server ASIK selesai memuat daftar siswa kelas baru
                return true;
            }
            await sleep(500);
        }

        console.warn("[CKG] Tombol 'Tampilkan Pencarian' tidak ditemukan setelah 6 percobaan.");
        return false;
    }

    // --
    // HELPER: CEK APAKAH LAYANAN SUDAH SELESAI / LENGKAP
    // --
    function layananSelesai(row) {
        if (!row) return true;
        const rawTeks = row.innerText || row.textContent || '';
        const teks = rawTeks.toLowerCase().replace(/\s+/g, ' ');
        if (teks.includes("selesai diperiksa") || teks.includes("selesai pemeriksaan") || teks.includes("sudah diperiksa") || teks.includes("lengkap") || teks.includes("selesai")) return true;
        if (teks.includes("tidak diperiksa") && !teks.includes("belum diperiksa")) return true;

        // Cek badge atau ikon hijau standar Kemenkes (icon-success.svg)
        const hijau = row.querySelector('img[src*="icon-success.svg"], .text-teal-500, [class*="text-teal"], [class*="bg-teal"], [class*="bg-[#E1F7F0]"]');
        const abu = row.querySelector('img[src*="icon-success-gray.svg"]');
        if (hijau && !abu) return true;

        // Cek innerHTML langsung untuk pola warna hijau ASIK (#40AE87 atau base64 IzQwQUU4NyI)
        const innerHTML = row.innerHTML || '';
        if (innerHTML.includes('IzQwQUU4NyI') || innerHTML.includes('#40AE87') || innerHTML.includes('#40ae87') || innerHTML.includes('text-[#40ae87]') || innerHTML.includes('text-[#40AE87]')) {
            return true;
        }

        // Cek badge class hijau (layout baru Sep 2026): text-[#40AE87] bg-[#E1F7F0] border-[#9CD9C3]
        const greenBadge = row.querySelector('[class*="text-[#40AE87]"], [class*="text-[#40ae87]"], [class*="bg-[#E1F7F0]"], [class*="border-[#9CD9C3]"]');
        if (greenBadge) return true;

        // KHUSUS STATUS TABEL PEMERIKSAAN MANDIRI (ASIK CKG):
        // Status di kolom ke-2 (td.text-center) menggunakan gambar SVG base64
        // Hijau (#40AE87 / #16b3ac) = SELESAI
        // Abu-abu (#CACCCF) = BELUM SELESAI
        const statusImg = row.querySelector('td:nth-child(2) img, td.text-center img, img');
        if (statusImg) {
            const src = statusImg.getAttribute('src') || statusImg.src || '';
            if (src.includes('data:image/svg+xml')) {
                // Deteksi cepat pola warna hijau #40AE87
                if (src.includes('IzQwQUU4NyI') || src.includes('40AE87') || src.includes('40ae87')) {
                    return true;
                }
                try {
                    const b64 = src.split(',')[1];
                    if (b64) {
                        const decoded = atob(b64).toLowerCase();
                        if (decoded.includes('40ae87') || decoded.includes('16b3ac') || decoded.includes('rgb(64, 174, 135)')) {
                            return true;
                        }
                    }
                } catch (e) { }
            }
        }

        return false;
    }

    // --
    // HELPER: CEK APAKAH LAYANAN WAJIB DILEWATI (SKIP TANPA ISI)
    // Layanan Nakes yang tidak perlu diisi: Pemeriksaan Kadar CO, RDT Malaria, Hepatitis Nakes, dll
    // Sejak update Sep 2026, toggle switch sudah dihapus, jadi cukup SKIP (jangan klik Input Data)
    // --
    function layananWajibLewati(rowOrText) {
        if (!rowOrText) return false;
        const raw = typeof rowOrText === 'string' ? rowOrText : (rowOrText.innerText || rowOrText.textContent || '');
        const teks = raw.toLowerCase().replace(/\s+/g, ' ');

        // PENTING: Jangan pernah lewati form kuesioner yang harus diisi (Input Data)!
        if (teks.includes("perilaku merokok") || teks.includes("faktor risiko hepatitis") || teks.includes("faktor risiko malaria")) {
            return false;
        }
        // PENTING: Jangan lewati form Anemia, Gula Darah Remaja, Skrining Telinga/Mata, Skrining Gigi, Gizi, dll
        if (teks.includes("anemia") || teks.includes("gula darah") || teks.includes("skrining telinga") || teks.includes("skrining gigi") || teks.includes("gizi anak") || teks.includes("tekanan darah") || teks.includes("kebugaran jasmani") || teks.includes("penyakit tropis") || teks.includes("kusta") || teks.includes("frambusia") || teks.includes("skabies")) {
            return false;
        }

        const list = [
            "kadar co",
            "rdt malaria",
            "pemeriksaan hepatitis",
            "sifilis", "lipid panel", "profil lipid",
            "payudara", "hpv", "inspekulo", "iva",
            "sirosis", "fibrosis"
        ];
        return list.some(k => teks.includes(k));
    }

    // --
    // RIWAYAT PROSES & LAPORAN STATUS SISWA
    // --
    let riwayatPasien = [];
    function muatRiwayat() {
        try {
            riwayatPasien = JSON.parse(localStorage.getItem('asik_ckg_riwayat') || '[]');
        } catch (e) { riwayatPasien = []; }
    }
    function simpanRiwayat() {
        try {
            if (riwayatPasien.length > 300) riwayatPasien = riwayatPasien.slice(-300);
            localStorage.setItem('asik_ckg_riwayat', JSON.stringify(riwayatPasien));
        } catch (e) { }
    }
    function catatRiwayat(nama, kelas, status, keterangan) {
        muatRiwayat();
        const waktu = new Date().toLocaleString('id-ID', { hour12: false });
        riwayatPasien.push({
            no: riwayatPasien.length + 1,
            nama: nama || '-',
            kelas: kelas || '-',
            status: status,
            keterangan: keterangan || '-',
            waktu: waktu
        });
        simpanRiwayat();
    }
    try { muatRiwayat(); } catch (e) { }

    function tampilkanDashboard() {
        const old = document.getElementById('asik-dashboard-overlay');
        if (old) { old.remove(); return; }

        muatRiwayat();

        const overlay = document.createElement('div');
        overlay.id = 'asik-dashboard-overlay';
        overlay.style.cssText = 'position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.65);z-index:9999999;display:flex;align-items:center;justify-content:center;backdrop-filter:blur(3px);';

        const total = riwayatPasien.length;
        const isSkip = (s) => {
            const st = (s || '').toLowerCase();
            return st.includes('skip') || st.includes('dilewati') || st.includes('lewati') || st.includes('tidak ditemukan');
        };
        const isGagal = (s) => {
            const st = (s || '').toLowerCase();
            return (st.includes('gagal') || st.includes('error')) && !isSkip(s);
        };
        const isSukses = (s) => {
            const st = (s || '').toLowerCase();
            return st.includes('selesai') && !isSkip(s) && !isGagal(s);
        };

        const sukses = riwayatPasien.filter(r => isSukses(r.status)).length;
        const gagal = riwayatPasien.filter(r => isGagal(r.status)).length;
        const skip = riwayatPasien.filter(r => isSkip(r.status)).length;

        const sorted = [...riwayatPasien].reverse();
        let tableRows = '';
        sorted.forEach((r, i) => {
            const st = r.status || '';
            const isSk = isSkip(st);
            const isGg = isGagal(st);
            const isOk = isSukses(st);

            const bgColor = isOk ? '#0d3320' : (isGg ? '#3b1111' : '#2a2000');
            const statusColor = isOk ? '#4ade80' : (isGg ? '#f87171' : '#fbbf24');
            tableRows += `<tr style="background:${bgColor}">
                <td style="padding:8px 12px;border-bottom:1px solid #334155;color:#94a3b8;">${sorted.length - i}</td>
                <td style="padding:8px 12px;border-bottom:1px solid #334155;color:#e2e8f0;font-weight:600;">${r.nama}</td>
                <td style="padding:8px 12px;border-bottom:1px solid #334155;color:#94a3b8;">${r.kelas}</td>
                <td style="padding:8px 12px;border-bottom:1px solid #334155;font-weight:700;color:${statusColor};">${r.status}</td>
                <td style="padding:8px 12px;border-bottom:1px solid #334155;color:#cbd5e1;">${r.keterangan}</td>
                <td style="padding:8px 12px;border-bottom:1px solid #334155;color:#64748b;font-size:11px;">${r.waktu}</td>
            </tr>`;
        });

        const panel = document.createElement('div');
        panel.style.cssText = 'background:#0f172a;border-radius:16px;width:92%;max-width:920px;max-height:85vh;overflow:hidden;display:flex;flex-direction:column;box-shadow:0 20px 60px rgba(0,0,0,0.6);border:1px solid #334155;font-family:system-ui,-apple-system,sans-serif;';
        panel.innerHTML = `
            <div style="padding:18px 24px;background:#1e293b;display:flex;align-items:center;justify-content:space-between;border-bottom:1px solid #334155;">
                <div>
                    <h2 style="margin:0;color:#f8fafc;font-size:17px;display:flex;align-items:center;gap:8px;">📊 Laporan Status Robot CKG Sekolah</h2>
                    <p style="margin:4px 0 0;color:#64748b;font-size:12px;">Riwayat pengisian kuesioner Nakes dan Tatalaksana Siswa</p>
                </div>
                <button id="asik-dash-close" style="background:none;border:none;color:#94a3b8;font-size:24px;cursor:pointer;padding:4px 8px;line-height:1;">&times;</button>
            </div>
            <div style="padding:16px 24px;display:flex;gap:12px;flex-wrap:wrap;">
                <div style="flex:1;min-width:110px;background:#1e293b;border-radius:10px;padding:12px 16px;text-align:center;border:1px solid #334155;">
                    <div style="font-size:26px;font-weight:800;color:#e2e8f0;">${total}</div>
                    <div style="font-size:11px;color:#94a3b8;margin-top:2px;">Total Siswa</div>
                </div>
                <div style="flex:1;min-width:110px;background:#052e16;border:1px solid #166534;border-radius:10px;padding:12px 16px;text-align:center;">
                    <div style="font-size:26px;font-weight:800;color:#4ade80;">${sukses}</div>
                    <div style="font-size:11px;color:#86efac;margin-top:2px;">Selesai</div>
                </div>
                <div style="flex:1;min-width:110px;background:#450a0a;border:1px solid #991b1b;border-radius:10px;padding:12px 16px;text-align:center;">
                    <div style="font-size:26px;font-weight:800;color:#f87171;">${gagal}</div>
                    <div style="font-size:11px;color:#fca5a5;margin-top:2px;">Gagal</div>
                </div>
                <div style="flex:1;min-width:110px;background:#422006;border:1px solid #92400e;border-radius:10px;padding:12px 16px;text-align:center;">
                    <div style="font-size:26px;font-weight:800;color:#fbbf24;">${skip}</div>
                    <div style="font-size:11px;color:#fde68a;margin-top:2px;">Skip</div>
                </div>
            </div>
            <div style="flex:1;overflow-y:auto;padding:0 24px 16px;">
                <table style="width:100%;border-collapse:collapse;font-size:12px;">
                    <thead>
                        <tr style="background:#1e293b;position:sticky;top:0;">
                            <th style="padding:10px 12px;text-align:left;color:#94a3b8;font-weight:600;border-bottom:2px solid #334155;">No</th>
                            <th style="padding:10px 12px;text-align:left;color:#94a3b8;font-weight:600;border-bottom:2px solid #334155;">Nama Siswa</th>
                            <th style="padding:10px 12px;text-align:left;color:#94a3b8;font-weight:600;border-bottom:2px solid #334155;">Kelas</th>
                            <th style="padding:10px 12px;text-align:left;color:#94a3b8;font-weight:600;border-bottom:2px solid #334155;">Status</th>
                            <th style="padding:10px 12px;text-align:left;color:#94a3b8;font-weight:600;border-bottom:2px solid #334155;">Keterangan</th>
                            <th style="padding:10px 12px;text-align:left;color:#94a3b8;font-weight:600;border-bottom:2px solid #334155;">Waktu</th>
                        </tr>
                    </thead>
                    <tbody>${tableRows || '<tr><td colspan="6" style="padding:28px;text-align:center;color:#64748b;">Belum ada riwayat proses</td></tr>'}</tbody>
                </table>
            </div>
            <div style="padding:14px 24px;background:#1e293b;display:flex;gap:10px;border-top:1px solid #334155;">
                <button id="asik-dash-download" style="flex:1;padding:10px;border:none;border-radius:8px;background:#2563eb;color:white;font-weight:700;font-size:12px;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:6px;">📥 Download CSV</button>
                <button id="asik-dash-clear" style="flex:1;padding:10px;border:none;border-radius:8px;background:#dc2626;color:white;font-weight:700;font-size:12px;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:6px;">🗑️ Hapus Riwayat</button>
            </div>
        `;

        overlay.appendChild(panel);
        document.body.appendChild(overlay);

        document.getElementById('asik-dash-close').onclick = () => overlay.remove();
        overlay.onclick = (e) => { if (e.target === overlay) overlay.remove(); };

        document.getElementById('asik-dash-download').onclick = () => {
            if (riwayatPasien.length === 0) { alert('Belum ada riwayat untuk diunduh!'); return; }
            const header = 'No,Nama Siswa,Kelas,Status,Keterangan,Waktu';
            const rows = riwayatPasien.map(r =>
                `${r.no},"${(r.nama || '').replace(/"/g, '""')}","${r.kelas}","${r.status}","${(r.keterangan || '').replace(/"/g, '""')}","${r.waktu}"`
            );
            const csv = '\uFEFF' + header + '\n' + rows.join('\n');
            const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `Laporan-CKG-Sekolah-${new Date().toISOString().slice(0, 10)}.csv`;
            a.click();
            URL.revokeObjectURL(url);
        };

        document.getElementById('asik-dash-clear').onclick = () => {
            if (confirm('Yakin ingin menghapus semua catatan riwayat proses?')) {
                riwayatPasien = [];
                simpanRiwayat();
                overlay.remove();
                alert('Riwayat berhasil dikosongkan!');
            }
        };
    }

    // --
    // HELPER: KEMBALI 2 KALI (Setelah Selesai Tatalaksana / Pasien)
    // --
    async function klikBackDuaKali() {
        console.log("[CKG] Eksekusi kembali 2x ke daftar pencarian siswa...");
        for (let i = 1; i <= 2; i++) {
            console.log(`[CKG] Menekan tombol kembali (${i}/2)...`);
            const backBtn = Array.from(document.querySelectorAll('img.cursor-pointer, .cursor-pointer')).find(el => {
                if (!isElementVisible(el)) return false;
                const cls = el.className || '';
                const src = el.src || '';
                return (cls.includes('w-[24px]') && cls.includes('h-[24px]')) || src.includes('icon-arrow-left') || src.includes('arrow-left');
            });

            if (backBtn) {
                triggerClick(backBtn);
            } else {
                window.history.back();
            }
            await sleep(2000);
        }
    }

    // --
    // AUTOMATION LOOP
    // --
    async function jalankanLoop() {
        const state = getState();
        if (!state.running || state.data.length === 0) return;

        try {
            if (!state.sweepMode && state.index >= state.data.length) {
                let wbData = null;
                try {
                    const raw = GM_getValue('asik_ckg_workbook_data') || localStorage.getItem('asik_ckg_workbook_data');
                    if (raw) wbData = typeof raw === 'string' ? JSON.parse(raw) : raw;
                } catch (e) { }

                // Cari nama sheet terakhir yang baru saja selesai
                const lastProcessed = state.data && state.data.length > 0 ? state.data[state.data.length - 1] : null;
                const currentTab = (lastProcessed ? lastProcessed.TAB : '') || GM_getValue('asik_ckg_selected_tab', '');

                if (wbData && wbData.sheetNames && wbData.sheetNames.length > 0) {
                    // Temukan index sheet saat ini di daftar sheetNames
                    let curIdx = wbData.sheetNames.indexOf(currentTab);
                    if (curIdx === -1 && currentTab) {
                        curIdx = wbData.sheetNames.findIndex(n => {
                            const cn = bersih(n);
                            const ct = bersih(currentTab);
                            const numN = n.replace(/[^0-9]/g, '');
                            const numT = currentTab.replace(/[^0-9]/g, '');
                            return cn === ct || cn.includes(ct) || ct.includes(cn) || (numN && numN === numT);
                        });
                    }

                    // Jika masih ada sheet berikutnya (misal Kelas 6 setelah Kelas 5)
                    if (curIdx !== -1 && curIdx < wbData.sheetNames.length - 1) {
                        const nextTab = wbData.sheetNames[curIdx + 1];
                        console.log(`[CKG] 🎉 Tab "${wbData.sheetNames[curIdx]}" selesai! Otomatis melanjutkan ke Tab berikutnya: "${nextTab}"...`);

                        const nextStudents = wbData.sheets[nextTab] || [];
                        if (nextStudents.length > 0) {
                            GM_setValue('asik_ckg_selected_tab', nextTab);
                            const sheetSelect = document.getElementById('asik-ckg-sheet');
                            if (sheetSelect) sheetSelect.value = nextTab;

                            // Periksa apakah halaman saat ini masih di Detail Siswa, jika iya kembali ke tabel
                            const searchInput = document.querySelector('input#searchNik') || document.querySelector('input[placeholder*="cari"], input[placeholder*="nik"]');
                            if (!searchInput) {
                                console.log("[CKG] Kembali ke halaman tabel pencarian siswa untuk kelas baru...");
                                const backBtn = Array.from(document.querySelectorAll('img.cursor-pointer, .cursor-pointer, button')).find(el => {
                                    if (!isElementVisible(el) || el.closest('#asik-ckg-panel, #asik-dashboard-overlay')) return false;
                                    const cls = el.className || '';
                                    const src = el.src || '';
                                    return (cls.includes('w-[24px]') && cls.includes('h-[24px]')) || src.includes('icon-arrow-left') || src.includes('arrow-left');
                                });
                                if (backBtn) triggerClick(backBtn);
                                else window.history.back();
                                await sleep(2000);
                            }

                            // Update state dengan data kelas baru dan aktifkan flag needsClassSearch
                            setState({
                                ...state,
                                data: nextStudents,
                                index: 0,
                                running: true,
                                phase: 'MENCARI',
                                needsClassSearch: true,
                                searchRetry: 0,
                                tatalaksanaSelesai: false,
                                isTatalaksana: false
                            });
                            await sleep(1000);
                            return;
                        }
                    }
                }

                // Setelah seluruh sheet/antrean normal selesai, otomatis masuk ke mode penyisiran (sweep) tab "Sedang Pemeriksaan"
                console.log("[CKG] 🔄 Seluruh antrean Excel telah selesai diproses! Memulai mode penyisiran tab 'Sedang Pemeriksaan' dari kelas awal...");
                catatRiwayat('Antrean Normal Selesai', '-', 'Info', 'Memulai penyisiran data yang tertinggal di tab Sedang Pemeriksaan...');

                // Pastikan kembali ke halaman tabel pencarian siswa jika saat ini di detail
                const searchInput = document.querySelector('input#searchNik') || document.querySelector('input[placeholder*="cari"], input[placeholder*="nik"]');
                if (!searchInput) {
                    console.log("[CKG] Kembali ke halaman tabel pencarian siswa untuk mode penyisiran...");
                    const backBtn = Array.from(document.querySelectorAll('img.cursor-pointer, .cursor-pointer, button')).find(el => {
                        if (!isElementVisible(el) || el.closest('#asik-ckg-panel, #asik-dashboard-overlay')) return false;
                        const cls = el.className || '';
                        const src = el.src || '';
                        return (cls.includes('w-[24px]') && cls.includes('h-[24px]')) || src.includes('icon-arrow-left') || src.includes('arrow-left');
                    });
                    if (backBtn) triggerClick(backBtn);
                    else window.history.back();
                    await sleep(2000);
                }

                GM_setValue('asik_ckg_sweep_attempts', {});
                setState({
                    ...state,
                    sweepMode: true,
                    sweepSheetIndex: 0,
                    sweepCurrentStudent: null,
                    sweepEmptyCheckCount: 0,
                    running: true,
                    phase: 'MENCARI',
                    needsClassSearch: true,
                    searchRetry: 0,
                    tatalaksanaSelesai: false,
                    isTatalaksana: false
                });
                await sleep(1000);
                return;
            }

            let currentData = state.sweepMode
                ? (state.sweepCurrentStudent || { NAMA: 'Siswa Sedang Pemeriksaan', KELAS: '1', TB: '160', BB: '46', GDS: '90', HB: '12', MATA: 'normal', TELINGA: 'normal', GIGI: 'normal' })
                : (state.data && state.data[state.index] ? state.data[state.index] : null);

            // Sinkronisasi data sesi aktif (sessionStorage)
            try {
                const storedRaw = sessionStorage.getItem('asik_ckg_active_student');
                if (storedRaw) {
                    const parsed = JSON.parse(storedRaw);
                    if (parsed) {
                        if (state.sweepMode) {
                            currentData = { ...(currentData || {}), ...parsed };
                        } else if (currentData && currentData.NAMA) {
                            // Di mode normal: HANYA sinkronkan data klinis (GDS/HB) jika nama siswa memang COCOK!
                            // JANGAN PERNAH menimpa NAMA atau KELAS dari data Excel siswa saat ini!
                            if (isNamaCocok(currentData.NAMA, parsed.NAMA)) {
                                currentData = { ...parsed, ...currentData };
                            } else {
                                // Data di sessionStorage adalah milik siswa sebelumnya yang sudah selesai/dilewati
                                sessionStorage.removeItem('asik_ckg_active_student');
                                sessionStorage.removeItem('asik_ckg_active_student_name');
                            }
                        }
                    }
                }
            } catch (e) { }

            if (!currentData) {
                console.warn("[CKG] Data siswa saat ini tidak tersedia. Menunggu...");
                await sleep(1500);
                return;
            }
            const html = document.body.innerText.toLowerCase();
            // ============================================================
            // 0. CEK HALAMAN ERROR 404 ("This page could not be found")
            // ============================================================
            const bodyText = document.body ? document.body.innerText.toLowerCase() : "";
            const is404 = bodyText.includes('this page could not be found') ||
                bodyText.includes('page not found') ||
                bodyText.includes('halaman tidak ditemukan') ||
                (bodyText.includes('404') && (bodyText.includes('could not be found') || bodyText.includes('not found')));

            if (is404) {
                const lastRowId = GM_getValue('asik_ckg_last_clicked_row', '') || sessionStorage.getItem('asik_ckg_last_clicked_row') || '';
                if (lastRowId) {
                    const dilewati = new Set(GM_getValue('asik_ckg_dilewati', []));
                    dilewati.add(lastRowId);
                    GM_setValue('asik_ckg_dilewati', Array.from(dilewati));
                    console.warn(`[CKG 404] Item "${lastRowId}" ditandai BLACKLIST (404).`);
                }
                if (currentData) {
                    catatRiwayat(currentData.NAMA, currentData.KELAS, 'Dilewati (404)', 'Halaman 404 tidak ditemukan, ditandai selesai/dilewati otomatis');
                }
                state.phase = 'DETAIL_PELAYANAN';
                setState(state);
                window.history.back();
                await sleep(2500);
                return;
            }

            // ============================================================
            // 0.05 CEK ERROR MENAMPILKAN FORMULIR ("Terjadi kesalahan saat menampilkan formulir...")
            // ============================================================
            const isFormError = bodyText.includes('terjadi kesalahan saat menampilkan formulir') ||
                bodyText.includes('kesalahan saat menampilkan formulir') ||
                (bodyText.includes('menampilkan formulir') && (bodyText.includes('kembali ke halaman sebelumnya') || bodyText.includes('helpdesk'))) ||
                (bodyText.includes('laporkan kendala pada helpdesk') && (bodyText.includes('formulir') || bodyText.includes('kembali ke halaman')));

            if (isFormError) {
                console.warn("[CKG] ⚠️ Terdeteksi error: 'Terjadi kesalahan saat menampilkan formulir'. Melewati item ini seperti penanganan 404...");
                const lastRowId = GM_getValue('asik_ckg_last_clicked_row', '') || sessionStorage.getItem('asik_ckg_last_clicked_row') || '';
                if (lastRowId) {
                    const dilewati = new Set(GM_getValue('asik_ckg_dilewati', []));
                    dilewati.add(lastRowId);
                    GM_setValue('asik_ckg_dilewati', Array.from(dilewati));

                    const attemptCounts = GM_getValue('asik_ckg_attempts', {});
                    attemptCounts[lastRowId] = 99; // Tandai sudah dicoba maksimal agar di-skip di daftar
                    GM_setValue('asik_ckg_attempts', attemptCounts);

                    console.warn(`[CKG Error Form] Item "${lastRowId}" ditandai BLACKLIST (Dilewati).`);
                }
                if (currentData) {
                    catatRiwayat(currentData.NAMA, currentData.KELAS, 'Dilewati (Error Form)', `Terjadi kesalahan saat menampilkan formulir: ${lastRowId || 'Tatalaksana'}`);
                }

                if (lastRowId === 'tatalaksana_utama' || state.isTatalaksana) {
                    const failCount = (GM_getValue('asik_ckg_tatalaksana_fail_count', 0) || 0) + 1;
                    GM_setValue('asik_ckg_tatalaksana_fail_count', failCount);
                    if (failCount >= 2 || lastRowId === 'tatalaksana_utama') {
                        console.warn("[CKG] Tatalaksana utama gagal menampilkan formulir. Menandai tatalaksana selesai untuk siswa ini.");
                        state.tatalaksanaSelesai = true;
                        GM_setValue('asik_ckg_tatalaksana_done', true);
                        state.isTatalaksana = false;
                    }
                }

                state.phase = 'DETAIL_PELAYANAN';
                setState(state);

                // Coba cari dan klik tombol "Kembali" / "Tutup" / "OK" jika ada di modal atau halaman error
                const btnKembali = Array.from(document.querySelectorAll('button, div[class*="cursor-pointer"], a, span')).find(b => {
                    if (b.closest('#asik-ckg-panel, #asik-dashboard-overlay')) return false;
                    if (!isElementVisible(b)) return false;
                    const t = bersih(b.textContent);
                    return t === 'kembali' || t.includes('kembali ke halaman') || t === 'tutup' || t === 'ok' || t === 'mengerti';
                });

                if (btnKembali) {
                    const targetBtn = btnKembali.closest('button') || btnKembali;
                    console.log("[CKG] Mengklik tombol kembali pada tampilan error formulir:", targetBtn.textContent.trim());
                    triggerClick(targetBtn);
                    if (typeof targetBtn.click === 'function') {
                        try { targetBtn.click(); } catch (e) { }
                    }
                    await sleep(1000);
                }

                console.log("[CKG] Menjalankan window.history.back() untuk kembali ke halaman sebelumnya...");
                window.history.back();
                await sleep(2500);
                return;
            }

            // ============================================================
            // 0.1 CEK POPUP PERINGATAN / "DATA PEMERIKSAAN SEDANG DIPROSES" / ERROR
            // ============================================================
            const diprosesHandled = await tanganiModalSedangDiproses();
            if (diprosesHandled) {
                console.log("[CKG] ✅ Popup 'Data pemeriksaan sedang diproses' ditutup pada awal siklus.");
                await sleep(1000);
                return;
            }

            const allModals = Array.from(document.querySelectorAll('div[role="dialog"], .modal, .swal2-container, .popup, div.fixed, div[class*="fixed"][class*="z-"]'))
                .filter(el => !el.closest('#asik-ckg-panel, #asik-dashboard-overlay'));

            const popupPeringatan = allModals.find(el => {
                const t = bersih(el.textContent);
                return t.includes('gagal perbarui data') || t.includes('gagal memperbarui') || t.includes('perbarui data') ||
                    t.includes('terjadi kesalahan') || t.includes('gagal') || t.includes('no data') || t.includes('tidak ada data');
            });

            if (popupPeringatan) {
                console.warn("[CKG] ⚠️ Popup Peringatan / Gagal Perbarui Data terdeteksi:", popupPeringatan.textContent.trim().substring(0, 70));
                const btnOk = Array.from(popupPeringatan.querySelectorAll('button, div[class*="cursor-pointer"], .btn-fill-primary, div.btn-fill-primary, div.btn-outline-primary, a, span')).find(b => {
                    if (!isElementVisible(b)) return false;
                    const t = bersih(b.textContent);
                    return t === 'ok' || t === 'tutup' || t === 'mengerti' || t === 'kembali' || t === 'batal' || t === 'ya' || t.includes('ok') || t.includes('tutup') || t.includes('mengerti');
                });
                if (btnOk) {
                    const targetBtn = btnOk.closest('button') || btnOk;
                    console.log("[CKG] Mengklik tombol OK/Tutup pada popup peringatan:", targetBtn.textContent.trim());
                    triggerClick(targetBtn);
                    if (typeof targetBtn.click === 'function') {
                        try { targetBtn.click(); } catch (e) { }
                    }
                    await sleep(1500);
                    return;
                }
            }

            // ============================================================
            // DETEKSI KONDISI HALAMAN SECARA OTOMATIS
            // ============================================================
            const adaFormSurvey = document.querySelector('.sd-question, .sv-question') !== null;
            const hasTableLayanan = html.includes('tablelayanan') || html.includes('pelayanan oleh nakes') || document.querySelector('#tableLayanan') !== null;
            const hasTataText = (html.includes('detail tatalaksana') || html.includes('table-ckg-tatalaksana') || html.includes('terakhir tatalaksana:') || html.includes('belum tatalaksana')) && !hasTableLayanan;

            const adaHalamanTatalaksana = !adaFormSurvey && hasTataText;
            const adaDetailPelayanan = !adaFormSurvey && hasTableLayanan && !adaHalamanTatalaksana;

            const adaFilterSekolah = Array.from(document.querySelectorAll('span, div')).some(el => {
                const t = bersih(el.textContent);
                return t === 'pilih sekolah' || t === 'pilih kelas' || t === 'pelayanan ckg sekolah' || t === 'tampilkan pencarian';
            }) || document.querySelector('#sekolah') !== null;

            const adaTabelSiswa = (document.querySelector('input#searchNik') !== null ||
                document.querySelector('input[placeholder*="cari"], input[placeholder*="nik"], input[placeholder*="tiket"], input[placeholder*="nama"], input[placeholder*="Masukkan"]') !== null ||
                document.querySelector('table tbody tr') !== null) && !adaDetailPelayanan && !adaHalamanTatalaksana && !adaFormSurvey;

            console.log(`[CKG] Deteksi: Form=${adaFormSurvey}, Tatalaksana=${adaHalamanTatalaksana}, Detail=${adaDetailPelayanan}, Tabel=${adaTabelSiswa}, FilterSekolah=${adaFilterSekolah}, Phase=${state.phase}`);

            // ============================================================
            // 1. HALAMAN FORM SURVEYJS (Sedang mengisi kuesioner layanan / tatalaksana)
            // ============================================================
            if (adaFormSurvey) {
                console.log("[CKG] >>> HALAMAN FORM SURVEYJS — Mengisi form untuk:", currentData.NAMA);
                state.phase = 'MENGISI';
                setState(state);

                await isiFormPelayanan(currentData);
                await sleep(500);

                // Pastikan tidak ada input yang masih fokus & klik area kosong agar SurveyJS memvalidasi form dan memunculkan tombol Kirim
                if (document.activeElement && typeof document.activeElement.blur === 'function') {
                    try { document.activeElement.blur(); } catch (e) { }
                }
                try {
                    document.body.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
                } catch (e) { }

                // Scroll ke bagian paling bawah form agar tombol Kirim terlihat jelas
                window.scrollTo({ top: document.body.scrollHeight, behavior: 'smooth' });
                await sleep(400);

                // Cari tombol simpan / kirim form SurveyJS
                const candidateButtons = Array.from(document.querySelectorAll('input.sd-navigation__complete-btn, button.sd-navigation__complete-btn, input[type="button"][value="Kirim"], input[value="Kirim"], input[title="Kirim"], input[value="Simpan"], button[title="Kirim"], .sd-navigation__complete-btn, .sd-btn--action, input.sd-btn, button.sd-btn')).filter(el => {
                    if (el.disabled || el.closest('#asik-ckg-panel, #asik-dashboard-overlay')) return false;
                    const t = bersih(el.textContent || el.value || el.title || '');
                    return t === 'simpan' || t === 'kirim' || t === 'selesai' || t === 'complete' || t === 'submit' || t.includes('kirim') || t.includes('simpan');
                });

                const btnSimpan = candidateButtons[0] || null;

                if (btnSimpan) {
                    console.log("[CKG] 🚀 Menekan tombol Kirim Form...", btnSimpan.value || btnSimpan.textContent.trim());
                    btnSimpan.scrollIntoView({ behavior: 'smooth', block: 'center' });
                    await sleep(300);
                    triggerClick(btnSimpan);
                    if (typeof btnSimpan.click === 'function') {
                        try { btnSimpan.click(); } catch (e) { }
                    }
                    state.phase = 'MENYIMPAN';
                    setState(state);
                    await sleep(2500); // Jeda transisi setelah kirim agar server memproses
                } else {
                    console.warn("[CKG] Tombol Kirim Form SurveyJS belum ditemukan pada DOM.");
                }
                return;

                // ============================================================
                // 2. FASE MENYIMPAN (Jeda transisi server & kembali ke halaman sebelumnya)
                // ============================================================
            } else if (state.phase === 'MENYIMPAN') {
                console.log("[CKG] >>> FASE MENYIMPAN — Menangani transisi setelah kirim...");

                // 1. Jika ada popup konfirmasi / notif, klik OK/Tutup
                const btnTutup = Array.from(document.querySelectorAll('button, div[role="dialog"] button, .modal button')).find(el => {
                    if (!isElementVisible(el) || el.closest('#asik-ckg-panel, #asik-dashboard-overlay')) return false;
                    const t = bersih(el.textContent);
                    return t === 'tutup' || t === 'ok' || t === 'selesai' || t === 'lanjut';
                });
                if (btnTutup) {
                    triggerClick(btnTutup);
                    await sleep(800);
                }

                // 2. Jika muncul tombol 'Kembali ke Halaman Utama' setelah form terkirim
                const btnKembaliUtama = Array.from(document.querySelectorAll('button')).find(b => {
                    const t = bersih(b.textContent);
                    return (t.includes('kembali ke halaman utama') || t.includes('halaman utama')) && isElementVisible(b);
                });
                if (btnKembaliUtama) {
                    console.log("[CKG] Menekan tombol Kembali ke Halaman Utama...");
                    triggerClick(btnKembaliUtama);
                    await sleep(1500);
                }

                // 3. Kembali ke alur pemeriksaan/tatalaksana
                console.log("[CKG] Kembali ke daftar layanan/tatalaksana pasien...");
                state.phase = 'DETAIL_PELAYANAN';
                setState(state);
                await sleep(1500);
                return;

                // ============================================================
                // 3. HALAMAN DAFTAR TATALAKSANA PASIEN
                // ============================================================
            } else if (adaHalamanTatalaksana) {
                console.log("[CKG] >>> HALAMAN DAFTAR TATALAKSANA");

                const skippedList = GM_getValue('asik_ckg_dilewati', []);
                const attemptCounts = GM_getValue('asik_ckg_attempts', {});

                // 1. Cari baris tatalaksana yang belum selesai dan BELUM pernah dicoba / dilewati
                const rows = Array.from(document.querySelectorAll('tbody tr, .grid-cols-5, div.py-4, tr')).filter(r => !r.closest('#asik-ckg-panel, #asik-dashboard-overlay'));

                let targetBtn = null;
                let targetRowId = null;
                for (const row of rows) {
                    const teks = bersih(row.textContent);
                    const rowId = teks.replace(/[^a-z0-9]/gi, '').substring(0, 35);

                    // Cek apakah baris ini sudah selesai, sudah dicoba >= 2 kali, atau ada di daftar 404
                    const isDilewati = skippedList.some(k => teks.includes(bersih(k)) || bersih(k).includes(teks.substring(0, 15)));
                    const sudahDicoba = (attemptCounts[rowId] || 0) >= 2;

                    if (teks.includes('selesai tatalaksana') || teks.includes('selesai diperiksa') || isDilewati || sudahDicoba) {
                        console.log(`[CKG] Melewati baris tatalaksana (Selesai/404/Pernah Dicoba): "${teks.substring(0, 30)}..."`);
                        continue;
                    }

                    const btnMatches = Array.from(row.querySelectorAll('button, div, a, span')).filter(b => {
                        if (!isElementVisible(b) || b.disabled || b.classList.contains('cursor-not-allowed') || b.classList.contains('bg-disabled')) return false;
                        const t = bersih(b.textContent);
                        return t === 'mulai tatalaksana' || t.includes('mulai tatalaksana') || t === 'input data' || t === 'lanjutkan' || t === 'sedang tatalaksana' || t === 'isi data';
                    });
                    const btn = btnMatches.find(b => b.tagName === 'BUTTON') || (btnMatches.length > 0 ? btnMatches[btnMatches.length - 1] : null);
                    if (btn) {
                        row.scrollIntoView({ behavior: 'smooth', block: 'center' });
                        await sleep(300);
                        targetBtn = btn.closest('button') || btn;
                        targetRowId = rowId;
                        break;
                    }
                }

                if (targetBtn && targetRowId) {
                    console.log(`[CKG] Menemukan item Tatalaksana yang belum diisi: "${targetRowId}". Mengklik tombol...`);
                    attemptCounts[targetRowId] = (attemptCounts[targetRowId] || 0) + 1;
                    GM_setValue('asik_ckg_attempts', attemptCounts);
                    GM_setValue('asik_ckg_last_clicked_row', targetRowId);

                    state.phase = 'MENGISI';
                    setState(state);
                    triggerClick(targetBtn);
                    if (typeof targetBtn.click === 'function') {
                        try { targetBtn.click(); } catch (e) { }
                    }
                    await sleep(2000);
                    return;
                }

                // 2. Jika seluruh tatalaksana telah selesai atau dilewati (404), klik tombol Kembali di header
                console.log("[CKG] Seluruh item Tatalaksana selesai / dilewati! Mengklik tombol Back Header Tatalaksana...");
                state.tatalaksanaSelesai = true;
                GM_setValue('asik_ckg_tatalaksana_done', true);
                setState(state);

                const btnBack = document.querySelector('button div.text-lg.font-bold, div.text-lg.font-bold')?.closest('button') ||
                    document.querySelector('img[src*="arrow-left"], [class*="arrow-left"]')?.closest('button, div.cursor-pointer');
                if (btnBack) {
                    console.log("[CKG] Menemukan tombol Back Header Tatalaksana! Mengklik...");
                    triggerClick(btnBack);
                    await sleep(1500);
                } else {
                    console.log("[CKG] Tombol Back Header tidak ditemukan, menggunakan history.back()...");
                    window.history.back();
                    await sleep(1500);
                }
                state.phase = 'DETAIL_PELAYANAN';
                setState(state);
                return;

                // ============================================================
                // 3. HALAMAN DETAIL SISWA: "Pemeriksaan Mandiri" & "Pelayanan oleh Nakes"
                // ============================================================
            } else if (adaDetailPelayanan || state.phase === 'DETAIL_PELAYANAN') {
                console.log("[CKG] >>> HALAMAN DETAIL PEMERIKSAAN SISWA (Mandiri & Nakes)");

                // Tangani popup 'Data pemeriksaan sedang diproses' jika masih ada di halaman
                await tanganiModalSedangDiproses();

                // Deteksi nama siswa dari halaman Detail Pemeriksaan untuk sinkronisasi otomatis ke Excel
                const studentNameEl = document.querySelector('.text-\\[18px\\].font-bold, div[class*="text-[18px]"].font-bold');
                if (studentNameEl) {
                    const studentName = studentNameEl.textContent.trim();
                    if (studentName) {
                        sessionStorage.setItem('asik_ckg_active_student_name', studentName);
                        let wbData = null;
                        try {
                            const raw = GM_getValue('asik_ckg_workbook_data') || localStorage.getItem('asik_ckg_workbook_data');
                            if (raw) wbData = typeof raw === 'string' ? JSON.parse(raw) : raw;
                        } catch (e) { }
                        if (wbData && wbData.sheets) {
                            const matched = findStudentInWorkbook(studentName, wbData);
                            if (matched && matched.NAMA !== 'Siswa') {
                                state.sweepCurrentStudent = matched;
                                sessionStorage.setItem('asik_ckg_active_student', JSON.stringify(matched));
                                setState(state);
                                console.log(`[CKG] 👤 Siswa terdeteksi di halaman detail: "${matched.NAMA}" (GDS: ${matched.GDS || 'kosong'}, HB: ${matched.HB || 'kosong'})`);
                            }
                        }
                    }
                }
                // 1. Pastikan seluruh accordion di #tableLayanan atau area konten terbuka
                const tableLayanan = document.querySelector('#tableLayanan') || document.querySelector('main') || document.body;
                if (tableLayanan) {
                    const accordions = Array.from(tableLayanan.querySelectorAll('button[aria-controls="dropdown-content"], button')).filter(b => {
                        if (b.closest('aside, nav, #sidebar, [class*="sidebar"], #asik-ckg-panel, #asik-dashboard-overlay')) return false;
                        const t = bersih(b.textContent || '');
                        // Jangan skip accordion judul seperti "Pemeriksaan Kadar CO (Tatalaksana Merokok)"
                        if (t.includes('selesaikan layanan') || t === 'mulai tatalaksana' || t.includes('kirim rapor') || t.includes('kirim rapot') || t === 'input data' || t === 'lanjutkan' || t === 'ubah data') return false;
                        const svg = b.querySelector('svg');
                        return svg && !(svg.getAttribute('class') || '').includes('rotate-180');
                    });
                    for (const acc of accordions) {
                        triggerClick(acc);
                        await sleep(150);
                    }
                }

                // 1B. Cek status keseluruhan Pemeriksaan Mandiri dari header "Jumlah Pemeriksaan (X/Y)"
                let allMandiriDone = false;
                const allHeaderCandidates = Array.from(document.querySelectorAll('div, span, p, h1, h2, h3, h4, button'));
                const mandiriHeader = allHeaderCandidates.find(el => {
                    const txt = el.textContent || '';
                    return txt.includes('Jumlah Pemeriksaan') && /\(\s*\d+\s*\/\s*\d+\s*\)/.test(txt);
                });
                if (mandiriHeader) {
                    const match = mandiriHeader.textContent.match(/Jumlah Pemeriksaan\s*\(\s*(\d+)\s*\/\s*(\d+)\s*\)/i);
                    if (match) {
                        const cur = parseInt(match[1], 10);
                        const tot = parseInt(match[2], 10);
                        if (tot > 0 && cur >= tot) {
                            allMandiriDone = true;
                            console.log(`[CKG] ✅ Seluruh Pemeriksaan Mandiri sudah lengkap (${cur}/${tot}).`);
                        }
                    }
                }

                // 2. CARI SEMUA BARIS LAYANAN YANG ADA
                const rawRows = Array.from(document.querySelectorAll('.table-pemeriksaan-mandiri tbody tr, .table-pemeriksaan-mandiri tr, #tableLayanan .grid-cols-5, #tableLayanan .grid-cols-4, div.grid-cols-5, div.grid-cols-4, tbody tr'));
                const allRows = [];
                const seenRows = new Set();
                for (const r of rawRows) {
                    if (r.closest('aside, nav, #sidebar, [class*="sidebar"], #asik-ckg-panel, #asik-dashboard-overlay')) continue;
                    if (r.closest('thead') || r.querySelector('th')) continue;
                    if (seenRows.has(r)) continue;
                    seenRows.add(r);
                    allRows.push(r);
                }

                // ============================================================
                // 2. PROSES BARIS LAYANAN SECARA MENGALIR (IN-LINE DARI ATAS KE BAWAH)
                // - Jika baris form biasa -> Klik "Input Data" dan isi form
                // - Jika baris layanan wajib matikan (misal Kadar CO) -> Matikan toggle switch & konfirmasi popup
                // ============================================================
                let pendingRow = null;
                let pendingBtn = null;
                let pendingServiceKey = '';
                const serviceAttempts = GM_getValue('asik_ckg_service_attempts', {});

                for (const r of allRows) {
                    // Jika seluruh Pemeriksaan Mandiri sudah selesai (misal 7/7), lewati baris di tabel mandiri
                    if (allMandiriDone && r.closest('.table-pemeriksaan-mandiri')) {
                        continue;
                    }

                    // Periksa apakah layanan pada baris ini sudah selesai/lengkap (termasuk status "Tidak diperiksa")
                    if (layananSelesai(r)) {
                        continue;
                    }

                    const teks = bersih(r.textContent || '');
                    if (!teks) continue;

                    const sName = teks.split('\n')[0].replace(/input data|mulai|lanjutkan|isi data|sedang pemeriksaan/gi, '').trim().substring(0, 45);
                    const sKey = sName.toLowerCase();

                    // KASUS A: Layanan wajib dilewati (seperti Kadar CO, RDT Malaria, Pemeriksaan Hepatitis Nakes)
                    // Sejak update Sep 2026: toggle switch sudah dihapus, cukup SKIP baris ini tanpa klik apa pun
                    if (layananWajibLewati(teks)) {
                        console.log(`[CKG] ⏭️ Baris layanan "${sName}" wajib dilewati (skip tanpa isi). Melanjutkan ke baris berikutnya...`);
                        continue;
                    }

                    // KASUS B: Layanan biasa yang harus diisi ("Input Data", "Mulai", "Lanjutkan")
                    if (sKey && (serviceAttempts[sKey] || 0) >= 2) {
                        console.warn(`[CKG] ⚠️ Layanan "${sName}" sudah dicoba 2x dan belum selesai/tersangkut. Melewati baris ini agar tidak stuck loop...`);
                        continue;
                    }

                    // Cari tombol aksi pengisian aktif di baris ini
                    const actionButtons = Array.from(r.querySelectorAll('button, div, a, span')).filter(b => {
                        if (!isElementVisible(b) || b.disabled || b.classList.contains('cursor-not-allowed') || b.classList.contains('bg-disabled')) return false;
                        const bt = bersih(b.textContent || b.getAttribute('title') || b.getAttribute('aria-label') || '');
                        // Jangan ambil tombol umum di luar form layanan
                        if (bt.includes('tatalaksana') || bt.includes('selesaikan') || bt.includes('rapor') || bt.includes('rapot')) return false;
                        return bt === 'input data' || bt === 'mulai' || bt === 'lanjutkan' || bt === 'isi data' || bt === 'sedang pemeriksaan' ||
                            bt.includes('input data') || bt.includes('lanjutkan');
                    });

                    const btn = actionButtons.find(b => b.tagName === 'BUTTON') || actionButtons[actionButtons.length - 1];
                    if (btn) {
                        pendingRow = r;
                        pendingBtn = btn.closest('button') || btn;
                        pendingServiceKey = sKey;
                        break;
                    }
                }

                // JIKA ADA LAYANAN BELUM LENGKAP: BUKA FORM DAN ISI DENGAN NILAI DEFAULT / EXCEL
                if (pendingRow && pendingBtn) {
                    if (pendingServiceKey) {
                        serviceAttempts[pendingServiceKey] = (serviceAttempts[pendingServiceKey] || 0) + 1;
                        GM_setValue('asik_ckg_service_attempts', serviceAttempts);
                    }
                    const rowText = bersih(pendingRow.textContent || '').substring(0, 40);
                    console.log(`[CKG] 📝 Menemukan layanan yang belum selesai: "${rowText}...". Mengklik tombol "${bersih(pendingBtn.textContent)}"...`);
                    pendingRow.scrollIntoView({ behavior: 'smooth', block: 'center' });
                    await sleep(300);
                    state.phase = 'MENGISI';
                    setState(state);
                    triggerClick(pendingBtn);
                    if (typeof pendingBtn.click === 'function') {
                        try { pendingBtn.click(); } catch (e) { }
                    }
                    await sleep(2000);
                    return;
                }

                // 3. KIRIM RAPOR (Hanya jika masih berstatus 0/3 dan belum dikirim)
                const raporDone = state.raporTerkirim || GM_getValue('asik_ckg_rapor_sent', false);
                if (!raporDone) {
                    const allRaporEls = Array.from(document.querySelectorAll('.btn-sent-report, div[class*="btn-sent-report"], button[class*="btn-sent-report"], div, button, a'))
                        .filter(el => {
                            if (el.closest('aside, nav, #sidebar, [class*="sidebar"], #asik-ckg-panel, #asik-dashboard-overlay')) return false;
                            if (!isElementVisible(el)) return false;
                            const t = bersih(el.textContent);
                            return t.includes('kirim rapor') || t.includes('kirim rapot');
                        });

                    const btnRapor = allRaporEls.find(el => (el.className || '').includes('btn-sent-report')) ||
                        allRaporEls.find(el => el.tagName === 'BUTTON') ||
                        allRaporEls[allRaporEls.length - 1];

                    if (btnRapor) {
                        const teksBtn = bersih(btnRapor.textContent);
                        const sudahTerkirim = teksBtn.includes('(1/3)') || teksBtn.includes('(2/3)') || teksBtn.includes('(3/3)');

                        if (sudahTerkirim) {
                            console.log(`[CKG] 📄 Tombol Kirim Rapor sudah berstatus "${teksBtn}". Tidak mengklik lagi.`);
                            state.raporTerkirim = true;
                            GM_setValue('asik_ckg_rapor_sent', true);
                            setState(state);
                        } else {
                            console.log(`[CKG] 📄 Menemukan tombol Kirim Rapor: "${teksBtn}". Mengklik untuk mengirim rapor...`);
                            btnRapor.scrollIntoView({ behavior: 'smooth', block: 'center' });
                            await sleep(400);
                            triggerClick(btnRapor);
                            if (typeof btnRapor.click === 'function') {
                                try { btnRapor.click(); } catch (e) { }
                            }
                            await sleep(1000);
                            await tanganiModalKirimRapor();
                            // Polling tunggu konfirmasi popup "Data pemeriksaan sedang diproses"
                            for (let p = 0; p < 8; p++) {
                                const closed = await tanganiModalSedangDiproses();
                                if (closed) {
                                    console.log("[CKG] ✅ Popup 'Data pemeriksaan sedang diproses' berhasil ditutup setelah Kirim Rapor.");
                                    break;
                                }
                                await sleep(600);
                            }
                            state.raporTerkirim = true;
                            GM_setValue('asik_ckg_rapor_sent', true);
                            setState(state);
                            await sleep(1500);
                            return;
                        }
                    } else {
                        console.log("[CKG] Tombol Kirim Rapor tidak ditemukan di halaman. Melanjutkan ke tahap berikutnya...");
                        state.raporTerkirim = true;
                        GM_setValue('asik_ckg_rapor_sent', true);
                        setState(state);
                    }
                }

                // 4. MULAI TATALAKSANA
                const tatalaksanaDone = state.tatalaksanaSelesai || GM_getValue('asik_ckg_tatalaksana_done', false);
                if (!tatalaksanaDone) {
                    const btnTatalaksana = Array.from(document.querySelectorAll('button.btn-fill-primary, button')).find(b => {
                        if (b.closest('aside, nav, #sidebar, [class*="sidebar"], #asik-ckg-panel, #asik-dashboard-overlay')) return false;
                        const text = bersih(b.innerText || b.textContent || '');
                        return (text === 'mulai' || text === 'mulai tatalaksana' || text.includes('mulai tatalaksana')) &&
                            !b.disabled &&
                            !b.classList.contains('cursor-not-allowed') &&
                            !b.classList.contains('bg-disabled') &&
                            isElementVisible(b);
                    });

                    if (btnTatalaksana) {
                        console.log("[CKG] 🩺 Ditemukan tombol Mulai Tatalaksana! Mengklik untuk masuk ke form Tatalaksana...");
                        btnTatalaksana.scrollIntoView({ behavior: 'smooth', block: 'center' });
                        await sleep(400);
                        state.isTatalaksana = true;
                        state.phase = 'MENGISI';
                        GM_setValue('asik_ckg_last_clicked_row', 'tatalaksana_utama');
                        setState(state);
                        triggerClick(btnTatalaksana);
                        if (typeof btnTatalaksana.click === 'function') {
                            try { btnTatalaksana.click(); } catch (e) { }
                        }
                        await sleep(2000);
                        return;
                    } else {
                        console.log("[CKG] Tombol Mulai Tatalaksana tidak aktif/tidak ada. Menandai tatalaksana selesai...");
                        state.tatalaksanaSelesai = true;
                        GM_setValue('asik_ckg_tatalaksana_done', true);
                        setState(state);
                        await sleep(400);
                    }
                }

                // 5. SELESAIKAN LAYANAN (SEBELUM KEMBALI)
                console.log("[CKG] 🏁 Menekan tombol Selesaikan Layanan...");
                const allSelesaikan = Array.from(document.querySelectorAll('button.btn-outline-error, .btn-outline-error, button[class*="btn-outline-error"], button, div, a')).filter(el => {
                    if (el.closest('aside, nav, #sidebar, [class*="sidebar"], #asik-ckg-panel, #asik-dashboard-overlay')) return false;
                    if (!isElementVisible(el)) return false;
                    const t = bersih(el.textContent);
                    return t === 'selesaikan layanan' || t.includes('selesaikan layanan');
                });

                const btnSelesaikan = allSelesaikan.find(el => (el.className || '').includes('btn-outline-error')) ||
                    allSelesaikan.find(el => el.tagName === 'BUTTON') ||
                    allSelesaikan[0];

                if (btnSelesaikan) {
                    const targetBtn = btnSelesaikan.closest('button') || btnSelesaikan;
                    console.log("[CKG] Mengklik tombol Selesaikan Layanan:", targetBtn.textContent.trim());
                    targetBtn.scrollIntoView({ behavior: 'smooth', block: 'center' });
                    await sleep(400);
                    triggerClick(targetBtn);
                    if (typeof targetBtn.click === 'function') {
                        try { targetBtn.click(); } catch (e) { }
                    }
                    await sleep(1000);
                    await tanganiModalSelesaikanLayanan();
                    await sleep(2000);
                }

                // 6. SELESAIKAN PROGRESS SISWA DAN KEMBALI KE PENCARIAN
                catatRiwayat(currentData.NAMA, currentData.KELAS, 'Selesai', 'Layanan, Rapot, Tatalaksana & Selesaikan Layanan Berhasil');
                sessionStorage.removeItem('asik_ckg_dilewati');
                sessionStorage.removeItem('asik_ckg_last_clicked_row');
                sessionStorage.removeItem('asik_ckg_active_student');
                sessionStorage.removeItem('asik_ckg_active_student_name');
                state.sweepCurrentStudent = null;
                GM_setValue('asik_ckg_dilewati', []);
                GM_setValue('asik_ckg_attempts', {});
                GM_setValue('asik_ckg_last_clicked_row', '');
                GM_setValue('asik_ckg_tatalaksana_fail_count', 0);
                GM_setValue('asik_ckg_tatalaksana_done', false);
                GM_setValue('asik_ckg_rapor_sent', false);
                GM_setValue('asik_ckg_service_attempts', {});
                state.tatalaksanaSelesai = false;
                state.raporTerkirim = false;
                state.isTatalaksana = false;

                // Kembali 1x ke halaman pencarian siswa
                const searchInput = document.querySelector('input#searchNik') || document.querySelector('input[placeholder*="cari"], input[placeholder*="nik"], input[placeholder*="tiket"]');
                if (!searchInput) {
                    console.log("[CKG] Menekan tombol kembali 1x ke halaman pencarian siswa...");
                    const backBtn = Array.from(document.querySelectorAll('img.cursor-pointer, .cursor-pointer, button')).find(el => {
                        if (!isElementVisible(el) || el.closest('#asik-ckg-panel, #asik-dashboard-overlay')) return false;
                        const cls = el.className || '';
                        const src = el.src || '';
                        return (cls.includes('w-[24px]') && cls.includes('h-[24px]')) || src.includes('icon-arrow-left') || src.includes('arrow-left');
                    });
                    if (backBtn) {
                        triggerClick(backBtn);
                    } else {
                        window.history.back();
                    }
                    await sleep(2000);
                }

                if (state.sweepMode) {
                    console.log(`[CKG Sweep] ✅ Siswa "${currentData.NAMA}" selesai diproses dalam mode penyisiran!`);
                    const sweepAttempts = GM_getValue('asik_ckg_sweep_attempts', {});
                    delete sweepAttempts[currentData.NAMA];
                    GM_setValue('asik_ckg_sweep_attempts', sweepAttempts);
                    state.sweepCurrentStudent = null;
                    state.phase = 'MENCARI';
                    state.needsClassSearch = true; // Wajib klik 'Tampilkan Pencarian' lagi setelah kembali
                    setState(state);
                } else {
                    state.index += 1;
                    state.phase = 'MENCARI';
                    state.needsClassSearch = true; // Wajib klik 'Tampilkan Pencarian' lagi setelah kembali
                    setState(state);
                }
                return;

                // ============================================================
                // 4. HALAMAN TABEL PENCARIAN SISWA (Daftar Kelas)
                // ============================================================
            } else if (adaTabelSiswa || state.phase === 'MENCARI') {
                state.phase = 'MENCARI';
                setState(state);

                // ============================================================
                // 4A. MODE PENYISIRAN (SWEEP TAB SEDANG PEMERIKSAAN)
                // ============================================================
                if (state.sweepMode) {
                    let wbData = null;
                    try {
                        const raw = GM_getValue('asik_ckg_workbook_data') || localStorage.getItem('asik_ckg_workbook_data');
                        if (raw) wbData = typeof raw === 'string' ? JSON.parse(raw) : raw;
                    } catch (e) { }

                    const sheetNames = (wbData && wbData.sheetNames && wbData.sheetNames.length > 0) ? wbData.sheetNames : ['1'];
                    const sweepIdx = state.sweepSheetIndex || 0;

                    // Jika seluruh sheet/kelas telah disapu
                    if (sweepIdx >= sheetNames.length) {
                        console.log("[CKG Sweep] 🏆 Seluruh kelas telah selesai disapu tab Sedang Pemeriksaan!");
                        catatRiwayat('Penyisiran Selesai', '-', 'Selesai', 'Seluruh tab Sedang Pemeriksaan di semua kelas telah selesai dikerjakan.');
                        alert("✅ SELESAI SEMUA!\n\nSeluruh antrean siswa dan data di tab 'Sedang Pemeriksaan' telah berhasil dikerjakan.");
                        setState({ ...state, running: false, sweepMode: false, sweepCurrentStudent: null });
                        return;
                    }

                    const currentSheetName = sheetNames[sweepIdx];
                    const targetKelas = extractClassNum(currentSheetName) || currentSheetName;
                    const targetNum = extractClassNum(targetKelas);
                    console.log(`[CKG Sweep] >>> MODE PENYISIRAN TAB SEDANG PEMERIKSAAN — Kelas ${targetKelas} (${currentSheetName}) [${sweepIdx + 1}/${sheetNames.length}]`);

                    // 1. Pilih Kelas Sweep jika belum sesuai
                    const mainArea = document.querySelector('main, section, div[class*="min-h-screen"]') || document.body;
                    const allMainTriggers = Array.from(mainArea.querySelectorAll('div[class*="cursor-pointer"], div[class*="h-[2.9rem]"]'))
                        .filter(el => {
                            if (el.closest('#asik-ckg-panel, #asik-dashboard-overlay, aside, nav, #sidebar, [class*="sidebar"]')) return false;
                            if (!isElementVisible(el)) return false;
                            const span = el.querySelector('span');
                            if (!span) return false;
                            const t = bersih(span.textContent);
                            if (!t || t === 'ckg umum' || t === 'ckg sekolah' || t === 'tatalaksana' || t === 'pelayanan') return false;
                            return true;
                        });

                    const kelasTrigger = allMainTriggers[1] || allMainTriggers.find(el => {
                        const t = bersih(el.querySelector('span')?.textContent || el.textContent);
                        return t === 'pilih kelas' || t.includes('kelas') || /^(sd|smp|sma|smk)\s+kelas/i.test(t);
                    });
                    const kelasSpan = kelasTrigger ? (kelasTrigger.querySelector('span') || kelasTrigger) : null;
                    const currentSelectedText = kelasSpan ? bersih(kelasSpan.textContent) : '';
                    const activeAsik = detectActiveClassOnAsikPage();
                    const currentNum = activeAsik.num || extractClassNum(currentSelectedText);

                    const lastSearchedKelas = GM_getValue('asik_ckg_last_kelas', '') || state.lastSearchedKelas || '';
                    const lastNum = extractClassNum(lastSearchedKelas);

                    const inputSearchReady = document.querySelector('input#searchNik') !== null ||
                        document.querySelector('input[placeholder*="cari"], input[placeholder*="nik"], input[placeholder*="tiket"], input[placeholder*="nama"], input[placeholder*="Masukkan"]') !== null;

                    let mustSwitchClass = false;
                    if (currentSelectedText === 'pilih kelas' || !currentNum) {
                        mustSwitchClass = true;
                    } else if (targetNum && currentNum && targetNum !== currentNum) {
                        mustSwitchClass = true;
                    } else if (state.needsClassSearch || !inputSearchReady) {
                        // Butuh re-trigger Tampilkan Pencarian (setelah kembali dari detail siswa atau input belum muncul)
                        mustSwitchClass = true;
                    }

                    if (mustSwitchClass) {
                        console.log(`[CKG Sweep] 🔄 Ganti Kelas Sweep: Memilih Kelas ${targetKelas} (Saat ini: "${currentSelectedText || 'Belum dipilih'}")...`);
                        if (currentSelectedText === 'pilih kelas' || (targetNum && currentNum && targetNum !== currentNum)) {
                            const kelasOk = await pilihDropdownASIK('kelas', targetKelas);
                            if (!kelasOk) {
                                console.warn(`[CKG Sweep] Gagal memilih kelas ${targetKelas}. Mencoba lagi iterasi berikutnya...`);
                                return;
                            }
                            await sleep(1000);
                        }

                        console.log("[CKG Sweep] 🔘 Mengklik tombol 'Tampilkan Pencarian' untuk memuat kelas sweep...");
                        const searchOk = await klikTombolTampilkanPencarian();
                        if (searchOk) {
                            state.needsClassSearch = false;
                            state.lastSearchedKelas = targetKelas;
                            state.searchRetry = 0;
                            state.sweepEmptyCheckCount = 0;
                            GM_setValue('asik_ckg_last_kelas', targetKelas);
                            setState(state);
                        } else {
                            state.needsClassSearch = true;
                            setState(state);
                        }
                        return;
                    } else {
                        if (!lastSearchedKelas || lastSearchedKelas !== targetKelas || state.needsClassSearch) {
                            state.lastSearchedKelas = targetKelas;
                            state.needsClassSearch = false;
                            GM_setValue('asik_ckg_last_kelas', targetKelas);
                            setState(state);
                        }
                    }

                    // 2. Pastikan berpindah ke tab "Sedang Pemeriksaan"
                    const allTabs = Array.from(document.querySelectorAll('div.cursor-pointer, div, span')).filter(el => {
                        return !el.closest('#asik-ckg-panel, #asik-dashboard-overlay, [id*="asik-ckg"]') && isElementVisible(el);
                    });

                    const tabSedang = allTabs.find(el => {
                        const t = bersih(el.textContent);
                        return t.startsWith('sedang pemeriksaan');
                    });

                    if (tabSedang) {
                        const cls = tabSedang.className || '';
                        const isActive = cls.includes('border-b-[#16b3ac]') ||
                            cls.includes('text-teal-500') ||
                            cls.includes('text-teal') ||
                            cls.includes('active');
                        if (!isActive) {
                            console.log("[CKG Sweep] Berpindah ke tab 'Sedang Pemeriksaan'...");
                            triggerClick(tabSedang);
                            await sleep(1200);
                            return;
                        }
                    }

                    // 3. KOSONGKAN input pencarian agar seluruh daftar siswa di tab Sedang Pemeriksaan terlihat
                    const searchInput = document.querySelector('input#searchNik') ||
                        document.querySelector('input[placeholder*="cari"], input[placeholder*="nik"], input[type="text"]');

                    if (searchInput && searchInput.value && searchInput.value.trim() !== '') {
                        console.log("[CKG Sweep] 🧹 Mengosongkan kotak pencarian agar seluruh daftar tampil...");
                        isiInput(searchInput, '', false);
                        searchInput.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, cancelable: true, key: 'Enter', code: 'Enter', keyCode: 13 }));
                        searchInput.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true, cancelable: true, key: 'Enter', code: 'Enter', keyCode: 13 }));
                        await sleep(1500);
                        return;
                    }

                    // 4. Periksa baris di tabel tab Sedang Pemeriksaan
                    const tableRows = Array.from(document.querySelectorAll('tbody tr')).filter(r => !r.closest('#asik-ckg-panel, #asik-dashboard-overlay, [id*="asik-ckg"]'));
                    const pageText = document.body ? document.body.innerText.toLowerCase() : '';
                    const pageHtml = document.body ? document.body.innerHTML.toLowerCase() : '';

                    const hasEmptyIcon = document.querySelector('img[src*="empty"], img[src*="icon-empty"], [class*="icon-empty"], [id*="icon-empty"]') !== null ||
                        pageHtml.includes('icon-empty');

                    const hasEmptyText = pageText.includes('cek data di status pemeriksaan lain') ||
                        pageText.includes('cek data di status') ||
                        pageText.includes('status pemeriksaan lain') ||
                        pageText.includes('tidak ada data') ||
                        pageText.includes('no data') ||
                        pageText.includes('belum ada data') ||
                        pageText.includes('belum ada') ||
                        pageText.includes('data tidak ditemukan');

                    const isTableEmpty = tableRows.length === 0 ||
                        hasEmptyIcon ||
                        hasEmptyText ||
                        (tableRows.length === 1 && (
                            tableRows[0].innerText.toLowerCase().includes('tidak ada data') ||
                            tableRows[0].innerText.toLowerCase().includes('no data') ||
                            tableRows[0].innerText.toLowerCase().includes('cek data') ||
                            tableRows[0].innerText.toLowerCase().includes('status pemeriksaan lain') ||
                            tableRows[0].innerText.toLowerCase().includes('icon-empty')
                        ));

                    if (isTableEmpty) {
                        const emptyCount = (state.sweepEmptyCheckCount || 0) + 1;
                        console.log(`[CKG Sweep] ℹ️ Tab Sedang Pemeriksaan Kelas ${targetKelas} kosong / terdeteksi 'icon-empty' (Konfirmasi ${emptyCount}/3)...`);

                        if (emptyCount >= 3) {
                            console.log(`[CKG Sweep] ✅ Kelas ${targetKelas} (${currentSheetName}) telah BERSIH dari antrean Sedang Pemeriksaan! Lanjut ke kelas berikutnya...`);
                            catatRiwayat(`Kelas ${targetKelas}`, currentSheetName, 'Selesai', 'Tab Sedang Pemeriksaan sudah bersih.');

                            state.sweepSheetIndex = sweepIdx + 1;
                            state.needsClassSearch = true;
                            state.sweepEmptyCheckCount = 0;
                            setState(state);
                            await sleep(1000);
                        } else {
                            state.sweepEmptyCheckCount = emptyCount;
                            setState(state);
                            await sleep(1500);
                        }
                        return;
                    }

                    state.sweepEmptyCheckCount = 0;

                    // 5. Cari baris siswa pertama yang belum melebihi batas percobaan error
                    const sweepAttempts = GM_getValue('asik_ckg_sweep_attempts', {});
                    let chosenRow = null;
                    let chosenStudent = null;

                    for (let r of tableRows) {
                        const tds = Array.from(r.querySelectorAll('td'));
                        if (tds.length === 0) continue;

                        let rowNama = '';
                        if (tds[1]) {
                            const t1 = bersih(tds[1].textContent);
                            if (t1 && !t1.match(/^\d+$/) && t1.length > 2 && !/^(sd|smp|sma|kelas|tingkat)\b/i.test(t1)) {
                                rowNama = t1;
                            }
                        }
                        if (!rowNama) {
                            for (let td of tds) {
                                const txt = bersih(td.textContent);
                                if (txt && !txt.match(/^\d+$/) &&
                                    !txt.includes('sedang pemeriksaan') &&
                                    !txt.includes('belum pemeriksaan') &&
                                    !txt.includes('cek data di status') &&
                                    !txt.includes('status pemeriksaan lain') &&
                                    !txt.includes('icon-empty') &&
                                    !/^(sd|smp|sma|kelas|tingkat)\b/i.test(txt) &&
                                    txt.length > 2) {
                                    rowNama = txt;
                                    break;
                                }
                            }
                        }
                        if (!rowNama && tds[1]) rowNama = bersih(tds[1].textContent);
                        if (!rowNama && tds[0]) rowNama = bersih(tds[0].textContent);

                        const failCount = sweepAttempts[rowNama] || 0;
                        if (failCount >= 3) {
                            console.warn(`[CKG Sweep] ⚠️ Siswa "${rowNama}" dilewati karena sudah gagal ${failCount}x.`);
                            continue;
                        }

                        chosenRow = r;
                        chosenStudent = findStudentInWorkbook(rowNama, wbData, targetKelas);
                        break;
                    }

                    if (!chosenRow || !chosenStudent) {
                        const emptyWait = (state.sweepEmptyCheckCount || 0) + 1;
                        if (emptyWait < 3) {
                            console.log(`[CKG Sweep] Menunggu baris data siap dimuat di Kelas ${targetKelas} (${emptyWait}/3)...`);
                            state.sweepEmptyCheckCount = emptyWait;
                            setState(state);
                            await sleep(1500);
                            return;
                        }
                        console.warn(`[CKG Sweep] Semua siswa di Kelas ${targetKelas} telah dicoba / dilewati. Melanjutkan ke kelas berikutnya...`);
                        state.sweepSheetIndex = sweepIdx + 1;
                        state.needsClassSearch = true;
                        state.sweepEmptyCheckCount = 0;
                        setState(state);
                        return;
                    }

                    const btnInRow = chosenRow.querySelector('button, div.cursor-pointer, a');
                    console.log(`[CKG Sweep] 🎯 Memproses siswa dari tab Sedang Pemeriksaan: "${chosenStudent.NAMA}" (Kelas ${targetKelas})...`);

                    state.sweepCurrentStudent = chosenStudent;
                    state.searchRetry = 0;
                    state.phase = 'DETAIL_PELAYANAN';
                    state.raporTerkirim = false;
                    state.tatalaksanaSelesai = false;
                    state.isTatalaksana = false;
                    GM_setValue('asik_ckg_rapor_sent', false);
                    GM_setValue('asik_ckg_tatalaksana_done', false);
                    GM_setValue('asik_ckg_service_attempts', {});
                    setState(state);

                    if (btnInRow) triggerClick(btnInRow);
                    else triggerClick(chosenRow);
                    await sleep(1500);
                    return;
                }

                // ============================================================
                // 4B. MODE NORMAL (Mencari Siswa Berdasarkan Nama di Excel)
                // ============================================================
                console.log(`[CKG] >>> HALAMAN TABEL SISWA — Mencari: ${currentData.NAMA}`);
                const mainArea = document.querySelector('main, section, div[class*="min-h-screen"]') || document.body;
                const allMainTriggers = Array.from(mainArea.querySelectorAll('div[class*="cursor-pointer"], div[class*="h-[2.9rem]"]'))
                    .filter(el => {
                        if (el.closest('#asik-ckg-panel, #asik-dashboard-overlay, aside, nav, #sidebar, [class*="sidebar"]')) return false;
                        if (!isElementVisible(el)) return false;
                        const span = el.querySelector('span');
                        if (!span) return false;
                        const t = bersih(span.textContent);
                        if (!t || t === 'ckg umum' || t === 'ckg sekolah' || t === 'tatalaksana' || t === 'pelayanan') return false;
                        return true;
                    });

                // Dropdown 1 = Kelas (Dropdown 0 = Sekolah, Dropdown 2 = Filter Pencarian)
                const kelasTrigger = allMainTriggers[1] || allMainTriggers.find(el => {
                    const t = bersih(el.querySelector('span')?.textContent || el.textContent);
                    return t === 'pilih kelas' || t.includes('kelas') || /^(sd|smp|sma|smk)\s+kelas/i.test(t);
                });
                const kelasSpan = kelasTrigger ? (kelasTrigger.querySelector('span') || kelasTrigger) : null;
                const currentSelectedText = kelasSpan ? bersih(kelasSpan.textContent) : '';

                // Deteksi kelas yang sedang aktif di layar ASIK
                const activeAsik = detectActiveClassOnAsikPage();
                const currentNum = activeAsik.num || extractClassNum(currentSelectedText);

                const targetKelas = extractClassNum(currentData.KELAS) || currentData.KELAS || '1';
                const targetNum = extractClassNum(targetKelas);


                const lastSearchedKelas = GM_getValue('asik_ckg_last_kelas', '') || state.lastSearchedKelas || '';
                const lastNum = extractClassNum(lastSearchedKelas);

                // Cek apakah tabel/kotak pencarian siswa sudah tampil di layar
                const inputSearchReady = document.querySelector('input#searchNik') !== null ||
                    document.querySelector('input[placeholder*="cari"], input[placeholder*="nik"], input[placeholder*="tiket"], input[placeholder*="nama"], input[placeholder*="Masukkan"]') !== null;

                // Logika perpindahan kelas:
                // HANYA ganti kelas jika dropdown masih bertuliskan 'pilih kelas', atau jika kelas tujuan berbeda dari kelas yang sudah dimuat.
                let mustSwitchClass = false;
                if (currentSelectedText === 'pilih kelas' || !currentNum) {
                    mustSwitchClass = true;
                } else if (targetNum && currentNum && targetNum !== currentNum) {
                    mustSwitchClass = true;
                } else if (state.needsClassSearch || !inputSearchReady) {
                    // Butuh re-trigger Tampilkan Pencarian (setelah kembali dari detail siswa atau input belum muncul)
                    mustSwitchClass = true;
                }

                if (mustSwitchClass) {
                    console.log(`[CKG] 🔄 Ganti/Muat Kelas: Siswa "${currentData.NAMA}" berada di Kelas ${targetKelas} (Di ASIK: "${currentSelectedText || 'Belum dipilih'}", Terakhir Dimuat: "${lastSearchedKelas}").`);

                    // 1. Pilih dropdown kelas jika belum sesuai
                    if (currentSelectedText === 'pilih kelas' || (targetNum && currentNum && targetNum !== currentNum)) {
                        const kelasOk = await pilihDropdownASIK('kelas', targetKelas);
                        if (!kelasOk) {
                            console.warn(`[CKG] Gagal memilih kelas ${targetKelas}. Mencoba lagi iterasi berikutnya...`);
                            return;
                        }
                        await sleep(1000);
                    }

                    // 2. WAJIB SELALU KLIK TOMBOL TAMPILKAN PENCARIAN
                    console.log("[CKG] 🔘 Mengklik tombol 'Tampilkan Pencarian' untuk memuat data kelas baru...");
                    const searchOk = await klikTombolTampilkanPencarian();
                    if (searchOk) {
                        state.needsClassSearch = false;
                        state.lastSearchedKelas = targetKelas;
                        state.searchRetry = 0;
                        GM_setValue('asik_ckg_last_kelas', targetKelas);
                        setState(state);
                    } else {
                        state.needsClassSearch = true;
                        setState(state);
                    }
                    return;
                } else {
                    // Kelas sudah sesuai & kotak pencarian sudah aktif -> Sinkronkan state agar tidak memicu pencarian kelas lagi
                    if (!lastSearchedKelas || lastSearchedKelas !== targetKelas || state.needsClassSearch) {
                        state.lastSearchedKelas = targetKelas;
                        state.needsClassSearch = false;
                        GM_setValue('asik_ckg_last_kelas', targetKelas);
                        setState(state);
                    }
                }

                // 4a. Pastikan berpindah ke tab "Sedang Pemeriksaan" jika belum aktif
                const allTabs = Array.from(document.querySelectorAll('div.cursor-pointer, div, span')).filter(el => {
                    return !el.closest('#asik-ckg-panel, #asik-dashboard-overlay, [id*="asik-ckg"]') && isElementVisible(el);
                });

                const tabSedang = allTabs.find(el => {
                    const t = bersih(el.textContent);
                    return t.startsWith('sedang pemeriksaan');
                });

                if (tabSedang) {
                    const cls = tabSedang.className || '';
                    const isActive = cls.includes('border-b-[#16b3ac]') ||
                        cls.includes('text-teal-500') ||
                        cls.includes('text-teal') ||
                        cls.includes('active');
                    if (!isActive) {
                        console.log("[CKG] 📑 Berpindah ke tab 'Sedang Pemeriksaan'...");
                        triggerClick(tabSedang);
                        await sleep(1200);
                        return; // Jeda agar tab Sedang Pemeriksaan selesai dimuat oleh Vue
                    }
                }

                // 4b. Pastikan filter pencarian diubah ke "Nama" jika saat ini bukan "Nama" (misal: "Nomor Tiket" atau "NIK")
                const spanAktif = Array.from(document.querySelectorAll('span')).find(el => {
                    if (!isElementVisible(el) || el.closest('#asik-ckg-panel, #asik-dashboard-overlay, aside, nav, #sidebar')) return false;
                    const txt = (el.innerText || el.textContent || '').trim();
                    return (txt === 'Nama' || txt === 'NIK' || txt === 'Nomor Tiket') && (el.className || '').includes('line-clamp');
                });

                if (spanAktif && (spanAktif.innerText || spanAktif.textContent || '').trim() !== 'Nama') {
                    const namaFilterSaatIni = (spanAktif.innerText || spanAktif.textContent || '').trim();
                    console.log(`[CKG] 🔄 Membuka dropdown filter pencarian (saat ini: "${namaFilterSaatIni}"). Mengganti ke "Nama"...`);
                    const clickableBox = spanAktif.closest('.cursor-pointer') || spanAktif;
                    triggerClick(clickableBox);
                    await sleep(1500); // Tahan 1.5 detik agar dropdown benar-benar muncul di DOM

                    const opsiNama = Array.from(document.querySelectorAll('div.cursor-pointer, div.py-2, li, span, div')).find(el => {
                        if (el.closest('#asik-ckg-panel, #asik-dashboard-overlay, aside, nav, #sidebar')) return false;
                        const txt = (el.innerText || el.textContent || '').trim();
                        return (txt === 'Nama' || txt.toLowerCase() === 'nama') && (el.offsetWidth > 0 || el.getBoundingClientRect().width > 0 || isElementVisible(el));
                    });

                    if (opsiNama) {
                        console.log("[CKG] ✅ Memilih opsi filter 'Nama'...");
                        triggerClick(opsiNama);
                        if (typeof opsiNama.click === 'function') {
                            try { opsiNama.click(); } catch (e) { }
                        }
                        await sleep(2000); // Beri jeda agar Vue selesai me-render ulang kotak input baru
                        return; // Langsung kembali agar iterasi berikutnya langsung mengambil input yang sudah mode Nama
                    } else {
                        console.warn("[CKG] Opsi filter 'Nama' tidak ditemukan di dropdown!");
                    }
                }

                // 4c. Cari input text NIK/Nama Siswa di halaman
                const searchInput = document.querySelector('input#searchNik') ||
                    document.querySelector('input[placeholder*="nama"], input[placeholder*="cari"], input[placeholder*="nik"], input[placeholder*="tiket"], input[placeholder*="Masukkan"], input[type="text"]');

                if (searchInput) {
                    const namaSiswa = currentData.NAMA;
                    const queryNama = getSearchQuery(namaSiswa, state.searchRetry || 0);
                    console.log(`[CKG] Mengetik pencarian nama: "${queryNama}" (Nama Lengkap: "${namaSiswa}")...`);
                    isiInput(searchInput, queryNama, false);
                    await sleep(300);

                    // 4c. Trigger pencarian (tekan Enter)
                    searchInput.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, cancelable: true, key: 'Enter', code: 'Enter', keyCode: 13 }));
                    searchInput.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true, cancelable: true, key: 'Enter', code: 'Enter', keyCode: 13 }));
                    await sleep(1500); // Tunggu respons pencarian server

                    // 4d. Cari baris siswa yang cocok di tabel
                    const allTds = Array.from(document.querySelectorAll('tbody td')).filter(el => !el.closest('#asik-ckg-panel, #asik-dashboard-overlay, [id*="asik-ckg"]'));

                    const matchTd = allTds.find(td => {
                        const t = bersih(td.textContent);
                        if (!t || t.includes('tidak ada data') || t.includes('no data') || t.includes('belum ada') || t.includes('cek data di status') || t.includes('status pemeriksaan lain') || t.includes('icon-empty')) return false;
                        return isNamaCocok(t, currentData.NAMA);
                    });

                    if (matchTd) {
                        const row = matchTd.closest('tr') || matchTd;
                        const btnInRow = row.querySelector('button, div.cursor-pointer, a');
                        console.log(`[CKG] Siswa ditemukan: "${matchTd.textContent.trim()}". Membuka layanan...`);
                        state.searchRetry = 0;
                        state.phase = 'DETAIL_PELAYANAN';
                        state.raporTerkirim = false;
                        state.tatalaksanaSelesai = false;
                        state.isTatalaksana = false;
                        GM_setValue('asik_ckg_rapor_sent', false);
                        GM_setValue('asik_ckg_tatalaksana_done', false);
                        GM_setValue('asik_ckg_service_attempts', {});
                        setState(state);
                        if (btnInRow) triggerClick(btnInRow);
                        else triggerClick(row);
                        await sleep(1500);
                    } else {
                        // Jika nama tidak ditemukan di pencarian ASIK
                        const searchRetry = (state.searchRetry || 0) + 1;
                        console.warn(`[CKG] Siswa "${currentData.NAMA}" belum ditemukan di tabel (percobaan ${searchRetry}/2)...`);

                        // HANYA lewati jika sudah mencoba 2x (percobaan 1: 2 kata, percobaan 2: 1 kata)
                        if (searchRetry >= 2) {
                            console.warn(`[CKG] Siswa "${currentData.NAMA}" TIDAK DITEMUKAN di ASIK setelah 2x pencarian. Melewati siswa ini dan lanjut ke nama berikutnya...`);
                            catatRiwayat(currentData.NAMA, currentData.KELAS, 'Dilewati', 'Nama tidak ditemukan di pencarian ASIK setelah 2x percobaan');

                            // Bersihkan input pencarian dan reset filter tabel dengan Enter
                            isiInput(searchInput, '', true);
                            await sleep(600);

                            // Bersihkan sesi aktif siswa yang dilewati agar tidak meracuni siswa berikutnya
                            sessionStorage.removeItem('asik_ckg_active_student');
                            sessionStorage.removeItem('asik_ckg_active_student_name');
                            state.sweepCurrentStudent = null;

                            state.searchRetry = 0;
                            state.index += 1;
                            state.phase = 'MENCARI';
                            setState(state);
                        } else {
                            // Percobaan 1 belum ketemu, bersihkan input dan coba lagi dengan kata kunci 1 kata
                            console.log(`[CKG] Mencoba pencarian ulang untuk "${currentData.NAMA}" dengan kata kunci lebih fleksibel (1 kata)...`);
                            isiInput(searchInput, '', false);
                            await sleep(400);
                            state.searchRetry = searchRetry;
                            setState(state);
                        }
                    }
                } else {
                    console.log("[CKG] Kotak input pencarian belum muncul di layar.");
                }
                return;

                // ============================================================
                // 5. HALAMAN PILIH SEKOLAH & KELAS AWAL
                // ============================================================
            } else if (adaFilterSekolah && !adaTabelSiswa) {
                console.log("[CKG] >>> HALAMAN PILIH SEKOLAH & KELAS");
                state.phase = 'PILIH_SEKOLAH';
                setState(state);

                // 1. Pilih Sekolah terlebih dahulu
                const targetSekolah = state.sekolahTarget || 'PALMERIAM 01';
                console.log(`[CKG] Target Sekolah: "${targetSekolah}"`);
                const sekolahOk = await pilihDropdownASIK('sekolah', targetSekolah);
                if (!sekolahOk) {
                    console.log(`[CKG] Sekolah "${targetSekolah}" belum berhasil terpilih. Menunggu antarmuka...`);
                    return;
                }

                // Jeda 1.2 detik agar ASIK selesai memuat opsi kelas dari server setelah sekolah terpilih
                await sleep(1200);

                // 2. Pilih Kelas setelah sekolah berhasil terpilih
                const targetKelas = currentData.KELAS || '1';
                console.log(`[CKG] Target Kelas: "${targetKelas}"`);
                const kelasOk = await pilihDropdownASIK('kelas', targetKelas);
                if (!kelasOk) {
                    console.log(`[CKG] Kelas "${targetKelas}" belum berhasil terpilih. Menunggu antarmuka...`);
                    return;
                }

                // Jeda 1 detik sebelum klik tombol tampilkan pencarian
                await sleep(1000);

                // 3. Klik Tampilkan Pencarian setelah sekolah & kelas terpilih
                const searchOk = await klikTombolTampilkanPencarian();
                if (searchOk) {
                    state.phase = 'MENCARI';
                    state.needsClassSearch = false;
                    state.lastSearchedKelas = targetKelas;
                    state.searchRetry = 0;
                    GM_setValue('asik_ckg_last_kelas', targetKelas);
                    setState(state);
                } else {
                    console.log("[CKG] Tombol Tampilkan Pencarian belum aktif. Memeriksa filter...");
                }
                return;

                // ============================================================
                // 6. HALAMAN TIDAK DIKENALI
                // ============================================================
            } else {
                console.log(`[CKG] Halaman sedang loading / tidak dikenali. Phase: ${state.phase}. Menunggu...`);
                await sleep(1500);
            }

        } catch (error) {
            console.error("[CKG Error]", error);
            await sleep(1500);
        } finally {
            // Jadwalkan iterasi loop berikutnya secara teratur (1.5 detik) jika robot masih RUNNING
            if (getState().running) {
                if (loopTimerId) clearTimeout(loopTimerId);
                loopTimerId = setTimeout(jalankanLoop, 1500);
            }
        }
    }


    // --
    // USER INTERFACE (UI)
    // --
    function createUI() {
        if (document.getElementById('asik-ckg-panel')) return;

        const style = document.createElement('style');
        style.textContent = `
            #asik-ckg-panel {
                position: fixed; bottom: 20px; right: 20px; z-index: 999999;
                width: 260px; font-family: system-ui, -apple-system, sans-serif;
                border-radius: 12px; overflow: hidden;
                box-shadow: 0 8px 24px rgba(0,0,0,0.35);
                transition: all 0.3s ease; user-select: none;
            }
            #asik-ckg-panel.minimized { width: auto; }
            #asik-ckg-panel.minimized .asik-body { display: none; }
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
            .asik-status-val { color: #e2e8f0; font-weight: 600; text-align: right; max-width: 150px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
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

        const panel = document.createElement('div');
        panel.id = 'asik-ckg-panel';
        panel.innerHTML = `
            <div class="asik-header" style="background:#0f766e; color:#ffffff;">
                <span class="asik-header-title" style="font-size:13px; font-weight:800; display:flex; align-items:center; gap:6px;">
                    🏫 [ROBOT CKG SEKOLAH]
                </span>
                <button class="asik-header-btn" id="asik-ckg-btn-minimize">-</button>
            </div>
            <div class="asik-body">
                <div style="margin-bottom:10px; font-size:10px; color:#94a3b8; text-align:center;">
                    Format Excel: NAMA, TB, BB, MATA, TELINGA, GIGI
                </div>
                <div class="asik-status-row"><span>Status</span><span id="asik-ckg-s-robot" class="asik-status-val">-</span></div>
                <div class="asik-status-row"><span>Sekolah</span><span id="asik-ckg-s-sekolah" class="asik-status-val">-</span></div>
                <div class="asik-status-row"><span>Pasien</span><span id="asik-ckg-s-nama" class="asik-status-val">-</span></div>
                <div class="asik-status-row"><span>Progress</span><span id="asik-ckg-s-progress" class="asik-status-val" style="cursor:pointer;" title="Klik untuk melompat ke nomor siswa tertentu">-</span></div>
                
                <div style="display:flex; gap:4px; margin-top:6px; margin-bottom:6px;">
                    <button id="asik-ckg-btn-prev" class="asik-btn asik-btn-gray" style="flex:1; padding:5px 0; margin-bottom:0; font-size:11px;" title="Kembali ke siswa sebelumnya">⏮️ PREV</button>
                    <button id="asik-ckg-btn-jump" class="asik-btn asik-btn-yellow" style="flex:1; padding:5px 0; margin-bottom:0; font-size:11px; color:#000;" title="Pilih nomor urut siswa tertentu">🔢 NO</button>
                    <button id="asik-ckg-btn-next" class="asik-btn asik-btn-gray" style="flex:1; padding:5px 0; margin-bottom:0; font-size:11px;" title="Lewati ke siswa berikutnya">⏭️ NEXT</button>
                </div>

                <hr class="asik-divider">
                <select id="asik-ckg-sheet" style="margin-bottom:6px; width:100%; display:none; padding:4px; border-radius:4px; font-size:11px; background:#1e293b; color:white; border:1px solid #334155;"></select>
                
                <button id="asik-ckg-btn-main" class="asik-btn asik-btn-green">START ROBOT</button>
                <button id="asik-ckg-btn-sweep" class="asik-btn asik-btn-yellow" style="color:#000;">🧹 SAPU SEDANG PEMERIKSAAN</button>
                <button id="asik-ckg-btn-reset" class="asik-btn asik-btn-red">RESET DATA</button>
                <button id="asik-ckg-btn-laporan" class="asik-btn asik-btn-blue">📊 LAPORAN</button>
                <button id="asik-ckg-btn-upload" class="asik-btn asik-btn-gray">UPLOAD EXCEL</button>
                <input type="file" id="asik-ckg-file" accept=".xlsx, .xls" style="display:none;" />
            </div>
        `;
        document.body.appendChild(panel);

        // Draggable & Minimize
        let isDragging = false, offsetX = 0, offsetY = 0;
        const header = panel.querySelector('.asik-header');
        const btnMin = document.getElementById('asik-ckg-btn-minimize');

        btnMin.onclick = (e) => {
            e.stopPropagation();
            panel.classList.toggle('minimized');
            btnMin.textContent = panel.classList.contains('minimized') ? '+' : '-';
        };

        header.onmousedown = (e) => { if (e.target === btnMin) return; isDragging = true; offsetX = e.clientX - panel.getBoundingClientRect().left; offsetY = e.clientY - panel.getBoundingClientRect().top; panel.style.transition = 'none'; };
        document.addEventListener('mousemove', (e) => { if (!isDragging) return; panel.style.left = (e.clientX - offsetX) + 'px'; panel.style.top = (e.clientY - offsetY) + 'px'; panel.style.right = 'auto'; panel.style.bottom = 'auto'; });
        document.addEventListener('mouseup', () => {
            if (isDragging) {
                isDragging = false;
                panel.style.transition = 'all 0.3s ease';
            }
        });

        const parseSingleSheet = (ws, wsname) => {
            if (!ws) return { students: [], sekolah: "" };
            const rawRows = window.XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
            if (!rawRows || rawRows.length === 0) return { students: [], sekolah: "" };

            // 1. Ekstrak nama sekolah dari judul di baris awal
            let extractedSekolah = "";
            for (let i = 0; i < Math.min(15, rawRows.length); i++) {
                const rowStr = (rawRows[i] || []).map(c => String(c || '').trim()).join(' ').toUpperCase();
                if (rowStr.includes("SD ") || rowStr.includes("SMP ") || rowStr.includes("SMA ") || rowStr.includes("SMK ") || rowStr.includes("SDN ")) {
                    const match = rowStr.match(/(?:SDN|SD|SMP|SMA|SMK)\s+(.+?)(?:\s+TAHUN|\s+DI|\s+LINGKUNGAN|\s*$)/i);
                    if (match) {
                        extractedSekolah = match[1].trim();
                        break;
                    }
                }
            }

            // 2. Ekstrak angka kelas default dari nama sheet (misal: "Kelas 1" -> "1", "10" -> "10", "X" -> "10", "XII" -> "12")
            let defaultKelas = extractClassNum(wsname);
            if (!defaultKelas) defaultKelas = wsname.replace(/[^0-9]/g, '') || wsname;

            // 3. Cari baris Header Tabel yang sebenarnya
            let headerIdx = -1;
            let colNama = -1;
            let colKelas = -1;
            let colTb = -1;
            let colBb = -1;
            let colSistol = -1;
            let colDiastol = -1;
            let colMata = -1;
            let colTelinga = -1;
            let colGigi = -1;
            let colGds = -1;
            let colHb = -1;

            for (let r = 0; r < Math.min(30, rawRows.length); r++) {
                const row = rawRows[r] || [];
                const rowStrs = row.map(c => String(c || '').toUpperCase().trim());

                const hasNama = rowStrs.some(c => c.includes('NAMA') && !['ORANG', 'WALI', 'AYAH', 'IBU', 'SEKOLAH', 'DAFTAR'].some(x => c.includes(x)));
                const hasSupportCol = rowStrs.some(c => ['NO', 'NO.', 'NISN', 'NIK', 'NIK SISWA', 'KELAS', 'JK', 'JENIS KELAMIN', 'TGL LAHIR', 'ALAMAT'].some(x => c === x || c.startsWith(x)));

                if (hasNama && (hasSupportCol || rowStrs.filter(Boolean).length >= 3)) {
                    headerIdx = r;
                    for (let c = 0; c < rowStrs.length; c++) {
                        const cell = rowStrs[c];
                        if (cell.includes('NAMA') && !['ORANG', 'WALI', 'AYAH', 'IBU', 'SEKOLAH'].some(x => cell.includes(x))) {
                            colNama = c;
                            break;
                        }
                    }
                    break;
                }
            }

            // Fallback jika tidak menemukan kolom pendukung
            if (headerIdx === -1) {
                for (let r = 0; r < Math.min(30, rawRows.length); r++) {
                    const row = rawRows[r] || [];
                    for (let c = 0; c < row.length; c++) {
                        const cell = String(row[c] || '').toUpperCase().trim();
                        if (cell === 'NAMA SISWA' || cell === 'NAMA LENGKAP' || cell === 'NAMA PESERTA' || (cell.includes('NAMA') && !cell.includes('DAFTAR') && !cell.includes('SEKOLAH'))) {
                            headerIdx = r;
                            colNama = c;
                            break;
                        }
                    }
                    if (headerIdx !== -1) break;
                }
            }

            if (headerIdx === -1 || colNama === -1) {
                return { students: [], sekolah: extractedSekolah };
            }

            // 4. Scan baris header untuk memetakan kolom TB, BB, KELAS, SISTOLE, DIASTOLE, MATA, TELINGA, GIGI, GDS, HB
            for (let r = 0; r < Math.min(30, rawRows.length); r++) {
                const row = rawRows[r] || [];
                for (let c = 0; c < row.length; c++) {
                    const cell = String(row[c] || '').toUpperCase().trim();
                    if (!cell) continue;
                    // Abaikan sel berupa kalimat judul panjang agar tidak salah mendeteksi kolom
                    if (cell.length > 80 || cell.split(/\s+/).length > 12) continue;

                    const normCell = cell.replace(/[^A-Z0-9]/g, '');

                    if ((cell === 'KELAS' || normCell === 'KELAS' || cell.includes('KELAS') || cell.includes('ROMBEL')) && colKelas === -1) colKelas = c;
                    if ((cell === 'TB' || normCell === 'TB' || cell === 'TINGGI' || cell.includes('TINGGI BADAN') || normCell === 'TINGGIBADAN') && colTb === -1) colTb = c;
                    if ((cell === 'BB' || normCell === 'BB' || cell === 'BERAT' || cell.includes('BERAT BADAN') || normCell === 'BERATBADAN') && colBb === -1) colBb = c;
                    if ((cell.includes('SISTOL') || cell === 'TD' || cell === 'TENSI') && colSistol === -1) colSistol = c;
                    if (cell.includes('DIASTOL') && colDiastol === -1) colDiastol = c;
                    if ((cell === 'MATA' || cell === 'VISUS' || cell === 'PENGLIHATAN' || (cell.includes('MATA') && !['KECAMATAN', 'KACAMATA', 'BUTA'].some(x => cell.includes(x)))) && colMata === -1) colMata = c;
                    if ((cell === 'TELINGA' || cell.includes('TELINGA') || cell.includes('PENDENGARAN') || cell.includes('SERUMEN')) && colTelinga === -1) colTelinga = c;
                    if ((cell === 'GIGI' || cell.includes('GIGI') || cell.includes('KARIES')) && colGigi === -1) colGigi = c;

                    // Deteksi Kolom GDS (Gula Darah Sewaktu) - dukung segala variasi teks header
                    const isGds = (
                        cell === 'GDS' ||
                        normCell === 'GDS' ||
                        normCell === 'GDS1' ||
                        normCell.startsWith('GDS') ||
                        normCell.includes('GDS') ||
                        cell.includes('GDS') ||
                        normCell.includes('GULADARAH') ||
                        normCell.includes('GLUKOSADARAH') ||
                        normCell.includes('GULA') ||
                        cell.includes('GULA') ||
                        normCell.includes('GLUKOSA') ||
                        cell.includes('GLUKOSA') ||
                        normCell.includes('GLUCOSE') ||
                        cell.includes('GLUCOSE') ||
                        cell === 'DM'
                    ) && !cell.includes('GDS 2') && !normCell.includes('GDS2') && !cell.includes('KEDUA') && !cell.includes('KE-2') && !cell.includes('RIWAYAT');

                    if (isGds && colGds === -1) colGds = c;

                    // Deteksi Kolom Hb (Hemoglobin) - dukung segala variasi teks header (HB, Kadar HB, Hasil HB, dll.)
                    const isHb = (
                        cell === 'HB' ||
                        normCell === 'HB' ||
                        normCell.startsWith('HB') ||
                        normCell.includes('HB') ||
                        cell.includes('HB') ||
                        normCell.includes('HEMOGLOBIN') ||
                        cell.includes('HEMOGLOBIN') ||
                        normCell.includes('ANEMIA') ||
                        cell.includes('ANEMIA')
                    ) && !cell.includes('REHAB') && !cell.includes('RIWAYAT');

                    if (isHb && colHb === -1) colHb = c;
                }
            }

            console.log(`[CKG Excel] Sheet "${wsname}" header dipetakan (baris ${headerIdx}): NAMA=${colNama}, KELAS=${colKelas}, TB=${colTb}, BB=${colBb}, GDS=${colGds} (${colGds !== -1 ? 'DITEMUKAN di Excel' : 'TIDAK ADA'}), HB=${colHb} (${colHb !== -1 ? 'DITEMUKAN di Excel' : 'TIDAK ADA'}), MATA=${colMata}, TELINGA=${colTelinga}, GIGI=${colGigi}`);

            // 5. Ekstrak Data Siswa (Dengan filter ketat terhadap baris catatan/rekap/footer)
            const students = [];
            for (let r = headerIdx + 1; r < rawRows.length; r++) {
                const row = rawRows[r] || [];
                if (colNama >= row.length) continue;
                const nama = String(row[colNama] || '').trim();
                const namaUpper = nama.toUpperCase();
                const namaLower = nama.toLowerCase();

                // 1. Cek panjang dan hanya angka
                if (!nama || nama.length < 2 || /^\d+$/.test(nama)) {
                    continue;
                }

                // 2. Cek label header berulang
                if (['NAMA', 'NAMA SISWA', 'NAMA LENGKAP', 'NAMA PESERTA', 'NAMA ANAK', 'NAMA MURID', 'JUDUL'].includes(namaUpper)) {
                    continue;
                }

                // 3. Cek apakah baris ini adalah baris rekapitulasi / footer / tanda tangan (periksa seluruh sel di baris ini)
                const rowStrs = row.map(c => String(c || '').toUpperCase().trim());
                const isFooterRow = rowStrs.some(c => {
                    return c === 'JUMLAH SISWA' || c === 'JUMLAH' || c === 'TOTAL' || c === 'REKAP' || c === 'REKAPITULASI' ||
                        c === 'CATATAN' || c === 'CATATAN:' || c === 'KETERANGAN' || c === 'KETERANGAN:' ||
                        c === 'LAKI-LAKI' || c === 'PEREMPUAN' || c === 'WALI KELAS' || c === 'GURU KELAS' ||
                        c === 'KEPALA SEKOLAH' || c === 'MENGETAHUI' || c === 'MENGETAHUI,' || c.startsWith('NIP');
                });
                if (isFooterRow) continue;

                // 4. Cek pola teks catatan / status rekap pada kolom nama (contoh: "input 84", "tak ada nik2", "sudah ckg 27", dll.)
                const isCatatan = /^(input\s*\d+|tak\s*ada\s*nik|tidak\s*ada\s*nik|sudah\s*ckg|belum\s*ckg|invalid|catatan|keterangan|rekap|total|jumlah|laki[-\s]?laki|perempuan|wali\s*kelas|kepala\s*sekolah|guru|mengetahui|nip)\b/i.test(namaLower) ||
                    /^(input|tak ada nik|tidak ada nik|sudah input|belum input|invalid nik)/i.test(namaLower) ||
                    /(tak\s*ada\s*nik|tidak\s*ada\s*nik|invalid\s*nik|sudah\s*ckg|belum\s*ckg)/i.test(namaLower);
                if (isCatatan) continue;

                let rawTbVal = (colTb !== -1 && colTb < row.length) ? (row[colTb] !== null && row[colTb] !== undefined ? String(row[colTb]).trim() : '') : '';
                let rawBbVal = (colBb !== -1 && colBb < row.length) ? (row[colBb] !== null && row[colBb] !== undefined ? String(row[colBb]).trim() : '') : '';

                // Konversi desimal dan bulatkan ke integer terdekat agar sesuai tampilan Excel (misal 159.9 -> 160)
                let tb = '';
                if (rawTbVal && rawTbVal !== '-' && rawTbVal !== '--' && rawTbVal !== '0') {
                    let numTb = parseFloat(rawTbVal.replace(',', '.'));
                    if (!isNaN(numTb) && numTb > 0) {
                        if (numTb > 300 && numTb < 3000) numTb = numTb / 10;
                        tb = String(Math.round(numTb));
                    }
                }

                let bb = '';
                if (rawBbVal && rawBbVal !== '-' && rawBbVal !== '--' && rawBbVal !== '0') {
                    let numBb = parseFloat(rawBbVal.replace(',', '.'));
                    if (!isNaN(numBb) && numBb > 0) {
                        if (numBb > 250 && numBb < 2500) numBb = numBb / 10;
                        bb = String(Math.round(numBb));
                    }
                }
                const sistole = (colSistol !== -1 && colSistol < row.length) ? String(row[colSistol] || '').trim() : '';
                const diastole = (colDiastol !== -1 && colDiastol < row.length) ? String(row[colDiastol] || '').trim() : '';
                const mata = (colMata !== -1 && colMata < row.length) ? String(row[colMata] || '').trim() : '';
                const telinga = (colTelinga !== -1 && colTelinga < row.length) ? String(row[colTelinga] || '').trim() : '';
                const gigi = (colGigi !== -1 && colGigi < row.length) ? String(row[colGigi] || '').trim() : '';

                let rawGdsVal = (colGds !== -1 && colGds < row.length) ? (row[colGds] !== null && row[colGds] !== undefined ? String(row[colGds]).trim() : '') : '';
                let cleanGdsVal = rawGdsVal;
                if (cleanGdsVal === '-' || cleanGdsVal === '--' || cleanGdsVal === '0') {
                    cleanGdsVal = '';
                } else if (cleanGdsVal) {
                    if (cleanGdsVal.includes('.') || cleanGdsVal.includes(',')) {
                        const num = parseFloat(cleanGdsVal.replace(',', '.'));
                        if (!isNaN(num) && num > 0) {
                            cleanGdsVal = String(Math.round(num));
                        }
                    }
                    const digits = cleanGdsVal.replace(/[^0-9]/g, '');
                    cleanGdsVal = (digits && digits !== '0') ? digits : '';
                }
                const gds = cleanGdsVal;
                const rawHbVal = (colHb !== -1 && colHb < row.length) ? String(row[colHb] || '').trim() : '';
                const cleanHbVal = (rawHbVal === '-' || rawHbVal === '0' || rawHbVal === '--') ? '' : rawHbVal.replace(/[^0-9.,]/g, '').replace(',', '.').trim();
                const hb = cleanHbVal || '12';
                let rawKelas = (colKelas !== -1 && colKelas < row.length) ? String(row[colKelas] || '').trim() : defaultKelas;
                if (!rawKelas) rawKelas = defaultKelas;
                let kelas = extractClassNum(rawKelas) || defaultKelas;
                if (defaultKelas && (defaultKelas === '10' || defaultKelas === '11' || defaultKelas === '12') && kelas !== defaultKelas) {
                    kelas = defaultKelas;
                }

                const studentItem = {
                    NAMA: nama,
                    TB: tb,
                    BB: bb,
                    GDS: gds,
                    HB: hb,
                    SISTOLE: sistole,
                    DIASTOLE: diastole,
                    KELAS: kelas,
                    TAB: wsname,
                    MATA: mata,
                    TELINGA: telinga,
                    GIGI: gigi
                };

                // Masukkan semua siswa ke daftar agar kirim rapor, tatalaksana, dan selesai pemeriksaan tetap diproses untuk setiap siswa
                students.push(studentItem);
            }

            return { students, sekolah: extractedSekolah };
        };

        const renderSheetDropdown = (selectedTab) => {
            const sheetSelect = document.getElementById('asik-ckg-sheet');
            if (!sheetSelect) return;

            let wbData = null;
            try {
                const raw = GM_getValue('asik_ckg_workbook_data') || localStorage.getItem('asik_ckg_workbook_data');
                if (raw) wbData = typeof raw === 'string' ? JSON.parse(raw) : raw;
            } catch (e) { }

            if (!wbData || !wbData.sheetNames || wbData.sheetNames.length === 0) {
                sheetSelect.style.display = 'none';
                return;
            }

            sheetSelect.innerHTML = '';

            // Opsi 1: SEMUA TAB jika ada lebih dari 1 tab
            if (wbData.sheetNames.length > 1) {
                const allOpt = document.createElement('option');
                allOpt.value = '__ALL__';
                allOpt.textContent = `📂 SEMUA TAB (${wbData.sheetNames.length} Tab / Otomatis Lanjut)`;
                sheetSelect.appendChild(allOpt);
            }

            wbData.sheetNames.forEach(name => {
                const opt = document.createElement('option');
                opt.value = name;
                const count = (wbData.sheets[name] || []).length;
                opt.textContent = `Tab: ${name} (${count} siswa)`;
                sheetSelect.appendChild(opt);
            });

            sheetSelect.style.display = 'block';

            const activeTab = selectedTab || GM_getValue('asik_ckg_selected_tab', wbData.sheetNames.length > 1 ? '__ALL__' : wbData.sheetNames[0]);
            sheetSelect.value = activeTab;
        };

        const loadSavedSheetData = (tabName, autoAlert = true) => {
            let wbData = null;
            try {
                const raw = GM_getValue('asik_ckg_workbook_data') || localStorage.getItem('asik_ckg_workbook_data');
                if (raw) wbData = typeof raw === 'string' ? JSON.parse(raw) : raw;
            } catch (e) { }

            if (!wbData) {
                if (autoAlert) alert("Data Excel belum tersedia. Silakan upload Excel terlebih dahulu.");
                return;
            }

            let finalData = [];
            if (tabName === '__ALL__') {
                wbData.sheetNames.forEach(name => {
                    const list = wbData.sheets[name] || [];
                    finalData = finalData.concat(list);
                });
            } else {
                finalData = wbData.sheets[tabName] || [];
            }

            if (finalData.length === 0) {
                if (autoAlert) alert(`Tidak ada siswa pada tab "${tabName}".`);
                return;
            }

            setState({
                data: finalData,
                index: 0,
                running: false,
                phase: 'PILIH_SEKOLAH',
                needsClassSearch: true,
                sekolahTarget: wbData.sekolah || '',
                sweepMode: false,
                sweepSheetIndex: 0,
                sweepCurrentStudent: null,
                sweepEmptyCheckCount: 0
            });

            GM_setValue('asik_ckg_selected_tab', tabName);
            console.log(`[CKG] Berhasil memuat ${finalData.length} data siswa untuk tab "${tabName}".`);
            if (autoAlert) {
                const infoText = tabName === '__ALL__' ? `${finalData.length} data siswa dari ${wbData.sheetNames.length} tab Excel` : `${finalData.length} data siswa dari tab "${tabName}"`;
                const sampleGds = finalData.find(s => s.GDS && s.GDS !== '90') || finalData.find(s => s.GDS) || {};
                const sampleHb = finalData.find(s => s.HB && s.HB !== '12') || finalData.find(s => s.HB) || {};
                const gdsStatus = sampleGds.GDS ? `Terdeteksi (contoh: ${sampleGds.GDS})` : 'Tidak Terdeteksi';
                const hbStatus = sampleHb.HB ? `Terdeteksi (contoh: ${sampleHb.HB})` : 'Tidak Terdeteksi';
                alert(`✅ Berhasil memuat ${infoText}.\nSekolah: ${wbData.sekolah || 'Tidak diketahui'}\nKolom GDS: ${gdsStatus}\nKolom HB: ${hbStatus}`);
            }
        };

        // Render sheet dropdown on panel creation
        renderSheetDropdown();

        // Event Listeners
        document.getElementById('asik-ckg-btn-upload').addEventListener('click', () => {
            document.getElementById('asik-ckg-file').click();
        });

        document.getElementById('asik-ckg-file').addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (!file) return;

            const reader = new FileReader();
            reader.onload = (evt) => {
                try {
                    const data = new Uint8Array(evt.target.result);
                    const wb = window.XLSX.read(data, { type: 'array', raw: true });

                    const wbData = {
                        sheetNames: wb.SheetNames,
                        sheets: {},
                        sekolah: ""
                    };

                    wb.SheetNames.forEach(name => {
                        const ws = wb.Sheets[name];
                        const res = parseSingleSheet(ws, name);
                        wbData.sheets[name] = res.students;
                        if (!wbData.sekolah && res.sekolah) {
                            wbData.sekolah = res.sekolah;
                        }
                    });

                    // Persist workbook data
                    GM_setValue('asik_ckg_workbook_data', JSON.stringify(wbData));
                    localStorage.setItem('asik_ckg_workbook_data', JSON.stringify(wbData));

                    const defaultTab = wb.SheetNames.length > 1 ? '__ALL__' : wb.SheetNames[0];
                    renderSheetDropdown(defaultTab);
                    loadSavedSheetData(defaultTab, true);
                } catch (err) {
                    console.error("[CKG] Gagal membaca Excel:", err);
                    alert("Gagal membaca file Excel. Pastikan file berformat .xlsx atau .xls yang valid.");
                }
            };
            reader.readAsArrayBuffer(file);
        });

        document.getElementById('asik-ckg-sheet').addEventListener('change', (e) => {
            loadSavedSheetData(e.target.value, true);
        });

        document.getElementById('asik-ckg-btn-main').addEventListener('click', () => {
            const state = getState();
            if (state.data.length === 0) return alert("Data kosong! Silakan upload Excel terlebih dahulu.");

            // Toggle start/stop
            if (state.running) {
                // STOP: Batalkan timer yang sedang berjalan
                if (loopTimerId) {
                    clearTimeout(loopTimerId);
                    loopTimerId = null;
                }
                setState({ ...state, running: false });
                console.log('[CKG] Robot DIHENTIKAN.');
            } else {
                // START: Cek apakah di halaman ASIK sudah ada sekolah dan kelas yang terpilih
                const activeAsik = detectActiveClassOnAsikPage();
                console.log(`[CKG] Deteksi filter ASIK aktif: Kelas="${activeAsik.text}" (Kelas ${activeAsik.num})`);

                // Pastikan mode penyisiran (sweepMode) dimatikan saat tombol START ditekan
                state.sweepMode = false;
                state.sweepSheetIndex = 0;
                state.sweepCurrentStudent = null;
                state.sweepEmptyCheckCount = 0;

                const adaFilterAwal = Array.from(document.querySelectorAll('span, div')).some(el => {
                    const t = bersih(el.textContent);
                    return t === 'pilih sekolah' || t === 'pilih kelas';
                }) || document.querySelector('#sekolah') !== null;

                const newPhase = (activeAsik.num && !adaFilterAwal) ? 'MENCARI' : (adaFilterAwal ? 'PILIH_SEKOLAH' : (state.phase === 'IDLE' ? 'MENCARI' : state.phase));
                const updatedState = getState();
                sessionStorage.removeItem('asik_ckg_active_student');
                sessionStorage.removeItem('asik_ckg_active_student_name');
                setState({ ...updatedState, sweepMode: false, sweepSheetIndex: 0, sweepCurrentStudent: null, running: true, phase: newPhase, needsClassSearch: true });
                console.log('[CKG] Robot DIMULAI. Phase:', newPhase);
                jalankanLoop();
            }
        });

        document.getElementById('asik-ckg-btn-sweep').addEventListener('click', () => {
            const state = getState();
            let wbData = null;
            try {
                const raw = GM_getValue('asik_ckg_workbook_data') || localStorage.getItem('asik_ckg_workbook_data');
                if (raw) wbData = typeof raw === 'string' ? JSON.parse(raw) : raw;
            } catch (e) { }

            if (!wbData || !wbData.sheetNames || wbData.sheetNames.length === 0) {
                if (state.data.length === 0) {
                    return alert("Silakan upload file Excel terlebih dahulu agar robot mengetahui daftar kelas.");
                }
            }

            const activeAsik = detectActiveClassOnAsikPage();
            let startSweepIdx = 0;
            if (activeAsik.num && wbData && wbData.sheetNames) {
                const idx = wbData.sheetNames.findIndex(sn => extractClassNum(sn) === activeAsik.num);
                if (idx !== -1) startSweepIdx = idx;
            }

            const kelasAwalInfo = (wbData && wbData.sheetNames && wbData.sheetNames[startSweepIdx])
                ? `Kelas ${extractClassNum(wbData.sheetNames[startSweepIdx]) || wbData.sheetNames[startSweepIdx]}`
                : 'kelas awal';

            const confirmSweep = confirm(`Mulai pengerjaan ulang (penyisiran) seluruh antrean di tab 'Sedang Pemeriksaan' mulai dari ${kelasAwalInfo}?`);
            if (!confirmSweep) return;

            // Pastikan kembali ke halaman tabel jika saat ini sedang di detail
            const searchInput = document.querySelector('input#searchNik') || document.querySelector('input[placeholder*="cari"], input[placeholder*="nik"]');
            if (!searchInput) {
                const backBtn = Array.from(document.querySelectorAll('img.cursor-pointer, .cursor-pointer, button')).find(el => {
                    if (!isElementVisible(el) || el.closest('#asik-ckg-panel, #asik-dashboard-overlay')) return false;
                    const cls = el.className || '';
                    const src = el.src || '';
                    return (cls.includes('w-[24px]') && cls.includes('h-[24px]')) || src.includes('icon-arrow-left') || src.includes('arrow-left');
                });
                if (backBtn) triggerClick(backBtn);
                else window.history.back();
            }

            GM_setValue('asik_ckg_sweep_attempts', {});
            setState({
                ...state,
                sweepMode: true,
                sweepSheetIndex: startSweepIdx,
                sweepCurrentStudent: null,
                sweepEmptyCheckCount: 0,
                running: true,
                phase: 'MENCARI',
                needsClassSearch: true,
                searchRetry: 0,
                tatalaksanaSelesai: false,
                isTatalaksana: false
            });
            console.log(`[CKG] 🧹 Memulai mode penyisiran tab 'Sedang Pemeriksaan' mulai dari indeks ${startSweepIdx} (${kelasAwalInfo})...`);
            jalankanLoop();
        });

        document.getElementById('asik-ckg-btn-reset').addEventListener('click', () => {
            if (confirm("Reset seluruh data dan progress?")) {
                GM_setValue('asik_ckg_sweep_attempts', {});
                setState({ data: [], index: 0, running: false, phase: 'IDLE', sekolahTarget: '', sweepMode: false, sweepSheetIndex: 0, sweepCurrentStudent: null });
                document.getElementById('asik-ckg-file').value = '';
                document.getElementById('asik-ckg-sheet').style.display = 'none';
            }
        });

        document.getElementById('asik-ckg-s-sekolah').addEventListener('click', () => {
            const state = getState();
            const baru = prompt("Masukkan nama sekolah target (misal: PALMERIAM 01):", state.sekolahTarget || "");
            if (baru !== null) {
                state.sekolahTarget = baru.trim();
                setState(state);
            }
        });

        document.getElementById('asik-ckg-btn-prev').addEventListener('click', () => {
            const state = getState();
            if (state.data.length === 0) return alert("Belum ada data Excel yang dimuat!");
            if (state.index > 0) {
                sessionStorage.removeItem('asik_ckg_active_student');
                sessionStorage.removeItem('asik_ckg_active_student_name');
                state.sweepCurrentStudent = null;
                state.index -= 1;
                state.phase = 'MENCARI';
                state.searchRetry = 0;
                state.tatalaksanaSelesai = false;
                state.isTatalaksana = false;
                setState(state);
                console.log(`[CKG] Mundur ke Siswa #${state.index + 1}: ${state.data[state.index].NAMA}`);
            } else {
                alert("Sudah berada di siswa pertama (No. 1).");
            }
        });

        document.getElementById('asik-ckg-btn-next').addEventListener('click', () => {
            const state = getState();
            if (state.data.length === 0) return alert("Belum ada data Excel yang dimuat!");
            if (state.index < state.data.length - 1) {
                sessionStorage.removeItem('asik_ckg_active_student');
                sessionStorage.removeItem('asik_ckg_active_student_name');
                state.sweepCurrentStudent = null;
                const skippedStudent = state.data[state.index];
                catatRiwayat(skippedStudent.NAMA, skippedStudent.KELAS, 'Skip (Manual)', 'Dilewati secara manual melalui tombol NEXT');
                state.index += 1;
                state.phase = 'MENCARI';
                state.searchRetry = 0;
                state.tatalaksanaSelesai = false;
                state.isTatalaksana = false;
                setState(state);
                console.log(`[CKG] Maju/Skip ke Siswa #${state.index + 1}: ${state.data[state.index].NAMA}`);
            } else {
                alert("Sudah berada di siswa terakhir.");
            }
        });

        document.getElementById('asik-ckg-btn-jump').addEventListener('click', () => {
            const state = getState();
            if (state.data.length === 0) return alert("Belum ada data Excel yang dimuat!");
            const input = prompt(`Masukkan nomor urut siswa yang ingin dikerjakan (1 - ${state.data.length}):`, (state.index + 1).toString());
            if (!input) return;
            const num = parseInt(input.trim(), 10);
            if (!isNaN(num) && num >= 1 && num <= state.data.length) {
                sessionStorage.removeItem('asik_ckg_active_student');
                sessionStorage.removeItem('asik_ckg_active_student_name');
                state.sweepCurrentStudent = null;
                state.index = num - 1;
                state.phase = 'MENCARI';
                state.searchRetry = 0;
                state.tatalaksanaSelesai = false;
                state.isTatalaksana = false;
                setState(state);
                console.log(`[CKG] Berpindah ke Siswa #${num}: ${state.data[state.index].NAMA}`);
            } else {
                alert(`Nomor tidak valid! Harap masukkan angka antara 1 dan ${state.data.length}.`);
            }
        });

        document.getElementById('asik-ckg-s-progress').addEventListener('click', () => {
            document.getElementById('asik-ckg-btn-jump').click();
        });

        document.getElementById('asik-ckg-btn-laporan').addEventListener('click', () => {
            tampilkanDashboard();
        });
    }

    function updateUI() {
        const sRobot = document.getElementById('asik-ckg-s-robot');
        const sSekolah = document.getElementById('asik-ckg-s-sekolah');
        const sNama = document.getElementById('asik-ckg-s-nama');
        const sProgress = document.getElementById('asik-ckg-s-progress');
        const btnMain = document.getElementById('asik-ckg-btn-main');

        if (!sRobot) return;
        const state = getState();

        if (state.sweepMode) {
            sRobot.innerHTML = state.running ? '<span class="asik-dot asik-dot-green"></span>SAPU SEDANG PEMERIKSAAN' : '<span class="asik-dot asik-dot-yellow"></span>PAUSED (SWEEP)';
            sSekolah.textContent = state.sekolahTarget || '-';
            sNama.textContent = state.sweepCurrentStudent ? state.sweepCurrentStudent.NAMA : 'Menyisir Tab...';

            let wbData = null;
            try {
                const raw = GM_getValue('asik_ckg_workbook_data') || localStorage.getItem('asik_ckg_workbook_data');
                if (raw) wbData = typeof raw === 'string' ? JSON.parse(raw) : raw;
            } catch (e) { }
            const totalSheets = (wbData && wbData.sheetNames) ? wbData.sheetNames.length : 1;
            sProgress.textContent = `Kelas ${(state.sweepSheetIndex || 0) + 1} / ${totalSheets}`;
        } else if (state.data.length === 0) {
            sRobot.innerHTML = '<span class="asik-dot asik-dot-red"></span>NO DATA';
            sSekolah.textContent = '-';
            sNama.textContent = '-';
            sProgress.textContent = '0 / 0';
        } else {
            sRobot.innerHTML = state.running ? '<span class="asik-dot asik-dot-green"></span>RUNNING' : '<span class="asik-dot asik-dot-yellow"></span>PAUSED';
            sSekolah.textContent = state.sekolahTarget || '-';

            if (state.index < state.data.length) {
                const cur = state.data[state.index];
                sNama.textContent = cur.NAMA || '-';
            } else {
                sNama.textContent = 'Selesai';
            }

            sProgress.textContent = `${state.index} / ${state.data.length}`;
        }

        if (btnMain) {
            btnMain.className = `asik-btn ${state.running ? 'asik-btn-red' : 'asik-btn-green'}`;
            btnMain.textContent = state.running ? 'STOP ROBOT' : 'START ROBOT';
        }
    }

    // Pantau URL untuk memunculkan UI
    setInterval(() => {
        const html = document.body.innerText.toLowerCase();
        // Munculkan UI selama di menu CKG (Pelayanan atau Form)
        if (window.location.href.toLowerCase().includes('ckg') || html.includes('belum/sedang pemeriksaan') || html.includes('skrining mandiri')) {
            createUI();
            updateUI();
        } else {
            const panel = document.getElementById('asik-ckg-panel');
            if (panel) panel.style.display = 'none'; // Sembunyikan jika keluar dari menu CKG
        }
    }, 2000);

    // Auto-resume jika di-refresh / pindah halaman
    setTimeout(() => {
        const state = getState();
        if (state.running && state.data.length > 0) {
            console.log("[CKG] Melanjutkan proses otomatisasi...");
            jalankanLoop();
        }
    }, 1500);

})();
