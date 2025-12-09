// app.js — نسخه مقاوم برای شناسایی هدرها و خواندن بدهکار/بستانکار

(function(){
  const el = id => document.getElementById(id);
  const fileInput = el('fileInput');

  // تبدیل ارقام فارسی به لاتین
  function persianToLatinDigits(s){
    if (s == null) return s;
    return String(s).replace(/[۰-۹]/g, d => String('۰۱۲۳۴۵۶۷۸۹'.indexOf(d)));
  }

  // حذف کاراکترهای Zero-width و trim
  function normalizeKey(k){
    if (k == null) return k;
    return String(k).replace(/[\u200B\u200C\uFEFF]/g, '').trim();
  }

  // تبدیل مقدار به عدد (مقاوم به جداکننده هزار و ارقام فارسی)
  function toNumber(v){
    if (v == null) return null;
    let s = String(v);
    s = persianToLatinDigits(s);
    // حذف هر چیزی به جز ارقام، منفی و نقطه و ویرگول
    s = s.replace(/[^\d\-\.,]/g, '');
    // اگر ویرگول به عنوان جداکننده اعشار استفاده شده، تبدیل به نقطه
    const commaCount = (s.match(/,/g) || []).length;
    const dotCount = (s.match(/\./g) || []).length;
    if (commaCount > 0 && dotCount === 0) {
      s = s.replace(/,/g, '.');
    } else {
      s = s.replace(/,/g, '');
    }
    const n = Number(s);
    return Number.isFinite(n) ? n : null;
  }

  // جستجوی کلید مناسب در ردیف با چند الگو
  function findValueByPossibleNames(row, names){
    if (!row) return null;
    for (const k of Object.keys(row)) {
      const nk = normalizeKey(k);
      for (const name of names) {
        if (nk === normalizeKey(name)) return row[k];
      }
    }
    // fallback: جستجوی جزئی (مثلاً فقط 'بدهکار' داخل هدر)
    for (const k of Object.keys(row)) {
      const nk = normalizeKey(k);
      for (const name of names) {
        if (nk.includes(normalizeKey(name))) return row[k];
      }
    }
    return null;
  }

  if (!fileInput) {
    console.warn('عنصر fileInput پیدا نشد. مطمئن شو id="fileInput" در index.html وجود دارد.');
    return;
  }

  fileInput.addEventListener('change', async function(e){
    const file = e.target.files && e.target.files[0];
    if (!file) return;

    try {
      const data = await file.arrayBuffer();

      if (typeof XLSX === 'undefined') {
        alert('کتابخانه XLSX لود نشده. لطفاً xlsx.full.min.js را قبل از app.js اضافه کن.');
        return;
      }

      const wb = XLSX.read(data, { type: 'array', cellDates: true, dateNF: 'yyyy-mm-dd' });
      const firstSheetName = wb.SheetNames && wb.SheetNames[0];
      if (!firstSheetName) {
        alert('هیچ شیتی در فایل پیدا نشد.');
        return;
      }

      const ws = wb.Sheets[firstSheetName];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: null, raw: false });

      console.log('تعداد ردیف‌ها:', rows.length);

      if (rows.length === 0) {
        console.warn('شیت اول خالی است یا داده‌ای برای تبدیل وجود ندارد.');
        return;
      }

      // فقط ردیف اول را چاپ می‌کنیم تا هدرها را ببینی
      console.log('کلیدهای ردیف اول (هدرها یا کلیدهای خروجی sheet_to_json):', Object.keys(rows[0]).map(k => normalizeKey(k)));

      // نام‌های ممکن برای ستون‌ها — اگر نام دقیق دیگری داری اینجا اضافه کن
      const debitNames = ['بدهکار (ریال)', 'بدهکار', 'بدهکار(ریال)', 'بدهکار ریال'];
      const creditNames = ['بستانکار (ریال)', 'بستانکار', 'بستانکار(ریال)', 'بستانکار ریال'];

      // پردازش همه ردیف‌ها و چاپ نمونه
      rows.forEach((r, idx) => {
        const rawDebit = findValueByPossibleNames(r, debitNames);
        const rawCredit = findValueByPossibleNames(r, creditNames);
        const debit = toNumber(rawDebit);
        const credit = toNumber(rawCredit);

        // فقط چند ردیف اول را با جزئیات بیشتر چاپ کن
        if (idx < 20) {
          console.log(
            'ردیف ' + (idx+1) + ' → کلیدها:',
            Object.keys(r).map(k => normalizeKey(k))
          );
          console.log('  مقدار خام بدهکار:', rawDebit, '→ عدد:', debit);
          console.log('  مقدار خام بستانکار:', rawCredit, '→ عدد:', credit);
        }
      });

      alert('چاپ کنسول انجام شد. کلیدهای ردیف اول را بررسی کن و برایم بفرست.');

    } catch (err) {
      alert('خطا در پردازش فایل: ' + err.message);
      console.error(err);
    }
  });
})();
