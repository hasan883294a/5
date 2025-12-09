// app.js — خواندن دقیق ستون‌های بدهکار/بستانکار با هدرهای «(﷼)»

(function(){
  const el = id => document.getElementById(id);
  const fileInput = el('fileInput');

  // تبدیل ارقام فارسی/عربی به لاتین
  function toLatinDigits(s){
    if (s == null) return s;
    const map = {
      '۰': '0','۱':'1','۲':'2','۳':'3','۴':'4','۵':'5','۶':'6','۷':'7','۸':'8','۹':'9',
      '٠': '0','١':'1','٢':'2','٣':'3','٤':'4','٥':'5','٦':'6','٧':'7','٨':'8','٩':'9'
    };
    return String(s).replace(/[۰-۹٠-٩]/g, ch => map[ch] ?? ch);
  }

  // نرمال‌سازی کلیدها: حذف نیم‌فاصله و کاراکترهای صفر-عرض
  function normalizeKey(k){
    if (k == null) return k;
    return String(k).replace(/[\u200B\u200C\u200D\u2060\uFEFF]/g, '').trim();
  }

  // تبدیل مقدار به عدد (پذیرش جداکننده هزار، ارقام فارسی، ﷼ و ...)
  function toNumber(v){
    if (v == null) return null;
    let s = toLatinDigits(String(v));
    // حذف نماد ریال و هر حرف غیرعددی بجز -, . , ,
    s = s.replace(/[^\d\-\.,]/g, '');
    // اگر فقط ویرگول هست و نقطه نیست، ویرگول را اعشار فرض کن
    const hasComma = s.includes(',');
    const hasDot = s.includes('.');
    if (hasComma && !hasDot) {
      s = s.replace(/,/g, '.');
    } else {
      // کاما را جداکننده هزار فرض و حذف کن
      s = s.replace(/,/g, '');
    }
    const n = Number(s);
    return Number.isFinite(n) ? n : null;
  }

  // پیدا کردن مقدار با نام‌های ممکن
  function findValue(row, candidates){
    if (!row) return null;
    const keys = Object.keys(row);
    // تطابق دقیق
    for (const k of keys) {
      const nk = normalizeKey(k);
      for (const name of candidates) {
        if (nk === normalizeKey(name)) return row[k];
      }
    }
    // تطابق جزئی (مثلاً شامل «بدهکار» + «﷼»)
    for (const k of keys) {
      const nk = normalizeKey(k);
      for (const name of candidates) {
        const nn = normalizeKey(name);
        if (nk.includes(nn)) return row[k];
      }
    }
    return null;
  }

  if (!fileInput) {
    alert('عنصر fileInput یافت نشد. مطمئن شو id="fileInput" در index.html وجود دارد.');
    return;
  }

  fileInput.addEventListener('change', async (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;

    try {
      const data = await file.arrayBuffer();

      if (typeof XLSX === 'undefined') {
        alert('کتابخانه XLSX لود نشده است. ابتدا xlsx.full.min.js را لود کن.');
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
        console.warn('شیت اول خالی است یا داده‌ای قابل تبدیل ندارد.');
        return;
      }

      // هدرهای واقعی که sheet_to_json تولید کرده
      console.log('کلیدهای ردیف اول:', Object.keys(rows[0]).map(k => normalizeKey(k)));

      // نام‌های دقیق و چند حالت نزدیک برای ایمنی
      const creditNames = ['بستانکار (﷼)', 'بستانکار(﷼)', 'بستانکار ‌(﷼)'];
      const debitNames  = ['بدهکار (﷼)',  'بدهکار(﷼)',  'بدهکار ‌(﷼)'];

      // نمایش 20 ردیف اول برای بررسی
      let sumDebit = 0, sumCredit = 0;
      rows.forEach((r, idx) => {
        const rawDebit  = findValue(r, debitNames);
        const rawCredit = findValue(r, creditNames);
        const debit  = toNumber(rawDebit);
        const credit = toNumber(rawCredit);

        if (idx < 20) {
          console.log('ردیف ' + (idx+1) + ' → بدهکار خام:', rawDebit, '| بدهکار عددی:', debit);
          console.log('ردیف ' + (idx+1) + ' → بستانکار خام:', rawCredit, '| بستانکار عددی:', credit);
        }

        if (debit != null)  sumDebit  += debit;
        if (credit != null) sumCredit += credit;
      });

      console.log('جمع بدهکار:', sumDebit, ' | جمع بستانکار:', sumCredit);
      alert('خواندن ستون‌ها انجام شد. جمع بدهکار/بستانکار در کنسول چاپ شد.');
      } catch (err) {
      alert('خطا در پردازش فایل: ' + err.message);
      console.error(err);
    }
  });
})();
