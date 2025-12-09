// app.js — استخراج ستون‌های 3,4,6,7 از سمت راست از ردیف دوم به بعد
(function(){
  const el = id => document.getElementById(id);
  const fileInput = el('fileInput');

  function toLatinDigits(s){
    if (s == null) return s;
    const map = {'۰':'0','۱':'1','۲':'2','۳':'3','۴':'4','۵':'5','۶':'6','۷':'7','۸':'8','۹':'9',
                 '٠':'0','١':'1','٢':'2','٣':'3','٤':'4','٥':'5','٦':'6','٧':'7','٨':'8','٩':'9'};
    return String(s).replace(/[۰-۹٠-٩]/g, ch => map[ch] ?? ch);
  }

  function normalizeString(s){
    if (s == null) return s;
    return String(s).replace(/[\u200B\u200C\u200D\u2060\uFEFF]/g, '').trim();
  }

  function toNumber(v){
    if (v == null) return null;
    let s = toLatinDigits(String(v));
    // حذف نماد ریال و هر کاراکتر غیرعددی به جز - . ,
    s = s.replace(/[^\d\-\.,]/g, '');
    const hasComma = s.includes(','), hasDot = s.includes('.');
    if (hasComma && !hasDot) s = s.replace(/,/g, '.'); else s = s.replace(/,/g, '');
    const n = Number(s);
    return Number.isFinite(n) ? n : null;
  }

  if (!fileInput) {
    console.warn('عنصر fileInput پیدا نشد. مطمئن شو id="fileInput" در index.html وجود دارد.');
    return;
  }

  fileInput.addEventListener('change', async (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;

    try {
      const data = await file.arrayBuffer();

      if (typeof XLSX === 'undefined') {
        alert('کتابخانه XLSX لود نشده. ابتدا xlsx.full.min.js را اضافه کن.');
        return;
      }

      // خواندن شیت به صورت آرایه‌ای (header:1)
      const wb = XLSX.read(data, { type: 'array', cellDates: true });
      const firstSheetName = wb.SheetNames && wb.SheetNames[0];
      if (!firstSheetName) {
        alert('هیچ شیتی پیدا نشد.');
        return;
      }

      const ws = wb.Sheets[firstSheetName];
      const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null }); // آرایه آرایه‌ها
      if (!raw || raw.length === 0) {
        alert('شیت خالی است.');
        return;
      }

      // نمایش هدرها (ردیف اول)
      const headerRow = raw[0].map(h => normalizeString(h));
      console.log('هدرها از چپ به راست:', headerRow);

      // تعداد ستون‌ها
      const colCount = headerRow.length;

      // موقعیت‌های مورد نظر از سمت راست: 3,4,6,7
      const positionsFromRight = [3,4,6,7];

      // تبدیل به اندیس از چپ (0-based)
      const indices = positionsFromRight.map(pos => {
        const idx = colCount - pos;
        return idx >= 0 ? idx : null;
      });

      console.log('اندیس‌های استخراج شده از چپ (0-based):', indices);

      // پردازش ردیف‌ها از ردیف دوم به بعد (یعنی raw[1] به بعد)
      const results = [];
      for (let r = 1; r < raw.length; r++) {
        const row = raw[r];
        // اگر ردیف کوتاه است، آن را با null پر کن
        const fullRow = Array.from({length: colCount}, (_, i) => row[i] ?? null);
        const out = {};
        indices.forEach((idx, i) => {
          const pos = positionsFromRight[i];
          if (idx === null) {
            out[pos_${pos}] = { header: null, raw: null, value: null };
          } else {
            const rawVal = fullRow[idx];
            out[pos_${pos}] = {
              header: headerRow[idx],
              raw: rawVal,
              value: toNumber(rawVal)
            };
          }
        });
        results.push(out);
      }

      // چاپ نمونه 20 ردیف اول استخراج شده
      console.log('نمونه استخراج ستون‌های 3,4,6,7 از سمت راست (20 ردیف اول):', results.slice(0,20));

      // جمع‌زدن مقادیر هر ستون (در صورت نیاز)
      const sums = {};
      positionsFromRight.forEach(pos => sums[pos_${pos}] = 0);
      results.forEach(r => {
        positionsFromRight.forEach(pos => {
          const v = r[pos_${pos}].value;
          if (v != null) sums[pos_${pos}] += v;
        });
      });
      console.log('جمع مقادیر هر ستون استخراج شده:', sums);

      alert('استخراج انجام شد. خروجی نمونه و جمع‌ها در کنسول چاپ شدند.');
      // اگر خواستی می‌تونم همین results را به فرمت دلخواه آماده کنم یا فایل خروجی بسازم
    } catch (err) {
      console.error(err);
      alert('خطا در پردازش فایل: ' + err.message);
    }
  });
})();
