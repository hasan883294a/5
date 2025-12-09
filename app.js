// app.js — پردازش اکسل و تبدیل ستون‌های 6 و 7 از سمت راست به عدد
(function(){
  const fileInput = document.getElementById('fileInput');

  // تبدیل رشته‌های عددی فارسی/عربی با جداکننده هزارگان به عدد جاوااسکریپت
  function convertPersianNumberStringToNumber(str) {
    if (!str || typeof str !== 'string') return str;
    const cleaned = str
      .replace(/٬|,/g, '') // حذف جداکننده هزارگان فارسی و انگلیسی
      .replace(/[۰-۹]/g, d => '۰۱۲۳۴۵۶۷۸۹'.indexOf(d)) // فارسی
      .replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d)); // عربی
    const num = parseFloat(cleaned);
    return isNaN(num) ? str : num;
  }

  if (!fileInput) {
    console.warn('عنصر fileInput پیدا نشد. مطمئن شو در index.html وجود دارد.');
    return;
  }

  fileInput.addEventListener('change', async (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;

    try {
      const data = await file.arrayBuffer();
      if (typeof XLSX === 'undefined') {
        alert('کتابخانه XLSX لود نشده است. ابتدا xlsx.full.min.js را اضافه کن.');
        return;
      }

      const wb = XLSX.read(data, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

      if (!raw || raw.length < 2) {
        alert('شیت خالی یا بدون داده است.');
        return;
      }

      const headers = raw[0].map(h => h == null ? null : String(h).trim());
      const colCount = headers.length;
      console.log('هدرها از چپ به راست:', headers);

      // محاسبه اندیس ستون‌های 6 و 7 از سمت راست
      const idx6 = colCount - 6;
      const idx7 = colCount - 7;

      // پردازش ردیف‌ها از ردیف دوم به بعد
      const results = [];
      for (let r = 1; r < raw.length; r++) {
        const row = raw[r];
        if (row.length < colCount) {
          for (let k = row.length; k < colCount; k++) row[k] = null;
        }
        const val6 = convertPersianNumberStringToNumber(row[idx6]);
        const val7 = convertPersianNumberStringToNumber(row[idx7]);
        results.push({
          rowIndex: r+1,
          col6: { header: headers[idx6], raw: row[idx6], value: val6 },
          col7: { header: headers[idx7], raw: row[idx7], value: val7 }
        });
      }

      // چاپ نمونه 20 ردیف اول
      console.log('نمونه ستون‌های 6 و 7 (20 ردیف اول):', results.slice(0,20));

      // جمع مقادیر
      let sum6 = 0, sum7 = 0;
      results.forEach(r => {
        if (typeof r.col6.value === 'number') sum6 += r.col6.value;
        if (typeof r.col7.value === 'number') sum7 += r.col7.value;
      });
      console.log('جمع ستون 6:', sum6, 'جمع ستون 7:', sum7);

      alert('تبدیل ستون‌های 6 و 7 به عدد انجام شد. خروجی در کنسول است.');
    } catch (err) {
      console.error(err);
      alert('خطا در پردازش فایل: ' + err.message);
    }
  });
})();
