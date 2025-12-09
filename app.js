fileInput.addEventListener('change', async (e) => {
  const file = e.target.files?.[0];
  if (!file) return;

  try {
    const data = await file.arrayBuffer();
    const wb = XLSX.read(data, { type: 'array', cellDates: true, dateNF: 'yyyy-mm-dd' });

    const outWb = XLSX.utils.book_new();

    // فقط شیت اول برای تست
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: null, raw: false });

    // چاپ کلیدهای هر ردیف
    rows.forEach((r, idx) => {
      console.log(ردیف ${idx+1}:, Object.keys(r));
    });

    // تست خواندن ستون‌ها
    rows.forEach((r, idx) => {
      const debit = r['بدهکار (ریال)'];
      const credit = r['بستانکار (ریال)'];
      console.log(ردیف ${idx+1} → بدهکار:, debit, 'بستانکار:', credit);
    });

    alert('کلیدهای ستون‌ها در کنسول مرورگر چاپ شدند. لطفاً بررسی کن.');
  } catch (err) {
    alert('خطا در پردازش فایل: ' + err.message);
    console.error(err);
  }
});
