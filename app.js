// نسخه‌ی تستی امن بدون template string برای جلوگیری از خطای نحو

(function(){
  var el = function(id){ return document.getElementById(id); };
  var fileInput = el('fileInput');

  if (!fileInput) {
    alert('عنصر fileInput پیدا نشد. مطمئن شو id="fileInput" در index.html وجود دارد.');
    return;
  }

  fileInput.addEventListener('change', function(e){
    var file = e.target.files && e.target.files[0];
    if (!file) return;

    file.arrayBuffer().then(function(data){
      // مطمئن شو که XLSX قبلاً لود شده
      if (typeof XLSX === 'undefined') {
        alert('کتابخانه XLSX لود نشده است. اسکریپت xlsx.full.min.js را قبل از app.js اضافه کن.');
        return;
      }

      var wb = XLSX.read(data, { type: 'array', cellDates: true, dateNF: 'yyyy-mm-dd' });
      var firstName = wb.SheetNames && wb.SheetNames[0];
      if (!firstName) {
        alert('هیچ شیتی در فایل پیدا نشد.');
        return;
      }
      var ws = wb.Sheets[firstName];
      var rows = XLSX.utils.sheet_to_json(ws, { defval: null, raw: false });

      console.log('تعداد ردیف‌ها: ' + rows.length);

      // چاپ کلیدهای هر ردیف
      rows.forEach(function(r, idx){
        console.log('ردیف ' + (idx+1) + ':', Object.keys(r));
      });

      // تست خواندن ستون‌های بدهکار/بستانکار
      rows.forEach(function(r, idx){
        var debit = r['بدهکار (ریال)'];
        var credit = r['بستانکار (ریال)'];
        console.log('ردیف ' + (idx+1) + ' → بدهکار:', debit, 'بستانکار:', credit);
      });

      alert('کلیدهای ستون‌ها و مقادیر در کنسول چاپ شدند. لطفاً بررسی کن.');
    }).catch(function(err){
      alert('خطا در خواندن فایل: ' + err.message);
      console.error(err);
    });
  });
})();
