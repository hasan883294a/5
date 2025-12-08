
const el = (id) => document.getElementById(id);
const fileInput = el('fileInput');
const summaryEl = el('summary');
const summaryContentEl = el('summaryContent');
const downloadEl = el('download');
const downloadBtn = el('downloadBtn');

let config = null;
let analyzedBlob = null;

(async function init(){
  try {
    config = await fetch('templates/config.json').then(r => r.json());
  } catch {
    alert('عدم دسترسی به templates/config.json');
  }
})();

fileInput.addEventListener('change', async (e) => {
  const file = e.target.files?.[0];
  if (!file) return;

  try {
    const data = await file.arrayBuffer();
    const wb = XLSX.read(data, { type: 'array', cellDates: true, dateNF: 'yyyy-mm-dd' });

    const invoiceName = config.invoiceSheetName;
    const aggregateName = config.aggregateSheetName;
    const summaryName = config.summarySheetName;

    // 1) کپی مستقیم شیت «صورت حساب»
    const outWb = XLSX.utils.book_new();
    const invoiceWs = wb.Sheets[invoiceName] || wb.Sheets[wb.SheetNames[0]];
    if (!invoiceWs) throw new Error('شیت «صورت حساب» یافت نشد.');
    XLSX.utils.book_append_sheet(outWb, deepCopySheet(invoiceWs), invoiceName);

    // 2) تجمیع سایر شیت‌ها
    const aggRows = [];
    const headerMap = config.headers || {};
    const sheetCategories = config.categories;

    for (let i = 0; i < wb.SheetNames.length; i++) {
      const name = wb.SheetNames[i];
      if (i === 0 && name === invoiceName) continue; // شیت اول: صورت حساب
      const ws = wb.Sheets[name];
      if (!ws) continue;
      const rows = XLSX.utils.sheet_to_json(ws, { defval: null, raw: false });
      for (const r of rows) {
        const variantCode = pickValue(r, headerMap.variantCode);
        const variantTitle = pickValue(r, headerMap.variantTitle);
        const debit = toNumber(r['F']);
        const credit = toNumber(r['G']);
        aggRows.push({
          'کد تنوع': variantCode ?? null,
          'عنوان تنوع': variantTitle ?? null,
          'بدهکار': debit,
          'بستانکار': credit,
          'نام شیت منبع': name
        });
      }
    }

    const aggWs = XLSX.utils.json_to_sheet(aggRows);
    applyThousandFormat(aggWs, ['C', 'D']); // C=بدهکار, D=بستانکار
    XLSX.utils.book_append_sheet(outWb, aggWs, aggregateName);

    // 3) ساخت «خلاصه داده»
    const summaryRows = buildSummary(aggRows, sheetCategories);
    const summaryWs = XLSX.utils.json_to_sheet(summaryRows);
    applyThousandFormat(summaryWs, ['F', 'G', 'H', 'I']); // جمع بدهکار/بستانکار/درآمد/درآمد به ازای هر فروش
    XLSX.utils.book_append_sheet(outWb, summaryWs, summaryName);

    // 4) خروجی
    const out = XLSX.write(outWb, { type: 'array', bookType: 'xlsx' });
    analyzedBlob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

    renderSummaryUI(summaryRows);
    downloadEl.classList.remove('hidden');
    summaryEl.classList.remove('hidden');

  } catch (err) {
    alert('خطا در پردازش فایل: ' + err.message);
    console.error(err);
  }
});

downloadBtn.addEventListener('click', () => {
  if (!analyzedBlob) return;
  const url = URL.createObjectURL(analyzedBlob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'invoice_analysis.xlsx';
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
});

// ——— توابع کمکی ———

function deepCopySheet(ws){
  const json = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: true });
  const out = XLSX.utils.aoa_to_sheet(json);
  return out;
}

function pickValue(row, candidates){
  if (!candidates) return null;
  for (const key of candidates) {
    if (key in row && row[key] != null && row[key] !== '') return row[key];
  }
  return null;
}

function toNumber(v){
  if (v == null) return null;
  const s = String(v).replace(/[^\d\-., ]/g, '').replace(/[, ]/g, '');
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}

function applyThousandFormat(ws, cols){
  const range = XLSX.utils.decode_range(ws['!ref']);
  for (const col of cols) {
    for (let R = range.s.r + 1; R <= range.e.r; R++) {
      const addr = `${col}${R+1}`;
      const cell = ws[addr];
      if (cell && typeof cell.v === 'number') {
        cell.t = 'n';
        cell.z = config.numberFormat || '#,##0';
      }
    }
  }
}

function buildSummary(aggRows, categories){
  const catNames = {
    sale: new Set(categories.sale || []),
    sale_credit: new Set(categories.sale_credit || []),
    return_sale: new Set(categories.return_sale || []),
    return_sale_credit: new Set(categories.return_sale_credit || [])
  };

  const groups = {};
  for (const r of aggRows) {
    const code = r['کد تنوع'];
    const title = r['عنوان تنوع'];
    const src = r['نام شیت منبع'];
    const debit = r['بدهکار'] ?? 0;
    const credit = r['بستانکار'] ?? 0;

    if (!groups[code]) {
      groups[code] = {
        'کد تنوع': code,
        'عنوان تنوع': title,
        'تعداد در فروش': 0,
        'تعداد در فروش اعتباری': 0,
        'تعداد در برگشت از فروش': 0,
        'تعداد در برگشت از فروش اعتباری': 0,
        'فروش خالص': 0,
        'جمع بدهکار': 0,
        'جمع بستانکار': 0,
        'درآمد': 0,
        'درآمد به ازای هر فروش': 0
      };
    }

    const g = groups[code];
    if (catNames.sale.has(src)) g['تعداد در فروش']++;
    if (catNames.sale_credit.has(src)) g['تعداد در فروش اعتباری']++;
    if (catNames.return_sale.has(src)) g['تعداد در برگشت از فروش']++;
    if (catNames.return_sale_credit.has(src)) g['تعداد در برگشت از فروش اعتباری']++;

    g['جمع بدهکار'] += debit;
    g['جمع بستانکار'] += credit;
  }

  for (const code of Object.keys(groups)) {
    const g = groups[code];
    g['فروش خالص'] = g['تعداد در فروش'] + g['تعداد در فروش اعتباری'] - g['تعداد در برگشت از فروش'] - g['تعداد در برگشت از فروش اعتباری'];
    g['درآمد'] = g['جمع بستانکار'] - g['جمع بدهکار'];
    g['درآمد به ازای هر فروش'] = g['فروش خالص'] > 0 ? (g['درآمد'] / g['فروش خالص']) : 0;
  }

  return Object.values(groups).sort((a,b) => b['درآمد'] - a['درآمد']);
}

function renderSummaryUI(summary){
  summaryContentEl.innerHTML = '';
  const totals = {
    countCodes: summary.length,
    totalDebit: summary.reduce((s,r)=> s + r['جمع بدهکار'], 0),
    totalCredit: summary.reduce((s,r)=> s + r['جمع بستانکار'], 0),
    totalNetSales: summary.reduce((s,r)=> s + r['فروش خالص'], 0),
    totalRevenue: summary.reduce((s,r)=> s + r['درآمد'], 0)
  };
  const items = [
    { k: 'تعداد کدهای تنوع', v: formatThousand(totals.countCodes) },
    { k: 'جمع بدهکار', v: formatThousand(totals.totalDebit) },
    { k: 'جمع بستانکار', v: formatThousand(totals.totalCredit) },
    { k: 'فروش خالص (تعداد)', v: formatThousand(totals.totalNetSales) },
    { k: 'جمع درآمد', v: formatThousand(totals.totalRevenue) }
  ];
  for (const it of items) {
    const div = document.createElement('div');
    div.className = 'kv';
    div.innerHTML = `<b>${it.k}</b><span>${it.v}</span>`;
    summaryContentEl.appendChild(div);
  }
}

function formatThousand(n){
  if (typeof n !== 'number') return String(n ?? '');
  return n.toLocaleString('fa-IR');
}
