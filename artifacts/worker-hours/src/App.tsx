import { useEffect, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import { Calculator, Check, CircleHelp, Download, FileSpreadsheet, FileUp, Plus, Trash2, Users, X } from 'lucide-react';

type Worker = { id: string; name: string; totalHours: number; hourlyRate: number };
type SavedState = { rows: Worker[]; fileName: string };
const STORAGE_KEY = 'daftar-worker-hours-v1';

function makeWorker(name = '', totalHours = 0, hourlyRate = 0): Worker {
  return { id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`, name, totalHours, hourlyRate };
}

function formatNumber(value: number) {
  return new Intl.NumberFormat('ar-IQ', { maximumFractionDigits: 2 }).format(Number.isFinite(value) ? value : 0);
}

function numericCell(value: unknown): number {
  if (typeof value === 'number') return value;
  const raw = String(value ?? '').trim().replace(/[٫،]/g, '.').replace(/,/g, '');
  if (!raw) return 0;
  if (/^\d{1,3}:\d{2}(:\d{2})?$/.test(raw)) {
    const parts = raw.split(':').map(Number);
    const hours = parts[0] ?? 0;
    const minutes = parts[1] ?? 0;
    const seconds = parts[2] ?? 0;
    return hours + minutes / 60 + seconds / 3600;
  }
  const parsed = Number(raw);
  return Number.isFinite(parsed) ? parsed : 0;
}

function normalizeHeader(value: unknown) {
  return String(value ?? '')
    .toLowerCase()
    .normalize('NFKC')
    .replace(/[إأآ]/g, 'ا')
    .replace(/ى/g, 'ي')
    .replace(/[\s_\-ـ()[\]]/g, '');
}

async function importWorkbook(file: File): Promise<Worker[]> {
  const bytes = await file.arrayBuffer();
  const workbook = XLSX.read(bytes, { type: 'array', cellDates: true });
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  if (!firstSheet) throw new Error('لم نعثر على ورقة بيانات في هذا الملف.');
  const grid = XLSX.utils.sheet_to_json<unknown[]>(firstSheet, { header: 1, defval: '' });
  let headerRow = -1;
  let nameCol = -1;
  let hoursCol = -1;
  let fallbackHours: number[] = [];
  for (let r = 0; r < Math.min(grid.length, 30); r += 1) {
    const headers = (grid[r] ?? []).map(normalizeHeader);
    const possibleName = headers.findIndex((h) => /^(الإسم|الاسم|اسمالعامل|اسم|employee|worker|name)/.test(h));
    if (possibleName < 0) continue;
    // Match complete total-hours labels only. A loose substring search can
    // accidentally select "كسر إجمالي الساعات" (hours fraction) before
    // the actual "مجموع الساعات" column in attendance exports.
    const total = headers.findIndex((h) =>
      /^(مجموعالساعات|اجماليالساعات|الساعاتالاجمالية|اجماليساعاتالعمل|ساعاتالعملالاجمالية|totalhours|totalworkinghours|sumofhours)$/.test(h),
    );
    if (total < 0) continue;
    headerRow = r;
    nameCol = possibleName;
    hoursCol = total;
    break;
  }
  // Some attendance exports omit a grand-total field; in that case sum only
  // columns explicitly labeled as work/activity hours, never clock-time fields.
  if (headerRow < 0) {
    for (let r = 0; r < Math.min(grid.length, 30); r += 1) {
      const headers = (grid[r] ?? []).map(normalizeHeader);
      const possibleName = headers.findIndex((h) => /^(الإسم|الاسم|اسمالعامل|اسم|employee|worker|name)/.test(h));
      if (possibleName < 0) continue;
      const columns = headers.map((h, i) => /ساعاتالعمل|اعتيادي\(h\)|regularhours|workhours/.test(h) ? i : -1).filter((i) => i >= 0);
      if (columns.length) { headerRow = r; nameCol = possibleName; fallbackHours = columns; break; }
    }
  }
  if (headerRow < 0) {
    throw new Error('لم يتم العثور على عمود الاسم وإجمالي الساعات. تحقق من عناوين الأعمدة في الملف.');
  }
  const totals = new Map<string, number>();
  for (const row of grid.slice(headerRow + 1)) {
    const name = String(row?.[nameCol] ?? '').trim();
    if (!name || /^الإجمالي|^المجموع/.test(name)) continue;
    const hours = hoursCol >= 0 ? numericCell(row?.[hoursCol]) : fallbackHours.reduce((sum, col) => sum + numericCell(row?.[col]), 0);
    if (!Number.isFinite(hours) || hours <= 0) continue;
    totals.set(name, (totals.get(name) ?? 0) + hours);
  }
  if (!totals.size) throw new Error('تم فتح الملف لكن لم نجد أسماء عمال وساعات قابلة للاستيراد.');
  return Array.from(totals, ([name, totalHours]) => makeWorker(name, Math.round(totalHours * 100) / 100, 0));
}

function initialSavedState(): SavedState {
  try {
    const parsed = JSON.parse(localStorage.getItem(STORAGE_KEY) ?? 'null') as SavedState | null;
    if (parsed && Array.isArray(parsed.rows)) return { rows: parsed.rows, fileName: parsed.fileName ?? '' };
  } catch { /* Use a clean state if browser storage is unavailable or malformed. */ }
  return { rows: [], fileName: '' };
}

function App() {
  const [saved] = useState(initialSavedState);
  const [rows, setRows] = useState<Worker[]>(saved.rows);
  const [fileName, setFileName] = useState(saved.fileName);
  const [error, setError] = useState('');
  const [dragging, setDragging] = useState(false);
  const [busy, setBusy] = useState(false);
  const fileRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    try { localStorage.setItem(STORAGE_KEY, JSON.stringify({ rows, fileName })); } catch { /* Storage may be disabled. */ }
  }, [rows, fileName]);

  const totalPayroll = rows.reduce((sum, row) => sum + (Number(row.totalHours) || 0) * (Number(row.hourlyRate) || 0), 0);
  const totalHours = rows.reduce((sum, row) => sum + (Number(row.totalHours) || 0), 0);

  function updateRow(id: string, field: 'name' | 'totalHours' | 'hourlyRate', value: string) {
    setRows((current) => current.map((row) => row.id === id
      ? { ...row, [field]: field === 'name' ? value : value === '' ? 0 : Math.max(0, Number(value) || 0) }
      : row));
  }

  async function handleFile(file?: File) {
    if (!file) return;
    setBusy(true);
    setError('');
    try {
      const imported = await importWorkbook(file);
      setRows(imported);
      setFileName(file.name);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : 'تعذر قراءة الملف. حاول استخدام ملف Excel آخر.');
    } finally {
      setBusy(false);
      if (fileRef.current) fileRef.current.value = '';
    }
  }

  function clearAll() {
    setRows([]);
    setFileName('');
    setError('');
  }

  function exportExcel() {
    if (!rows.length) return;
    const data = rows.map((row, i) => ({
      'ت': i + 1,
      'اسم العامل': row.name,
      'إجمالي الساعات': Number(row.totalHours) || 0,
      'أجر الساعة (دينار)': Number(row.hourlyRate) || 0,
      'الأجر المستحق (دينار)': (Number(row.totalHours) || 0) * (Number(row.hourlyRate) || 0),
    }));
    const worksheet = XLSX.utils.json_to_sheet(data);
    worksheet['!cols'] = [{ wch: 6 }, { wch: 28 }, { wch: 18 }, { wch: 22 }, { wch: 24 }];
    worksheet['!views'] = [{ rightToLeft: true }];
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'أجور العمال');
    XLSX.writeFile(workbook, 'كشف_أجور_العمال.xlsx');
  }

  return (
    <div className="app-shell">
      <header className="topbar">
        <div className="brand">
          <span className="brand-mark"><Calculator size={19} strokeWidth={2.2} /></span>
          <span>حساب ساعات وأجور العمال</span>
        </div>
        <span className="currency-chip">العملة: دينار عراقي</span>
      </header>

      <main className="page">
        <section className="intro">
          <div>
            <div className="eyebrow">دفتر الأجور · IRAQ</div>
            <h1>من ساعات الحضور إلى أجر واضح.</h1>
            <p>استورد ملف الدوام، راجع ساعات كل عامل، واحسب المستحقات بسهولة.</p>
          </div>
          {rows.length > 0 && (
            <div className="summary-pill">
              <div><small>إجمالي الأجور</small><strong>{formatNumber(totalPayroll)}</strong></div>
              <span className="currency-chip">د.ع</span>
            </div>
          )}
        </section>

        {error && <div className="notice" role="alert">{error}</div>}

        <section
          className={`upload-card${dragging ? ' dragging' : ''}`}
          onDragOver={(event) => { event.preventDefault(); setDragging(true); }}
          onDragLeave={() => setDragging(false)}
          onDrop={(event) => { event.preventDefault(); setDragging(false); void handleFile(event.dataTransfer.files[0]); }}
          aria-label="استيراد ملف Excel"
        >
          <input
            ref={fileRef}
            className="file-input"
            type="file"
            accept=".xlsx,.xls,.csv"
            onChange={(event) => void handleFile(event.target.files?.[0])}
          />
          <div className="upload-content">
            <div className="upload-icon">{fileName ? <Check size={22} /> : <FileSpreadsheet size={21} />}</div>
            {fileName ? (
              <>
                <div className="upload-title">تم استيراد الملف <span className="file-name">{fileName}</span></div>
                <div className="upload-sub">تم جمع إجمالي الساعات حسب اسم العامل · البيانات محفوظة على هذا الجهاز</div>
                <div className="actions">
                  <button className="btn" type="button" onClick={() => fileRef.current?.click()}><FileUp size={14} /> تغيير الملف</button>
                  <button className="btn subtle" type="button" onClick={clearAll}><X size={14} /> مسح البيانات</button>
                </div>
              </>
            ) : (
              <>
                <div className="upload-title">{busy ? 'جارٍ قراءة الملف…' : 'ابدأ باستيراد كشف الحضور'}</div>
                <div className="upload-sub">ملفات Excel بصيغة .xlsx أو .xls · تبقى بياناتك على جهازك</div>
                <div className="actions">
                  <button className="btn primary" type="button" disabled={busy} onClick={() => fileRef.current?.click()}><FileUp size={15} /> اختيار ملف Excel</button>
                </div>
              </>
            )}
          </div>
        </section>

        <section className="table-card">
          <div className="table-head">
            <div className="table-title">الخانات الرئيسية لحساب ساعات وأجور العمال <span className="count-badge">{formatNumber(rows.length)} عامل</span></div>
            <button className="btn primary" type="button" onClick={() => setRows((current) => [...current, makeWorker()])}><Plus size={15} /> إضافة عامل</button>
          </div>
          {rows.length ? (
            <>
              <div className="table-scroll">
                <table>
                  <thead><tr><th>#</th><th>اسم العامل</th><th>إجمالي الساعات</th><th>أجر الساعة (دينار)</th><th>الأجر المستحق</th><th>حذف</th></tr></thead>
                  <tbody>
                    {rows.map((row, index) => (
                      <tr key={row.id}>
                        <td className="row-number">{formatNumber(index + 1)}</td>
                        <td><input aria-label={`اسم العامل ${index + 1}`} className="field" type="text" placeholder="اكتب اسم العامل" value={row.name} onChange={(event) => updateRow(row.id, 'name', event.target.value)} /></td>
                        <td><input aria-label={`إجمالي ساعات ${index + 1}`} className="field numeric" type="number" min="0" step="0.25" value={row.totalHours || ''} placeholder="0" onChange={(event) => updateRow(row.id, 'totalHours', event.target.value)} /></td>
                        <td><input aria-label={`أجر الساعة ${index + 1}`} className="field numeric" type="number" min="0" step="250" value={row.hourlyRate || ''} placeholder="0" onChange={(event) => updateRow(row.id, 'hourlyRate', event.target.value)} /></td>
                        <td className="pay-cell">{formatNumber((Number(row.totalHours) || 0) * (Number(row.hourlyRate) || 0))}<span>د.ع</span></td>
                        <td><button className="btn danger" type="button" aria-label={`حذف ${row.name || 'العامل'}`} onClick={() => setRows((current) => current.filter((worker) => worker.id !== row.id))}><Trash2 size={15} /></button></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div className="table-foot">
                <div className="foot-note">تُحسب المستحقات تلقائياً: إجمالي الساعات × أجر الساعة</div>
                <div className="total-box"><span>الساعات الإجمالية <bdi>{formatNumber(totalHours)}</bdi></span><strong>{formatNumber(totalPayroll)} <small>د.ع</small></strong></div>
                <button className="btn" type="button" onClick={exportExcel}><Download size={15} /> تنزيل كشف Excel</button>
              </div>
            </>
          ) : (
            <div className="empty-table">
              <div>
                <div className="empty-symbol"><Users size={21} /></div>
                <h3>لا يوجد عمال بعد</h3>
                <p>استورد ملف الدوام أعلاه أو أضف عاملاً يدوياً للبدء بحساب الأجور.</p>
                <button className="btn" type="button" onClick={() => setRows((current) => [...current, makeWorker()])}><Plus size={14} /> إضافة أول عامل</button>
              </div>
            </div>
          )}
        </section>
        <div className="workbook-note"><CircleHelp size={13} style={{ verticalAlign: 'middle', marginLeft: 5 }} /> يدعم الاستيراد عمود «مجموع الساعات» وأسماء العمال، مع جمع الساعات حسب كل عامل.</div>
      </main>
      <footer className="footer">حساب واضح. بيانات تبقى عندك.</footer>
    </div>
  );
}

export default App;
