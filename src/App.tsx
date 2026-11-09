import { useEffect, useMemo, useRef, useState } from 'react';
import type { FormEvent, ReactNode } from 'react';
import {
  Activity, ArrowDownToLine, ArrowUpRight, BarChart3, Boxes,
  CheckCircle2, ChevronDown, CircleHelp, Cloud, FileSpreadsheet, FileText,
  LayoutDashboard, Menu, PackageSearch, Plus, Search, Settings, ShieldCheck,
  Upload, Wallet, X, TrendingUp, Receipt, RefreshCw
} from 'lucide-react';
import { Area, AreaChart, CartesianGrid, ResponsiveContainer, Tooltip, XAxis, YAxis } from 'recharts';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';

type Batch = {
  id: string; landingDate: string; purchaseDate: string; omraniNo: string;
  invoiceNo: string; containerNo: string; name: string; purchaseEgg: number;
  omraniShipping: number; shohobShipping: number; driver: number; discount: number;
  storage: number; damage: number; forklift: number; sales: number; closingStock: number;
  remarks: string; createdAt: string;
};
type Page = 'overview' | 'batches' | 'lookup' | 'import' | 'reports' | 'settings';
const STORAGE_KEY = 'morvarid.batches.v1';
const money = (n: number) => new Intl.NumberFormat('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(n);
const totalCost = (b: Batch) => b.purchaseEgg + b.omraniShipping + b.shohobShipping + b.driver + b.storage + b.damage + b.forklift - b.discount;
const profit = (b: Batch) => b.sales - totalCost(b);
const blankBatch = (): Batch => ({
  id: crypto.randomUUID(), landingDate: '', purchaseDate: '', omraniNo: '', invoiceNo: '',
  containerNo: '', name: '', purchaseEgg: 0, omraniShipping: 0, shohobShipping: 0,
  driver: 0, discount: 0, storage: 0, damage: 0, forklift: 0, sales: 0,
  closingStock: 0, remarks: '', createdAt: new Date().toISOString()
});
const readBatches = (): Batch[] => {
  try { return JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]') as Batch[]; }
  catch { return []; }
};

function App() {
  const [page, setPage] = useState<Page>('overview');
  const [batches, setBatches] = useState<Batch[]>(readBatches);
  const [query, setQuery] = useState('');
  const [editing, setEditing] = useState<Batch | null>(null);
  const [form, setForm] = useState<Batch>(blankBatch());
  const [notice, setNotice] = useState('');
  const [mobileNav, setMobileNav] = useState(false);
  const [importRows, setImportRows] = useState<Record<string, unknown>[]>([]);
  const [sheetEndpoint, setSheetEndpoint] = useState(localStorage.getItem('morvarid.sheetEndpoint') || '');
  const fileRef = useRef<HTMLInputElement>(null);

  useEffect(() => { localStorage.setItem(STORAGE_KEY, JSON.stringify(batches)); }, [batches]);
  useEffect(() => { if (notice) { const t = window.setTimeout(() => setNotice(''), 3500); return () => window.clearTimeout(t); } }, [notice]);

  const totals = useMemo(() => batches.reduce((a, b) => ({
    sales: a.sales + b.sales, cost: a.cost + totalCost(b), profit: a.profit + profit(b),
    expenses: a.expenses + b.omraniShipping + b.shohobShipping + b.driver + b.storage + b.damage + b.forklift
  }), { sales: 0, cost: 0, profit: 0, expenses: 0 }), [batches]);
  const filtered = batches.filter(b => [b.omraniNo, b.invoiceNo, b.containerNo, b.name].join(' ').toLowerCase().includes(query.toLowerCase()));
  const chartData = useMemo(() => {
    const map = new Map<string, { month: string; sales: number; cost: number; profit: number }>();
    batches.forEach(b => {
      const key = (b.landingDate || b.purchaseDate || b.createdAt).slice(0, 7);
      const row = map.get(key) || { month: key, sales: 0, cost: 0, profit: 0 };
      row.sales += b.sales; row.cost += totalCost(b); row.profit += profit(b); map.set(key, row);
    });
    return [...map.values()].sort((a, b) => a.month.localeCompare(b.month)).slice(-8);
  }, [batches]);

  const openNew = () => { setEditing(null); setForm(blankBatch()); setPage('batches'); };
  const saveBatch = (event: FormEvent) => {
    event.preventDefault();
    if (!form.name.trim() || !form.invoiceNo.trim()) { setNotice('Batch name and supplier invoice number are required.'); return; }
    setBatches(current => editing ? current.map(b => b.id === editing.id ? form : b) : [form, ...current]);
    setNotice(editing ? 'Batch updated successfully.' : 'Batch saved on this device.');
    setEditing(null); setForm(blankBatch()); setPage('batches');
  };
  const editBatch = (batch: Batch) => { setEditing(batch); setForm(batch); setPage('batches'); };
  const deleteBatch = (id: string) => {
    if (window.confirm('Delete this batch? This action cannot be undone.')) setBatches(current => current.filter(b => b.id !== id));
  };
  const exportPdf = (records = batches, title = 'Morvarid Business Report') => {
    const doc = new jsPDF({ orientation: 'landscape' });
    doc.setFontSize(18); doc.text('MORVARID', 14, 17);
    doc.setFontSize(10); doc.text(title, 14, 24);
    doc.text(`Generated: ${new Date().toLocaleString()}  |  Currency: QAR`, 14, 30);
    autoTable(doc, {
      startY: 36,
      head: [['Landing date', 'Omrani No.', 'Invoice', 'Container', 'Batch', 'Sales (QAR)', 'Cost (QAR)', 'Profit (QAR)']],
      body: records.map(b => [b.landingDate, b.omraniNo, b.invoiceNo, b.containerNo, b.name, money(b.sales), money(totalCost(b)), money(profit(b))]),
      styles: { fontSize: 8 }, headStyles: { fillColor: [20, 34, 57] }
    });
    doc.save('morvarid-report.pdf');
  };
  const handleExcel = async (file?: File) => {
    if (!file) return;
    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: 'array', cellDates: true });
      const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(workbook.Sheets[workbook.SheetNames[0]], { defval: '' });
      setImportRows(rows); setPage('import');
      setNotice(`Loaded ${rows.length} rows. Review them before importing.`);
    } catch { setNotice('Could not read this file. Please upload a valid .xlsx, .xls or .csv file.'); }
  };
  const importExcelRows = () => {
    const mapped = importRows.map(row => {
      const get = (...keys: string[]) => {
        const found = Object.keys(row).find(k => keys.some(key => k.toLowerCase().trim() === key.toLowerCase()));
        return found ? row[found] : '';
      };
      const num = (...keys: string[]) => Number(get(...keys)) || 0;
      const date = (v: unknown) => v instanceof Date ? v.toISOString().slice(0, 10) : String(v || '').slice(0, 10);
      return {
        ...blankBatch(),
        landingDate: date(get('Date of Landing', 'Landing Date')),
        purchaseDate: date(get('Date of Purchase', 'Purchase Date')),
        omraniNo: String(get('Omrani No', 'Omrani No.')),
        invoiceNo: String(get('Supplier Invoice No', 'Invoice No', 'Invoice')),
        containerNo: String(get('Container No/ TY', 'Container No', 'Container')),
        name: String(get('Particular / Batch Name', 'Batch Name', 'Particular')),
        purchaseEgg: num('Purchase Egg', 'Purchase'),
        omraniShipping: num('Omrani Shipping expenses', 'Omrani Shipping'),
        shohobShipping: num('Al Shohob Shipping expenses', 'Shohob Shipping'),
        driver: num('Container Driver Expenses', 'Driver Expenses'),
        discount: num('Discount Allowed (-)', 'Discount'),
        storage: num('Storage rent expenses', 'Storage'),
        damage: num('Damage Amount', 'Damage'),
        forklift: num('Forklift Charges', 'Forklift'),
        sales: num('Sales', 'Sale'),
        closingStock: num('Closing Stock Amount', 'Closing Stock'),
        remarks: String(get('Remarks / Notes', 'Remarks', 'Notes'))
      } satisfies Batch;
    }).filter(b => b.invoiceNo.trim() && b.name.trim());
    if (!mapped.length) { setNotice('No valid rows found. Ensure invoice number and batch name columns exist.'); return; }
    setBatches(current => [...mapped, ...current]); setImportRows([]); setNotice(`Imported ${mapped.length} batches to this device.`);
  };
  const saveEndpoint = () => {
    localStorage.setItem('morvarid.sheetEndpoint', sheetEndpoint.trim());
    setNotice('Integration endpoint saved on this device. Live Google authorization is not configured yet.');
  };

  const navItems: { id: Page; label: string; icon: ReactNode; group: string }[] = [
    { id: 'overview', label: 'Executive overview', icon: <LayoutDashboard size={18}/>, group: 'WORKSPACE' },
    { id: 'batches', label: 'Batch entry & data log', icon: <Boxes size={18}/>, group: 'WORKSPACE' },
    { id: 'lookup', label: 'Container lookup', icon: <PackageSearch size={18}/>, group: 'WORKSPACE' },
    { id: 'import', label: 'Excel import center', icon: <FileSpreadsheet size={18}/>, group: 'TOOLS' },
    { id: 'reports', label: 'Financial reports', icon: <FileText size={18}/>, group: 'TOOLS' },
    { id: 'settings', label: 'Google Sheets setup', icon: <Settings size={18}/>, group: 'SYSTEM' }
  ];
  const title = navItems.find(n => n.id === page)?.label || 'Executive overview';

  return (
    <div className="app-layout">
      {mobileNav && <button className="mobile-scrim" aria-label="Close navigation" onClick={() => setMobileNav(false)} />}
      <aside className={`sidebar glass-panel ${mobileNav ? 'sidebar-open' : ''}`}>
        <div className="brand"><div className="brand-mark">M</div><div><strong>Morvarid</strong><span>QAR BUSINESS SUITE</span></div><button className="icon-button mobile-close" onClick={() => setMobileNav(false)} aria-label="Close menu"><X size={18}/></button></div>
        <div className="workspace-pill"><span className="online-dot"/> Local workspace <ChevronDown size={14}/></div>
        {['WORKSPACE','TOOLS','SYSTEM'].map(group => <div className="nav-group" key={group}><p>{group}</p>{navItems.filter(n => n.group === group).map(item => <button key={item.id} className={`nav-link ${page === item.id ? 'active' : ''}`} onClick={() => { setPage(item.id); setMobileNav(false); }}>{item.icon}<span>{item.label}</span></button>)}</div>)}
        <div className="sidebar-bottom"><div className="help-card"><CircleHelp size={18}/><div><strong>Need a hand?</strong><span>Review your workflow</span></div></div><div className="profile-row"><div className="avatar">M</div><div><strong>Morvarid Admin</strong><span>Administrator</span></div><Settings size={16}/></div></div>
      </aside>

      <main className="main-area">
        <header className="topbar">
          <button className="icon-button menu-trigger" onClick={() => setMobileNav(true)} aria-label="Open navigation"><Menu/></button>
          <div className="breadcrumb">Morvarid <span>/</span> <strong>{title}</strong></div>
          <div className="top-actions"><span className="saved-indicator"><span className="online-dot"/> Saved locally</span><button className="icon-button" title="Refresh data" onClick={() => setBatches(readBatches())}><RefreshCw size={17}/></button><div className="avatar">M</div></div>
        </header>

        <div className="page-content">
          <div className="page-heading"><div><p className="eyebrow">EGG IMPORT OPERATIONS & FINANCIAL MANAGEMENT</p><h1>{page === 'overview' ? 'Good day, Admin' : title}</h1><p className="subheading">{page === 'overview' ? 'Here is what is happening across your business.' : 'Manage your business information in one place.'}</p></div><button className="primary-button" onClick={openNew}><Plus size={17}/> New batch entry</button></div>

          {page === 'overview' && <>
            <div className="metric-grid">
              <Metric label="Total sales" value={totals.sales} icon={<TrendingUp/>} tone="blue" detail={`${batches.length} recorded batches`}/>
              <Metric label="Purchase & landed cost" value={totals.cost} icon={<Wallet/>} tone="violet" detail="All recorded batch costs"/>
              <Metric label="Operating expenses" value={totals.expenses} icon={<Receipt/>} tone="amber" detail="Shipping, storage & handling"/>
              <Metric label="Gross profit" value={totals.profit} icon={<Activity/>} tone="green" detail={totals.sales ? `${(totals.profit / totals.sales * 100).toFixed(1)}% gross margin` : 'Awaiting sales data'}/>
            </div>
            <div className="content-grid">
              <section className="panel chart-panel"><div className="panel-heading"><div><h2>Financial performance</h2><p>Sales, cost and profit over time</p></div><span className="tag">QAR</span></div>
                {chartData.length ? <div className="chart-wrap"><ResponsiveContainer width="100%" height="100%"><AreaChart data={chartData}><defs><linearGradient id="salesFill" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor="#72a7ff" stopOpacity={0.28}/><stop offset="95%" stopColor="#72a7ff" stopOpacity={0}/></linearGradient></defs><CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,.08)"/><XAxis dataKey="month" stroke="#8290a5" fontSize={11}/><YAxis stroke="#8290a5" fontSize={11}/><Tooltip contentStyle={{background:'#172338',border:'1px solid #31435d',borderRadius:12,color:'#fff'}}/><Area type="monotone" dataKey="sales" name="Sales" stroke="#72a7ff" fill="url(#salesFill)"/><Area type="monotone" dataKey="cost" name="Cost" stroke="#f5b85b" fill="transparent"/><Area type="monotone" dataKey="profit" name="Profit" stroke="#52d6a4" fill="transparent"/></AreaChart></ResponsiveContainer></div> : <EmptyState icon={<BarChart3/>} title="Your financial story starts here" text="Add your first batch to see sales, costs and profit trends." action={<button className="text-button" onClick={openNew}>Add first batch <ArrowUpRight size={15}/></button>}/>}
              </section>
              <section className="panel activity-panel"><div className="panel-heading"><div><h2>Recent batches</h2><p>Your latest recorded imports</p></div><button className="text-button" onClick={() => setPage('batches')}>View all <ArrowUpRight size={15}/></button></div>
                {batches.slice(0,5).length ? batches.slice(0,5).map(b => <div className="recent-row" key={b.id}><div className="batch-icon"><Boxes size={17}/></div><div className="recent-info"><strong>{b.name}</strong><span>{b.invoiceNo} · {b.landingDate || 'No landing date'}</span></div><div className="recent-value"><strong>{money(profit(b))}</strong><span>QAR profit</span></div></div>) : <EmptyState icon={<Boxes/>} title="No batches yet" text="Your latest batch entries will appear here."/>}
              </section>
            </div>
            <section className="panel quick-panel"><div className="panel-heading"><div><h2>Quick actions</h2><p>Common tasks for your daily workflow</p></div></div><div className="quick-actions"><QuickAction icon={<Plus/>} title="Add a batch" text="Enter import and cost details" onClick={openNew}/><QuickAction icon={<Upload/>} title="Import Excel" text="Upload existing batch records" onClick={() => {setPage('import');fileRef.current?.click();}}/><QuickAction icon={<FileText/>} title="Export report" text="Download a PDF summary" onClick={() => exportPdf()}/><QuickAction icon={<PackageSearch/>} title="Find a container" text="Search invoice or container ID" onClick={() => setPage('lookup')}/></div></section>
          </>}

          {page === 'batches' && <section className="panel form-panel"><div className="panel-heading"><div><h2>{editing ? 'Edit batch details' : 'Batch data entry'}</h2><p>Enter import details and financial figures. Costs and profit are calculated automatically.</p></div><span className="tag">Currency · QAR</span></div>
            <form onSubmit={saveBatch}><div className="form-section-title">Import information</div><div className="form-grid">
              <Field label="Date of landing"><input type="date" value={form.landingDate} onChange={e=>setForm({...form,landingDate:e.target.value})}/></Field>
              <Field label="Date of purchase"><input type="date" value={form.purchaseDate} onChange={e=>setForm({...form,purchaseDate:e.target.value})}/></Field>
              <Field label="Omrani number"><input value={form.omraniNo} onChange={e=>setForm({...form,omraniNo:e.target.value})} placeholder="e.g. OM2026/131"/></Field>
              <Field label="Supplier invoice number" required><input required value={form.invoiceNo} onChange={e=>setForm({...form,invoiceNo:e.target.value})} placeholder="Invoice reference"/></Field>
              <Field label="Container number / TY"><input value={form.containerNo} onChange={e=>setForm({...form,containerNo:e.target.value})} placeholder="Container reference"/></Field>
              <Field label="Particular / batch name" required><input required value={form.name} onChange={e=>setForm({...form,name:e.target.value})} placeholder="Batch name"/></Field>
            </div><div className="form-section-title">Purchase & operating costs <span>Enter amounts in QAR</span></div><div className="form-grid">
              {([
                ['purchaseEgg','Purchase egg'],['omraniShipping','Omrani shipping expenses'],['shohobShipping','Al Shohob shipping expenses'],['driver','Container driver expenses'],['discount','Discount allowed (-)'],['storage','Storage rent expenses'],['damage','Damage amount'],['forklift','Forklift charges'],['sales','Sales'],['closingStock','Closing stock amount']
              ] as [keyof Batch,string][]).map(([key,label])=><Field key={key} label={label}><input type="number" min="0" step="0.01" value={form[key] as number} onChange={e=>setForm({...form,[key]:Number(e.target.value)})}/></Field>)}
            </div><Field label="Remarks / notes"><textarea rows={3} value={form.remarks} onChange={e=>setForm({...form,remarks:e.target.value})} placeholder="Add any notes about this batch..."/></Field>
            <div className="calculation-strip"><div><span>Calculated cost (A)</span><strong>{money(totalCost(form))} QAR</strong></div><div><span>Gross profit (B-A)</span><strong className={profit(form)>=0?'positive':'negative'}>{money(profit(form))} QAR</strong></div><div><span>Profit margin</span><strong>{form.sales ? (profit(form)/form.sales*100).toFixed(1) : '0.0'}%</strong></div></div>
            <div className="form-footer"><button type="button" className="secondary-button" onClick={()=>{setForm(blankBatch());setEditing(null);}}>Reset form</button><button className="primary-button" type="submit"><CheckCircle2 size={17}/>{editing?'Update batch':'Save batch'}</button></div></form>
          </section>
          }
          {page === 'batches' && <section className="panel table-panel"><div className="panel-heading"><div><h2>Master data log</h2><p>{batches.length} saved batch records on this device</p></div><button className="secondary-button small" onClick={()=>exportPdf(batches,'Master batch register')}><ArrowDownToLine size={15}/> Export PDF</button></div><BatchTable batches={batches} onEdit={editBatch} onDelete={deleteBatch}/></section>}

          {page === 'lookup' && <section className="panel lookup-panel"><div className="panel-heading"><div><h2>Container search & inspection</h2><p>Search by Omrani number, supplier invoice, container or batch name.</p></div></div><div className="search-box"><Search size={19}/><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Start typing an invoice, container or batch name..."/></div><BatchTable batches={filtered} onEdit={editBatch} onDelete={deleteBatch}/></section>}

          {page === 'import' && <section className="panel import-panel"><div className="panel-heading"><div><h2>Import batch records</h2><p>Upload Excel or CSV. Review the preview before importing records.</p></div><span className="tag">.xlsx · .xls · .csv</span></div><div className="dropzone" onClick={()=>fileRef.current?.click()} onDragOver={e=>e.preventDefault()} onDrop={e=>{e.preventDefault();void handleExcel(e.dataTransfer.files[0]);}}><div className="upload-icon"><Upload/></div><strong>Drop your spreadsheet here or browse files</strong><span>Supported: Excel workbooks and CSV files</span><button type="button" className="secondary-button">Choose file</button><input ref={fileRef} hidden type="file" accept=".xlsx,.xls,.csv" onChange={e=>void handleExcel(e.target.files?.[0])}/></div>{importRows.length>0&&<><div className="import-summary"><strong>{importRows.length} rows detected</strong><span>Only rows with invoice number and batch name will be imported. Check your column headings.</span><button className="primary-button" onClick={importExcelRows}>Import valid rows</button></div><div className="table-scroll"><table><thead><tr>{Object.keys(importRows[0]).slice(0,7).map(k=><th key={k}>{k}</th>)}</tr></thead><tbody>{importRows.slice(0,8).map((row,i)=><tr key={i}>{Object.keys(importRows[0]).slice(0,7).map(k=><td key={k}>{String(row[k]??'')}</td>)}</tr>)}</tbody></table></div></>}</section>}

          {page === 'reports' && <section className="panel reports-panel"><div className="panel-heading"><div><h2>Financial reports</h2><p>Generate a portable report from the records currently stored on this device.</p></div><span className="tag">QAR</span></div><div className="report-summary"><div><span>Total sales</span><strong>{money(totals.sales)}</strong></div><div><span>Total cost</span><strong>{money(totals.cost)}</strong></div><div><span>Gross profit</span><strong className={totals.profit>=0?'positive':'negative'}>{money(totals.profit)}</strong></div><div><span>Batch count</span><strong>{batches.length}</strong></div></div><div className="report-download"><div className="report-file"><div className="pdf-icon"><FileText/></div><div><strong>Business performance report</strong><span>Batch register, sales, costs and gross profit · PDF</span></div></div><button className="primary-button" onClick={()=>exportPdf(batches)}><ArrowDownToLine size={17}/> Download PDF</button></div><p className="helper-text">Note: This report reflects the records available in this browser. Google Sheets synchronization is not active yet.</p></section>}

          {page === 'settings' && <section className="panel settings-panel"><div className="panel-heading"><div><h2>Google Sheets connection</h2><p>Prepare a secure connection for centralized business records and backup.</p></div><span className="status-chip pending">Not connected</span></div><div className="connection-card"><div className="google-sheet-icon"><FileSpreadsheet/></div><div className="connection-copy"><strong>Connect your Google workspace</strong><span>Morvarid will use an authorized Google integration to read and write your selected spreadsheet.</span></div><span className="status-chip pending">Setup required</span></div><div className="notice-box"><ShieldCheck size={19}/><div><strong>Authorization is not configured</strong><p>Do not enter your Google password here. A proper OAuth client and a server-side Apps Script/API integration must be configured before account sign-in and sheet access can be enabled.</p></div></div><Field label="Apps Script web app endpoint (optional setup value)"><input type="url" value={sheetEndpoint} onChange={e=>setSheetEndpoint(e.target.value)} placeholder="https://script.google.com/macros/s/.../exec"/></Field><div className="form-footer"><span className="helper-text">Saved only in this browser. No data is sent from this screen.</span><button className="primary-button" onClick={saveEndpoint}>Save configuration</button></div><div className="settings-list"><div><Cloud/><span><strong>Google account sign-in</strong><small>Requires OAuth consent screen and client configuration.</small></span><span className="status-chip pending">Pending</span></div><div><RefreshCw/><span><strong>Automatic Google Sheets sync</strong><small>Requires a deployed integration endpoint and sheet schema.</small></span><span className="status-chip pending">Pending</span></div><div><ShieldCheck/><span><strong>Backup and recovery</strong><small>Will be enabled with verified sync and restore handling.</small></span><span className="status-chip pending">Pending</span></div></div></section>}
          <footer className="app-footer"><span>© {new Date().getFullYear()} Morvarid Business Suite</span><span><span className="online-dot"/> Local data storage active</span></footer>
        </div>
      </main>
      {notice && <div className="toast" role="status"><CheckCircle2 size={17}/>{notice}<button onClick={()=>setNotice('')} aria-label="Dismiss notification"><X size={15}/></button></div>}
    </div>
  );
}

function Metric({label,value,icon,tone,detail}:{label:string;value:number;icon:ReactNode;tone:string;detail:string}) {
  return <article className="metric-card glass-panel"><div className="metric-top"><span>{label}</span><div className={`metric-icon ${tone}`}>{icon}</div></div><strong className="metric-value">{money(value)} <small>QAR</small></strong><span className="metric-detail">{detail}</span></article>;
}
function Field({label,required,children}:{label:string;required?:boolean;children:ReactNode}) { return <label className="field"><span>{label}{required&&<i> *</i>}</span>{children}</label>; }
function EmptyState({icon,title,text,action}:{icon:ReactNode;title:string;text:string;action?:ReactNode}) { return <div className="empty-state"><div>{icon}</div><strong>{title}</strong><p>{text}</p>{action}</div>; }
function QuickAction({icon,title,text,onClick}:{icon:ReactNode;title:string;text:string;onClick:()=>void}) { return <button className="quick-action" onClick={onClick}><span>{icon}</span><strong>{title}</strong><small>{text}</small><ArrowUpRight size={15}/></button>; }
function BatchTable({batches,onEdit,onDelete}:{batches:Batch[];onEdit:(b:Batch)=>void;onDelete:(id:string)=>void}) {
  if (!batches.length) return <EmptyState icon={<Boxes/>} title="No records found" text="Add a batch or import an Excel file to see records here."/>;
  return <div className="table-scroll"><table><thead><tr><th>Batch / container</th><th>Invoice</th><th>Landing date</th><th>Sales</th><th>Cost (A)</th><th>Profit (B-A)</th><th>Actions</th></tr></thead><tbody>{batches.map(b=><tr key={b.id}><td><strong>{b.name}</strong><small>{b.containerNo||b.omraniNo||'No container reference'}</small></td><td>{b.invoiceNo}</td><td>{b.landingDate||'—'}</td><td>{money(b.sales)}</td><td>{money(totalCost(b))}</td><td className={profit(b)>=0?'positive':'negative'}>{money(profit(b))}</td><td><div className="row-actions"><button onClick={()=>onEdit(b)}>Edit</button><button className="danger-text" onClick={()=>onDelete(b.id)}>Delete</button></div></td></tr>)}</tbody></table></div>;
}
export default App;
