import { useMemo, useRef, useState, type ChangeEvent, type DragEvent, type FormEvent } from "react";
import {
  AlertCircle, ArrowDownToLine, ArrowUpRight, BadgeCheck, CalendarDays, Check,
  ChevronRight, Clipboard, FileSpreadsheet, FileText, Filter, Landmark, LockKeyhole,
  Mail, MessageCircle, Plus, Search, Settings2, ShieldCheck, SlidersHorizontal, Sparkles, Upload,
  X,
} from "lucide-react";
import { format, parseISO } from "date-fns";
import type { jsPDF as JsPDFDocument } from "jspdf";
import {
  calculateInvoice, formatDate, formatINR, isBlankRow, parseLedgerDate, todayISO,
  type Invoice,
} from "./engine";

type MappingKey = "buyer" | "gstin" | "number" | "invoiceDate" | "amount" | "writtenAgreement" | "agreedDays";
type Mapping = Record<MappingKey, number>;
type NoticeTone = "A" | "B" | "C";
type FilterType = "All invoices" | "Overdue" | "Safe";

const fields: { key: MappingKey; label: string; aliases: string[]; required: boolean }[] = [
  { key: "buyer", label: "Buyer name", aliases: ["buyername", "partyname", "customername", "debtor", "buyer", "party", "customer"], required: true },
  { key: "gstin", label: "GSTIN", aliases: ["gstin", "gstnumber", "gstno"], required: false },
  { key: "number", label: "Invoice number", aliases: ["invoicenumber", "invoiceno", "billnumber", "billno", "voucher", "number"], required: true },
  { key: "invoiceDate", label: "Invoice date", aliases: ["invoicedate", "billdate", "date"], required: true },
  { key: "amount", label: "Invoice amount", aliases: ["invoiceamount", "amount", "total", "billamount", "netamount", "value"], required: true },
  { key: "writtenAgreement", label: "Written agreement", aliases: ["writtenagreement", "agreement", "agreementyesno", "haswrittenagreement"], required: false },
  { key: "agreedDays", label: "Agreed days", aliases: ["agreeddays", "creditdays", "paymentdays", "termsdays"], required: false },
];

const emptyMapping = (): Mapping => ({
  buyer: -1, gstin: -1, number: -1, invoiceDate: -1, amount: -1, writtenAgreement: -1, agreedDays: -1,
});

const today = todayISO();
const blankInvoice = (): Invoice => ({
  id: crypto.randomUUID(), buyer: "", gstin: "", number: "", invoiceDate: today,
  amount: 0, writtenAgreement: false, agreedDays: null,
});

function normalizeHeader(value: unknown) {
  return String(value ?? "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

function suggestedMapping(headers: unknown[]): Mapping {
  const mapping = emptyMapping();
  fields.forEach(({ key, aliases }) => {
    const index = headers.findIndex((header) => aliases.includes(normalizeHeader(header)));
    if (index >= 0) mapping[key] = index;
  });
  return mapping;
}

function money(value: number) {
  return formatINR(value).replace("₹", "₹ ");
}

function App() {
  const [invoices, setInvoices] = useState<Invoice[]>([]);
  const [rbiRate, setRbiRate] = useState(5.75);
  const [cutoff, setCutoff] = useState(today);
  const [msmeType, setMsmeType] = useState<"Micro" | "Small">("Micro");
  const [clientName, setClientName] = useState("");
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [query, setQuery] = useState("");
  const [filter, setFilter] = useState<FilterType>("All invoices");
  const [sortAsc, setSortAsc] = useState(false);
  const [fileRows, setFileRows] = useState<unknown[][] | null>(null);
  const [fileName, setFileName] = useState("");
  const [headers, setHeaders] = useState<unknown[]>([]);
  const [mapping, setMapping] = useState<Mapping>(emptyMapping());
  const [mappingOpen, setMappingOpen] = useState(false);
  const [manualOpen, setManualOpen] = useState(false);
  const [draft, setDraft] = useState<Invoice>(blankInvoice());
  const [noticeBuyer, setNoticeBuyer] = useState("");
  const [noticeOpen, setNoticeOpen] = useState(false);
  const [tone, setTone] = useState<NoticeTone>("A");
  const [toast, setToast] = useState("");
  const [dragging, setDragging] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  const calculated = useMemo(
    () => invoices.map((invoice) => calculateInvoice(invoice, cutoff, rbiRate, true)),
    [invoices, cutoff, rbiRate],
  );
  const filtered = useMemo(() => {
    const search = query.trim().toLowerCase();
    return calculated
      .filter((row) => filter === "All invoices" || (filter === "Overdue" ? row.daysDelayed > 0 : row.daysDelayed === 0))
      .filter((row) => !search || [row.buyer, row.number, row.gstin].some((value) => value.toLowerCase().includes(search)))
      .sort((a, b) => {
        const comparison = a.daysDelayed - b.daysDelayed || a.dueDate.localeCompare(b.dueDate);
        return sortAsc ? comparison : -comparison;
      });
  }, [calculated, filter, query, sortAsc]);
  const overdue = calculated.filter((row) => row.daysDelayed > 0);
  const totalPrincipal = calculated.reduce((sum, row) => sum + row.amount, 0);
  const overduePrincipal = overdue.reduce((sum, row) => sum + row.amount, 0);
  const interestAccrued = overdue.reduce((sum, row) => sum + row.interest, 0);
  const buyersAtRisk = new Set(overdue.filter((row) => row.taxRisk).map((row) => row.buyer.trim().toLowerCase())).size;
  const noticeRows = calculated.filter((row) => row.buyer === noticeBuyer && row.daysDelayed > 0);
  const statutoryRate = 3 * rbiRate;

  function notify(message: string) {
    setToast(message);
    window.setTimeout(() => setToast(""), 3600);
  }

  async function runWithFeedback(action: () => Promise<void>, label: string) {
    try {
      await action();
    } catch (error) {
      console.error(`${label} failed`, error);
      notify(`${label} failed. Please try again.`);
    }
  }

  async function loadFile(file: File) {
    if (!/\.(xlsx|xls|csv)$/i.test(file.name)) {
      notify("Choose an Excel (.xlsx, .xls) or CSV file.");
      return;
    }
    try {
      const data = await file.arrayBuffer();
      const XLSX = await import("@e965/xlsx");
      const workbook = XLSX.read(data, { type: "array", cellDates: true });
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json<unknown[]>(firstSheet, { header: 1, defval: "", raw: true });
      const populated = rows.filter((row) => !isBlankRow(row));
      if (populated.length < 2) {
        notify("This file has no invoice rows. Add a header row and at least one invoice.");
        return;
      }
      const fileHeaders = populated[0];
      setHeaders(fileHeaders);
      setFileRows(populated.slice(1));
      setFileName(file.name);
      setMapping(suggestedMapping(fileHeaders));
      setMappingOpen(true);
    } catch (error) {
      console.error("Unable to read ledger file", error);
      notify("Could not read this file. Check that it is a valid, unprotected workbook.");
    }
  }

  function onFileChange(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (file) void loadFile(file);
    event.target.value = "";
  }

  function onDrop(event: DragEvent<HTMLDivElement>) {
    event.preventDefault();
    setDragging(false);
    const file = event.dataTransfer.files[0];
    if (file) void loadFile(file);
  }

  function importMappedRows() {
    if (!fileRows) return;
    const required = fields.filter((field) => field.required && mapping[field.key] < 0);
    if (required.length) {
      notify(`Map the required column${required.length > 1 ? "s" : ""}: ${required.map((field) => field.label).join(", ")}.`);
      return;
    }
    const imported: Invoice[] = [];
    const errors: string[] = [];
    fileRows.forEach((row, index) => {
      if (isBlankRow(row)) return;
      const get = (key: MappingKey) => mapping[key] >= 0 ? row[mapping[key]] : "";
      const buyer = String(get("buyer") ?? "").trim();
      const number = String(get("number") ?? "").trim();
      const invoiceDate = parseLedgerDate(get("invoiceDate"));
      const rawAmount = get("amount");
      const amount = typeof rawAmount === "number" ? rawAmount : Number(String(rawAmount).replace(/[₹,\s]/g, ""));
      if (!buyer || !number || !invoiceDate || !Number.isFinite(amount) || amount <= 0) {
        errors.push(`Row ${index + 2}: ${!buyer ? "buyer name" : !number ? "invoice number" : !invoiceDate ? "valid invoice date" : "positive invoice amount"} missing or invalid`);
        return;
      }
      const agreement = String(get("writtenAgreement") ?? "").trim().toLowerCase();
      const writtenAgreement = ["yes", "true", "y", "1"].includes(agreement);
      const rawDays = String(get("agreedDays") ?? "").trim().match(/\d+(?:\.\d+)?/)?.[0];
      const agreedDays = rawDays ? Number(rawDays) : null;
      imported.push({
        id: crypto.randomUUID(), buyer, number, invoiceDate, amount,
        gstin: String(get("gstin") ?? "").trim(),
        writtenAgreement,
        agreedDays: Number.isFinite(agreedDays) ? agreedDays : null,
      });
    });
    if (imported.length === 0) {
      notify(errors[0] ?? "No valid invoice rows were found.");
      return;
    }
    setInvoices((current) => [...current, ...imported]);
    setMappingOpen(false);
    setFileRows(null);
    notify(`${imported.length} invoice${imported.length === 1 ? "" : "s"} imported${errors.length ? `; ${errors.length} invalid row${errors.length === 1 ? "" : "s"} skipped (first: ${errors[0]})` : ""}.`);
  }

  async function downloadTemplate() {
    const XLSX = await import("@e965/xlsx");
    const worksheet = XLSX.utils.aoa_to_sheet([
      ["Buyer Name", "GSTIN", "Invoice Number", "Invoice Date (YYYY-MM-DD)", "Invoice Amount", "Written Agreement (Yes/No)", "Agreed Days (Optional)"],
      ["Example Buyer Pvt Ltd", "27ABCDE1234F1Z5", "INV-001", today, 125000, "Yes", 30],
    ]);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Invoices");
    XLSX.writeFile(workbook, "MSME-AuditFlow-Template.xlsx");
    notify("Sample Excel template downloaded.");
  }

  async function downloadExcel() {
    if (!calculated.length) {
      notify("Add or import invoices before exporting.");
      return;
    }
    const XLSX = await import("@e965/xlsx");
    const rows = calculated.map((row) => ({
      "Buyer Name": row.buyer,
      GSTIN: row.gstin,
      "Invoice Number": row.number,
      "Invoice Date": row.invoiceDate,
      "Invoice Amount (INR)": row.amount,
      "Written Agreement": row.writtenAgreement ? "Yes" : "No",
      "Agreed Days": row.agreedDays ?? "",
      "Statutory Credit Window (Days)": row.creditDays,
      "Statutory Due Date": row.dueDate,
      "Days Delayed": row.daysDelayed,
      "Accrued Section 16 Interest (INR)": Number(row.interest.toFixed(2)),
      "Total Recoverable (INR)": Number(row.total.toFixed(2)),
      "Section 43B(h) Tax Disallowance Triggered": row.taxRisk ? "Yes" : "No",
    }));
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(rows), "Audit Reconciliation");
    XLSX.writeFile(workbook, "MSME-AuditFlow-Reconciliation.xlsx");
    notify("Clean Excel reconciliation exported.");
  }

  async function exportPdf() {
    if (!calculated.length) {
      notify("Add or import invoices before exporting.");
      return;
    }
    const [{ jsPDF }, { default: autoTable }] = await Promise.all([
      import("jspdf"),
      import("jspdf-autotable"),
    ]);
    const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
    doc.setFont("helvetica", "bold");
    doc.setFontSize(15);
    doc.text("Working Paper: MSME Receivables & Statutory Section 16/43B(h) Aging Schedule", 14, 16);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(9);
    doc.text(`Client: ${clientName.trim() || "____________________________"}    |    Audit cutoff: ${formatDate(cutoff)}    |    Supplier classification: Registered ${msmeType} Enterprise`, 14, 23);
    doc.text(`RBI Bank Rate: ${rbiRate}%    |    Section 16 rate: ${statutoryRate}% p.a. (3x RBI Bank Rate)    |    Currency: INR`, 14, 29);
    autoTable(doc, {
      startY: 34,
      head: [["Buyer", "Invoice #", "Invoice date", "Credit window", "Statutory due", "Delay (days)", "Principal (INR)", "Interest (INR)", "Total (INR)", "43B(h)"]],
      body: calculated.map((row) => [
        row.buyer, row.number, formatDate(row.invoiceDate), `${row.creditDays} days`, formatDate(row.dueDate),
        String(row.daysDelayed), row.amount.toFixed(2), row.interest.toFixed(2), row.total.toFixed(2), row.taxRisk ? "Triggered" : "No",
      ]),
      styles: { fontSize: 7, cellPadding: 2.2 },
      headStyles: { fillColor: [24, 62, 51] },
      alternateRowStyles: { fillColor: [244, 248, 246] },
      margin: { left: 14, right: 14 },
    });
    const tableBottom = (doc as JsPDFDocument & { lastAutoTable?: { finalY: number } }).lastAutoTable?.finalY ?? 40;
    let noteY = tableBottom + 9;
    if (noteY + 48 > doc.internal.pageSize.getHeight() - 10) {
      doc.addPage();
      noteY = 18;
    }
    doc.setFontSize(8);
    doc.text("Statutory notes: Section 15 of the MSMED Act, 2006 sets the payment period (maximum 45 days by written agreement).", 14, noteY);
    doc.text("Section 16 provides compound interest with monthly rests at three times the RBI Bank Rate. Section 23 addresses non-deductibility of such interest.", 14, noteY + 5);
    doc.text("Section 43B(h), Income-tax Act, 1961: qualifying MSE dues unpaid beyond the Section 15 period are deductible only on actual payment, subject to statutory provisions.", 14, noteY + 10);
    doc.text("Prepared from client-provided ledger data. Verify supplier Udyam status, payments, applicable law and cutoff before reliance.", 14, noteY + 15);
    doc.line(24, noteY + 38, 105, noteY + 38);
    doc.line(190, noteY + 38, 271, noteY + 38);
    doc.text("Auditor's signature / date", 24, noteY + 43);
    doc.text("Client / authorised signatory / date", 190, noteY + 43);
    doc.save("MSME-AuditFlow-Audit-Schedule.pdf");
    notify("CA audit schedule PDF exported.");
  }

  function saveManual(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!draft.buyer.trim() || !draft.number.trim() || !draft.invoiceDate || draft.amount <= 0) {
      notify("Enter a buyer, invoice number, date and positive amount.");
      return;
    }
    setInvoices((current) => [...current, { ...draft, buyer: draft.buyer.trim(), number: draft.number.trim(), gstin: draft.gstin.trim() }]);
    setManualOpen(false);
    setDraft(blankInvoice());
    notify("Invoice added to the audit ledger.");
  }

  function openDemand(buyer?: string) {
    const buyers = Array.from(new Set(calculated.filter((row) => row.daysDelayed > 0).map((row) => row.buyer)));
    if (!buyers.length) {
      notify("There are no overdue invoices to include in a statutory demand.");
      return;
    }
    setNoticeBuyer(buyer && buyers.includes(buyer) ? buyer : buyers[0]);
    setNoticeOpen(true);
  }

  function noticeText() {
    const buyer = noticeBuyer || "the Buyer";
    const dueAmount = noticeRows.reduce((sum, row) => sum + row.amount, 0);
    const interest = noticeRows.reduce((sum, row) => sum + row.interest, 0);
    const invoiceLines = noticeRows.map((row) =>
      `• ${row.number} dated ${formatDate(row.invoiceDate)} — principal ${money(row.amount)}, due ${formatDate(row.dueDate)}, ${row.daysDelayed} days overdue`,
    ).join("\n");
    const opening = tone === "A"
      ? `Dear ${buyer},\n\nAs part of our mutual audit and MSME payment-compliance review, we request your confirmation and prompt settlement of the following outstanding invoices.`
      : tone === "B"
        ? `Dear ${buyer},\n\nFORMAL STATUTORY PAYMENT REMINDER\n\nThe following invoices have passed their applicable payment dates under Section 15 of the Micro, Small and Medium Enterprises Development Act, 2006. Section 16 provides for compound interest with monthly rests at three times the RBI Bank Rate (currently configured at ${statutoryRate.toFixed(2)}% p.a.).`
        : `To: ${buyer}\n\nSUBJECT: FORMAL DEMAND FOR PAYMENT OF OVERDUE MSME INVOICES\n\nThis is a formal demand for payment of the amounts set out below. The applicable payment period is governed by Section 15 of the Micro, Small and Medium Enterprises Development Act, 2006. Pursuant to Section 16, delayed payment attracts compound interest with monthly rests at three times the RBI Bank Rate. If payment is not made, the matter may be referred to the Micro and Small Enterprises Facilitation Council / MSME Samadhaan, subject to applicable law and eligibility.`;
    return `${opening}\n\nOutstanding invoices:\n${invoiceLines}\n\nPrincipal outstanding: ${money(dueAmount)}\nAccrued interest as of ${formatDate(cutoff)}: ${money(interest)}\nTotal currently due: ${money(dueAmount + interest)}\n\nPlease arrange settlement and share remittance confirmation. This statement is based on the records presently available and is without prejudice to applicable legal rights. Kindly contact us promptly to reconcile any discrepancy.\n\nRegards,\n${clientName.trim() || "[Supplier / Authorised Representative]"}`;
  }

  async function copyNotice() {
    try {
      await navigator.clipboard.writeText(noticeText());
      notify("Demand statement copied to clipboard.");
    } catch (error) {
      console.error("Unable to copy demand statement", error);
      notify("Clipboard access was blocked. Select and copy the statement manually.");
    }
  }

  function setMap(key: MappingKey, value: string) {
    setMapping((current) => ({ ...current, [key]: value === "" ? -1 : Number(value) }));
  }

  const filters: FilterType[] = ["All invoices", "Overdue", "Safe"];

  return (
    <main className="app-shell">
      <header className="topbar">
        <a className="brand" href="#" aria-label="MSME AuditFlow home">
          <span className="brand-mark"><Landmark size={20} strokeWidth={2.1} /></span>
          <span className="brand-copy"><strong>MSME <em>AuditFlow</em></strong><small>STATUTORY INTEREST ENGINE</small></span>
        </a>
        <div className="topbar-right">
          <div className="privacy-pill"><LockKeyhole size={14} /> 100% Private &amp; Client-Side</div>
          <button className={`icon-button settings-trigger ${settingsOpen ? "active" : ""}`} onClick={() => setSettingsOpen(!settingsOpen)} aria-label="Open global settings"><Settings2 size={18} /></button>
        </div>
        {settingsOpen && <section className="settings-popover">
          <div className="popover-heading"><div><strong>Audit settings</strong><span>Applied across calculations and exports</span></div><button className="bare-icon" onClick={() => setSettingsOpen(false)} aria-label="Close settings"><X size={16} /></button></div>
          <label>Client / supplier name<input value={clientName} onChange={(event) => setClientName(event.target.value)} placeholder="Enter entity name" /></label>
          <label>RBI Bank Rate <span className="input-suffix"><input type="number" min="0" max="100" step="0.01" value={rbiRate} onChange={(event) => setRbiRate(Math.min(100, Math.max(0, Number(event.target.value) || 0)))} /><i>%</i></span></label>
          <div className="rate-summary"><span>Section 16 statutory rate</span><b>{statutoryRate.toFixed(2)}% <small>p.a.</small></b></div>
          <label>Audit cutoff date<input type="date" value={cutoff} onChange={(event) => event.target.value && setCutoff(event.target.value)} /></label>
          <label>Supplier MSME classification<select value={msmeType} onChange={(event) => setMsmeType(event.target.value as "Micro" | "Small")}><option>Micro</option><option>Small</option></select></label>
          <p className="settings-note">43B(h) tax flags assume all invoices in this receivables ledger remain unpaid through the cutoff date.</p>
        </section>}
      </header>

      <section className="page-heading">
        <div><div className="eyebrow"><span className="eyebrow-line" /> MSME ACT, 2006 <span className="eyebrow-dot" /> SECTION 43B(h)</div>
          <h1>Receivables <span>command center.</span></h1>
          <p>Statutory timelines, compounding interest &amp; tax exposure — all in one audit-ready view.</p>
        </div>
        <div className="heading-asof"><CalendarDays size={15} /><span>CALCULATED AS OF</span><strong>{format(parseISO(cutoff), "dd MMM yyyy")}</strong></div>
      </section>

      <section className="metrics-grid" aria-label="Executive summary">
        <Metric icon={<FileText size={17} />} label="RECEIVABLES ANALYZED" value={String(calculated.length)} detail={`${money(totalPrincipal)} principal`} tone="ink" />
        <Metric icon={<AlertCircle size={17} />} label="STATUTORY OVERDUE" value={money(overduePrincipal)} detail={`${overdue.length} overdue invoice${overdue.length === 1 ? "" : "s"}`} tone="amber" />
        <Metric icon={<Sparkles size={17} />} label="COMPOUND INTEREST" value={money(interestAccrued)} detail={`${statutoryRate.toFixed(2)}% p.a. · monthly rests`} tone="green" />
        <Metric icon={<ShieldCheck size={17} />} label="BUYERS AT TAX RISK" value={String(buyersAtRisk)} detail="43B(h) disallowance trigger" tone="red" />
      </section>

      <section className="workbench">
        <div className="panel ledger-panel">
          <div className="panel-header">
            <div className="panel-title"><span className="panel-icon"><FileSpreadsheet size={17} /></span><div><h2>Audit ledger</h2><p>Review statutory exposure invoice by invoice</p></div></div>
            <button className="primary-button compact" onClick={() => { setDraft(blankInvoice()); setManualOpen(true); }}><Plus size={16} /> Add invoice</button>
          </div>
          <div className={`dropzone ${dragging ? "dragging" : ""}`} onDragOver={(event) => { event.preventDefault(); setDragging(true); }} onDragLeave={() => setDragging(false)} onDrop={onDrop} onClick={() => inputRef.current?.click()}>
            <div className="upload-icon"><Upload size={18} /></div>
            <div className="upload-copy"><strong>Drop your ledger here <span>or browse files</span></strong><small>{fileName || "Excel or CSV · .xlsx, .xls, .csv"}</small></div>
            <button className="text-button" onClick={(event) => { event.stopPropagation(); void runWithFeedback(downloadTemplate, "Template download"); }}><ArrowDownToLine size={14} /> Sample template</button>
            <input ref={inputRef} type="file" accept=".xlsx,.xls,.csv" hidden onChange={onFileChange} />
          </div>
          <div className="table-toolbar">
            <div className="filter-tabs" role="tablist" aria-label="Filter invoices">
              {filters.map((item) => <button key={item} className={filter === item ? "selected" : ""} onClick={() => setFilter(item)} role="tab" aria-selected={filter === item}>{item}{item === "Overdue" && overdue.length > 0 && <span>{overdue.length}</span>}</button>)}
            </div>
            <label className="search-box"><Search size={15} /><input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Search buyer or invoice..." /></label>
          </div>
          <div className="table-scroll">
            <table>
              <thead><tr>
                <th>BUYER / INVOICE</th><th>CREDIT</th><th><button className="sort-button" onClick={() => setSortAsc(!sortAsc)}>DUE DATE <SlidersHorizontal size={12} /></button></th>
                <th>DELAY</th><th>PRINCIPAL</th><th>INTEREST</th><th>TOTAL RECOVERABLE</th><th aria-label="Actions" />
              </tr></thead>
              <tbody>
                {filtered.map((row) => <tr key={row.id}>
                  <td><div className="buyer-cell"><span className="buyer-avatar">{row.buyer.trim().charAt(0).toUpperCase()}</span><span><strong>{row.buyer}</strong><small>{row.number}{row.gstin ? ` · ${row.gstin}` : ""}{row.taxRisk && <i className="tax-tag">43B(h)</i>}</small></span></div></td>
                  <td><span className="credit-tag">{row.creditDays} days</span></td>
                  <td><span className="date-cell">{formatDate(row.dueDate)}</span><small className="sub-date">{row.daysDelayed > 0 ? "Payment overdue" : row.daysUntilDue < 7 ? "Due shortly" : "Payment due"}</small></td>
                  <td>{row.daysDelayed > 0 ? <span className="delay-pill">{row.daysDelayed}d</span> : <span className={`safe-pill ${row.daysUntilDue < 7 ? "soon" : ""}`}>{row.daysUntilDue < 7 ? `${row.daysUntilDue}d` : "Safe"}</span>}</td>
                  <td className="numeric">{money(row.amount)}</td><td className={`numeric ${row.interest > 0 ? "interest-value" : "muted"}`}>{money(row.interest)}</td>
                  <td className="numeric total-value">{money(row.total)}</td>
                  <td><button className="row-action" title="Draft statutory demand" aria-label={`Draft notice for ${row.buyer}`} onClick={() => openDemand(row.buyer)}><MessageCircle size={15} /></button></td>
                </tr>)}
                {filtered.length === 0 && <tr><td colSpan={8}><div className="empty-state"><span className="empty-illustration"><FileSpreadsheet size={23} /></span><strong>{calculated.length ? "No invoices match this view" : "Your audit workspace is ready"}</strong><p>{calculated.length ? "Try another filter or search term." : "Import your receivables ledger or add an invoice to calculate statutory due dates, interest and 43B(h) exposure."}</p>{!calculated.length && <button className="primary-button compact" onClick={() => { setDraft(blankInvoice()); setManualOpen(true); }}><Plus size={15} /> Add your first invoice</button>}</div></td></tr>}
              </tbody>
            </table>
          </div>
          <div className="table-foot"><span><span className="legend-dot red-dot" /> Statutory overdue <span className="legend-dot amber-dot" /> Due in less than 7 days <span className="legend-dot green-dot" /> Safe</span><span>{filtered.length} of {calculated.length} invoices</span></div>
        </div>

        <aside className="sidebar">
          <div className="section-card">
            <div className="side-card-heading"><span className="side-icon"><BadgeCheck size={17} /></span><div><h3>Statutory framework</h3><p>Rules applied to your ledger</p></div></div>
            <div className="law-item"><span className="law-number">15</span><div><b>Payment timeline</b><p>15 days without written agreement; up to 45 days by agreement.</p></div></div>
            <div className="law-item"><span className="law-number">16</span><div><b>Compound interest</b><p>3× RBI Bank Rate with monthly rests on overdue amounts.</p></div></div>
            <div className="law-item"><span className="law-number tax">43</span><div><b>Income tax exposure</b><p>Qualifying unpaid MSE expense may be disallowed under 43B(h), increasing taxable profit by the disallowed expense.</p></div></div>
            <div className="privacy-note"><LockKeyhole size={14} /><span><b>Your data stays on this device.</b> Files are read and processed locally in your browser.</span></div>
          </div>
          <div className="export-card">
            <div className="export-card-head"><div><span className="eyebrow-mini">READY TO CLOSE THE LOOP</span><h3>Audit &amp; action</h3></div><span className="export-spark"><Sparkles size={16} /></span></div>
            <p>Generate a CA-ready workpaper or take the next step with your buyer.</p>
            <button className="export-button pdf-button" onClick={() => void runWithFeedback(exportPdf, "PDF export")}><span><FileText size={17} /></span><span><b>Export CA audit schedule</b><small>Publication-ready PDF · A4</small></span><ArrowUpRight size={16} /></button>
            <button className="export-button excel-button" onClick={() => void runWithFeedback(downloadExcel, "Excel export")}><span><FileSpreadsheet size={17} /></span><span><b>Export reconciliation</b><small>Calculated workbook · .xlsx</small></span><ArrowUpRight size={16} /></button>
            <button className="notice-launch" onClick={() => openDemand()}><MessageCircle size={16} /> Draft statutory demand <ChevronRight size={15} /></button>
          </div>
          <div className="method-note"><Filter size={14} /><span>Interest calculated through <b>{formatDate(cutoff)}</b>, using the statutory annual rate and actual delay days.</span></div>
        </aside>
      </section>
      <footer className="page-footer"><span>MSME <b>AuditFlow</b> <span className="footer-separator">·</span> A private workspace for better compliance</span><span><ShieldCheck size={13} /> CLIENT-SIDE ONLY <span className="footer-separator">·</span> NOT LEGAL ADVICE</span></footer>

      {mappingOpen && <Modal title="Map your ledger columns" subtitle={`Review detected headers from ${fileName} before importing.`} onClose={() => setMappingOpen(false)}>
        <div className="mapping-intro"><span className="mapping-icon"><FileSpreadsheet size={17} /></span><span>We matched familiar column names automatically. Confirm the fields below to process your ledger.</span></div>
        <div className="mapping-list">{fields.map((field) => <label key={field.key} className="mapping-row"><span>{field.label}{field.required && <i>Required</i>}</span><select value={mapping[field.key] < 0 ? "" : mapping[field.key]} onChange={(event) => setMap(field.key, event.target.value)}><option value="">— Skip this column —</option>{headers.map((header, index) => <option key={index} value={index}>{String(header || `Column ${index + 1}`)}</option>)}</select></label>)}</div>
        <div className="modal-actions"><button className="secondary-button" onClick={() => setMappingOpen(false)}>Cancel</button><button className="primary-button" onClick={importMappedRows}><Upload size={15} /> Import ledger</button></div>
      </Modal>}

      {manualOpen && <Modal title="Add an invoice" subtitle="Enter outstanding invoice details for statutory review." onClose={() => setManualOpen(false)}>
        <form className="invoice-form" onSubmit={saveManual}>
          <label>Buyer name<input autoFocus value={draft.buyer} onChange={(event) => setDraft({ ...draft, buyer: event.target.value })} placeholder="Business / buyer name" /></label>
          <div className="form-row"><label>Invoice number<input value={draft.number} onChange={(event) => setDraft({ ...draft, number: event.target.value })} placeholder="INV-001" /></label><label>GSTIN <span className="optional">Optional</span><input value={draft.gstin} onChange={(event) => setDraft({ ...draft, gstin: event.target.value })} placeholder="Buyer GSTIN" /></label></div>
          <div className="form-row"><label>Invoice date<input type="date" value={draft.invoiceDate} onChange={(event) => setDraft({ ...draft, invoiceDate: event.target.value })} /></label><label>Invoice amount (₹)<input type="number" min="0.01" step="0.01" value={draft.amount || ""} onChange={(event) => setDraft({ ...draft, amount: Number(event.target.value) })} placeholder="0.00" /></label></div>
          <label className="agreement-check"><input type="checkbox" checked={draft.writtenAgreement} onChange={(event) => setDraft({ ...draft, writtenAgreement: event.target.checked, agreedDays: event.target.checked ? (draft.agreedDays ?? 45) : null })} /><span><b>Written payment agreement</b><small>Without one, statutory credit period is 15 days.</small></span></label>
          {draft.writtenAgreement && <label>Agreed credit days <span className="optional">Maximum statutory window: 45 days</span><input type="number" min="0" max="365" value={draft.agreedDays ?? 45} onChange={(event) => setDraft({ ...draft, agreedDays: Number(event.target.value) })} /></label>}
          <div className="modal-actions"><button type="button" className="secondary-button" onClick={() => setManualOpen(false)}>Cancel</button><button className="primary-button" type="submit"><Plus size={15} /> Add to ledger</button></div>
        </form>
      </Modal>}

      {noticeOpen && <Modal title="Statutory demand statement" subtitle="Prepare a buyer-ready notice using your current ledger and audit settings." onClose={() => setNoticeOpen(false)} wide>
        <div className="notice-controls"><label>Buyer<select value={noticeBuyer} onChange={(event) => setNoticeBuyer(event.target.value)}>{Array.from(new Set(calculated.filter((row) => row.daysDelayed > 0).map((row) => row.buyer))).map((buyer) => <option key={buyer}>{buyer}</option>)}</select></label><div className="tone-selector" aria-label="Notice tone">{(["A", "B", "C"] as const).map((item) => <button key={item} className={tone === item ? "active" : ""} onClick={() => setTone(item)}><b>Tone {item}</b><small>{item === "A" ? "Diplomatic" : item === "B" ? "Statutory warning" : "Formal demand"}</small></button>)}</div></div>
        <div className="notice-preview"><div className="preview-label"><span>PREVIEW</span><span>Based on {noticeRows.length} invoice{noticeRows.length === 1 ? "" : "s"}</span></div><pre>{noticeText()}</pre></div>
        <div className="modal-actions notice-actions"><button className="secondary-button" onClick={() => { window.location.href = `mailto:?subject=${encodeURIComponent(`Payment demand — ${noticeBuyer}`)}&body=${encodeURIComponent(noticeText())}`; }}><Mail size={15} /> Send via email</button><button className="secondary-button" onClick={() => window.open(`https://wa.me/?text=${encodeURIComponent(noticeText())}`, "_blank", "noopener,noreferrer")}><MessageCircle size={15} /> WhatsApp Web</button><button className="primary-button" onClick={() => void copyNotice()}><Clipboard size={15} /> Copy statement</button></div>
      </Modal>}
      {toast && <div className="toast"><Check size={16} />{toast}</div>}
    </main>
  );
}

function Metric({ icon, label, value, detail, tone }: { icon: React.ReactNode; label: string; value: string; detail: string; tone: string }) {
  return <article className={`metric-card ${tone}`}><div className="metric-top"><span className="metric-icon">{icon}</span><span className="metric-label">{label}</span><span className="metric-dots">•••</span></div><div className="metric-value">{value}</div><div className="metric-detail">{detail}</div></article>;
}

function Modal({ title, subtitle, children, onClose, wide = false }: { title: string; subtitle: string; children: React.ReactNode; onClose: () => void; wide?: boolean }) {
  return <div className="modal-backdrop" onMouseDown={(event) => { if (event.target === event.currentTarget) onClose(); }}><section className={`modal ${wide ? "wide-modal" : ""}`} role="dialog" aria-modal="true" aria-label={title}>
    <div className="modal-heading"><div><h2>{title}</h2><p>{subtitle}</p></div><button className="bare-icon" onClick={onClose} aria-label="Close dialog"><X size={18} /></button></div>{children}
  </section></div>;
}

export default App;
