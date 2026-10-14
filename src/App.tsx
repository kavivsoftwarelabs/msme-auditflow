import { useMemo, useRef, useState, type ChangeEvent, type DragEvent, type FormEvent } from "react";
import {
  AlertCircle, ArrowDownToLine, ArrowUpRight, BadgeCheck, CalendarDays, Check, ExternalLink,
  ChevronRight, Clipboard, FileSpreadsheet, FileText, Filter, Landmark, LockKeyhole,
  Mail, MessageCircle, Plus, Search, Settings2, ShieldCheck, SlidersHorizontal, Sparkles, Upload,
  X,
} from "lucide-react";
import { addDays, format, parseISO } from "date-fns";
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
import * as XLSX from "@e965/xlsx";
import type { jsPDF as JsPDFDocument } from "jspdf";
import {
  calculateInvoice, formatDate, formatINR, isBlankRow, todayISO,
  financialYearFor, type Invoice,
} from "./engine";
import { emptyMapping, ledgerFields, mapLedgerRows, suggestedLedgerMapping, type MappingKey, type LedgerMapping } from "./ledger";

type NoticeTone = "A" | "B" | "C";
type Perspective = "payables" | "receivables";
type FilterType = "All" | "Overdue" | "Within statutory period";
type EngagementMetadata = {
  auditeeName: string;
  pan: string;
  udyam: string;
  financialYear: string;
  assessmentYear: string;
  cutoff: string;
  auditorFirm: string;
  auditorRegistration: string;
};

const today = todayISO();
const initialFinancialYear = financialYearFor(today);
const blankInvoice = (): Invoice => ({
  id: crypto.randomUUID(), buyer: "", gstin: "", number: "", invoiceDate: today,
  amount: 0, writtenAgreement: false, agreedDays: null,
});

function money(value: number) {
  return formatINR(value).replace("₹", "₹ ");
}

function rupeeSymbolImage() {
  const canvas = document.createElement("canvas");
  canvas.width = 80;
  canvas.height = 80;
  const context = canvas.getContext("2d");
  if (!context) throw new Error("Canvas is unavailable for PDF currency rendering.");
  context.font = "700 72px sans-serif";
  context.fillStyle = "#ffffff";
  context.textAlign = "center";
  context.textBaseline = "middle";
  context.fillText("₹", 40, 40);
  return canvas.toDataURL("image/png");
}

function App() {
  const [payableInvoices, setPayableInvoices] = useState<Invoice[]>([]);
  const [receivableInvoices, setReceivableInvoices] = useState<Invoice[]>([]);
  const [perspective, setPerspective] = useState<Perspective>("receivables");
  const invoices = perspective === "payables" ? payableInvoices : receivableInvoices;
  const [rbiRate, setRbiRate] = useState(5.75);
  const [engagement, setEngagement] = useState<EngagementMetadata>({
    auditeeName: "",
    pan: "",
    udyam: "",
    financialYear: initialFinancialYear.financialYear,
    assessmentYear: initialFinancialYear.assessmentYear,
    cutoff: today,
    auditorFirm: "",
    auditorRegistration: "",
  });
  const [msmeType, setMsmeType] = useState<"Micro" | "Small">("Micro");
  const cutoff = engagement.cutoff;
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [query, setQuery] = useState("");
  const [filter, setFilter] = useState<FilterType>("All");
  const [sortAsc, setSortAsc] = useState(false);
  const [fileRows, setFileRows] = useState<unknown[][] | null>(null);
  const [fileName, setFileName] = useState("");
  const [headers, setHeaders] = useState<unknown[]>([]);
  const [mapping, setMapping] = useState<LedgerMapping>(emptyMapping());
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
    () => invoices.map((invoice) => {
      const supplierIsMse = invoice.supplierMseType
        ? invoice.supplierMseType !== "Not MSE"
        : msmeType === "Micro" || msmeType === "Small";
      const statutoryApplicable = perspective === "receivables" || supplierIsMse;
      return calculateInvoice(invoice, cutoff, rbiRate, statutoryApplicable, perspective === "payables" && supplierIsMse);
    }),
    [invoices, cutoff, rbiRate, perspective, msmeType],
  );
  const filtered = useMemo(() => {
    const search = query.trim().toLowerCase();
    return calculated
      .filter((row) => filter === "All" || (filter === "Overdue" ? row.daysDelayed > 0 : row.statutoryApplicable && row.daysDelayed === 0))
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
  const disallowedLiability = perspective === "payables" ? overdue.reduce((sum, row) => sum + row.disallowedAmount, 0) : 0;
  const counterpartiesAtRisk = new Set(overdue.map((row) => row.buyer.trim().toLowerCase())).size;
  const totalRecoverable = calculated.reduce((sum, row) => sum + row.total, 0);
  const noticeRows = calculated.filter((row) => row.buyer === noticeBuyer && row.daysDelayed > 0);
  const statutoryRate = 3 * rbiRate;

  function updateEngagement<K extends keyof EngagementMetadata>(key: K, value: EngagementMetadata[K]) {
    setEngagement((current) => ({ ...current, [key]: value }));
  }

  function appendInvoices(rows: Invoice[]) {
    const update = (current: Invoice[]) => [...current, ...rows];
    if (perspective === "payables") setPayableInvoices(update);
    else setReceivableInvoices(update);
  }

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
      setMapping(suggestedLedgerMapping(fileHeaders));
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
    const { invoices: imported, errors } = mapLedgerRows(
      fileRows,
      mapping,
      perspective === "payables" ? "sundry creditor / supplier" : "sundry debtor / trade receivable",
    );
    if (imported.length === 0) {
      notify(errors[0] ?? "No valid invoice rows were found.");
      return;
    }
    appendInvoices(imported);
    setMappingOpen(false);
    setFileRows(null);
    notify(`${imported.length} invoice${imported.length === 1 ? "" : "s"} imported${errors.length ? `; ${errors.length} invalid row${errors.length === 1 ? "" : "s"} skipped (first: ${errors[0]})` : ""}.`);
  }

  function createSampleInvoices(): Invoice[] {
    const sampleDate = (daysBeforeCutoff: number) =>
      format(addDays(parseISO(cutoff), -daysBeforeCutoff), "yyyy-MM-dd");
    return [
      {
        id: crypto.randomUUID(), buyer: perspective === "payables" ? "Kaveri Precision Components" : "Zenith Auto Components Ltd", gstin: "07AAAAA0000A1Z5",
        number: "INV-26-27/014", invoiceDate: sampleDate(5), amount: 87500,
        writtenAgreement: false, agreedDays: null, supplierMseType: "Micro",
      },
      {
        id: crypto.randomUUID(), buyer: perspective === "payables" ? "Bharat Industrial Gases" : "Apex Industrial Systems Pvt Ltd", gstin: "27BBBBB1111B1Z4",
        number: "INV-26-27/089", invoiceDate: sampleDate(60), amount: 250000,
        writtenAgreement: true, agreedDays: 60, supplierMseType: "Small",
      },
      {
        id: crypto.randomUUID(), buyer: perspective === "payables" ? "Shakti Machine Tools Pvt Ltd" : "Pragati Machine Tools", gstin: "29CCCCC2222C1Z3",
        number: "INV-26-27/132", invoiceDate: sampleDate(40), amount: 143750,
        writtenAgreement: false, agreedDays: null, supplierMseType: "Micro",
      },
    ];
  }

  function loadSampleInvoices() {
    appendInvoices(createSampleInvoices());
    notify(`3 sample ${perspective === "payables" ? "supplier bills" : "receivables"} added. The 60-day agreement is capped at the 45-day statutory maximum.`);
  }

  function engagementSheet() {
    return [
      ["AUDIT ENGAGEMENT METADATA", ""],
      ["Audit Perspective", perspective === "payables" ? "Trade Payables / Creditors (Auditee as Buyer)" : "Trade Receivables / Debtors (Auditee as Supplier)"],
      ["Auditee / Client Legal Name", engagement.auditeeName],
      ["Client PAN", engagement.pan],
      ["Udyam Registration No.", engagement.udyam],
      ["Financial Year (F.Y.)", engagement.financialYear],
      ["Assessment Year (A.Y.)", engagement.assessmentYear],
      ["Audit Cutoff Date", cutoff],
      ["Auditor / CA Firm Name", engagement.auditorFirm],
      ["Auditor FRN / M.No.", engagement.auditorRegistration],
      ["Supplier MSE Classification (fallback)", msmeType],
      ["RBI Bank Rate", `${rbiRate}%`],
      ["Section 16 Statutory Rate", `${statutoryRate}% p.a. (3x RBI Bank Rate)`],
      ["Confidentiality", "Processed locally in browser; no financial data transmitted."],
    ];
  }

  async function downloadTemplate() {
    const sample = createSampleInvoices();
    const worksheet = XLSX.utils.aoa_to_sheet([
      [
        perspective === "payables" ? "Supplier Name" : "Party Name",
        "GSTIN/UIN",
        "Voucher Ref/Inv No",
        "Bill Date",
        "Pending Amount",
        "Agreed Credit Period (Days)",
        ...(perspective === "payables" ? ["Supplier MSE Classification"] : []),
      ],
      ...sample.map((invoice) => [
        invoice.buyer,
        invoice.gstin,
        invoice.number,
        invoice.invoiceDate,
        invoice.amount,
        invoice.agreedDays ?? "",
        ...(perspective === "payables" ? [invoice.supplierMseType] : []),
      ]),
    ]);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Ledger Outstandings");
    XLSX.writeFile(workbook, "MSME-AuditFlow-Tally-Ledger-Template.xlsx");
    notify(`Tally-ready ${perspective} template downloaded with 3 sample rows.`);
  }

  async function downloadExcel() {
    if (!calculated.length) {
      notify("Add or import invoices before exporting.");
      return;
    }
    const rows = calculated.map((row) => ({
      [perspective === "payables" ? "Sundry Creditor / Supplier Name" : "Trade Receivable / Debtor"]: row.buyer,
      GSTIN: row.gstin,
      ...(perspective === "payables" ? { "Supplier MSE Classification": row.supplierMseType ?? msmeType } : {}),
      "Voucher Ref / Invoice Number": row.number,
      "Invoice Date": row.invoiceDate,
      "Principal (INR)": row.amount,
      "Written Agreement": row.writtenAgreement ? "Yes" : "No",
      "Agreed Credit Period (Days)": row.agreedDays ?? "",
      "Statutory Credit Window (Days)": row.statutoryApplicable ? row.creditDays : "",
      "MSE Statutory Applicability": row.statutoryApplicable ? "Yes" : "No",
      "Appointed Due Date u/s 15": row.statutoryApplicable ? row.dueDate : "",
      "Overdue Beyond Appointed Day (Days)": row.daysDelayed,
      "Statutory Penal Interest u/s 16 (INR)": Number(row.interest.toFixed(2)),
      ...(perspective === "payables" ? { "Tax Disallowance Liability u/s 43B(h) (INR)": row.disallowedAmount } : {}),
      [perspective === "payables" ? "Total Payable (INR)" : "Total Recoverable (INR)"]: Number(row.total.toFixed(2)),
      "Within Statutory Period": !row.statutoryApplicable ? "Not Applicable" : row.daysDelayed === 0 ? "Yes" : "No",
    }));
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(engagementSheet()), "Engagement Metadata");
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(rows), "Audit Annexure");
    XLSX.writeFile(workbook, "MSME-AuditFlow-Audit-Annexure.xlsx");
    notify(`${perspective === "payables" ? "Payables disallowance schedule" : "Receivables recovery schedule"} and engagement metadata exported to Excel.`);
  }

  async function exportPdf() {
    if (!calculated.length) {
      notify("Add or import invoices before exporting.");
      return;
    }
    const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
    const pageWidth = doc.internal.pageSize.getWidth();
    doc.setFont("helvetica", "bold");
    doc.setFontSize(12);
    doc.text(
      perspective === "payables"
        ? "ANNEXURE OF SUNDRY CREDITORS (PAYABLES - FORM 3CD CL. 22)"
        : "ANNEXURE OF SUNDRY DEBTORS (RECEIVABLES - MSMED SEC 16)",
      pageWidth / 2,
      15,
      { align: "center", maxWidth: pageWidth - 28 },
    );
    doc.setFont("helvetica", "normal");
    doc.setFontSize(8);
    doc.text(
      perspective === "payables"
        ? "ANNEXURE: SCHEDULE OF TRADE PAYABLES TO MSEs & TAX DISALLOWANCE u/s 43B(h) (Form 3CD Clause 22)"
        : "ANNEXURE: SCHEDULE OF OVERDUE TRADE RECEIVABLES & ACCRUED INTEREST u/s 16 (MSMED Act)",
      pageWidth / 2,
      20,
      { align: "center", maxWidth: pageWidth - 28 },
    );
    doc.setFontSize(7.5);
    doc.text("Prepared for Statutory Audit under Section 44AB of the Income-tax Act, 1961 / Companies Act, 2013", pageWidth / 2, 25, { align: "center" });
    doc.setFillColor(245, 247, 249);
    doc.setDrawColor(164, 174, 184);
    doc.roundedRect(14, 29, pageWidth - 28, 37, 1.5, 1.5, "FD");
    doc.line(pageWidth / 2, 32, pageWidth / 2, 63);
    const leftMeta = [
      ["Auditee", engagement.auditeeName || "—"],
      ["PAN", engagement.pan || "—"],
      ["Udyam No.", engagement.udyam || "—"],
      ["F.Y. / A.Y.", `${engagement.financialYear || "—"} / ${engagement.assessmentYear || "—"}`],
    ];
    const rightMeta = [
      ["Audit Working Paper Ref", "MSME-WP-01"],
      ["Cutoff Date", formatDate(cutoff)],
      ["Auditor", engagement.auditorFirm || "—"],
      ["Auditor FRN / M.No.", engagement.auditorRegistration || "—"],
    ];
    doc.setFontSize(7);
    leftMeta.forEach(([label, value], index) => {
      const y = 36 + index * 7.5;
      doc.setFont("helvetica", "bold");
      doc.text(`${label}:`, 18, y);
      doc.setFont("helvetica", "normal");
      doc.text(value, 46, y, { maxWidth: 90 });
    });
    rightMeta.forEach(([label, value], index) => {
      const y = 36 + index * 7.5;
      doc.setFont("helvetica", "bold");
      doc.text(`${label}:`, pageWidth / 2 + 5, y);
      doc.setFont("helvetica", "normal");
      doc.text(value, pageWidth / 2 + 42, y, { maxWidth: 87 });
    });
    doc.setFont("helvetica", "normal");
    const rupeeMark = rupeeSymbolImage();
    autoTable(doc, {
      startY: 70,
      head: [["Sl.", perspective === "payables" ? "Sundry Creditor / Supplier" : "Sundry Debtor / Trade Receivable", "GSTIN", "Inv No.", "Inv Date", "Appointed Due Date", "Delay (Days)", "Principal (     )", perspective === "payables" ? "Sec 16 Interest Payable (     )" : "Sec 16 Interest Receivable (     )", perspective === "payables" ? "43B(h) Disallowed Amount (     )" : "Total Recoverable (     )"]],
      body: calculated.map((row, index) => [
        String(index + 1), row.buyer, row.gstin || "—", row.number, formatDate(row.invoiceDate),
        row.statutoryApplicable ? formatDate(row.dueDate) : "Not applicable", row.statutoryApplicable ? String(row.daysDelayed) : "—", row.amount.toFixed(2),
        row.interest.toFixed(2), perspective === "payables" ? row.disallowedAmount.toFixed(2) : row.total.toFixed(2),
      ]),
      foot: [[
        { content: "TOTALS", colSpan: 7, styles: { halign: "right", fontStyle: "bold" } },
        { content: totalPrincipal.toFixed(2), styles: { fontStyle: "bold", halign: "right" } },
        { content: interestAccrued.toFixed(2), styles: { fontStyle: "bold", halign: "right" } },
        { content: (perspective === "payables" ? disallowedLiability : totalRecoverable).toFixed(2), styles: { fontStyle: "bold", halign: "right" } },
      ]],
      showFoot: "lastPage",
      styles: { font: "helvetica", fontSize: 6.4, cellPadding: 2, textColor: [35, 42, 49], lineColor: [194, 201, 207], lineWidth: 0.15 },
      headStyles: { fillColor: [55, 65, 81], textColor: [255, 255, 255], fontStyle: "bold", halign: "center", fontSize: 6.2 },
      footStyles: { fillColor: [229, 233, 238], textColor: [25, 32, 40], lineColor: [137, 148, 159], lineWidth: 0.25 },
      alternateRowStyles: { fillColor: [247, 248, 250] },
      margin: { left: 14, right: 14, bottom: 16 },
      didDrawCell: (data) => {
        const prefixes: Record<number, string> = {
          7: "Principal (",
          8: perspective === "payables" ? "Sec 16 Interest Payable (" : "Sec 16 Interest Receivable (",
          9: perspective === "payables" ? "43B(h) Disallowed Amount (" : "Total Recoverable (",
        };
        const prefix = data.section === "head" && data.row.index === 0 ? prefixes[data.column.index] : undefined;
        if (!prefix) return;
        doc.setFont("helvetica", "bold");
        doc.setFontSize(6.2);
        const symbolX = data.cell.x + 2 + doc.getTextWidth(prefix);
        doc.addImage(rupeeMark, "PNG", symbolX, data.cell.y + (data.cell.height - 2.5) / 2, 2.5, 2.5);
      },
      didDrawPage: () => {
        doc.setFontSize(7);
        doc.setTextColor(100, 110, 120);
        doc.text("MSME-WP-01", 14, doc.internal.pageSize.getHeight() - 7);
        doc.text(`Page ${doc.getNumberOfPages()}`, pageWidth - 14, doc.internal.pageSize.getHeight() - 7, { align: "right" });
      },
    });
    const tableBottom = (doc as JsPDFDocument & { lastAutoTable?: { finalY: number } }).lastAutoTable?.finalY ?? 40;
    let noteY = tableBottom + 7;
    if (noteY + 68 > doc.internal.pageSize.getHeight() - 11) {
      doc.addPage();
      noteY = 18;
    }
    doc.setFontSize(7);
    doc.setTextColor(45, 52, 60);
    const citations = [
      "Basis: Section 15, MSMED Act, 2006 — appointed day / payment period is subject to the statutory 15-day limit without written agreement and 45-day maximum by agreement.",
      `Section 16 — compound interest with monthly rests at three times the RBI Bank Rate; configured rate ${statutoryRate.toFixed(2)}% p.a. ${perspective === "payables" ? "Section 23 — such penal interest is not deductible." : "This schedule presents interest receivable by the auditee as supplier."}`,
      ...(perspective === "payables" ? ["Section 43B(h), Income-tax Act, 1961 — qualifying sums payable to Micro / Small Enterprises are allowed as a deduction on actual payment, subject to statutory provisions."] : []),
      "All monetary amounts are stated in Indian Rupees (INR).",
      `Schedule prepared from ledger balances as at ${formatDate(cutoff)}. Verify Udyam status, actual payments, appointed day, applicable law and tax treatment before reliance.`,
    ];
    citations.forEach((text, index) => {
      doc.text(text, 14, noteY + index * 4.5, { maxWidth: pageWidth - 28 });
    });
    const signatureY = noteY + 23;
    doc.setDrawColor(148, 157, 166);
    doc.line(14, signatureY, 145, signatureY);
    doc.line(152, signatureY, pageWidth - 14, signatureY);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(7.5);
    doc.text("Auditee Declaration", 14, signatureY + 4);
    doc.text("Auditor's Working Paper Reference", 152, signatureY + 4);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(7);
    doc.text("Certified that the above classification is based on valid Udyam certificates verified.", 14, signatureY + 9, { maxWidth: 128 });
    doc.text("Verified with reference to books of accounts, Tally ledgers, and delivery challans.", 152, signatureY + 9, { maxWidth: 128 });
    doc.line(24, signatureY + 23, 94, signatureY + 23);
    doc.line(181, signatureY + 23, 251, signatureY + 23);
    doc.text("Authorised Signatory", 24, signatureY + 27);
    doc.text("Chartered Accountant", 181, signatureY + 27);
    doc.text("(Client Stamp & Date)", 24, signatureY + 31);
    doc.text(`(${engagement.auditorRegistration || "M.No. / FRN"})`, 181, signatureY + 31);
    doc.save(`MSME-AuditFlow-${perspective === "payables" ? "Payables-43Bh" : "Receivables-Sec16"}-Annexure.pdf`);
    notify("Statutory audit annexure PDF generated.");
  }

  function saveManual(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!draft.buyer.trim() || !draft.number.trim() || !draft.invoiceDate || draft.amount <= 0) {
      notify("Enter a buyer, invoice number, date and positive amount.");
      return;
    }
    appendInvoices([{ ...draft, buyer: draft.buyer.trim(), number: draft.number.trim(), gstin: draft.gstin.trim() }]);
    setManualOpen(false);
    setDraft(blankInvoice());
    notify("Invoice added to the audit ledger.");
  }

  function openDemand(buyer?: string) {
    if (perspective !== "receivables") return;
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
      `• ${row.number} dated ${formatDate(row.invoiceDate)} — principal ${money(row.amount)}, appointed day ${formatDate(row.dueDate)}, ${row.daysDelayed} days beyond the statutory appointed day`,
    ).join("\n");
    const opening = tone === "A"
      ? `Dear ${buyer},\n\nAs part of our mutual audit and MSME payment-compliance review, we request your confirmation and prompt settlement of the following outstanding trade receivables.`
      : tone === "B"
        ? `Dear ${buyer},\n\nFORMAL STATUTORY PAYMENT REMINDER\n\nThe following trade receivables have passed their applicable appointed days under Section 15 of the Micro, Small and Medium Enterprises Development Act, 2006. Section 16 provides for compound interest with monthly rests at three times the RBI Bank Rate (currently configured at ${statutoryRate.toFixed(2)}% p.a.).`
        : `To: ${buyer}\n\nSUBJECT: FORMAL DEMAND FOR PAYMENT OF OVERDUE MSME TRADE RECEIVABLES\n\nThis is a formal demand for payment of the amounts set out below. The applicable payment period is governed by Section 15 of the Micro, Small and Medium Enterprises Development Act, 2006. Pursuant to Section 16, delayed payment attracts compound interest with monthly rests at three times the RBI Bank Rate. If payment is not made, the matter may be referred to the Micro and Small Enterprises Facilitation Council / MSME Samadhaan, subject to applicable law and eligibility.`;
    return `${opening}\n\nTrade receivables beyond appointed day:\n${invoiceLines}\n\nPrincipal outstanding: ${money(dueAmount)}\nStatutory penal interest u/s 16 as of ${formatDate(cutoff)}: ${money(interest)}\nTotal currently due: ${money(dueAmount + interest)}\n\nPlease arrange settlement and share remittance confirmation. This statement is based on the records presently available and is without prejudice to applicable legal rights. Kindly contact us promptly to reconcile any discrepancy.\n\nRegards,\n[Supplier / Authorised Representative]`;
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

  const filters: FilterType[] = ["All", "Overdue", "Within statutory period"];

  return (
    <main className="app-shell">
      <header className="topbar">
        <a className="brand" href="#" aria-label="MSME AuditFlow home">
          <span className="brand-mark"><Landmark size={20} strokeWidth={2.1} /></span>
          <span className="brand-copy"><strong>MSME <em>AuditFlow</em></strong><small>STATUTORY INTEREST ENGINE</small></span>
        </a>
        <div className="topbar-right">
          <a className="feedback-link" href="https://docs.google.com/forms/d/e/1FAIpQLSd2oONI0bNEyR-0VFMmDTdhc5GzvPoPooJlQt_ZHJfYzzTsbw/viewform" target="_blank" rel="noopener noreferrer"><MessageCircle size={14} /> Share feedback <ExternalLink size={12} /></a>
          <div className="privacy-pill"><LockKeyhole size={14} /> 100% Private &amp; Client-Side</div>
          <button className={`icon-button settings-trigger ${settingsOpen ? "active" : ""}`} onClick={() => setSettingsOpen(!settingsOpen)} aria-label="Open global settings"><Settings2 size={18} /></button>
        </div>
        {settingsOpen && <section className="settings-popover">
          <div className="popover-heading"><div><strong>Audit settings</strong><span>Applied across calculations and exports</span></div><button className="bare-icon" onClick={() => setSettingsOpen(false)} aria-label="Close settings"><X size={16} /></button></div>
          <label>RBI Bank Rate <span className="input-suffix"><input type="number" min="0" max="100" step="0.01" value={rbiRate} onChange={(event) => setRbiRate(Math.min(100, Math.max(0, Number(event.target.value) || 0)))} /><i>%</i></span></label>
          <div className="rate-summary"><span>Section 16 statutory rate</span><b>{statutoryRate.toFixed(2)}% <small>p.a.</small></b></div>
          <label>Default supplier MSE classification<select value={msmeType} onChange={(event) => setMsmeType(event.target.value as "Micro" | "Small")}><option>Micro</option><option>Small</option></select></label>
          <p className="settings-note">{perspective === "payables" ? "43B(h) flags use each supplier's imported classification, falling back to this default, and assume balances remain unpaid through the cutoff." : "The 43B(h) disallowance calculation is disabled for trade receivables."}</p>
        </section>}
      </header>

      <section className="page-heading">
        <div><div className="eyebrow"><span className="eyebrow-line" /> MSME ACT, 2006 <span className="eyebrow-dot" /> SECTION 43B(h)</div>
          <h1>Audit Working Paper: <span>Sec 43B(h) &amp; MSMED Compliance Schedule</span></h1>
          <p>{perspective === "payables" ? "Tax Audit Form 3CD (Clause 22) and Section 43B(h) compliance for supplier payables." : "MSMED Act Sections 15/16 working-capital recovery and statutory receivables schedule."}</p>
        </div>
        <div className="heading-asof"><CalendarDays size={15} /><span>CALCULATED AS OF</span><strong>{format(parseISO(cutoff), "dd MMM yyyy")}</strong></div>
      </section>

      <section className="metrics-grid" aria-label="Executive summary">
        <Metric icon={<FileText size={17} />} label={perspective === "payables" ? "PAYABLES ANALYZED" : "RECEIVABLES ANALYZED"} value={String(calculated.length)} detail={`${money(totalPrincipal)} principal`} tone="ink" />
        <Metric icon={<AlertCircle size={17} />} label="BEYOND APPOINTED DAY (SEC 15)" value={money(overduePrincipal)} detail={`${overdue.length} overdue ${perspective === "payables" ? "payable" : "receivable"}${overdue.length === 1 ? "" : "s"}`} tone="amber" />
        <Metric icon={<Sparkles size={17} />} label={perspective === "payables" ? "INTEREST PAYABLE u/s 16" : "STATUTORY PENAL INTEREST RECEIVABLE u/s 16"} value={money(interestAccrued)} detail={`${statutoryRate.toFixed(2)}% p.a. compounded monthly at 3x RBI Rate (${rbiRate}% Bank Rate)`} tone="green" />
        {perspective === "payables"
          ? <Metric icon={<ShieldCheck size={17} />} label="TAX DISALLOWANCE LIABILITY u/s 43B(h)" value={money(disallowedLiability)} detail={`${counterpartiesAtRisk} sundry creditor${counterpartiesAtRisk === 1 ? "" : "s"} · unpaid MSE balances`} tone="red" />
          : <Metric icon={<ShieldCheck size={17} />} label="TOTAL RECOVERABLE WITH STATUTORY INTEREST" value={money(totalRecoverable)} detail={`${counterpartiesAtRisk} sundry debtor${counterpartiesAtRisk === 1 ? "" : "s"} at risk · ${money(overduePrincipal)} overdue principal`} tone="red" />}
      </section>

      <section className="workbench">
        <div className="panel ledger-panel">
          <div className="panel-header">
            <div className="panel-title"><span className="panel-icon"><FileSpreadsheet size={17} /></span><div><h2>{perspective === "payables" ? "Trade payables verification schedule" : "Trade receivables verification schedule"}</h2><p>Working paper detail by voucher / invoice reference</p></div></div>
            <button className="primary-button compact" onClick={() => { setDraft(blankInvoice()); setManualOpen(true); }}><Plus size={16} /> Add {perspective === "payables" ? "payable" : "receivable"}</button>
          </div>
          <section className="engagement-card" aria-label="Audit engagement metadata">
            <div className="engagement-heading"><div><span className="engagement-kicker">AUDIT FILE</span><h3>Engagement metadata</h3></div><span className="wp-reference">WP REF. MSME-WP-01</span></div>
            <div className="engagement-grid">
              <label>Auditee / Client Legal Name<input value={engagement.auditeeName} onChange={(event) => updateEngagement("auditeeName", event.target.value)} placeholder="ABC Precision Engineering Pvt Ltd" /></label>
              <label>Client PAN<input value={engagement.pan} onChange={(event) => updateEngagement("pan", event.target.value.toUpperCase())} placeholder="ABCDE1234F" /></label>
              <label>Udyam Registration No.<input value={engagement.udyam} onChange={(event) => updateEngagement("udyam", event.target.value.toUpperCase())} placeholder="UDYAM-UP-28-0000000" /></label>
              <label>Financial Year (F.Y.)<input value={engagement.financialYear} onChange={(event) => updateEngagement("financialYear", event.target.value)} placeholder="2025–26" /></label>
              <label>Assessment Year (A.Y.)<input value={engagement.assessmentYear} onChange={(event) => updateEngagement("assessmentYear", event.target.value)} placeholder="2026–27" /></label>
              <label>Audit Cutoff Date<input type="date" value={cutoff} onChange={(event) => event.target.value && updateEngagement("cutoff", event.target.value)} /></label>
              <label>Auditor / CA Firm Name<input value={engagement.auditorFirm} onChange={(event) => updateEngagement("auditorFirm", event.target.value)} placeholder="M/s [Firm Name] & Co., Chartered Accountants" /></label>
              <label>Auditor FRN / M.No.<input value={engagement.auditorRegistration} onChange={(event) => updateEngagement("auditorRegistration", event.target.value)} placeholder="FRN / Membership No." /></label>
            </div>
          </section>
          <div className="perspective-switch" role="group" aria-label="Audit perspective">
            <span className="perspective-label">AUDIT PERSPECTIVE</span>
            <div className="segmented-control">
              <button type="button" className={perspective === "payables" ? "active" : ""} aria-pressed={perspective === "payables"} onClick={() => { setPerspective("payables"); setFilter("All"); }}>Trade Payables / Creditors <small>Auditee as Buyer</small></button>
              <button type="button" className={perspective === "receivables" ? "active" : ""} aria-pressed={perspective === "receivables"} onClick={() => { setPerspective("receivables"); setFilter("All"); }}>Trade Receivables / Debtors <small>Auditee as Supplier</small></button>
            </div>
            <p>{perspective === "payables" ? "For Form 3CD Clause 22 and Section 43B(h): identify qualifying MSE supplier invoices unpaid beyond their appointed day." : "For MSMED Act Sections 15/16: calculate statutory interest receivable on overdue buyer balances."}</p>
          </div>
          <div className="confidentiality-callout"><LockKeyhole size={18} /><div><strong>ICAI Confidentiality &amp; Data Security Note (Clause 1, Part I, Second Schedule)</strong><p>This utility executes 100% locally in your web browser memory. No debtor names, GSTINs, turnover values, or ledger balances are transmitted to or stored on any external server. After the application assets are loaded, disconnecting your internet will not affect ledger processing, calculations or file generation.</p></div></div>
          <div className={`dropzone ${dragging ? "dragging" : ""}`} onDragOver={(event) => { event.preventDefault(); setDragging(true); }} onDragLeave={() => setDragging(false)} onDrop={onDrop} onClick={() => inputRef.current?.click()}>
            <div className="upload-icon"><Upload size={18} /></div>
            <div className="upload-copy"><strong>Drop your ledger here <span>or browse files</span></strong><small>{fileName || "Excel or CSV · .xlsx, .xls, .csv"}</small></div>
            <input ref={inputRef} type="file" accept=".xlsx,.xls,.csv" hidden onChange={onFileChange} />
          </div>
          <div className="template-actions"><button className="text-button" onClick={() => void runWithFeedback(downloadTemplate, "Tally template download")}><ArrowDownToLine size={14} /> Download Tally-Ready Ledger Template (.xlsx)</button><button className="text-button" onClick={loadSampleInvoices}><Sparkles size={14} /> Load 3 sample {perspective === "payables" ? "supplier bills" : "receivables"}</button></div>
          <details className="statutory-accordion"><summary><span><Landmark size={15} /> Statutory References &amp; Basis of Audit Verification</span><ChevronRight size={15} /></summary><div className="statutory-citations">
            <p><b>Section 15, MSMED Act, 2006:</b> Payment within the statutory ceiling of 15 days where no written agreement exists, or up to a maximum statutory ceiling of 45 days by written agreement, notwithstanding a longer contractual period.</p>
            <p><b>Section 16, MSMED Act, 2006:</b> Mandatory penal compound interest with monthly rests at 3x the RBI Bank Rate for the period of delay commencing from the appointed day.</p>
            <p><b>Section 23, MSMED Act, 2006:</b> Penal interest paid or payable under the Act is not deductible as an expense when computing income tax.</p>
            {perspective === "payables" && <p><b>Section 43B(h), Income-tax Act, 1961:</b> A sum payable to a registered Micro or Small Enterprise beyond the statutory due date is allowed as a deduction only on actual payment, subject to the Act.</p>}
          </div></details>
          <div className="table-toolbar">
            <div className="filter-tabs" role="tablist" aria-label={`Filter trade ${perspective}`}>
              {filters.map((item) => <button key={item} className={filter === item ? "selected" : ""} onClick={() => setFilter(item)} role="tab" aria-selected={filter === item}>{item === "All" ? `All ${perspective}` : item === "Overdue" ? "Beyond appointed day" : item}{item === "Overdue" && overdue.length > 0 && <span>{overdue.length}</span>}</button>)}
            </div>
            <label className="search-box"><Search size={15} /><input value={query} onChange={(event) => setQuery(event.target.value)} placeholder={`Search ${perspective === "payables" ? "supplier" : "buyer"} or invoice...`} /></label>
          </div>
          <div className="table-scroll">
            <table>
              <thead><tr>
                <th>{perspective === "payables" ? "SUNDRY CREDITOR / SUPPLIER NAME" : "SUNDRY DEBTOR / TRADE RECEIVABLE"}</th><th>APPOINTED PERIOD</th><th><button className="sort-button" onClick={() => setSortAsc(!sortAsc)}>APPOINTED DUE DATE <SlidersHorizontal size={12} /></button></th>
                <th>DELAY (DAYS)</th><th>PRINCIPAL (₹)</th><th>STATUTORY PENAL INTEREST u/s 16 (₹)</th><th>{perspective === "payables" ? "43B(h) DISALLOWED (₹)" : "TOTAL RECOVERABLE (₹)"}</th><th aria-label="Actions" />
              </tr></thead>
              <tbody>
                {filtered.map((row) => <tr key={row.id}>
                  <td><div className="buyer-cell"><span className="buyer-avatar">{row.buyer.trim().charAt(0).toUpperCase()}</span><span><strong>{row.buyer}</strong><small>{row.number}{row.gstin ? ` · ${row.gstin}` : ""}{perspective === "payables" && row.supplierMseType && <i className="tax-tag">{row.supplierMseType === "Not MSE" ? "Not MSE" : `${row.supplierMseType} MSE`}</i>}{perspective === "payables" && row.taxRisk && <i className="tax-tag">43B(h)</i>}</small></span></div></td>
                  <td><span className="credit-tag">{row.statutoryApplicable ? `${row.creditDays} days` : "Not applicable"}</span></td>
                  <td>{row.statutoryApplicable ? <><span className="date-cell">{formatDate(row.dueDate)}</span><small className="sub-date">{row.daysDelayed > 0 ? "Overdue beyond appointed day (Sec 15)" : row.daysUntilDue < 7 ? "Appointed day approaching" : "Within statutory period"}</small></> : <span className="muted">Not applicable</span>}</td>
                  <td>{!row.statutoryApplicable ? <span className="muted">—</span> : row.daysDelayed > 0 ? <span className="delay-pill">{row.daysDelayed}d</span> : <span className={`safe-pill ${row.daysUntilDue < 7 ? "soon" : ""}`}>{row.daysUntilDue < 7 ? `${row.daysUntilDue}d` : "Within statutory period"}</span>}</td>
                  <td className="numeric">{money(row.amount)}</td><td className={`numeric ${row.interest > 0 ? "interest-value" : "muted"}`}>{money(row.interest)}</td>
                  <td className="numeric total-value">{money(perspective === "payables" ? row.disallowedAmount : row.total)}</td>
                  <td>{perspective === "receivables" && <button className="row-action" title="Draft statutory demand" aria-label={`Draft notice for ${row.buyer}`} onClick={() => openDemand(row.buyer)}><MessageCircle size={15} /></button>}</td>
                </tr>)}
                {filtered.length === 0 && <tr><td colSpan={8}><div className="empty-state"><span className="empty-illustration"><FileSpreadsheet size={23} /></span><strong>{calculated.length ? `No trade ${perspective} match this view` : "Your audit working paper is ready"}</strong><p>{calculated.length ? "Try another statutory-period filter or search term." : `Import the Tally ledger or add a ${perspective === "payables" ? "supplier bill" : "receivable"} to calculate the appointed day and Section 16 statutory interest${perspective === "payables" ? " and 43B(h) disallowance" : ""}.`}</p>{!calculated.length && <button className="primary-button compact" onClick={() => { setDraft(blankInvoice()); setManualOpen(true); }}><Plus size={15} /> Add first {perspective === "payables" ? "payable" : "receivable"}</button>}</div></td></tr>}
              </tbody>
            </table>
          </div>
          <div className="table-foot"><span><span className="legend-dot red-dot" /> Beyond appointed day (Sec 15) <span className="legend-dot amber-dot" /> Appointed day within 7 days <span className="legend-dot green-dot" /> Within statutory period</span><span>{filtered.length} of {calculated.length} trade {perspective}</span></div>
        </div>

        <aside className="sidebar">
          <div className="section-card">
            <div className="side-card-heading"><span className="side-icon"><BadgeCheck size={17} /></span><div><h3>Statutory framework</h3><p>Rules applied to your ledger</p></div></div>
            <div className="law-item"><span className="law-number">15</span><div><b>Payment timeline</b><p>15 days without written agreement; up to 45 days by agreement.</p></div></div>
            <div className="law-item"><span className="law-number">16</span><div><b>Compound interest</b><p>3× RBI Bank Rate with monthly rests on overdue amounts.</p></div></div>
            {perspective === "payables"
              ? <div className="law-item"><span className="law-number tax">43</span><div><b>Income tax exposure</b><p>Qualifying unpaid MSE trade payables may be disallowed under Section 43B(h), increasing taxable profit by the disallowed expense.</p></div></div>
              : <div className="law-item"><span className="law-number tax">22</span><div><b>Receivables disclosure</b><p>Overdue balances due from buyers are presented for MSMED Act Section 22 disclosure and Section 16 interest recovery; no 43B(h) disallowance is attributed to receivables.</p></div></div>}
            <div className="privacy-note"><LockKeyhole size={14} /><span><b>Your data stays on this device.</b> Files are read and processed locally in your browser.</span></div>
          </div>
          <div className="export-card">
            <div className="export-card-head"><div><span className="eyebrow-mini">AUDIT DOCUMENTATION</span><h3>Generate annexures</h3></div><span className="export-spark"><Sparkles size={16} /></span></div>
            <p>Generate the audit working paper {perspective === "receivables" ? "and statutory demand statement" : "and management disallowance schedule"} for review.</p>
            <button className="export-button pdf-button" onClick={() => void runWithFeedback(exportPdf, "PDF generation")}><span><FileText size={17} /></span><span><b>Generate Audit Annexure / Working Paper</b><small>Standard Statutory Audit Working Paper (Format 3CD / Schedule III Compliant)</small></span><ArrowUpRight size={16} /></button>
            <p className="audit-standard-note">Prepared in reference to SA 230 (Audit Documentation) &amp; Clause 22 of Form 3CD</p>
            <button className="export-button excel-button" onClick={() => void runWithFeedback(downloadExcel, "Excel generation")}><span><FileSpreadsheet size={17} /></span><span><b>Generate Audit Annexure (Excel)</b><small>Engagement metadata + reconciliation</small></span><ArrowUpRight size={16} /></button>
            {perspective === "receivables"
              ? <button className="notice-launch" onClick={() => openDemand()}><MessageCircle size={16} /> Statutory Balance Confirmation &amp; Sec 15 Demand Statement <ChevronRight size={15} /></button>
              : <div className="notice-launch internal-schedule"><ShieldCheck size={16} /> Internal Disallowance Schedule for Management Representation</div>}
          </div>
          <div className="method-note"><Filter size={14} /><span>Interest calculated through <b>{formatDate(cutoff)}</b>, using the statutory annual rate and actual delay days.</span></div>
        </aside>
      </section>
      <footer className="page-footer"><span>MSME <b>AuditFlow</b> <span className="footer-separator">·</span> A private workspace for better compliance</span><span><ShieldCheck size={13} /> CLIENT-SIDE ONLY <span className="footer-separator">·</span> NOT LEGAL ADVICE</span></footer>

      {mappingOpen && <Modal title="Map your ledger columns" subtitle={`Review detected headers from ${fileName} before importing.`} onClose={() => setMappingOpen(false)}>
        <div className="mapping-intro"><span className="mapping-icon"><FileSpreadsheet size={17} /></span><span>We matched familiar column names automatically. A positive agreed credit period is treated as a written agreement if no agreement column is supplied; verify this against the executed contract.</span></div>
        <div className="mapping-list">{ledgerFields.map((field) => <label key={field.key} className="mapping-row"><span>{field.key === "buyer" ? (perspective === "payables" ? "Sundry creditor / supplier name" : "Sundry debtor / trade receivable") : field.label}{field.required && <i>Required</i>}</span><select value={mapping[field.key] < 0 ? "" : mapping[field.key]} onChange={(event) => setMap(field.key, event.target.value)}><option value="">— Skip this column —</option>{headers.map((header, index) => <option key={index} value={index}>{String(header || `Column ${index + 1}`)}</option>)}</select></label>)}</div>
        <div className="modal-actions"><button className="secondary-button" onClick={() => setMappingOpen(false)}>Cancel</button><button className="primary-button" onClick={importMappedRows}><Upload size={15} /> Import ledger</button></div>
      </Modal>}

      {manualOpen && <Modal title="Add a trade receivable" subtitle="Enter outstanding voucher details for statutory verification." onClose={() => setManualOpen(false)}>
        <form className="invoice-form" onSubmit={saveManual}>
          <label>{perspective === "payables" ? "Sundry creditor / supplier" : "Sundry debtor / trade receivable"}<input autoFocus value={draft.buyer} onChange={(event) => setDraft({ ...draft, buyer: event.target.value })} placeholder={perspective === "payables" ? "Supplier legal name" : "Business / debtor legal name"} /></label>
          <div className="form-row"><label>Voucher reference / invoice number<input value={draft.number} onChange={(event) => setDraft({ ...draft, number: event.target.value })} placeholder="INV-001" /></label><label>GSTIN / UIN <span className="optional">Optional</span><input value={draft.gstin} onChange={(event) => setDraft({ ...draft, gstin: event.target.value })} placeholder={perspective === "payables" ? "Supplier GSTIN" : "Debtor GSTIN"} /></label></div>
          {perspective === "payables" && <label>Supplier MSE classification<select value={draft.supplierMseType ?? ""} onChange={(event) => setDraft({ ...draft, supplierMseType: event.target.value ? event.target.value as "Micro" | "Small" | "Not MSE" : undefined })}><option value="">Use global default ({msmeType})</option><option value="Micro">Micro</option><option value="Small">Small</option><option value="Not MSE">Not an MSE</option></select></label>}
          <div className="form-row"><label>Invoice date<input type="date" value={draft.invoiceDate} onChange={(event) => setDraft({ ...draft, invoiceDate: event.target.value })} /></label><label>Invoice amount (₹)<input type="number" min="0.01" step="0.01" value={draft.amount || ""} onChange={(event) => setDraft({ ...draft, amount: Number(event.target.value) })} placeholder="0.00" /></label></div>
          <label className="agreement-check"><input type="checkbox" checked={draft.writtenAgreement} onChange={(event) => setDraft({ ...draft, writtenAgreement: event.target.checked, agreedDays: event.target.checked ? (draft.agreedDays ?? 45) : null })} /><span><b>Written payment agreement</b><small>Without one, statutory credit period is 15 days.</small></span></label>
          {draft.writtenAgreement && <label>Agreed credit days <span className="optional">Maximum statutory window: 45 days</span><input type="number" min="0" max="365" value={draft.agreedDays ?? 45} onChange={(event) => setDraft({ ...draft, agreedDays: Number(event.target.value) })} /></label>}
          <div className="modal-actions"><button type="button" className="secondary-button" onClick={() => setManualOpen(false)}>Cancel</button><button className="primary-button" type="submit"><Plus size={15} /> Add to ledger</button></div>
        </form>
      </Modal>}

      {noticeOpen && perspective === "receivables" && <Modal title="Statutory Balance Confirmation & Sec 15 Demand Statement" subtitle="Prepare a sundry-debtor balance confirmation using the current working paper." onClose={() => setNoticeOpen(false)} wide>
        <div className="notice-controls"><label>Sundry debtor / trade receivable<select value={noticeBuyer} onChange={(event) => setNoticeBuyer(event.target.value)}>{Array.from(new Set(calculated.filter((row) => row.daysDelayed > 0).map((row) => row.buyer))).map((buyer) => <option key={buyer}>{buyer}</option>)}</select></label><div className="tone-selector" aria-label="Notice tone">{(["A", "B", "C"] as const).map((item) => <button key={item} className={tone === item ? "active" : ""} onClick={() => setTone(item)}><b>Tone {item}</b><small>{item === "A" ? "Diplomatic" : item === "B" ? "Statutory warning" : "Formal demand"}</small></button>)}</div></div>
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
