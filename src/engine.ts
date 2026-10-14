import { addDays, differenceInCalendarDays, format, isValid, parseISO } from "date-fns";

export type Invoice = {
  id: string;
  buyer: string;
  gstin: string;
  number: string;
  invoiceDate: string;
  amount: number;
  writtenAgreement: boolean;
  agreedDays: number | null;
};

export type CalculatedInvoice = Invoice & {
  creditDays: number;
  dueDate: string;
  daysDelayed: number;
  daysUntilDue: number;
  interest: number;
  total: number;
  taxRisk: boolean;
};

export const todayISO = () => format(new Date(), "yyyy-MM-dd");

export function parseLedgerDate(value: unknown): string | null {
  if (value instanceof Date && isValid(value)) return format(value, "yyyy-MM-dd");
  if (typeof value === "number" && Number.isFinite(value)) {
    const date = new Date(Date.UTC(1899, 11, 30) + Math.floor(value) * 86_400_000);
    return isValid(date) ? date.toISOString().slice(0, 10) : null;
  }
  if (typeof value !== "string" || !value.trim()) return null;
  const text = value.trim();
  const iso = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(text);
  if (iso) {
    const parsed = new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
    return parsed.getFullYear() === Number(iso[1]) && parsed.getMonth() === Number(iso[2]) - 1 && parsed.getDate() === Number(iso[3])
      ? format(parsed, "yyyy-MM-dd")
      : null;
  }
  const dmy = /^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})$/.exec(text);
  if (dmy) {
    const year = Number(dmy[3]) < 100 ? 2000 + Number(dmy[3]) : Number(dmy[3]);
    const parsed = new Date(year, Number(dmy[2]) - 1, Number(dmy[1]));
    return parsed.getFullYear() === year && parsed.getMonth() === Number(dmy[2]) - 1 && parsed.getDate() === Number(dmy[1])
      ? format(parsed, "yyyy-MM-dd")
      : null;
  }
  const parsed = new Date(text);
  return isValid(parsed) ? format(parsed, "yyyy-MM-dd") : null;
}

export function calculateInvoice(
  invoice: Invoice,
  cutoff: string,
  rbiRate: number,
  isMse: boolean,
): CalculatedInvoice {
  const creditDays = invoice.writtenAgreement
    ? Math.min(Math.max(Math.trunc(invoice.agreedDays ?? 45), 0), 45)
    : 15;
  const dueDate = format(addDays(parseISO(invoice.invoiceDate), creditDays), "yyyy-MM-dd");
  const daysDelayed = Math.max(0, differenceInCalendarDays(parseISO(cutoff), parseISO(dueDate)));
  const daysUntilDue = Math.max(0, differenceInCalendarDays(parseISO(dueDate), parseISO(cutoff)));
  const monthlyRate = (3 * rbiRate / 100) / 12;
  const total = invoice.amount * Math.pow(1 + monthlyRate, daysDelayed / 30.4375);
  return {
    ...invoice,
    creditDays,
    dueDate,
    daysDelayed,
    daysUntilDue,
    interest: daysDelayed > 0 ? total - invoice.amount : 0,
    total: daysDelayed > 0 ? total : invoice.amount,
    taxRisk: isMse && daysDelayed > 0,
  };
}

export const formatINR = (value: number) =>
  new Intl.NumberFormat("en-IN", { style: "currency", currency: "INR", maximumFractionDigits: 2 }).format(value);

export const formatDate = (value: string) => format(parseISO(value), "dd MMM yyyy");

export const isBlankRow = (row: unknown[]) =>
  row.every((cell) => cell === null || cell === undefined || String(cell).trim() === "");
