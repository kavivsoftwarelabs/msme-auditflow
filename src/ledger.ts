import { isBlankRow, parseLedgerDate } from "./engine.ts";
import type { Invoice } from "./engine.ts";

export type MappingKey = "buyer" | "gstin" | "number" | "invoiceDate" | "amount" | "writtenAgreement" | "agreedDays" | "supplierMseType";
export type LedgerMapping = Record<MappingKey, number>;

export const ledgerFields: { key: MappingKey; label: string; aliases: string[]; required: boolean }[] = [
  { key: "buyer", label: "Sundry debtor / trade receivable", aliases: ["buyername", "partyname", "customername", "debtor", "buyer", "party", "customer", "sundrydebtors", "tradereceivable", "suppliername", "vendorname", "creditor", "supplier", "vendor", "sundrycreditors", "tradepayable"], required: true },
  { key: "gstin", label: "GSTIN / UIN", aliases: ["gstin", "gstnumber", "gstno", "gstinuin"], required: false },
  { key: "number", label: "Voucher reference / invoice number", aliases: ["invoicenumber", "invoiceno", "billnumber", "billno", "voucher", "number", "voucherrefinvno", "voucherref", "invno"], required: true },
  { key: "invoiceDate", label: "Invoice date", aliases: ["invoicedate", "billdate", "date"], required: true },
  { key: "amount", label: "Pending amount", aliases: ["invoiceamount", "amount", "total", "billamount", "netamount", "value", "pendingamount"], required: true },
  { key: "writtenAgreement", label: "Written agreement", aliases: ["writtenagreement", "agreement", "agreementyesno", "haswrittenagreement"], required: false },
  { key: "agreedDays", label: "Agreed credit period (days)", aliases: ["agreeddays", "creditdays", "paymentdays", "termsdays", "agreedcreditperioddays"], required: false },
  { key: "supplierMseType", label: "Supplier MSE classification", aliases: ["suppliermsetype", "suppliermseclassification", "msmeclassification", "msmetype", "supplierclassification"], required: false },
];

export const emptyMapping = (): LedgerMapping => ({
  buyer: -1,
  gstin: -1,
  number: -1,
  invoiceDate: -1,
  amount: -1,
  writtenAgreement: -1,
  agreedDays: -1,
  supplierMseType: -1,
});

function normalizeHeader(value: unknown) {
  return String(value ?? "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

export function suggestedLedgerMapping(headers: unknown[]): LedgerMapping {
  const mapping = emptyMapping();
  ledgerFields.forEach(({ key, aliases }) => {
    const index = headers.findIndex((header) => aliases.includes(normalizeHeader(header)));
    if (index >= 0) mapping[key] = index;
  });
  return mapping;
}

export function mapLedgerRows(rows: unknown[][], mapping: LedgerMapping, counterpartyLabel = "Sundry debtor / trade receivable") {
  const required = ledgerFields.filter((field) => field.required && mapping[field.key] < 0);
  if (required.length) {
    return {
      invoices: [] as Invoice[],
      errors: [`Map the required columns: ${required.map((field) => field.key === "buyer" ? counterpartyLabel : field.label).join(", ")}.`],
    };
  }
  const invoices: Invoice[] = [];
  const errors: string[] = [];
  rows.forEach((row, index) => {
    if (isBlankRow(row)) return;
    const get = (key: MappingKey) => mapping[key] >= 0 ? row[mapping[key]] : "";
    const buyer = String(get("buyer") ?? "").trim();
    const number = String(get("number") ?? "").trim();
    const invoiceDate = parseLedgerDate(get("invoiceDate"));
    const rawAmount = get("amount");
    const amount = typeof rawAmount === "number" ? rawAmount : Number(String(rawAmount).replace(/[₹,\s]/g, ""));
    if (!buyer || !number || !invoiceDate || !Number.isFinite(amount) || amount <= 0) {
      errors.push(`Row ${index + 2}: ${!buyer ? counterpartyLabel : !number ? "voucher / invoice number" : !invoiceDate ? "valid bill date" : "positive pending amount"} missing or invalid`);
      return;
    }
    const agreement = String(get("writtenAgreement") ?? "").trim().toLowerCase();
    const rawDays = String(get("agreedDays") ?? "").trim().match(/\d+(?:\.\d+)?/)?.[0];
    const agreedDays = rawDays ? Number(rawDays) : null;
    const explicitAgreement = ["yes", "true", "y", "1"].includes(agreement);
    const writtenAgreement = mapping.writtenAgreement >= 0
      ? explicitAgreement
      : agreedDays !== null && agreedDays > 0;
    const rawMseType = String(get("supplierMseType") ?? "").trim().toLowerCase();
    const supplierMseType = rawMseType === "micro" ? "Micro"
      : rawMseType === "small" ? "Small"
        : ["not mse", "non-mse", "non mse", "no"].includes(rawMseType) ? "Not MSE"
          : undefined;
    invoices.push({
      id: crypto.randomUUID(),
      buyer,
      number,
      invoiceDate,
      amount,
      gstin: String(get("gstin") ?? "").trim(),
      writtenAgreement,
      agreedDays: Number.isFinite(agreedDays) ? agreedDays : null,
      supplierMseType,
    });
  });
  return { invoices, errors };
}
