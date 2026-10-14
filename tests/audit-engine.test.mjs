import assert from "node:assert/strict";
import test from "node:test";
import {
  calculateInvoice,
  financialYearFor,
  parseLedgerDate,
} from "../src/engine.ts";
import {
  ledgerFields,
  mapLedgerRows,
  suggestedLedgerMapping,
} from "../src/ledger.ts";

const invoice = {
  id: "test-invoice",
  buyer: "Test Engineering Ltd",
  gstin: "07AAAAA0000A1Z5",
  number: "INV-001",
  invoiceDate: "2026-01-01",
  amount: 100_000,
  writtenAgreement: false,
  agreedDays: null,
};

test("parses ISO, Indian day-first, and Excel serial dates without accepting invalid dates", () => {
  assert.equal(parseLedgerDate("2026-01-31"), "2026-01-31");
  assert.equal(parseLedgerDate("31/01/2026"), "2026-01-31");
  assert.equal(parseLedgerDate(46023), "2026-01-01");
  assert.equal(parseLedgerDate("31/02/2026"), null);
  assert.equal(parseLedgerDate(""), null);
});

test("derives Indian financial and assessment years across the April boundary", () => {
  assert.deepEqual(financialYearFor("2026-03-31"), {
    financialYear: "2025–26",
    assessmentYear: "2026–27",
  });
  assert.deepEqual(financialYearFor("2026-04-01"), {
    financialYear: "2026–27",
    assessmentYear: "2027–28",
  });
});

test("uses 15 days without agreement and triggers no interest on the appointed day", () => {
  const due = calculateInvoice(invoice, "2026-01-16", 5.75, true);
  assert.equal(due.creditDays, 15);
  assert.equal(due.dueDate, "2026-01-16");
  assert.equal(due.daysDelayed, 0);
  assert.equal(due.interest, 0);
  assert.equal(due.taxRisk, false);
  assert.equal(due.disallowedAmount, 0);
});

test("caps written credit periods at 45 days and compounds interest after the appointed day", () => {
  const cappedInvoice = { ...invoice, writtenAgreement: true, agreedDays: 90 };
  const due = calculateInvoice(cappedInvoice, "2026-02-15", 5.75, true);
  assert.equal(due.creditDays, 45);
  assert.equal(due.dueDate, "2026-02-15");

  const overdue = calculateInvoice(cappedInvoice, "2026-03-17", 5.75, true);
  const expected = invoice.amount * (Math.pow(1 + 0.1725 / 12, 30 / 30.4375) - 1);
  assert.equal(overdue.daysDelayed, 30);
  assert.ok(Math.abs(overdue.interest - expected) < 1e-8);
  assert.equal(overdue.disallowedAmount, invoice.amount);
});

test("does not flag disallowance for suppliers outside the Micro/Small classification", () => {
  const result = calculateInvoice(invoice, "2026-03-17", 5.75, false);
  assert.ok(result.daysDelayed > 0);
  assert.equal(result.taxRisk, false);
  assert.equal(result.disallowedAmount, 0);
});

test("auto-maps TallyPrime outstandings headers and infers written agreement from agreed days", () => {
  const headers = [
    "Party Name",
    "GSTIN/UIN",
    "Voucher Ref/Inv No",
    "Bill Date",
    "Pending Amount",
    "Agreed Credit Period (Days)",
  ];
  const mapping = suggestedLedgerMapping(headers);
  assert.deepEqual(ledgerFields.map(({ key }) => mapping[key]), [0, 1, 2, 3, 4, -1, 5]);

  const { invoices, errors } = mapLedgerRows([
    ["Zenith Auto Components Ltd", "07AAAAA0000A1Z5", "INV-01", "2026-08-15", 250_000, 60],
    ["Pragati Machine Tools", "29CCCCC2222C1Z3", "INV-02", "04/09/2026", 143_750, ""],
    ["", "", "", "", "", ""],
  ], mapping);

  assert.equal(errors.length, 0);
  assert.equal(invoices.length, 2);
  assert.equal(invoices[0].writtenAgreement, true);
  assert.equal(invoices[0].agreedDays, 60);
  assert.equal(invoices[1].writtenAgreement, false);
  assert.equal(invoices[1].invoiceDate, "2026-09-04");
});
