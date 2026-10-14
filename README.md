# MSME AuditFlow

A client-side CA audit working paper utility with separate, locally retained Trade Payables / Creditors and Trade Receivables / Debtors workflows. The payables view calculates Section 15 appointed days, Section 16 interest payable, and qualifying Section 43B(h) disallowances for registered Micro/Small suppliers; the receivables view calculates interest receivable and provides statutory demand statements without attributing buyer-side tax disallowances to receivables. TallyPrime-style supplier and debtor ledgers are auto-mapped, including optional supplier MSE classifications.

## Run locally

```sh
npm ci
npm run dev
```

## Regression tests

```sh
npm test
npm run build
```

## GitHub Pages

The production build is configured for the `kavivsoftwarelabs/msme-auditflow` project site. Pushes to `main` build and deploy the static app through the GitHub Pages workflow. In the repository settings, set **Pages → Build and deployment → Source** to **GitHub Actions**.

## Privacy and statutory review

Ledger files are parsed and processed in the browser; the app has no backend or financial-data upload. Payables and receivables ledgers remain separate when switching perspectives. WhatsApp and email actions are available only in the receivables view and open external services only when selected by the user. Confirm supplier registration, payment status, dates, current law and tax treatment with a qualified professional before relying on estimates. The Section 43B(h) flags assume qualifying supplier invoices remain unpaid through the selected cutoff date.
