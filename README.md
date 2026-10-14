# MSME AuditFlow

A client-side MSME receivables audit workspace for statutory due dates, Section 16 interest estimates, Section 43B(h) flags and audit exports.

## Run locally

```sh
npm ci
npm run dev
```

## GitHub Pages

The production build is configured for the `kavivsoftwarelabs/msme-auditflow` project site. Pushes to `main` build and deploy the static app through the GitHub Pages workflow. In the repository settings, set **Pages → Build and deployment → Source** to **GitHub Actions**.

## Privacy and statutory review

Ledger files are parsed and processed in the browser; the app has no backend or financial-data upload. WhatsApp and email actions open external services only when selected by the user. Confirm supplier registration, payment status, dates, current law and tax treatment with a qualified professional before relying on estimates. The Section 43B(h) flags assume ledger invoices remain unpaid through the selected cutoff date.
