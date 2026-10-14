# MSME AuditFlow

A client-side CA audit working paper utility for MSME statutory appointed-day verification, Section 16 penal interest computation, Section 43B(h) tax disallowance schedules and engagement-stamped audit annexures. TallyPrime Ledger Outstandings columns are auto-mapped, and the interface includes an editable engagement metadata card and statutory-reference basis.

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

Ledger files are parsed and processed in the browser; the app has no backend or financial-data upload. WhatsApp and email actions open external services only when selected by the user. Confirm supplier registration, payment status, dates, current law and tax treatment with a qualified professional before relying on estimates. The Section 43B(h) flags assume ledger invoices remain unpaid through the selected cutoff date.
