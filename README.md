# iExpense Compliance Automation Suite

A tool that automates a monthly finance compliance process end to end: it reconciles two
corporate-card data exports, classifies every transaction to decide who should be notified,
drafts the notices with the correct wording and recipients per aging bucket, and keeps a full
audit trail — turning a multi-day manual job into a ~15-minute run.

> **Note:** This is a sanitized, portfolio version. All names, NetIDs, emails, and
> transactions in the samples are fictional, and email **sending is disabled**.

---

## Try it in your browser (no install)

**▶ Live demo:** _enable GitHub Pages, then link it here_ — e.g. `https://ylolla.github.io/corporate-card-compliance-automation/`

Open the page, click **"load sample data"** (or drop your own two CSVs), and hit **Process files**.
You'll see the classification counts, a rendered email preview per cardholder, and the
flagged-for-review list — all running locally in your browser. Nothing is uploaded anywhere.

---

## What it does

1. **Reconciles** the "outstanding transactions" file against the "iExpense" export on `Trx Id`.
2. **Recovers missing NetIDs** — from a person's other transactions, and (in the full version)
   from a directory file — so every recipient resolves to a valid `netid@…` address.
3. **Classifies** each transaction. The newer iExpense status wins over the older snapshot.
   Only two things are held for review: `Pending System Administrator Action`, or *no NetID
   on file* (nothing to send to). Everything already resolved (Paid, Ready for Payment, etc.)
   is excluded automatically.
4. **Drafts notices** with per-bucket subject lines and wording, the right CC list
   (fiscal officers on 60+ day buckets, managers, and the process owner), and an HTML table
   of the transactions.
5. **Logs every run** to a timestamped audit workbook (emailed / excluded / flagged / all).

## The impact

A task previously spread across ~10 people with no consistent method, consolidated into one
repeatable run. Doing it *correctly* by hand is ~15 minutes per cardholder (status lookup,
recipient lookup, wording, drafting, logging) — roughly **1,250+ hours a year** at full volume.
The tool does it in minutes, and does it the same way every time.

---

## Repo contents

| Path | What it is |
|---|---|
| `index.html` | The browser demo — self-contained, mirrors the core classification + email logic |
| `generate_sample_data.py` | Regenerates the fictional sample CSVs |
| `sample_data/` | Pre-generated fictional inputs the demo loads |
| `app_demo.py` | The full Flask app (sanitized, **mock sender**) — run it locally for the complete experience |

## Run the full app locally

```bash
pip install flask pandas openpyxl
python app_demo.py
# open http://127.0.0.1:5000
```

The full app adds the review dashboard, audit-workbook export, and the multi-mode send screen —
but in this public version the sender is **mocked**: it logs what it *would* send and never
contacts anyone.

---

Built by **Yeshwanth Lolla** · Sr. Compliance and Data Analyst · Financial Systems & Services.
Sample data is fictional; this repository contains no institutional data.
