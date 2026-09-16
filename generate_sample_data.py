"""
Generate fictional sample data for the iExpense Compliance Automation demo.

Produces two CSVs that mirror the real inputs' structure — with entirely made-up
names, NetIDs, and transactions. Nothing here is real institutional data.

    python generate_sample_data.py

Writes:
    sample_data/outstanding_sample.csv   (the outstanding corporate-card file)
    sample_data/iexpense_sample.csv      (the "iExpense" export)
"""
import csv, os, random

random.seed(42)
OUT = os.path.join(os.path.dirname(__file__), "sample_data")
os.makedirs(OUT, exist_ok=True)

FIRST = ["Alex","Jordan","Sam","Taylor","Morgan","Casey","Riley","Jamie","Drew","Quinn",
         "Avery","Parker","Reese","Skyler","Devon","Harper","Rowan","Emerson","Blake","Sage"]
LAST = ["Rivera","Chen","Okafor","Nguyen","Patel","Silva","Kim","Harris","Novak","Flores",
        "Bauer","Mendez","Osei","Larsson","Yamamoto","Costa","Ahmed","Weber","Ivanov","Reyes"]
MERCHANTS = ["SQ *CORNER CAFE","AMAZON MKTPLACE","STAPLES 00123","DELTA AIR 006","UBER TRIP",
             "MARRIOTT HOTELS","OFFICE DEPOT","ZOOM.US","GITHUB INC","APPLE STORE R042"]
BUCKETS = ["< 30 Days","30 - 60 Days","60 - 90 Days","90 - 120 Days","120 - 150 Days","> 150 Days"]
STATUSES = ["Pending Manager Approval","Pending Your Resolution","Rejected","Withdrawn",
            "Paid","Ready for Payment","Pending System Administrator Action",""]

def netid(i): return f"D{1000+i}X"
def name(i): return f"{LAST[i % len(LAST)]}, {FIRST[i % len(FIRST)]} {chr(65 + (i % 26))}"

people = [(name(i), netid(i)) for i in range(40)]

# ---- iExpense export: has NetIDs (some blank / FAKE to exercise recovery) ----
iexp_rows, tid = [], 900000
for i,(nm,nid) in enumerate(people):
    for _ in range(random.randint(1,4)):
        tid += 1
        # 1 in 6 rows get a blank or FAKE netid, to show recovery working
        shown = nid
        roll = random.random()
        if roll < 0.10: shown = ""
        elif roll < 0.16: shown = "FAKE000002"
        iexp_rows.append({
            "Trx Id": tid,
            "Report Header Id": 2400000 + tid,
            "Employee Name": nm,
            "Employee NetID": shown,
            "Expense Status": random.choice(STATUSES),
            "Merchant Name": random.choice(MERCHANTS),
        })

# ---- Outstanding file: master list, joins on Trx Id ----
# Use most of iExpense's Trx Ids, plus a few orphans with no iExpense match.
tawnia_rows = []
sample = random.sample(iexp_rows, int(len(iexp_rows) * 0.85))
for r in sample:
    tawnia_rows.append({
        "Trx Id": r["Trx Id"],
        "Report Header Id": r["Report Header Id"],
        "Card Holder Name": r["Employee Name"],
        "Bucket": random.choice(BUCKETS),
        "Transaction Date": f"{random.randint(1,12)}/{random.randint(1,28)}/2026",
        "Billed Date": f"{random.randint(1,12)}/{random.randint(1,28)}/2026",
        "Billed Amount": round(random.uniform(12, 1800), 2),
        "Expense Status": random.choice(STATUSES),
        "Merchant Name": r["Merchant Name"],
        "GL Org": f"{random.randint(1000,9999)}",
        "GL Org L1": random.choice(["TDF-THAYER","GSM-GEISEL","TUCK-BUSINESS","LIBRARY","SAS-SCHOOL OF ARTS AND SCIENCES"]),
    })
# a few orphans (in outstanding, not in iExpense) -> exercise "no email found"
for j in range(4):
    tid += 1
    tawnia_rows.append({
        "Trx Id": tid, "Report Header Id": 2400000 + tid,
        "Card Holder Name": f"{LAST[(j+5)%len(LAST)]}, Ghost {chr(70+j)}",
        "Bucket": random.choice(BUCKETS),
        "Transaction Date": "3/3/2026", "Billed Date": "3/4/2026",
        "Billed Amount": round(random.uniform(20, 500), 2),
        "Expense Status": "Rejected", "Merchant Name": random.choice(MERCHANTS),
        "GL Org": f"{random.randint(1000,9999)}",
        "GL Org L1": "TDF-THAYER",
    })

def write(path, rows, cols):
    with open(path, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=cols); w.writeheader()
        for r in rows: w.writerow(r)

write(os.path.join(OUT, "iexpense_sample.csv"), iexp_rows,
      ["Trx Id","Report Header Id","Employee Name","Employee NetID","Expense Status","Merchant Name"])
write(os.path.join(OUT, "outstanding_sample.csv"), tawnia_rows,
      ["Trx Id","Report Header Id","Card Holder Name","Bucket","Transaction Date","Billed Date",
       "Billed Amount","Expense Status","Merchant Name","GL Org","GL Org L1"])

print(f"Wrote {len(tawnia_rows)} outstanding rows and {len(iexp_rows)} iExpense rows to {OUT}/")
