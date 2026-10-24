"""Generate a Superstore-style sample dataset (same columns as the Kaggle Superstore file).
Replace data/superstore_sample.csv with the real Superstore export to run on real data."""
import csv, random
from datetime import date, timedelta

random.seed(15)

catalog = {
    "Office Supplies": {
        "Binders": ["Avery Binder", "Cardinal Slant-D Binder", "GBC Binding Covers", "Ibico Binding System", "Wilson Jones Binder"],
        "Paper": ["Xerox 1881", "Easy-staple Paper", "Eaton Copy Paper", "Hammermill Paper", "Ampad Notebook"],
        "Storage": ["Fellowes Bankers Box", "Rogers File Case", "Akro-Mils Bin", "Safco Shelving", "Eldon File Cart"],
        "Art": ["Crayola Markers", "Boston Pencil Sharpener", "Newell Pens", "Sanford Highlighters", "Staples Scissors"],
        "Labels": ["Avery Address Labels", "Dymo Labels", "Hon Name Tags", "Smead Labels"],
        "Envelopes": ["Staples Envelopes", "Jiffy Mailers", "Poly String Envelopes", "Ames Mailers"],
        "Fasteners": ["Staples Paper Clips", "Acco Fasteners", "Advantus Push Pins", "OIC Binder Clips"],
        "Supplies": ["Fiskars Shears", "Acme Scissors", "Elmer's Glue", "Stanley Tape"],
        "Appliances": ["Holmes Fan", "Hoover Vacuum", "Kensington Surge Strip", "Belkin Power Strip"],
    },
    "Furniture": {
        "Chairs": ["Hon Executive Chair", "Global Task Chair", "Office Star Chair", "Novimex Stool"],
        "Tables": ["Bush Conference Table", "Lesro Round Table", "Chromcraft Desk", "Hon Folding Table"],
        "Bookcases": ["Sauder Bookcase", "O'Sullivan Shelf", "Bush Corner Unit", "Atlantic Bookcase"],
        "Furnishings": ["Eldon Desk Lamp", "Deflect-o Mat", "Howard Clock", "Artistic Frame"],
    },
    "Technology": {
        "Phones": ["Samsung Smart Phone", "Cisco IP Phone", "Motorola Headset", "Nokia Handset"],
        "Accessories": ["Logitech Mouse", "Kingston USB Drive", "Belkin Keyboard", "HP Webcam"],
        "Machines": ["Brother Fax Machine", "Canon Copier", "Epson Label Printer", "Hewlett Packard Printer"],
        "Copiers": ["Canon imageCLASS", "Sharp Copy Machine", "Xerox Workcentre"],
    },
}
weights = {"Office Supplies": 60, "Furniture": 21, "Technology": 19}
prefix = {"Office Supplies": "OFF", "Furniture": "FUR", "Technology": "TEC"}

products = []  # (category, sub, pid, name)
n = 1000
for cat, subs in catalog.items():
    for sub, names in subs.items():
        for nm in names:
            for variant in range(1, 4):
                products.append((cat, sub, f"{prefix[cat]}-{sub[:2].upper()}-{n}", f"{nm} {variant}"))
                n += 7
by_cat = {c: [p for p in products if p[0] == c] for c in catalog}

regions = {"West": ["Los Angeles", "Seattle", "San Francisco", "Portland"],
           "East": ["New York City", "Philadelphia", "Boston", "Newark"],
           "Central": ["Chicago", "Houston", "Dallas", "Detroit"],
           "South": ["Atlanta", "Miami", "Charlotte", "Nashville"]}
segments = ["Consumer", "Corporate", "Home Office"]
ship = ["Standard Class", "Second Class", "First Class", "Same Day"]
first = ["Aarav", "Maya", "John", "Priya", "Liam", "Sofia", "Noah", "Emma", "Ravi", "Olivia", "Ethan", "Ava"]
last = ["Shah", "Patel", "Smith", "Brown", "Lee", "Khan", "Davis", "Clark", "Mehta", "Young"]
customers = [(f"CU-{1000+i}", f"{random.choice(first)} {random.choice(last)}", random.choice(segments)) for i in range(120)]

rows = []
order_no = 0
rid = 1
start = date(2025, 1, 1)
while rid <= 1500:
    order_no += 1
    cust = random.choice(customers)
    region = random.choice(list(regions))
    city = random.choice(regions[region])
    od = start + timedelta(days=random.randint(0, 720))
    sd = od + timedelta(days=random.randint(0, 7))
    for _ in range(random.choice([1, 1, 2, 2, 3, 4])):
        if rid > 1500:
            break
        cat = random.choices(list(weights), weights=list(weights.values()))[0]
        c, sub, pid, name = random.choice(by_cat[cat])
        qty = random.randint(1, 8)
        price = {"Office Supplies": random.uniform(4, 90), "Furniture": random.uniform(40, 600), "Technology": random.uniform(25, 900)}[c]
        disc = random.choice([0, 0, 0, 0.1, 0.2, 0.3])
        sales = round(price * qty * (1 - disc), 2)
        profit = round(sales * random.uniform(-0.12, 0.35), 2)
        rows.append([rid, f"CA-{od.year}-{100000+order_no}", od.isoformat(), sd.isoformat(), random.choice(ship),
                     cust[0], cust[1], cust[2], "United States", city, region, pid, c, sub, name, sales, qty, disc, profit])
        rid += 1

hdr = ["Row ID", "Order ID", "Order Date", "Ship Date", "Ship Mode", "Customer ID", "Customer Name", "Segment",
       "Country", "City", "Region", "Product ID", "Category", "Sub-Category", "Product Name", "Sales", "Quantity", "Discount", "Profit"]
with open("data/superstore_sample.csv", "w", newline="") as f:
    w = csv.writer(f); w.writerow(hdr); w.writerows(rows)
print(len(rows), "rows,", len({r[11] for r in rows}), "unique products")
