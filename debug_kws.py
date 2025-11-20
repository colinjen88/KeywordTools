import csv

def load_keywords(path):
    kws = []
    try:
        with open(path, newline="", encoding="utf-8-sig") as fh:
            reader = csv.reader(fh)
            for row in reader:
                if not row:
                    continue
                parts = [p.strip() for p in row[0].split(',') if p.strip()]
                kws.extend(parts)
    except Exception as e:
        print(f"Error: {e}")
    return kws

kws = load_keywords('allKeyWord_normalized.csv')
print(f"Loaded {len(kws)} keywords.")
print(f"First 5: {kws[:5]}")
