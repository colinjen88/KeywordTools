#!/usr/bin/env python3
"""
簡單工具：將單行、逗號分隔或每行一個關鍵字的 allKeyWord.csv 轉成每行一個關鍵字的 CSV

輸出檔案：allKeyWord_normalized.csv（每行一個關鍵字，UTF-8-SIG）
"""
import csv
import sys
import os


def normalize(in_path, out_path):
    kws = []
    with open(in_path, newline="", encoding="utf-8-sig") as fh:
        reader = csv.reader(fh)
        rows = list(reader)
        if not rows:
            print("輸入檔案為空")
            return
        # 情形 A: 單列但 csv.reader 已把逗號分割成多個欄位
        if len(rows) == 1 and len(rows[0]) > 1:
            kws = [p.strip() for p in rows[0] if p.strip()]
        # 情形 B: 單列且整個欄位是用逗號串接的長字串
        elif len(rows) == 1 and "," in rows[0][0]:
            parts = [p.strip() for p in rows[0][0].split(",") if p.strip()]
            kws = parts
        else:
            for row in rows:
                if not row:
                    continue
                kws.append(row[0].strip())

    # 寫出每行一個關鍵字
    with open(out_path, "w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.writer(fh)
        for k in kws:
            writer.writerow([k])

    print(f"已寫出 {len(kws)} 個關鍵字到 {out_path}")


if __name__ == "__main__":
    in_path = "allKeyWord.csv"
    out_path = "allKeyWord_normalized.csv"
    if len(sys.argv) >= 2:
        in_path = sys.argv[1]
    if len(sys.argv) >= 3:
        out_path = sys.argv[2]
    if not os.path.exists(in_path):
        print(f"找不到輸入檔：{in_path}")
        sys.exit(2)
    normalize(in_path, out_path)
