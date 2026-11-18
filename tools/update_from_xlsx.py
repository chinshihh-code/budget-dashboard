"""從「經費登錄總表」xlsx 的「總表」工作表讀出 24 個項次的核定／主計／線上登帳，
與 index.html 內的 ITEMS 比對。

用法：
  python3 -I tools/update_from_xlsx.py <xlsx> <index.html>            # 只列出變動（不改檔）
  python3 -I tools/update_from_xlsx.py <xlsx> <index.html> --apply    # 寫回 index.html（含 AS_OF、SOURCE_NOTE）
檔名需以 _MMDD.xlsx 結尾（例如 _1124.xlsx），用來設定「資料截至」日期。
"""
import sys, re, json, os
import openpyxl

xlsx, html = sys.argv[1], sys.argv[2]
apply = '--apply' in sys.argv

m = re.search(r'_(\d{2})(\d{2})\.xlsx$', os.path.basename(xlsx))
if not m:
    sys.exit('檔名需以 _MMDD.xlsx 結尾')
as_of = f'2026-{m.group(1)}-{m.group(2)}'

wb = openpyxl.load_workbook(xlsx, read_only=True, data_only=True)
ws = wb['總表']
rows = list(ws.iter_rows(min_row=1, max_row=40, max_col=17, values_only=True))

def num(v):
    if v is None or v == '':
        return None
    return int(round(float(v)))

found = {}
section = None
for r in rows[1:]:
    head = (str(r[0]).strip() if r[0] else '')
    if head.startswith('人事費'):
        section = '人事費'
    elif head.startswith('業務費1'):
        section = '業務費1'
    elif head.startswith('業務費2'):
        section = '業務費2'
    elif head.startswith('業務費') and '小計' in head:
        section = None
    seq = (str(r[3]).strip().rstrip('、') if r[3] else '')
    if not section or not seq or r[11] is None:
        continue
    if section == '人事費':
        code = '人事費' + seq
    elif section == '業務費1':
        code = '業務費' + seq
    else:
        code = '業務費二-' + seq
    found[code] = {'單位': str(r[2]).strip(), '核定': num(r[11]), '主計': num(r[12]) or 0, '線上': num(r[14])}

t = open(html, encoding='utf-8').read()
s = t.index('const ITEMS = ') + len('const ITEMS = ')
e = t.index(';\n', s)
items = json.loads(t[s:e])

codes = [it['項次'] for it in items]
missing = [c for c in codes if c not in found]
extra = [c for c in found if c not in codes]
if missing or extra:
    sys.exit(f'項次對不上：缺少 {missing}，多出 {extra}。請人工確認總表格式是否改變。')

changes = []
for it in items:
    f = found[it['項次']]
    if f['單位'] != it['單位']:
        sys.exit(f"{it['項次']} 的單位不同：檔案 {f['單位']}，網頁 {it['單位']}。請人工確認。")
    for k in ('核定', '主計', '線上'):
        if f[k] != it[k]:
            changes.append((it['項次'], it['項目'], k, it[k], f[k]))
            it[k] = f[k]

old_asof = re.search(r"const AS_OF = '([\d-]+)';", t).group(1)
print(json.dumps({'as_of_old': old_asof, 'as_of_new': as_of, 'changes': changes,
                  'totals': {c: sum(x['主計'] for x in items if x['費用別'] == c) for c in ('人事費', '業務費')}},
                 ensure_ascii=False, indent=1))

if apply:
    t = t[:s] + json.dumps(items, ensure_ascii=False) + t[e:]
    t = re.sub(r"const AS_OF = '[\d-]+';", f"const AS_OF = '{as_of}';", t)
    note = f"核銷會計對帳總檔（{os.path.basename(xlsx)}）的「總表」工作表。"
    t = re.sub(r"const SOURCE_NOTE = '.*?';", "const SOURCE_NOTE = " + json.dumps(note, ensure_ascii=False).replace('"', "'") + ";", t)
    open(html, 'w', encoding='utf-8').write(t)
    print('已寫回', html)
