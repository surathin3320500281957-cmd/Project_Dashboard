"""
============================================================================
rebuild_dashboard.py — สคริปต์เดียวจบ: ดึงข้อมูล + เข้ารหัส + ฝังเข้าแดชบอร์ด
============================================================================
รวม 6 สคริปต์เดิม (extract_all_v2 / extract_active_v2 / extract_vacant_v2 /
build_comparison2 / extract_survey / encode_all_v2) เป็นไฟล์เดียว พร้อมแก้จุด
ที่เคยต้องทำมือแยก (CH64 5-กลุ่ม, SV64 รูปภาพ) ให้เป็นอัตโนมัติในตัว

วิธีใช้:
  python3 rebuild_dashboard.py

Input ที่ต้องมีก่อนรัน:
  - /mnt/user-data/uploads/DBหน่วยความรับผิดชอบ.xlsx  (ไฟล์ Excel ต้นทาง)
  - index.html ในโฟลเดอร์เดียวกับสคริปต์นี้ (ไฟล์แดชบอร์ดที่จะแทนที่ข้อมูลให้)

Output:
  - แทนที่ค่า DB64/PB64/CS64/PF64/VC64/CH64/SV64 ใน index.html ให้อัตโนมัติ
  - พิมพ์ผลตรวจสอบทุกขั้นตอนออกมาให้ไล่เช็คได้ (sum ตัวเลข, สมการอนุรักษ์ ฯลฯ)

⚠️ ก่อนรันทุกครั้ง: ตรวจสูตร COUNTIFS ในชีต "สรุปหน่วยเต็ม" ก่อนว่าเปลี่ยนไหม
   ถ้าเปลี่ยน ต้องแก้เงื่อนไขในส่วน STEP 2/3 ของสคริปต์นี้ให้ตรงก่อนรัน
============================================================================
"""
import openpyxl, json, gzip, base64, re, os
from collections import Counter

SRC = '/mnt/user-data/uploads/DBหน_วยความร_บผ_ดชอบ.xlsx'
INDEX_HTML = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'index.html')

def compress(payload):
    raw = json.dumps(payload, ensure_ascii=False, separators=(',', ':')).encode('utf-8')
    c = gzip.compress(raw, compresslevel=9)
    return base64.b64encode(c).decode('ascii')

print('='*70); print('เปิดไฟล์:', SRC); print('='*70)
wb = openpyxl.load_workbook(SRC, data_only=True)

# ============================================================================
# STEP 1: สรุปหน่วยเต็ม (project-level) → DB64
# ============================================================================
print('\n[1/7] สรุปหน่วยเต็ม (ระดับโครงการ)')
ws = wb['สรุปหน่วยเต็ม']
FIELD_MAP = [
    ('no',1),('dept',2),('div',3),('office_code',4),('office_name',5),('proj_type',6),
    ('income_model',7),('building_char',8),('building_form',9),('proj_code',10),('proj_name',11),
    ('address',12),('zone',13),('province',14),('income_level',15),('proj_group',16),
    ('community_care',17),('utility_transfer',18),('building_count',19),('juristic_count',20),
    ('units_total',21),('units_transferred',22),('units_rent',23),('units_lease_whole',24),
    ('units_hire_purchase',25),('units_pending_sale',26),('units_guarantee',27),
    ('units_pending_transfer',28),('staff_house',29),('vacant_ready_rent',30),
    ('vacant_notready_rent',31),('vacant_ready_sale',32),('vacant_notready_sale',33),
    ('check',34)
]
TEXT_FIELDS = {'no','dept','div','office_code','office_name','proj_type','income_model',
               'building_char','building_form','proj_code','proj_name','address','zone',
               'province','income_level','proj_group','community_care','utility_transfer'}
raw_rows = []
for row in range(2, ws.max_row+1):
    if ws.cell(row=row, column=10).value is None: continue
    rec = {}
    for name, col in FIELD_MAP:
        v = ws.cell(row=row, column=col).value
        rec[name] = v if v is not None else ('' if name in TEXT_FIELDS else 0)
    rec['vacant_total'] = rec['vacant_ready_rent']+rec['vacant_notready_rent']+rec['vacant_ready_sale']+rec['vacant_notready_sale']
    rec['active_total'] = rec['units_rent']+rec['units_lease_whole']+rec['units_hire_purchase']+rec['units_pending_sale']+rec['units_guarantee']+rec['units_pending_transfer']
    rec['vacant_ready_total'] = rec['vacant_ready_rent']+rec['vacant_ready_sale']
    rec['vacant_notready_total'] = rec['vacant_notready_rent']+rec['vacant_notready_sale']
    raw_rows.append(rec)

print('  Projects:', len(raw_rows))
print('  Sum active_total:', sum(r['active_total'] for r in raw_rows))
print('  Sum vacant_total:', sum(r['vacant_total'] for r in raw_rows))
print('  Sum units_hire_purchase:', sum(r['units_hire_purchase'] for r in raw_rows))
print('  Sum units_rent:', sum(r['units_rent'] for r in raw_rows))
print('  Sum units_lease_whole:', sum(r['units_lease_whole'] for r in raw_rows))
print('  Sum units_pending_sale:', sum(r['units_pending_sale'] for r in raw_rows))
print('  Sum units_guarantee:', sum(r['units_guarantee'] for r in raw_rows))
print('  Sum units_pending_transfer:', sum(r['units_pending_transfer'] for r in raw_rows))
print('  Distinct div:', sorted(set(r['div'] for r in raw_rows)))
print('  Distinct office:', len(set(r['office_name'] for r in raw_rows)))

valid_proj = set(r['proj_code'] for r in raw_rows)  # ← ขอบเขต 185 โครงการ ใช้กรองทุกจุดต่อจากนี้

# ============================================================================
# STEP 2: หน่วย Active — มีสัญญา / เช่าเหมา / เช่าซื้อรอโอน → PB64 บางส่วน, CS64, PF64
# ============================================================================
print('\n[2/7] หน่วย Active (มีสัญญา/เช่าเหมา/เช่าซื้อรอโอน)')
active_records = []
extra_cs = []
orphans = Counter()

# --- มีสัญญา: D=PROJ(4), L=MAT(12), N=HOUSE_NO_ACT(14), O=CONTRACT_TYPE_CODE(15), Y=CUST_STATUS_NAME(25)
ws = wb['มีสัญญา']
cnt_by_status = Counter()
for row in range(2, ws.max_row+1):
    proj = ws.cell(row=row, column=4).value
    if proj is None: continue
    proj = str(proj).strip()
    o = ws.cell(row=row, column=15).value
    o = str(o).strip() if o is not None else None
    y = ws.cell(row=row, column=25).value
    mat = ws.cell(row=row, column=12).value
    house = ws.cell(row=row, column=14).value
    status = None
    if o == '50': status = 'rent'
    elif o in ('20','80'): status = 'hire_purchase'
    elif o == '10': status = 'pending_sale'  # ⚠️ ไม่มีเงื่อนไข Y="กำลังผ่อนชำระ" แล้ว (สูตรปัจจุบันนับทุกหน่วย O="10")
    if status is None: continue
    if proj not in valid_proj:
        orphans[proj] += 1
        continue
    active_records.append({'proj_code':proj,'source':'C','mat':mat,'house_no':house,'status':status,'ref_status':y or ''})
    extra_cs.append([mat, y or '', o])
    cnt_by_status[status] += 1
print('  มีสัญญา matched:', dict(cnt_by_status))

# --- เช่าเหมา: E=PROJ_CODE(5), L=MATERIAL_NUMBER(12), N=HOUSE_NO_ACT(14)
ws = wb['เช่าเหมา']
cnt2 = 0
for row in range(2, ws.max_row+1):
    proj = ws.cell(row=row, column=5).value
    if proj is None: continue
    proj = str(proj).strip()
    mat = ws.cell(row=row, column=12).value
    house = ws.cell(row=row, column=14).value
    if proj not in valid_proj:
        orphans[proj] += 1
        continue
    active_records.append({'proj_code':proj,'source':'L','mat':mat,'house_no':house,'status':'lease_whole','ref_status':''})
    cnt2 += 1
print('  เช่าเหมา matched:', cnt2)

# --- เช่าซื้อรอโอน: F=PROJ_CODE(6), H=CL_MATERIAL_NO(8), L=HOUSE_NO_ACT(12), V=CONTRACT_STATUS(22), AA=PAY_FR(27)
ws = wb['เช่าซื้อรอโอน']
cnt3 = Counter()
extra_pf = []
for row in range(2, ws.max_row+1):
    proj = ws.cell(row=row, column=6).value
    if proj is None: continue
    proj = str(proj).strip()
    v = ws.cell(row=row, column=22).value
    mat = ws.cell(row=row, column=8).value
    house = ws.cell(row=row, column=12).value
    payfr = ws.cell(row=row, column=27).value
    status = None
    if v and 'ค้ำประกัน' in str(v): status = 'guarantee'
    elif v and 'รอโอน' in str(v): status = 'pending_transfer'
    if status is None: continue
    if proj not in valid_proj:
        orphans[proj] += 1
        continue
    active_records.append({'proj_code':proj,'source':'P','mat':mat,'house_no':house,'status':status,'ref_status':v or ''})
    extra_pf.append([mat, payfr or ''])
    cnt3[status] += 1
print('  เช่าซื้อรอโอน matched:', dict(cnt3))

status_sum = Counter(r['status'] for r in active_records)
print('  TOTAL active records:', len(active_records), '| ต้องตรงกับ Sum active_total ด้านบน')
print('  Orphan proj_codes (ไม่อยู่ใน 185 โครงการ ถูกตัดออก):', dict(orphans))

# ============================================================================
# STEP 3: หน่วยว่าง (ใหม่+เดิม) → PB64 ส่วนที่เหลือ, VC64
# ============================================================================
print('\n[3/7] หน่วยว่าง (ใหม่ + เดิม)')

def extract_vacant_sheet(sheet_name, check_valid=True):
    ws = wb[sheet_name]
    headers = {ws.cell(row=1,column=c).value: c for c in range(1, ws.max_column+1)}
    col_proj = headers['PROJ_CODE']; col_mat = headers['CL_MATERIAL_NO']
    col_submodel = headers['SUB_MODEL_NAME']; col_house = headers['HOUSE_NO_ACT']
    col_planno = headers['HOUSE_NO']; col_am = headers['STOCK_STATUS']; col_au = headers['TYPE_NAME']
    col_bb = headers['CONTRACT_STATUS']; col_bc = headers['END_GUARANTEE_DATE']
    col_bd = headers['REDATE']; col_be = headers['DATEK']; col_bf = headers['STOCK_RMK']; col_bg = headers['REF_NO']
    BUCKET_MAP = {
        ('ว่างพร้อมขาย','เช่า'):'ready_rent', ('ว่างไม่พร้อมขาย','เช่า'):'notready_rent',
        ('ว่างพร้อมขาย','เช่าซื้อ'):'ready_sale', ('ว่างไม่พร้อมขาย','เช่าซื้อ'):'notready_sale',
    }
    records, orphan_ct = [], 0
    for row in range(2, ws.max_row+1):
        proj = ws.cell(row=row, column=col_proj).value
        if proj is None: continue
        proj = str(proj).strip()
        bucket = BUCKET_MAP.get((ws.cell(row=row, column=col_am).value, ws.cell(row=row, column=col_au).value))
        if bucket is None: continue
        if check_valid and proj not in valid_proj:
            orphan_ct += 1
            continue
        records.append({
            'proj_code': proj, 'mat': ws.cell(row=row, column=col_mat).value or '',
            'sub_model_name': ws.cell(row=row, column=col_submodel).value or '',
            'house_no': ws.cell(row=row, column=col_house).value or '',
            'plan_no': str(ws.cell(row=row, column=col_planno).value or ''),
            'status': bucket,
            'contract_status': ws.cell(row=row, column=col_bb).value or '',
            'end_guarantee_date': ws.cell(row=row, column=col_bc).value or '',
            'redate': ws.cell(row=row, column=col_bd).value or '',
            'datek': ws.cell(row=row, column=col_be).value or 0,
            'stock_rmk': ws.cell(row=row, column=col_bf).value or '',
            'ref_no': ws.cell(row=row, column=col_bg).value or '',
        })
    return records, orphan_ct

vacant_new, orph1 = extract_vacant_sheet('หน่วยว่าง', check_valid=True)
print('  New vacant matched:', len(vacant_new), '| orphans:', orph1, '| ต้องตรงกับ Sum vacant_total ด้านบน')
print('  By bucket:', dict(Counter(r['status'] for r in vacant_new)))

vacant_old, orph2 = extract_vacant_sheet('หน่วยว่างเดิม', check_valid=False)
print('  Old vacant (หน่วยว่างเดิม) total (ไม่กรองโครงการ):', len(vacant_old))
print('  By bucket:', dict(Counter(r['status'] for r in vacant_old)))

# ตรวจ MAT ซ้ำระหว่างหน่วยว่างกับเช่าเหมา (ไม่ควรมีเลย — เคยเจอเคสโครงการรามอินทรา 408 หน่วย)
lease_whole_mats = set(r['mat'] for r in active_records if r['source']=='L')
vacant_mats = set(r['mat'] for r in vacant_new)
dup_mats = vacant_mats & lease_whole_mats
print(f'  MAT ซ้ำระหว่างหน่วยว่าง-เช่าเหมา: {len(dup_mats)}', '⚠️ ควรเป็น 0!' if dup_mats else '✅')

# ============================================================================
# STEP 4: เทียบหน่วยว่างเก่า-ใหม่ → CH64 (5 กลุ่ม)
# ============================================================================
print('\n[4/7] เปรียบเทียบหน่วยว่างเก่า-ใหม่ (5 กลุ่ม)')
new_vacant_by_mat = {r['mat']: r for r in vacant_new}
active_by_mat = {r['mat']: r for r in active_records}
old_mat_set = set(r['mat'] for r in vacant_old)

comparison = []
groups = Counter()
for r in vacant_old:
    mat, old_status = r['mat'], r['status']
    if mat in new_vacant_by_mat:
        nr = new_vacant_by_mat[mat]
        grp = 'still_same' if nr['status']==old_status else 'still_changed'
        comparison.append({'mat':mat,'proj_code':r['proj_code'],'group':grp,'old_status':old_status,'new_status':nr['status']})
    elif mat in active_by_mat:
        ar = active_by_mat[mat]
        comparison.append({'mat':mat,'proj_code':r['proj_code'],'group':'moved_active','old_status':old_status,'new_status':ar['status']})
    else:
        comparison.append({'mat':mat,'proj_code':r['proj_code'],'group':'not_found','old_status':old_status,'new_status':None})
    groups[comparison[-1]['group']] += 1
for r in vacant_new:
    if r['mat'] not in old_mat_set:
        comparison.append({'mat':r['mat'],'proj_code':r['proj_code'],'group':'newly_vacant','old_status':None,'new_status':r['status']})
        groups['newly_vacant'] += 1

print('  Group totals:', dict(groups))
eq1 = groups['still_same']+groups['still_changed']+groups['moved_active']+groups['not_found']
eq2 = groups['still_same']+groups['still_changed']+groups['newly_vacant']
print(f'  สมการ 1 (เก่า): {eq1} vs old_vacant={len(vacant_old)}', '✅' if eq1==len(vacant_old) else '❌ ไม่ลงตัว!')
print(f'  สมการ 2 (ใหม่): {eq2} vs new_vacant={len(vacant_new)}', '✅' if eq2==len(vacant_new) else '❌ ไม่ลงตัว!')

filtered_comparison = [c for c in comparison if c['proj_code'] in valid_proj]
print('  After filtering to 185 projects:', len(filtered_comparison), dict(Counter(c['group'] for c in filtered_comparison)))

# ============================================================================
# STEP 5: ผลสำรวจกายภาพ + รูปภาพ → SV64
# ============================================================================
print('\n[5/7] ผลสำรวจกายภาพ (ชีต "สภาพภาพอาคารว่างพร้อมขาย")')
ws = wb['สภาพภาพอาคารว่างพร้อมขาย']
headers = {ws.cell(row=1,column=c).value: c for c in range(1, ws.max_column+1)}
print('  Columns found:', list(headers.keys()))
col_mat = headers.get('MAT'); col_house = headers.get('บ้านเลขที่'); col_occ = headers.get('การอยู่อาศัย')
col_room = headers.get('สภาพห้อง'); col_asset = headers.get('ทรัพย์สิน')
col_water = headers.get('มิเตอร์น้ำ'); col_elec = headers.get('มิเตอร์ไฟ'); col_remark = headers.get('หมายเหตุเพิ่มเติม')
col_links = [headers.get(f'linkรูป{i}') for i in range(1, 7)]

survey_records = []
for row in range(2, ws.max_row+1):
    mat = ws.cell(row=row, column=col_mat).value if col_mat else None
    if not mat: continue
    links = [str(ws.cell(row=row, column=c).value).strip() if (c and ws.cell(row=row, column=c).value) else '' for c in col_links]
    survey_records.append({
        'mat': str(mat),
        'house_no_survey': str(ws.cell(row=row, column=col_house).value or '') if col_house else '',
        'occupancy': (ws.cell(row=row, column=col_occ).value or '') if col_occ else '',
        'room_condition': (ws.cell(row=row, column=col_room).value or '') if col_room else '',
        'asset': (ws.cell(row=row, column=col_asset).value or '') if col_asset else '',
        'water_meter': (ws.cell(row=row, column=col_water).value or '') if col_water else '',
        'elec_meter': (ws.cell(row=row, column=col_elec).value or '') if col_elec else '',
        'remark_survey': (ws.cell(row=row, column=col_remark).value or '') if col_remark else '',
        'links': links,
    })
print('  Total survey records:', len(survey_records))
with_links = [r for r in survey_records if any(r['links'])]
print('  Records with >=1 photo:', len(with_links))

survey_mats = set(r['mat'] for r in survey_records)
all_vacant_mats = set(r['mat'] for r in vacant_new)
print(f'  Join กับหน่วยว่างปัจจุบัน: {len(survey_mats & all_vacant_mats)}/{len(survey_mats)} ตรงกัน',
      f'({len(survey_mats - all_vacant_mats)} หาไม่เจอ — อาจย้ายเป็น Active แล้วหรือ MAT พิมพ์ผิด)')

# ============================================================================
# STEP 6: เข้ารหัส gzip+base64 ทั้ง 7 ก้อน
# ============================================================================
print('\n[6/7] เข้ารหัสข้อมูล (gzip + base64)')
blobs = {}

cols_db = list(raw_rows[0].keys())
blobs['DB64'] = compress({'cols':cols_db, 'rows':[[r[c] for c in cols_db] for r in raw_rows]})

BUCKET_LABEL = {
    'rent':'เช่า','lease_whole':'เช่าเหมา','hire_purchase':'เช่าซื้อ','pending_sale':'จะซื้อจะขาย',
    'guarantee':'ติดค้ำประกันฯ','pending_transfer':'รอโอน',
    'ready_rent':'ว่างพร้อมขาย(เพื่อเช่า)','notready_rent':'ว่างไม่พร้อมขาย(เพื่อเช่า)',
    'ready_sale':'ว่างพร้อมขาย(เพื่อขาย)','notready_sale':'ว่างไม่พร้อมขาย(เพื่อขาย)'
}
pb_rows = [[r['proj_code'], r['source'], r['house_no'] or '', r['mat'] or '', BUCKET_LABEL[r['status']], r['ref_status'] or ''] for r in active_records]
pb_rows += [[r['proj_code'], 'V', r['house_no'] or '', r['mat'] or '', BUCKET_LABEL[r['status']], ''] for r in vacant_new]
blobs['PB64'] = compress({'cols':['pc','s','h','m','b','n'], 'rows':pb_rows})

blobs['CS64'] = compress({'cols':['m','y','o'], 'rows':extra_cs})
blobs['PF64'] = compress({'cols':['m','y'], 'rows':extra_pf})

vc_rows = [[r['proj_code'], r['mat'], r['house_no'], r['status'], r['contract_status'], r['end_guarantee_date'],
            r['redate'], r['datek'], r['stock_rmk'], r['ref_no'], r['sub_model_name'], r['plan_no']] for r in vacant_new]
blobs['VC64'] = compress({'cols':['p','m','h','t','c','e','r','k','s','f','sm','pn'], 'rows':vc_rows})

ch_rows = [[c['mat'], c['proj_code'], c['group'], c['old_status'] or '', c['new_status'] or ''] for c in comparison]
blobs['CH64'] = compress({'cols':['m','p','g','os','ns'], 'rows':ch_rows})

sv_rows = []
for r in survey_records:
    row = [r['mat'], r['room_condition'], r['asset'], r['water_meter'], r['elec_meter'],
           r['remark_survey'], r['house_no_survey'], r['occupancy']]
    row.extend(r['links'])
    sv_rows.append(row)
blobs['SV64'] = compress({'cols':['m','rc','as','wm','em','rk','hs','oc','l1','l2','l3','l4','l5','l6'], 'rows':sv_rows})

for name, b64 in blobs.items():
    print(f'  {name}: {len(b64)/1024:.1f} KB (base64)')

# ============================================================================
# STEP 7: ฝังเข้า index.html
# ============================================================================
print('\n[7/7] ฝังข้อมูลเข้า index.html')
if not os.path.exists(INDEX_HTML):
    print(f'  ⚠️ ไม่พบไฟล์ {INDEX_HTML} — ข้ามขั้นตอนนี้ ไปเข้ารหัสเองด้วยมือแทน')
    for name, b64 in blobs.items():
        open(f'{name}_v2.txt', 'w').write(b64)
    print('  บันทึกไฟล์ .txt ไว้ให้แทนแล้ว (แทนที่ const XXX64 = "..."; ในไฟล์เองตามชื่อ)')
else:
    with open(INDEX_HTML, encoding='utf-8') as f:
        html = f.read()
    for name, b64 in blobs.items():
        pattern = re.compile(r'const ' + name + r' = "[^"]+";')
        replacement = f'const {name} = "{b64}";'
        html, n = pattern.subn(replacement, html)
        print(f'  {name}: แทนที่ {n} จุด', '✅' if n==1 else '⚠️ ควรเจอ 1 จุดเท่านั้น ตรวจสอบด้วย')
    with open(INDEX_HTML, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f'  บันทึก {INDEX_HTML} เรียบร้อย')

print('\n' + '='*70)
print('เสร็จสิ้น — อย่าลืมตรวจ syntax (node -e) + ทดสอบคลิกจริงด้วย jsdom ก่อนส่งมอบ')
print('='*70)
