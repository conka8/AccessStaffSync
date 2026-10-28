"""Contract checks for sample XLSX batches and idempotence/staleness logic.

This is a platform-independent reference model, NOT a VBA/Access execution test.
Use setup/Run-Demo.ps1 on Windows with Access + Excel for the integration test.
"""
import unittest
from datetime import datetime
from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
NS = {'x':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
HEADER = ('EventID','StaffID','EventTimeUTC','JobOfferStatus','Availability','Address','MessagingConsent')
FIELDS = ('JobOfferStatus','Availability','Address','MessagingConsent')


def rows_from_xlsx(file):
    """Read the deliberately string-only generated sample workbook via OpenXML."""
    with zipfile.ZipFile(file) as archive:
        root = ET.fromstring(archive.read('xl/worksheets/sheet1.xml'))
    rows = []
    for row in root.findall('.//x:sheetData/x:row', NS):
        record = [''] * 7
        for cell in row.findall('x:c', NS):
            address = cell.get('r')
            col = ''.join(ch for ch in address if ch.isalpha())
            idx = ord(col[0]) - ord('A')
            assert cell.get('t') in ('str','inlineStr'), f'fixture must preserve text: {address}'
            value = cell.find('x:v', NS)
            if value is not None:
                record[idx] = value.text or ''
        rows.append([v.strip() for v in record])
    assert rows[0] == list(HEADER), rows[0]
    return [dict(zip(HEADER, r)) for r in rows[1:]]


def apply_batch(staff, ledger, events):
    totals = {'total':0,'applied':0,'stale':0,'duplicate':0,'rejected':0}
    seen = set()
    for row in events:
        totals['total'] += 1
        eid, staff_id = row['EventID'], row['StaffID']
        try:
            ts = datetime.strptime(row['EventTimeUTC'], '%Y-%m-%dT%H:%M:%SZ')
            if not eid or not staff_id or not any(row[f] for f in FIELDS): raise ValueError('empty')
            consent = row['MessagingConsent'].upper()
            if consent and consent not in ('YES','NO','TRUE','FALSE','1','0','SI'):
                raise ValueError('invalid consent')
        except ValueError:
            totals['rejected'] += 1
            continue
        if eid in seen or eid in ledger:
            totals['duplicate'] += 1
            continue
        seen.add(eid)
        if staff_id not in staff:
            totals['rejected'] += 1
            continue
        record = staff[staff_id]
        changed = False
        for name in FIELDS:
            patch = row[name]
            if patch and (name not in record['timestamps'] or ts >= record['timestamps'][name]):
                if name == 'MessagingConsent':
                    patch = patch.upper() in ('YES','TRUE','1','SI')
                elif patch == '__CLEAR__':
                    patch = None
                record[name] = patch
                record['timestamps'][name] = ts
                changed = True
        ledger[eid] = 'APPLIED' if changed else 'STALE'
        totals['applied' if changed else 'stale'] += 1
    return totals


class ContractTests(unittest.TestCase):
    def setUp(self):
        self.staff = {f'{i:06d}': {'JobOfferStatus':'Pending','Availability':'Unknown',
                          'Address':f'Example address {i}','MessagingConsent':False,
                          'timestamps':{}} for i in range(1,4001)}
        self.ledger = {}

    def test_two_fixture_batches(self):
        first = rows_from_xlsx(ROOT/'samples/inbox/updates_001.xlsx')
        second = rows_from_xlsx(ROOT/'samples/inbox/updates_002.xlsx')
        a = apply_batch(self.staff,self.ledger,first)
        b = apply_batch(self.staff,self.ledger,second)
        self.assertEqual(a, {'total':6,'applied':4,'stale':0,'duplicate':1,'rejected':1})
        self.assertEqual(b, {'total':7,'applied':3,'stale':1,'duplicate':1,'rejected':2})
        self.assertEqual(self.staff['000042']['JobOfferStatus'],'Accepted')
        self.assertEqual(self.staff['000042']['Availability'],'Available next week')
        self.assertEqual(self.staff['000042']['Address'],'Carrer de Mallorca 10, Barcelona')
        self.assertEqual(self.staff['000314']['MessagingConsent'],True)
        self.assertEqual(self.staff['001234']['Address'],None)
        self.assertEqual(len(self.ledger),8)

    def test_duplicate_delivery_is_idempotent(self):
        first = rows_from_xlsx(ROOT/'samples/inbox/updates_001.xlsx')
        apply_batch(self.staff,self.ledger,first)
        before = {k:v.copy() for k,v in self.staff.items()}
        check = apply_batch(self.staff,self.ledger,first)
        self.assertEqual(check['applied'],0)
        self.assertEqual(check['duplicate'],5) # 5 valid rows; unknown remains rejected
        self.assertEqual(self.staff,before)

    def test_workbooks_contain_text_staff_ids_and_utc(self):
        for p in (ROOT/'samples/inbox').glob('*.xlsx'):
            rows = rows_from_xlsx(p)
            for r in rows:
                self.assertFalse(r['StaffID'] and len(r['StaffID']) != 6)
                self.assertTrue(r['EventTimeUTC'])  # Includes one intentionally invalid timestamp

    def test_vba_contains_essential_safety_features(self):
        code = (ROOT/'vba/modStaffSync.bas').read_text()
        for required in ('BeginTrans','CommitTrans','Rollback', 'Lock Read Write',
                         'ReconcileCompletedFiles','NullablePatch','IsNewEnough',
                         'Recordset','AutomationSecurity = 3','EventID','SyncTick'):
            with self.subTest(required=required): self.assertIn(required,code)

if __name__ == '__main__':
    unittest.main()
