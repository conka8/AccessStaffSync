"""Generate Excel fixtures with the required string columns (artifact_tool dependency)."""
from pathlib import Path
from artifact_tool import Workbook, SpreadsheetFile

ROOT = Path(__file__).resolve().parents[1]
HEADERS = ['EventID','StaffID','EventTimeUTC','JobOfferStatus','Availability','Address','MessagingConsent']
DATA = {
    'updates_001.xlsx': [
        ['E-1001','000042','2026-10-27T09:00:00Z','Offered','Available','','YES'],
        ['E-1002','000314','2026-10-27T09:15:00Z','','','Calle Alcalá 21, Madrid','NO'],
        ['E-1003','001234','2026-10-27T09:30:00Z','Interviewing','Weekends','',''],
        ['E-1004','999999','2026-10-27T09:45:00Z','Available','','',''],
        ['E-1001','000042','2026-10-27T09:00:00Z','Offered','Available','','YES'],
        ['E-1005','000042','2026-10-28T08:00:00Z','','','Carrer de Mallorca 10, Barcelona',''],
    ],
    'updates_002.xlsx': [
        ['E-1006','000042','2026-10-26T09:00:00Z','','Unavailable','',''],
        ['E-1007','000042','2026-10-28T10:00:00Z','Accepted','Available next week','',''],
        ['E-1002','000314','2026-10-27T09:15:00Z','','','Calle Alcalá 21, Madrid','NO'],
        ['E-1008','000314','2026-10-28T10:00:00Z','','','','MAYBE'],
        ['E-1009','001234','2026/10/28 10:00','','','','YES'],
        ['E-1010','001234','2026-10-28T10:30:00Z','','','__CLEAR__',''],
        ['E-1011','000314','2026-10-28T11:00:00Z','','','','YES'],
    ],
    'updates_template.xlsx': [],
}

for filename, rows in DATA.items():
    w = Workbook.create()
    s = w.worksheets.add('WhatsAppUpdates')
    s.get_range('A:G').set_number_format('@')
    # Trailing whitespace forces ISO timestamps to remain text in artifact_tool.
    # The VBA importer trims it before strict ISO-UTC validation.
    text_rows = [r[:2] + [r[2] + ' '] + r[3:] for r in rows]
    s.get_range_by_indexes(0, 0, len(rows)+1, 7).values = [HEADERS] + text_rows
    s.freeze_panes.freeze_rows(1)
    s.get_range('A1:G1').format = {
        'fill': '#172C45',
        'font': {'color': '#FFFFFF','bold': True, 'size': 11},
        'row_height': 29,
        'vertical_alignment': 'center',
    }
    s.get_range('A:A').format.column_width = 16
    s.get_range('B:B').format.column_width = 15
    s.get_range('C:C').format.column_width = 25
    s.get_range('D:D').format.column_width = 23
    s.get_range('E:E').format.column_width = 26
    s.get_range('F:F').format.column_width = 43
    s.get_range('G:G').format.column_width = 22
    s.get_range('A:G').set_number_format('@')
    if rows:
        s.get_range_by_indexes(1, 0, len(rows), 7).format.row_height = 25
        s.get_range_by_indexes(1, 0, len(rows), 7).format.vertical_alignment = 'center'
    out = ROOT/'samples'/'inbox'/filename
    SpreadsheetFile.export_xlsx(w).save(str(out))
    print(filename, out.stat().st_size)
    if filename == 'updates_001.xlsx':
        print(w.inspect({'kind':'table','range':'WhatsAppUpdates!A1:G3','include':'values','table_max_rows':3,'table_max_cols':7}).ndjson[:1600])
