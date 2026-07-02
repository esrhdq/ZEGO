import sqlite3
conn = sqlite3.connect(r'C:\Users\admin\Desktop\claude\ZEGO\zego.db')
conn.row_factory = sqlite3.Row

tbls = conn.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()
print('Tables:', [t['name'] for t in tbls])

try:
    cols = conn.execute('PRAGMA table_info(form_types)').fetchall()
    print('form_types cols:', [c['name'] for c in cols])
    rows = conn.execute('SELECT id, name, is_active, sort_order FROM form_types ORDER BY sort_order').fetchall()
    for r in rows:
        print(dict(r))
except Exception as e:
    print('form_types error:', e)

try:
    print('--- branches ---')
    rows = conn.execute('SELECT id, code, name, type FROM branches ORDER BY type, code').fetchall()
    for r in rows:
        print(dict(r))
except Exception as e:
    print('branches error:', e)
