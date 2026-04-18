from flask import Flask, render_template, request, redirect, url_for, session, jsonify, flash, send_file
import psycopg2
import psycopg2.extras
import os
import hashlib
import sqlite3
import re
from functools import wraps
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

app = Flask(__name__, template_folder=os.path.join(os.path.dirname(os.path.abspath(__file__)), 'templates'))
app.secret_key = os.environ.get('SECRET_KEY', 'dev-secret-key-change-in-production')

DATABASE_URL = os.environ.get('DATABASE_URL', '')
# 비밀번호 특수문자 URL 인코딩 + Supabase SSL 처리
if DATABASE_URL:
    from urllib.parse import urlparse, urlunparse, quote
    _p = urlparse(DATABASE_URL)
    if _p.password:
        _encoded_pw = quote(_p.password, safe='')
        _netloc = f"{_p.username}:{_encoded_pw}@{_p.hostname}"
        if _p.port:
            _netloc += f":{_p.port}"
        DATABASE_URL = urlunparse((_p.scheme, _netloc, _p.path, '', 'sslmode=require', ''))
    elif 'sslmode' not in DATABASE_URL:
        DATABASE_URL += ('&' if '?' in DATABASE_URL else '?') + 'sslmode=require'
ALLOWED_IPS  = [ip.strip() for ip in os.environ.get('ALLOWED_IPS', '').split(',') if ip.strip()]
USE_SQLITE   = not DATABASE_URL
# Vercel 환경에서는 /tmp만 쓰기 가능 — 로컬은 프로젝트 폴더 사용
_local_db    = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'inventory.db')
SQLITE_DB    = '/tmp/inventory.db' if not os.access(os.path.dirname(_local_db), os.W_OK) else _local_db


# ── DB 연결 ───────────────────────────────────────────────────────────────────

class _DB:
    """psycopg2를 sqlite3 인터페이스처럼 사용하기 위한 래퍼"""
    def __init__(self, conn):
        self._conn = conn

    def execute(self, sql, params=()):
        cur = self._conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        cur.execute(sql, params if params else None)
        return cur

    def commit(self):
        self._conn.commit()

    def rollback(self):
        self._conn.rollback()

    def close(self):
        self._conn.close()


def _pg_to_sqlite(sql):
    """PostgreSQL SQL을 SQLite 호환 구문으로 변환"""
    # TO_CHAR 변환 (INTERVAL 포함 케이스 먼저)
    sql = re.sub(
        r"TO_CHAR\(CURRENT_DATE\s*-\s*INTERVAL\s*'(\d+)\s*months',\s*'YYYY-MM'\)",
        r"strftime('%Y-%m', date('now', '-\1 months'))", sql)
    sql = re.sub(r"TO_CHAR\(CURRENT_DATE,\s*'YYYY-MM'\)", "strftime('%Y-%m', 'now')", sql)
    sql = re.sub(r"TO_CHAR\(([^,]+),\s*'YYYY-MM'\)", r"strftime('%Y-%m', \1)", sql)
    # INTERVAL 날짜 산술
    sql = re.sub(r"CURRENT_DATE\s*-\s*INTERVAL\s*'(\d+)\s*months'", r"date('now', '-\1 months')", sql)
    sql = re.sub(r"CURRENT_DATE\s*-\s*INTERVAL\s*'(\d+)\s*days'", r"date('now', '-\1 days')", sql)
    # 타입 캐스트
    sql = re.sub(r'(\w+(?:\.\w+)*)::date\b', r'date(\1)', sql)
    sql = re.sub(r'(\w+(?:\.\w+)*)::text\b', r'\1', sql)
    # NOW() → CURRENT_TIMESTAMP
    sql = re.sub(r'\bNOW\(\)', 'CURRENT_TIMESTAMP', sql, flags=re.IGNORECASE)
    # 플레이스홀더
    sql = sql.replace('%s', '?')
    return sql


class _SQLiteDB:
    def __init__(self, conn):
        self._conn = conn

    def execute(self, sql, params=()):
        sql = _pg_to_sqlite(sql)
        cur = self._conn.cursor()
        cur.execute(sql, list(params) if params else [])
        return cur

    def commit(self):   self._conn.commit()
    def rollback(self): self._conn.rollback()
    def close(self):    self._conn.close()


def get_db():
    if USE_SQLITE:
        conn = sqlite3.connect(SQLITE_DB)
        conn.row_factory = sqlite3.Row
        return _SQLiteDB(conn)
    return _DB(psycopg2.connect(DATABASE_URL))


def hash_pw(pw):
    return hashlib.sha256(pw.encode()).hexdigest()


def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated


# ── Jinja2 필터 ──────────────────────────────────────────────────────────────

@app.template_filter('dt')
def dt_filter(value):
    if value is None:
        return '—'
    return str(value)[:16]


# ── 오류 핸들러 (디버그용) ────────────────────────────────────────────────────

@app.errorhandler(Exception)
def handle_exception(e):
    import traceback
    tb = traceback.format_exc()
    print(f"[UNHANDLED ERROR] {tb}")
    return f"<pre style='color:red'>{tb}</pre>", 500


# ── IP 화이트리스트 ───────────────────────────────────────────────────────────

@app.before_request
def check_ip():
    if ALLOWED_IPS:
        forwarded = request.headers.get('X-Forwarded-For', '')
        client_ip = forwarded.split(',')[0].strip() if forwarded else request.remote_addr
        if client_ip not in ALLOWED_IPS:
            return render_template('403.html'), 403


# ── DB 초기화 ─────────────────────────────────────────────────────────────────

def init_db():
    conn = get_db()
    try:
        if USE_SQLITE:
            conn.execute('''
                CREATE TABLE IF NOT EXISTS branches (
                    id   INTEGER PRIMARY KEY,
                    code TEXT NOT NULL UNIQUE,
                    name TEXT NOT NULL,
                    type TEXT NOT NULL
                )
            ''')
            conn.execute('''
                CREATE TABLE IF NOT EXISTS form_types (
                    id            INTEGER PRIMARY KEY,
                    name          TEXT NOT NULL UNIQUE,
                    unit          TEXT NOT NULL,
                    unit_detail   TEXT,
                    unit_price    INTEGER DEFAULT 0,
                    min_threshold INTEGER DEFAULT 2
                )
            ''')
            conn.execute('''
                CREATE TABLE IF NOT EXISTS inventory (
                    id            INTEGER PRIMARY KEY,
                    branch_id     INTEGER NOT NULL,
                    form_type_id  INTEGER NOT NULL,
                    quantity      INTEGER DEFAULT 0,
                    last_updated  TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (branch_id)    REFERENCES branches(id),
                    FOREIGN KEY (form_type_id) REFERENCES form_types(id),
                    UNIQUE(branch_id, form_type_id)
                )
            ''')
            conn.execute('''
                CREATE TABLE IF NOT EXISTS transactions (
                    id             INTEGER PRIMARY KEY,
                    type           TEXT NOT NULL,
                    form_type_id   INTEGER NOT NULL,
                    from_branch_id INTEGER,
                    to_branch_id   INTEGER,
                    quantity       INTEGER NOT NULL,
                    notes          TEXT,
                    created_by     TEXT,
                    created_at     TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    period_month   TEXT,
                    FOREIGN KEY (form_type_id)   REFERENCES form_types(id),
                    FOREIGN KEY (from_branch_id) REFERENCES branches(id),
                    FOREIGN KEY (to_branch_id)   REFERENCES branches(id)
                )
            ''')
            conn.execute('''
                CREATE TABLE IF NOT EXISTS users (
                    id        INTEGER PRIMARY KEY,
                    username  TEXT NOT NULL UNIQUE,
                    password  TEXT NOT NULL,
                    branch_id INTEGER,
                    role      TEXT DEFAULT 'staff',
                    FOREIGN KEY (branch_id) REFERENCES branches(id)
                )
            ''')
        else:
            conn.execute('''
                CREATE TABLE IF NOT EXISTS branches (
                    id   SERIAL PRIMARY KEY,
                    code TEXT NOT NULL UNIQUE,
                    name TEXT NOT NULL,
                    type TEXT NOT NULL
                )
            ''')
            conn.execute('''
                CREATE TABLE IF NOT EXISTS form_types (
                    id            SERIAL PRIMARY KEY,
                    name          TEXT NOT NULL UNIQUE,
                    unit          TEXT NOT NULL,
                    unit_detail   TEXT,
                    unit_price    INTEGER DEFAULT 0,
                    min_threshold INTEGER DEFAULT 2
                )
            ''')
            conn.execute('''
                CREATE TABLE IF NOT EXISTS inventory (
                    id            SERIAL PRIMARY KEY,
                    branch_id     INTEGER NOT NULL,
                    form_type_id  INTEGER NOT NULL,
                    quantity      INTEGER DEFAULT 0,
                    last_updated  TIMESTAMP DEFAULT NOW(),
                    FOREIGN KEY (branch_id)    REFERENCES branches(id),
                    FOREIGN KEY (form_type_id) REFERENCES form_types(id),
                    UNIQUE(branch_id, form_type_id)
                )
            ''')
            conn.execute('''
                CREATE TABLE IF NOT EXISTS transactions (
                    id             SERIAL PRIMARY KEY,
                    type           TEXT NOT NULL,
                    form_type_id   INTEGER NOT NULL,
                    from_branch_id INTEGER,
                    to_branch_id   INTEGER,
                    quantity       INTEGER NOT NULL,
                    notes          TEXT,
                    created_by     TEXT,
                    created_at     TIMESTAMP DEFAULT NOW(),
                    FOREIGN KEY (form_type_id)   REFERENCES form_types(id),
                    FOREIGN KEY (from_branch_id) REFERENCES branches(id),
                    FOREIGN KEY (to_branch_id)   REFERENCES branches(id)
                )
            ''')
            conn.execute('''
                CREATE TABLE IF NOT EXISTS users (
                    id        SERIAL PRIMARY KEY,
                    username  TEXT NOT NULL UNIQUE,
                    password  TEXT NOT NULL,
                    branch_id INTEGER,
                    role      TEXT DEFAULT 'staff',
                    FOREIGN KEY (branch_id) REFERENCES branches(id)
                )
            ''')

        branches = [
            ('GMP', '김포', 'DOM'), ('CJU', '제주', 'DOM'), ('CJJ', '청주', 'DOM'),
            ('PUS', '부산', 'DOM'), ('TSA', '송산', 'DOM'), ('KMJ', '광주', 'DOM'),
            ('서울역', '서울역', 'DOM'), ('광명역', '광명역', 'DOM'), ('이지드랍', '이지드랍', 'DOM'),
            ('ICN', '인천', 'INTL'), ('TPE', '타이페이', 'INTL'), ('NRT', '나리타', 'INTL'),
            ('KIX', '간사이', 'INTL'), ('FUK', '후쿠오카', 'INTL'), ('CTS', '삿포로', 'INTL'),
            ('OKA', '오키나와', 'INTL'), ('TKS', '도쿠시마', 'INTL'), ('BKK', '방콕', 'INTL'),
            ('CNX', '치앙마이', 'INTL'), ('DAD', '다낭', 'INTL'), ('CXR', '나트랑', 'INTL'),
            ('PQC', '푸꾸옥', 'INTL'), ('MDC', '마나도', 'INTL'), ('PVG', '상하이', 'INTL'),
            ('YNJ', '연길', 'INTL'), ('CGO', '정저우', 'INTL'), ('YNT', '옌타이', 'INTL'),
            ('HKG', '홍콩', 'INTL'), ('ALA', '알마티', 'INTL'),
            ('CARGO', '화물파트', 'CARGO'),
        ]
        for code, name, btype in branches:
            conn.execute(
                'INSERT INTO branches (code, name, type) VALUES (%s,%s,%s) ON CONFLICT(code) DO NOTHING',
                (code, name, btype)
            )

        form_types = [
            ('DOM BOARDING PASS (롤)',        'BOX', '50롤',    60000,  3),
            ('INTL BOARDING PASS(QR)',         'BOX', '5,000장', 175000, 3),
            ('INTL BOARDING PASS(QR, ICN)',    'BOX', '5,000장', 175000, 2),
            ('AUTO BAG TAG',                   'BOX', '10롤',   118000,  5),
            ('BAG TIPS',                       'BOX', '5,000장',  50000, 3),
            ('BAG TIPS (SNOOPY, DOM)',         'BOX', '5,000장',  57000, 2),
            ('MANUAL BAG TAG',                 'BOX', '5,000장', 180000, 2),
            ('Carry on Bag TAG (INTL)',        'BOX', '5,000장', 145000, 2),
            ('SRI 봉투(大)',                   'BOX', '500장',   330000, 1),
            ('CO-MAIL 봉투(NEW)',              'BOX', '200장',   160000, 1),
            ('유상비닐(小/PPS)',               '포대', '100개',  110000, 2),
            ('유상비닐(大/PPL)',               '포대', '100개',  130000, 2),
            ('BOX TAPE',                       'BOX', '50개',    55000, 3),
            ('PREMIUM TAG(D/S)',               'BOX', '5,000장',  80000, 2),
            ('FRAGILE TAG(NEW)',               'BOX', '5,000장',  80000, 2),
            ('HEAVY TAG',                      'BOX', '5,000장',  80000, 2),
            ('GTOG TAG',                       'BOX', '5,000장',  80000, 2),
            ('TRANSFER TAG',                   'BOX', '5,000장',  80000, 1),
            ('비상구열 스티커',                'BOX', '20,000장',110000, 1),
            ('AOC LABEL',                      'BOX', '5,000장',  80000, 1),
            ('POB LABEL',                      'BOX', '5,000장',  80000, 1),
            ('COB LABEL',                      'BOX', '5,000장',  80000, 1),
            ('UP SIDE LABEL',                  'BOX', '5,000장',  80000, 1),
            ('휠체어 배터리 분리 L/B',         'BOX', '5,000장',  80000, 1),
            ('CORROSIVE LABEL',                'BOX', '5,000장',  80000, 1),
            ('Dry Ice LABEL',                  'BOX', '5,000장',  80000, 1),
            ('한국 입국신고서 (ENG/CNA)',       'BOX', '5,000장', 150000, 1),
            ('제주 E/D카드',                   'BOX', '5,000장', 150000, 1),
            ('한국 세관신고서 (ENG/CNA)',       'BOX', '5,000장', 150000, 1),
            ('한국 세관신고서 (ENG/JPN)',       'BOX', '5,000장', 150000, 1),
            ('서약서',                         '권',  '100조',    8500, 3),
            ('합의서',                         '권',  '100조',    8500, 2),
            ('반려동물 서약서',                '권',  '100조',   10000, 2),
            ('악기서약서',                     '권',  '100조',   10000, 1),
            ('보호자 서약서',                  '권',  '100조',    8500, 1),
            ('총기인수인계서',                 '권',  '100조',   10000, 1),
            ('PIR',                            '권',  '100조',   10000, 2),
            ('SHR',                            '권',  '100조',   14000, 1),
            ('NOTOC',                          '권',  '100조',   10000, 1),
            ('BAG BINGO CHART(양면)',          '권',  '100조',    3200, 5),
        ]
        for name, unit, ud, price, thr in form_types:
            conn.execute(
                'INSERT INTO form_types (name, unit, unit_detail, unit_price, min_threshold) '
                'VALUES (%s,%s,%s,%s,%s) ON CONFLICT(name) DO NOTHING',
                (name, unit, ud, price, thr)
            )

        admin_pw = hash_pw('admin1234')
        conn.execute(
            'INSERT INTO users (username, password, branch_id, role) VALUES (%s,%s,NULL,%s) '
            'ON CONFLICT(username) DO NOTHING',
            ('admin', admin_pw, 'admin')
        )
        staff_pw = hash_pw('staff1234')
        conn.execute(
            "INSERT INTO users (username, password, branch_id, role) "
            "SELECT %s, %s, id, 'staff' FROM branches WHERE code='GMP' "
            "ON CONFLICT(username) DO NOTHING",
            ('gmp', staff_pw)
        )
        conn.execute(
            "INSERT INTO users (username, password, branch_id, role) "
            "SELECT %s, %s, id, 'staff' FROM branches WHERE code='ICN' "
            "ON CONFLICT(username) DO NOTHING",
            ('icn', staff_pw)
        )
        conn.commit()

        if USE_SQLITE:
            # SQLite: period_month 컬럼이 없으면 추가
            cols = [r[1] for r in conn.execute('PRAGMA table_info(transactions)').fetchall()]
            if 'period_month' not in cols:
                conn.execute('ALTER TABLE transactions ADD COLUMN period_month TEXT')
                conn.execute(
                    "UPDATE transactions SET period_month = strftime('%Y-%m', created_at) WHERE type='OUT'"
                )
        else:
            # PostgreSQL: DO $$ 블록으로 마이그레이션
            conn.execute('''
                DO $$
                BEGIN
                    IF NOT EXISTS (
                        SELECT 1 FROM information_schema.columns
                        WHERE table_name='transactions' AND column_name='period_month'
                    ) THEN
                        ALTER TABLE transactions ADD COLUMN period_month TEXT;
                        UPDATE transactions
                           SET period_month = TO_CHAR(created_at, 'YYYY-MM')
                         WHERE type = 'OUT';
                    END IF;
                END $$
            ''')
        conn.commit()

    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return redirect(url_for('dashboard') if 'user_id' in session else url_for('login'))


@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        conn = get_db()
        user = conn.execute(
            'SELECT u.*, b.name branch_name FROM users u LEFT JOIN branches b ON u.branch_id=b.id '
            'WHERE u.username=%s AND u.password=%s',
            (username, hash_pw(password))
        ).fetchone()
        conn.close()
        if user:
            session.update(user_id=user['id'], username=user['username'],
                           role=user['role'], branch_id=user['branch_id'],
                           branch_name=user['branch_name'])
            return redirect(url_for('dashboard'))
        flash('아이디 또는 비밀번호가 올바르지 않습니다.', 'danger')
    return render_template('login.html')


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


@app.route('/dashboard')
@login_required
def dashboard():
    conn = get_db()
    bid = session.get('branch_id')
    role = session.get('role')

    branch_cond = f'AND i.branch_id = {bid}' if role != 'admin' and bid else ''

    low_stock = conn.execute(f'''
        SELECT i.quantity, f.name form_name, f.min_threshold,
               b.name branch_name, b.code branch_code
        FROM inventory i
        JOIN form_types f ON i.form_type_id = f.id
        JOIN branches b ON i.branch_id = b.id
        WHERE i.quantity > 0 AND i.quantity <= f.min_threshold {branch_cond}
        ORDER BY i.quantity ASC LIMIT 20
    ''').fetchall()

    empty_stock = conn.execute(f'''
        SELECT f.name form_name, b.name branch_name, b.code branch_code
        FROM inventory i
        JOIN form_types f ON i.form_type_id = f.id
        JOIN branches b ON i.branch_id = b.id
        WHERE i.quantity = 0 {branch_cond}
        ORDER BY b.code, f.name LIMIT 20
    ''').fetchall()

    tx_cond = f'WHERE (t.from_branch_id={bid} OR t.to_branch_id={bid})' if role != 'admin' and bid else ''
    recent_tx = conn.execute(f'''
        SELECT t.*, f.name form_name,
               fb.name from_branch, tb.name to_branch
        FROM transactions t
        JOIN form_types f ON t.form_type_id = f.id
        LEFT JOIN branches fb ON t.from_branch_id = fb.id
        LEFT JOIN branches tb ON t.to_branch_id = tb.id
        {tx_cond}
        ORDER BY t.created_at DESC LIMIT 10
    ''').fetchall()

    stats = {
        'branches': conn.execute('SELECT COUNT(*) AS cnt FROM branches').fetchone()['cnt'],
        'forms':    conn.execute('SELECT COUNT(*) AS cnt FROM form_types').fetchone()['cnt'],
        'today_tx': conn.execute(
            "SELECT COUNT(*) AS cnt FROM transactions WHERE created_at::date = CURRENT_DATE"
        ).fetchone()['cnt'],
        'alerts':   len(low_stock) + len(empty_stock),
    }
    conn.close()
    return render_template('dashboard.html', low_stock=low_stock, empty_stock=empty_stock,
                           recent_tx=recent_tx, stats=stats)


@app.route('/inventory')
@login_required
def inventory():
    conn = get_db()
    branches = conn.execute('SELECT * FROM branches ORDER BY type, code').fetchall()
    form_types = conn.execute('SELECT * FROM form_types ORDER BY name').fetchall()

    bf = request.args.get('branch_id', '')
    ff = request.args.get('form_type_id', '')

    conditions, params = ['1=1'], []
    if session.get('role') != 'admin' and session.get('branch_id'):
        conditions.append('i.branch_id=%s'); params.append(session['branch_id'])
    elif bf:
        conditions.append('i.branch_id=%s'); params.append(bf)
    if ff:
        conditions.append('i.form_type_id=%s'); params.append(ff)

    where_clause = " AND ".join(conditions)

    rows = conn.execute(f'''
        SELECT i.*, f.name form_name, f.unit, f.unit_detail, f.unit_price, f.min_threshold,
               b.name branch_name, b.code branch_code, b.type branch_type
        FROM inventory i
        JOIN form_types f ON i.form_type_id = f.id
        JOIN branches b ON i.branch_id = b.id
        WHERE {where_clause}
        ORDER BY b.type, b.code, f.name
    ''', params).fetchall()

    summary_rows = conn.execute(f'''
        SELECT f.id form_type_id,
               f.name form_name, f.unit, f.unit_detail, f.unit_price, f.min_threshold,
               SUM(i.quantity)              total_qty,
               COUNT(DISTINCT i.branch_id) branch_cnt,
               SUM(CASE WHEN i.quantity = 0 THEN 1 ELSE 0 END)               empty_cnt,
               SUM(CASE WHEN i.quantity > 0 AND i.quantity <= f.min_threshold
                        THEN 1 ELSE 0 END)                                    low_cnt,
               SUM(i.quantity * f.unit_price)                                 total_value
        FROM inventory i
        JOIN form_types f ON i.form_type_id = f.id
        JOIN branches b ON i.branch_id = b.id
        WHERE {where_clause}
        GROUP BY f.id
        ORDER BY f.name
    ''', params).fetchall()

    grand_total_qty   = sum(r['total_qty']   or 0 for r in summary_rows)
    grand_total_value = sum(r['total_value'] or 0 for r in summary_rows)

    from datetime import date
    from collections import defaultdict
    today = date.today()
    month_labels = []
    y, m = today.year, today.month
    for _ in range(6):
        month_labels.insert(0, f'{y}-{m:02d}')
        m -= 1
        if m == 0: m, y = 12, y - 1

    if USE_SQLITE:
        _ph = ','.join(['?' for _ in month_labels])
        mo_conditions = [
            "t.type = 'OUT'",
            f"COALESCE(t.period_month, strftime('%Y-%m', t.created_at)) IN ({_ph})"
        ]
        mo_params = list(month_labels)
    else:
        mo_conditions = [
            "t.type = 'OUT'",
            f"COALESCE(t.period_month, TO_CHAR(t.created_at, 'YYYY-MM')) = ANY(%s)"
        ]
        mo_params = [month_labels]
    if session.get('role') != 'admin' and session.get('branch_id'):
        mo_conditions.append('t.from_branch_id = %s')
        mo_params.append(session['branch_id'])
    elif bf:
        mo_conditions.append('t.from_branch_id = %s')
        mo_params.append(bf)
    if ff:
        mo_conditions.append('t.form_type_id = %s')
        mo_params.append(ff)

    monthly_raw = conn.execute(f'''
        SELECT f.id form_type_id, f.name form_name, f.unit,
               COALESCE(t.period_month, TO_CHAR(t.created_at, 'YYYY-MM')) mo,
               SUM(t.quantity) qty
        FROM transactions t
        JOIN form_types f ON t.form_type_id = f.id
        WHERE {" AND ".join(mo_conditions)}
        GROUP BY f.id, COALESCE(t.period_month, TO_CHAR(t.created_at, 'YYYY-MM'))
        ORDER BY f.name, mo
    ''', mo_params).fetchall()

    monthly_pivot = defaultdict(lambda: defaultdict(int))
    monthly_meta = {}
    for r in monthly_raw:
        monthly_pivot[r['form_type_id']][r['mo']] += r['qty']
        monthly_meta[r['form_type_id']] = (r['form_name'], r['unit'])

    monthly_rows = []
    for fid_key in sorted(monthly_pivot, key=lambda x: monthly_meta[x][0]):
        fname, unit = monthly_meta[fid_key]
        vals = [monthly_pivot[fid_key].get(mo, 0) for mo in month_labels]
        monthly_rows.append({'form_name': fname, 'unit': unit,
                             'vals': vals, 'total': sum(vals)})
    monthly_totals = [sum(r['vals'][i] for r in monthly_rows) for i in range(6)]

    conn.close()
    return render_template('inventory.html',
                           rows=rows, summary_rows=summary_rows,
                           grand_total_qty=grand_total_qty,
                           grand_total_value=grand_total_value,
                           monthly_rows=monthly_rows,
                           monthly_totals=monthly_totals,
                           month_labels=month_labels,
                           branches=branches, form_types=form_types,
                           bf=bf, ff=ff)


@app.route('/inbound', methods=['GET', 'POST'])
@login_required
def inbound():
    conn = get_db()
    if request.method == 'POST':
        if session.get('role') != 'admin':
            if not session.get('branch_id'):
                flash('소속 지점이 없습니다. 관리자에게 문의하세요.', 'danger')
                conn.close()
                return redirect(url_for('dashboard'))
            bid = str(session['branch_id'])
        else:
            bid = request.form['branch_id']
        fid = request.form['form_type_id']
        qty = int(request.form['quantity'])
        notes = request.form.get('notes', '')

        conn.execute('''
            INSERT INTO inventory (branch_id, form_type_id, quantity, last_updated)
            VALUES (%s,%s,%s,NOW())
            ON CONFLICT(branch_id, form_type_id) DO UPDATE SET
              quantity = inventory.quantity + EXCLUDED.quantity,
              last_updated = NOW()
        ''', (bid, fid, qty))
        conn.execute(
            "INSERT INTO transactions (type, form_type_id, to_branch_id, quantity, notes, created_by) "
            "VALUES ('IN',%s,%s,%s,%s,%s)",
            (fid, bid, qty, notes, session['username'])
        )
        conn.commit()

        b = conn.execute('SELECT name FROM branches WHERE id=%s', (bid,)).fetchone()
        f = conn.execute('SELECT name FROM form_types WHERE id=%s', (fid,)).fetchone()
        flash(f'입고 완료 ✔ {b["name"]} — {f["name"]} {qty}개', 'success')
        conn.close()
        return redirect(url_for('inbound'))

    branches = conn.execute('SELECT * FROM branches ORDER BY type, code').fetchall()
    form_types = conn.execute('SELECT * FROM form_types ORDER BY name').fetchall()
    conn.close()
    return render_template('inbound.html', branches=branches, form_types=form_types,
                           selected_branch=session.get('branch_id'))


@app.route('/outbound', methods=['GET', 'POST'])
@login_required
def outbound():
    conn = get_db()
    if request.method == 'POST':
        if session.get('role') != 'admin':
            if not session.get('branch_id'):
                flash('소속 지점이 없습니다. 관리자에게 문의하세요.', 'danger')
                conn.close()
                return redirect(url_for('dashboard'))
            bid = str(session['branch_id'])
        else:
            bid = request.form['branch_id']
        fid = request.form['form_type_id']
        qty = int(request.form['quantity'])
        notes = request.form.get('notes', '')
        period_month = request.form.get('period_month', '')

        cur = conn.execute(
            'SELECT quantity FROM inventory WHERE branch_id=%s AND form_type_id=%s',
            (bid, fid)
        ).fetchone()
        if not cur or cur['quantity'] < qty:
            flash('재고가 부족합니다.', 'danger')
            conn.close()
            return redirect(url_for('outbound'))

        conn.execute(
            'UPDATE inventory SET quantity=quantity-%s, last_updated=NOW() WHERE branch_id=%s AND form_type_id=%s',
            (qty, bid, fid)
        )
        conn.execute(
            "INSERT INTO transactions (type, form_type_id, from_branch_id, quantity, notes, created_by, period_month) "
            "VALUES ('OUT',%s,%s,%s,%s,%s,%s)",
            (fid, bid, qty, notes, session['username'], period_month)
        )
        conn.commit()
        flash('출고 처리 완료 ✔', 'success')
        conn.close()
        return redirect(url_for('outbound'))

    branches = conn.execute('SELECT * FROM branches ORDER BY type, code').fetchall()
    form_types = conn.execute('SELECT * FROM form_types ORDER BY name').fetchall()
    conn.close()
    from datetime import date
    return render_template('outbound.html', branches=branches, form_types=form_types,
                           selected_branch=session.get('branch_id'),
                           today_ym=date.today().strftime('%Y-%m'))


@app.route('/transfer', methods=['GET', 'POST'])
@login_required
def transfer():
    conn = get_db()
    if request.method == 'POST':
        if session.get('role') != 'admin':
            if not session.get('branch_id'):
                flash('소속 지점이 없습니다. 관리자에게 문의하세요.', 'danger')
                conn.close()
                return redirect(url_for('dashboard'))
            from_bid = str(session['branch_id'])
        else:
            from_bid = request.form['from_branch_id']
        to_bid = request.form['to_branch_id']
        fid    = request.form['form_type_id']
        qty    = int(request.form['quantity'])
        notes  = request.form.get('notes', '')

        if from_bid == to_bid:
            flash('출발·도착 지점이 같습니다.', 'danger')
            conn.close()
            return redirect(url_for('transfer'))

        cur = conn.execute(
            'SELECT quantity FROM inventory WHERE branch_id=%s AND form_type_id=%s',
            (from_bid, fid)
        ).fetchone()
        if not cur or cur['quantity'] < qty:
            flash('출발 지점 재고가 부족합니다.', 'danger')
            conn.close()
            return redirect(url_for('transfer'))

        conn.execute(
            'UPDATE inventory SET quantity=quantity-%s, last_updated=NOW() WHERE branch_id=%s AND form_type_id=%s',
            (qty, from_bid, fid)
        )
        conn.execute('''
            INSERT INTO inventory (branch_id, form_type_id, quantity, last_updated)
            VALUES (%s,%s,%s,NOW())
            ON CONFLICT(branch_id, form_type_id) DO UPDATE SET
              quantity = inventory.quantity + EXCLUDED.quantity,
              last_updated = NOW()
        ''', (to_bid, fid, qty))
        conn.execute(
            "INSERT INTO transactions (type, form_type_id, from_branch_id, to_branch_id, quantity, notes, created_by) "
            "VALUES ('TRANSFER',%s,%s,%s,%s,%s,%s)",
            (fid, from_bid, to_bid, qty, notes, session['username'])
        )
        conn.commit()
        flash('이전 처리 완료 ✔', 'success')
        conn.close()
        return redirect(url_for('transfer'))

    branches = conn.execute('SELECT * FROM branches ORDER BY type, code').fetchall()
    form_types = conn.execute('SELECT * FROM form_types ORDER BY name').fetchall()
    conn.close()
    return render_template('transfer.html', branches=branches, form_types=form_types,
                           selected_branch=session.get('branch_id'))


@app.route('/transactions')
@login_required
def transactions():
    conn = get_db()
    branches   = conn.execute('SELECT * FROM branches ORDER BY type, code').fetchall()
    form_types = conn.execute('SELECT * FROM form_types ORDER BY name').fetchall()

    bf        = request.args.get('branch_id', '')
    ff        = request.args.get('form_type_id', '')
    tf        = request.args.get('type', '')
    date_from = request.args.get('date_from', '')
    date_to   = request.args.get('date_to', '')

    conditions, params = ['1=1'], []
    role = session.get('role')
    bid  = session.get('branch_id')
    if role != 'admin' and bid:
        conditions.append('(t.from_branch_id=%s OR t.to_branch_id=%s)')
        params.extend([bid, bid])
    elif bf:
        conditions.append('(t.from_branch_id=%s OR t.to_branch_id=%s)')
        params.extend([bf, bf])
    if ff:
        conditions.append('t.form_type_id=%s'); params.append(ff)
    if tf:
        conditions.append('t.type=%s'); params.append(tf)
    if date_from:
        conditions.append('t.created_at::date >= %s'); params.append(date_from)
    if date_to:
        conditions.append('t.created_at::date <= %s'); params.append(date_to)

    rows = conn.execute(f'''
        SELECT t.*, f.name form_name, f.unit,
               fb.name from_branch_name, fb.code from_branch_code,
               tb.name to_branch_name,   tb.code to_branch_code
        FROM transactions t
        JOIN form_types f ON t.form_type_id = f.id
        LEFT JOIN branches fb ON t.from_branch_id = fb.id
        LEFT JOIN branches tb ON t.to_branch_id   = tb.id
        WHERE {" AND ".join(conditions)}
        ORDER BY t.created_at DESC LIMIT 300
    ''', params).fetchall()
    conn.close()
    return render_template('transactions.html', rows=rows, branches=branches, form_types=form_types,
                           bf=bf, ff=ff, tf=tf, date_from=date_from, date_to=date_to)


@app.route('/report')
@login_required
def report():
    conn = get_db()
    bid  = session.get('branch_id')
    role = session.get('role')

    b_cond = f'AND b.id={bid}'          if role != 'admin' and bid else ''
    i_cond = f'AND i.branch_id={bid}'   if role != 'admin' and bid else ''
    t_cond = f'AND (t.from_branch_id={bid} OR t.to_branch_id={bid})' if role != 'admin' and bid else ''

    summary = conn.execute(f'''
        SELECT b.code, b.name, b.type,
               COUNT(CASE WHEN i.quantity=0 THEN 1 END) empty_cnt,
               COUNT(CASE WHEN i.quantity>0 AND i.quantity<=f.min_threshold THEN 1 END) low_cnt,
               COUNT(CASE WHEN i.quantity>f.min_threshold THEN 1 END) ok_cnt,
               COALESCE(SUM(i.quantity*f.unit_price),0) total_value
        FROM branches b
        LEFT JOIN inventory i  ON b.id=i.branch_id
        LEFT JOIN form_types f ON i.form_type_id=f.id
        WHERE 1=1 {b_cond}
        GROUP BY b.id, b.code, b.name, b.type ORDER BY b.type, b.code
    ''').fetchall()

    monthly_raw = conn.execute(f'''
        SELECT TO_CHAR(t.created_at, 'YYYY-MM') mo, t.type,
               COUNT(*) tx_count, COALESCE(SUM(t.quantity),0) total_qty
        FROM transactions t
        WHERE t.created_at >= CURRENT_DATE - INTERVAL '6 months' {t_cond}
        GROUP BY TO_CHAR(t.created_at, 'YYYY-MM'), t.type ORDER BY TO_CHAR(t.created_at, 'YYYY-MM') ASC
    ''').fetchall()

    months_set = sorted({r['mo'] for r in monthly_raw})
    def _monthly_vals(tx_type):
        m = {r['mo']: r['total_qty'] for r in monthly_raw if r['type'] == tx_type}
        return [m.get(mo, 0) for mo in months_set]
    chart_monthly = {
        'labels':   months_set,
        'in':       _monthly_vals('IN'),
        'out':      _monthly_vals('OUT'),
        'transfer': _monthly_vals('TRANSFER'),
    }

    status_row = conn.execute(f'''
        SELECT COUNT(CASE WHEN i.quantity=0 THEN 1 END) empty_cnt,
               COUNT(CASE WHEN i.quantity>0 AND i.quantity<=f.min_threshold THEN 1 END) low_cnt,
               COUNT(CASE WHEN i.quantity>f.min_threshold THEN 1 END) ok_cnt
        FROM inventory i
        JOIN form_types f ON i.form_type_id=f.id
        WHERE 1=1 {i_cond}
    ''').fetchone()
    chart_status = {
        'ok':    status_row['ok_cnt']    or 0,
        'low':   status_row['low_cnt']   or 0,
        'empty': status_row['empty_cnt'] or 0,
    }

    chart_branch_qty = conn.execute(f'''
        SELECT b.code label, b.type btype, COALESCE(SUM(i.quantity),0) qty
        FROM branches b
        LEFT JOIN inventory i ON b.id=i.branch_id
        WHERE 1=1 {b_cond}
        GROUP BY b.id, b.code, b.type HAVING COALESCE(SUM(i.quantity),0)>0 ORDER BY qty DESC LIMIT 20
    ''').fetchall()

    chart_branch_status = conn.execute(f'''
        SELECT b.code label,
               COALESCE(SUM(CASE WHEN i.quantity=0 THEN 1 ELSE 0 END),0) empty_cnt,
               COALESCE(SUM(CASE WHEN i.quantity>0 AND i.quantity<=f.min_threshold THEN 1 ELSE 0 END),0) low_cnt,
               COALESCE(SUM(CASE WHEN i.quantity>f.min_threshold THEN 1 ELSE 0 END),0) ok_cnt
        FROM branches b
        LEFT JOIN inventory i  ON b.id=i.branch_id
        LEFT JOIN form_types f ON i.form_type_id=f.id
        WHERE 1=1 {b_cond}
        GROUP BY b.id, b.code, b.type
        HAVING (COALESCE(SUM(CASE WHEN i.quantity=0 THEN 1 ELSE 0 END),0) +
                COALESCE(SUM(CASE WHEN i.quantity>0 AND i.quantity<=f.min_threshold THEN 1 ELSE 0 END),0) +
                COALESCE(SUM(CASE WHEN i.quantity>f.min_threshold THEN 1 ELSE 0 END),0)) > 0
        ORDER BY b.type, b.code
    ''').fetchall()

    chart_top_out = conn.execute(f'''
        SELECT f.name label, COALESCE(SUM(t.quantity),0) qty
        FROM transactions t
        JOIN form_types f ON t.form_type_id=f.id
        WHERE t.type='OUT' AND t.created_at >= CURRENT_DATE - INTERVAL '30 days' {t_cond}
        GROUP BY f.id, f.name ORDER BY qty DESC LIMIT 10
    ''').fetchall()

    chart_form_qty = conn.execute(f'''
        SELECT f.name label, COALESCE(SUM(i.quantity),0) qty
        FROM form_types f
        LEFT JOIN inventory i ON f.id=i.form_type_id
        WHERE 1=1 {i_cond}
        GROUP BY f.id, f.name HAVING COALESCE(SUM(i.quantity),0)>0 ORDER BY qty DESC LIMIT 20
    ''').fetchall()

    conn.close()

    def _rows_to_chart(rows):
        return {'labels': [r['label'] for r in rows],
                'data':   [r['qty']   for r in rows]}

    return render_template('report.html',
        summary=summary, monthly=monthly_raw,
        chart_monthly      = chart_monthly,
        chart_status       = chart_status,
        chart_branch_qty   = _rows_to_chart(chart_branch_qty),
        chart_branch_status= {'labels': [r['label'] for r in chart_branch_status],
                               'ok':    [r['ok_cnt']    for r in chart_branch_status],
                               'low':   [r['low_cnt']   for r in chart_branch_status],
                               'empty': [r['empty_cnt'] for r in chart_branch_status]},
        chart_top_out      = _rows_to_chart(chart_top_out),
        chart_form_qty     = _rows_to_chart(chart_form_qty),
    )


# ── 분기 신청 추정 ────────────────────────────────────────────────────────────

@app.route('/report/forecast')
@login_required
def forecast():
    import math
    from datetime import date
    from collections import defaultdict

    conn  = get_db()
    today = date.today()
    role  = session.get('role')
    bid   = session.get('branch_id')
    bf    = request.args.get('branch_id', '')

    cur_q        = (today.month - 1) // 3 + 1
    next_q       = cur_q % 4 + 1
    nq_y         = today.year + (1 if next_q == 1 else 0)
    nq_start_m   = (cur_q * 3) % 12 + 1
    nq_end_m     = nq_start_m + 2
    is_qend      = today.month in (3, 6, 9, 12)

    branches = conn.execute('SELECT * FROM branches ORDER BY type, code').fetchall()

    if role != 'admin':
        target_ids = [bid] if bid else []
    elif bf:
        target_ids = [int(bf)]
    else:
        target_ids = [b['id'] for b in branches]

    if not target_ids:
        conn.close()
        return render_template('forecast.html',
            forecast_sections=[], month_labels=[], branches=branches,
            today=today, cur_q=cur_q, next_q=next_q,
            nq_y=nq_y, nq_start_m=nq_start_m, nq_end_m=nq_end_m,
            is_qend=is_qend, bf=bf)

    id_str = ','.join(str(x) for x in target_ids)

    month_labels = []
    y, m = today.year, today.month - 1
    if m == 0: m, y = 12, y - 1
    for _ in range(6):
        month_labels.insert(0, f'{y}-{m:02d}')
        m -= 1
        if m == 0: m, y = 12, y - 1

    raw = conn.execute(f'''
        SELECT t.from_branch_id bid, f.id fid,
               f.name form_name, f.unit, f.unit_detail,
               COALESCE(t.period_month, TO_CHAR(t.created_at, 'YYYY-MM')) month,
               SUM(t.quantity) qty
        FROM transactions t
        JOIN form_types f ON t.form_type_id = f.id
        WHERE t.type = 'OUT'
          AND t.from_branch_id IN ({id_str})
          AND COALESCE(t.period_month, TO_CHAR(t.created_at, 'YYYY-MM'))
              BETWEEN TO_CHAR(CURRENT_DATE - INTERVAL '6 months', 'YYYY-MM')
                  AND TO_CHAR(CURRENT_DATE, 'YYYY-MM')
        GROUP BY t.from_branch_id, f.id, f.name, f.unit, f.unit_detail, COALESCE(t.period_month, TO_CHAR(t.created_at, 'YYYY-MM'))
    ''').fetchall()

    inv_map = {(r['branch_id'], r['form_type_id']): r['quantity']
               for r in conn.execute(
                   f'SELECT branch_id, form_type_id, quantity FROM inventory WHERE branch_id IN ({id_str})'
               ).fetchall()}

    branch_map = {b['id']: dict(b) for b in branches}

    bucket = defaultdict(lambda: defaultdict(int))
    meta   = {}
    for r in raw:
        key = (r['bid'], r['fid'])
        bucket[key][r['month']] += r['qty']
        meta[key] = (r['form_name'], r['unit'], r['unit_detail'])

    forecast_by_branch = defaultdict(list)
    for (b_id, f_id), monthly in bucket.items():
        vals     = [monthly.get(mo, 0) for mo in month_labels]
        total_6m = sum(vals)
        nonzero  = [v for v in vals if v > 0]
        avg_m    = total_6m / len(nonzero) if nonzero else 0

        r3 = sum(vals[3:]) / 3
        o3 = sum(vals[:3]) / 3
        if o3 > 0:
            trend_pct = (r3 - o3) / o3
        elif r3 > 0:
            trend_pct = 1.0
        else:
            trend_pct = 0.0

        if trend_pct > 0.15:
            trend_label, trend_icon, trend_factor = '증가', '↑', 1.15
        elif trend_pct < -0.15:
            trend_label, trend_icon, trend_factor = '감소', '↓', 0.9
        else:
            trend_label, trend_icon, trend_factor = '안정', '→', 1.0

        cur_stock   = inv_map.get((b_id, f_id), 0)
        nq_est      = math.ceil(avg_m * 3 * trend_factor)
        recommended = max(0, nq_est - cur_stock)

        fname, unit, udesc = meta.get((b_id, f_id), ('', '', ''))
        forecast_by_branch[b_id].append(dict(
            form_name=fname, unit=unit, unit_detail=udesc,
            vals=vals, total_6m=total_6m,
            avg_monthly=round(avg_m, 1),
            trend_pct=round(trend_pct * 100),
            trend_label=trend_label, trend_icon=trend_icon,
            cur_stock=cur_stock,
            nq_est=nq_est,
            recommended=recommended,
        ))

    for b_id in forecast_by_branch:
        forecast_by_branch[b_id].sort(key=lambda x: x['recommended'], reverse=True)

    forecast_sections = [
        dict(branch_id=b_id,
             branch_code=branch_map[b_id]['code'],
             branch_name=branch_map[b_id]['name'],
             branch_type=branch_map[b_id]['type'],
             form_items=forecast_by_branch[b_id],
             col_sums=[sum(it['vals'][i] for it in forecast_by_branch[b_id]) for i in range(6)])
        for b_id in sorted(forecast_by_branch)
        if forecast_by_branch[b_id]
    ]

    conn.close()
    return render_template('forecast.html',
        forecast_sections=forecast_sections,
        month_labels=month_labels,
        branches=branches, bf=bf,
        today=today, cur_q=cur_q, next_q=next_q,
        nq_y=nq_y, nq_start_m=nq_start_m, nq_end_m=nq_end_m,
        is_qend=is_qend)


@app.route('/report/forecast/download')
@login_required
def forecast_download():
    import math
    from datetime import date
    from collections import defaultdict

    conn  = get_db()
    today = date.today()
    role  = session.get('role')
    bid   = session.get('branch_id')
    bf    = request.args.get('branch_id', '')

    cur_q      = (today.month - 1) // 3 + 1
    next_q     = cur_q % 4 + 1
    nq_y       = today.year + (1 if next_q == 1 else 0)
    nq_start_m = (cur_q * 3) % 12 + 1

    branches = conn.execute('SELECT * FROM branches ORDER BY type, code').fetchall()
    if role != 'admin':
        target_ids = [bid] if bid else []
    elif bf:
        target_ids = [int(bf)]
    else:
        target_ids = [b['id'] for b in branches]

    if not target_ids:
        conn.close()
        flash('데이터 없음', 'warning')
        return redirect(url_for('forecast'))

    id_str = ','.join(str(x) for x in target_ids)

    month_labels = []
    y, m = today.year, today.month - 1
    if m == 0: m, y = 12, y - 1
    for _ in range(6):
        month_labels.insert(0, f'{y}-{m:02d}')
        m -= 1
        if m == 0: m, y = 12, y - 1

    raw = conn.execute(f'''
        SELECT t.from_branch_id bid, f.id fid, f.name form_name, f.unit,
               COALESCE(t.period_month, TO_CHAR(t.created_at, 'YYYY-MM')) month,
               SUM(t.quantity) qty
        FROM transactions t JOIN form_types f ON t.form_type_id=f.id
        WHERE t.type='OUT' AND t.from_branch_id IN ({id_str})
          AND COALESCE(t.period_month, TO_CHAR(t.created_at, 'YYYY-MM'))
              BETWEEN TO_CHAR(CURRENT_DATE - INTERVAL '6 months', 'YYYY-MM')
                  AND TO_CHAR(CURRENT_DATE, 'YYYY-MM')
        GROUP BY t.from_branch_id, f.id, f.name, f.unit, COALESCE(t.period_month, TO_CHAR(t.created_at, 'YYYY-MM'))
    ''').fetchall()

    inv_map = {(r['branch_id'], r['form_type_id']): r['quantity']
               for r in conn.execute(
                   f'SELECT branch_id, form_type_id, quantity FROM inventory WHERE branch_id IN ({id_str})'
               ).fetchall()}

    branch_map = {b['id']: dict(b) for b in branches}

    bucket = defaultdict(lambda: defaultdict(int))
    meta   = {}
    for r in raw:
        key = (r['bid'], r['fid'])
        bucket[key][r['month']] += r['qty']
        meta[key] = (r['form_name'], r['unit'])

    is_admin = role == 'admin'
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f'{nq_y}년 Q{next_q} 신청 추정'
    hfill, hfont, halign = _make_header_style()
    bd = _border()

    headers = (['지점코드','지점명'] if is_admin else []) + \
              ['양식명','단위'] + [f'{mo}출고' for mo in month_labels] + \
              ['월평균','추세','현재재고','다음분기예상','권장신청수량']
    col_widths = ([10,14] if is_admin else []) + \
                 [34,8] + [10]*6 + [10,8,10,12,12]
    for col,(h,w) in enumerate(zip(headers,col_widths),1):
        c = ws.cell(row=1,column=col,value=h)
        c.fill,c.font,c.alignment = hfill,hfont,halign
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.row_dimensions[1].height = 22

    green_f = PatternFill('solid',fgColor='D1FAE5')
    yellow_f= PatternFill('solid',fgColor='FEF9C3')
    red_f   = PatternFill('solid',fgColor='FEE2E2')

    row_num = 2
    for b_id in sorted(bucket, key=lambda x: branch_map.get(x, {}).get('code', '')):
        b = branch_map.get(b_id, {})
        for (bb,fid),monthly in {k:v for k,v in bucket.items() if k[0]==b_id}.items():
            fname,unit = meta.get((bb,fid),('',''))
            vals   = [monthly.get(mo,0) for mo in month_labels]
            nonzero= [v for v in vals if v > 0]
            avg_m  = sum(vals)/len(nonzero) if nonzero else 0
            r3,o3  = sum(vals[3:])/3, sum(vals[:3])/3
            tp     = (r3-o3)/o3 if o3>0 else (1.0 if r3>0 else 0.0)
            tf     = 1.15 if tp>0.15 else (0.9 if tp<-0.15 else 1.0)
            tl     = '증가' if tp>0.15 else ('감소' if tp<-0.15 else '안정')
            cur_s  = inv_map.get((bb,fid),0)
            nq_est = math.ceil(avg_m*3*tf)
            rec    = max(0, nq_est-cur_s)

            row_vals = ([b.get('code',''), b.get('name','')] if is_admin else []) + \
                       [fname, unit] + vals + \
                       [round(avg_m,1), tl, cur_s, nq_est, rec]
            fill = red_f if rec>5 else (yellow_f if rec>0 else green_f)
            for col,v in enumerate(row_vals,1):
                c = ws.cell(row=row_num,column=col,value=v)
                c.border = bd
                c.alignment = Alignment(vertical='center',
                    horizontal='center' if col>=(3+(2 if is_admin else 0)) else 'left')
            last_col = len(row_vals)
            ws.cell(row=row_num,column=last_col).fill = fill
            ws.row_dimensions[row_num].height = 16
            row_num += 1

    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = f'A1:{get_column_letter(len(headers))}1'
    conn.close()

    buf = BytesIO(); wb.save(buf); buf.seek(0)
    fname_dl = f'{nq_y}년_Q{next_q}_신청추정_{today.strftime("%Y%m%d")}.xlsx'
    return send_file(buf, as_attachment=True, download_name=fname_dl,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/api/stock')
@login_required
def api_stock():
    bid = request.args.get('branch_id')
    fid = request.args.get('form_type_id')
    conn = get_db()
    row = conn.execute(
        'SELECT quantity FROM inventory WHERE branch_id=%s AND form_type_id=%s',
        (bid, fid)
    ).fetchone()
    conn.close()
    return jsonify({'quantity': row['quantity'] if row else 0})


# ── 사용자 관리 (관리자 전용) ─────────────────────────────────────────────────

@app.route('/admin/users')
@login_required
def manage_users():
    if session.get('role') != 'admin':
        flash('관리자 권한이 필요합니다.', 'danger')
        return redirect(url_for('dashboard'))
    conn = get_db()
    users    = conn.execute('''
        SELECT u.*, b.name branch_name, b.code branch_code
        FROM users u LEFT JOIN branches b ON u.branch_id = b.id
        ORDER BY u.role DESC, u.username
    ''').fetchall()
    branches = conn.execute('SELECT * FROM branches ORDER BY type, code').fetchall()
    conn.close()
    return render_template('users.html', users=users, branches=branches)


@app.route('/admin/users/create', methods=['POST'])
@login_required
def create_user():
    if session.get('role') != 'admin':
        return jsonify({'ok': False, 'error': '권한 없음'}), 403

    username  = request.form.get('username', '').strip()
    password  = request.form.get('password', '').strip()
    branch_id = request.form.get('branch_id') or None
    role      = request.form.get('role', 'staff')

    if not username or not password:
        flash('아이디와 비밀번호를 입력하세요.', 'danger')
        return redirect(url_for('manage_users'))
    if len(password) < 6:
        flash('비밀번호는 6자 이상이어야 합니다.', 'danger')
        return redirect(url_for('manage_users'))

    conn = get_db()
    try:
        conn.execute(
            'INSERT INTO users (username, password, branch_id, role) VALUES (%s,%s,%s,%s)',
            (username, hash_pw(password), branch_id, role)
        )
        conn.commit()
        flash(f'계정 [{username}] 이(가) 생성되었습니다.', 'success')
    except (psycopg2.errors.UniqueViolation, sqlite3.IntegrityError):
        conn.rollback()
        flash(f'이미 존재하는 아이디입니다: {username}', 'danger')
    finally:
        conn.close()
    return redirect(url_for('manage_users'))


@app.route('/admin/users/<int:uid>/delete', methods=['POST'])
@login_required
def delete_user(uid):
    if session.get('role') != 'admin':
        return jsonify({'ok': False}), 403
    if uid == session.get('user_id'):
        flash('본인 계정은 삭제할 수 없습니다.', 'danger')
        return redirect(url_for('manage_users'))
    conn = get_db()
    conn.execute('DELETE FROM users WHERE id=%s', (uid,))
    conn.commit()
    conn.close()
    flash('계정이 삭제되었습니다.', 'success')
    return redirect(url_for('manage_users'))


@app.route('/admin/users/<int:uid>/reset-password', methods=['POST'])
@login_required
def reset_user_password(uid):
    if session.get('role') != 'admin':
        return jsonify({'ok': False}), 403
    new_pw = request.form.get('new_password', '').strip()
    if len(new_pw) < 6:
        flash('비밀번호는 6자 이상이어야 합니다.', 'danger')
        return redirect(url_for('manage_users'))
    conn = get_db()
    conn.execute('UPDATE users SET password=%s WHERE id=%s', (hash_pw(new_pw), uid))
    conn.commit()
    conn.close()
    flash('비밀번호가 초기화되었습니다.', 'success')
    return redirect(url_for('manage_users'))


# ── 비밀번호 변경 (본인) ───────────────────────────────────────────────────────

@app.route('/change-password', methods=['GET', 'POST'])
@login_required
def change_password():
    if request.method == 'POST':
        cur_pw  = request.form.get('current_password', '').strip()
        new_pw  = request.form.get('new_password', '').strip()
        conf_pw = request.form.get('confirm_password', '').strip()

        conn = get_db()
        user = conn.execute('SELECT * FROM users WHERE id=%s', (session['user_id'],)).fetchone()

        if user['password'] != hash_pw(cur_pw):
            flash('현재 비밀번호가 올바르지 않습니다.', 'danger')
            conn.close()
            return redirect(url_for('change_password'))
        if len(new_pw) < 6:
            flash('새 비밀번호는 6자 이상이어야 합니다.', 'danger')
            conn.close()
            return redirect(url_for('change_password'))
        if new_pw != conf_pw:
            flash('새 비밀번호가 일치하지 않습니다.', 'danger')
            conn.close()
            return redirect(url_for('change_password'))

        conn.execute('UPDATE users SET password=%s WHERE id=%s', (hash_pw(new_pw), session['user_id']))
        conn.commit()
        conn.close()
        flash('비밀번호가 변경되었습니다.', 'success')
        return redirect(url_for('dashboard'))

    return render_template('change_password.html')


# ── 입고 엑셀 템플릿 다운로드 ────────────────────────────────────────────────

@app.route('/inbound/template')
@login_required
def inbound_template():
    conn = get_db()
    branches   = conn.execute('SELECT code, name, type FROM branches ORDER BY type, code').fetchall()
    form_types = conn.execute('SELECT name, unit, unit_detail FROM form_types ORDER BY name').fetchall()

    my_branch_code = ''
    if session.get('role') != 'admin' and session.get('branch_id'):
        row = conn.execute('SELECT code FROM branches WHERE id=%s', (session['branch_id'],)).fetchone()
        my_branch_code = row['code'] if row else ''
    conn.close()

    wb  = openpyxl.Workbook()
    ws  = wb.active
    ws.title = '입고업로드'
    hfill, hfont, halign = _make_header_style()
    bd  = _border()

    headers    = ['지점코드', '양식명', '수량', '비고']
    col_widths = [14, 38, 10, 26]
    for col, (h, w) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill, cell.font, cell.alignment = hfill, hfont, halign
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.row_dimensions[1].height = 22

    guide = '※ 지점코드·양식명은 드롭다운 또는 "참고" 시트를 확인하세요. 수량은 BOX/포대/권 단위입니다.'
    ws.merge_cells('A2:D2')
    c = ws.cell(row=2, column=1, value=guide)
    c.font      = Font(color='5A6A8A', size=9, italic=True)
    c.alignment = Alignment(horizontal='left', vertical='center')
    c.fill      = PatternFill('solid', fgColor='EEF2FF')
    ws.row_dimensions[2].height = 18

    is_staff   = session.get('role') != 'admin'
    gray_fill  = PatternFill('solid', fgColor='F1F5F9')
    data_start = 3

    for r in range(data_start, data_start + 50):
        for col in range(1, 5):
            cell = ws.cell(row=r, column=col)
            cell.border    = bd
            cell.alignment = Alignment(vertical='center',
                                       horizontal='center' if col == 3 else 'left')
        if is_staff and my_branch_code:
            c = ws.cell(row=r, column=1, value=my_branch_code)
            c.fill      = gray_fill
            c.font      = Font(color='374151', bold=True)
            c.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[r].height = 17

    last_branch = len(branches) + 1
    last_form   = len(form_types) + 1
    sqref       = f'A{data_start}:A{data_start+49}'
    sqref_form  = f'B{data_start}:B{data_start+49}'

    if not is_staff:
        dv_b = DataValidation(type='list',
                              formula1=f"'참고'!$A$2:$A${last_branch}",
                              allow_blank=True,
                              showErrorMessage=True,
                              error='참고 시트의 지점코드 목록에서 선택하세요.',
                              errorTitle='잘못된 지점코드')
        dv_b.sqref = sqref
        ws.add_data_validation(dv_b)

    dv_f = DataValidation(type='list',
                          formula1=f"'참고'!$C$2:$C${last_form}",
                          allow_blank=True,
                          showErrorMessage=True,
                          error='참고 시트의 양식명 목록에서 선택하세요.',
                          errorTitle='잘못된 양식명')
    dv_f.sqref = sqref_form
    ws.add_data_validation(dv_f)

    ws.freeze_panes = 'A3'

    ws2 = wb.create_sheet('참고')
    ref_headers = [('A', '지점코드', 12), ('B', '지점명', 16), ('C', '양식명', 38),
                   ('D', '단위', 10), ('E', '단위상세', 14)]
    for col_letter, h, w in ref_headers:
        c = ws2[f'{col_letter}1']
        c.value, c.fill, c.font, c.alignment = h, hfill, hfont, halign
        ws2.column_dimensions[col_letter].width = w
    ws2.row_dimensions[1].height = 22

    type_kr = {'DOM': '국내', 'INTL': '국제', 'CARGO': '화물'}
    for i, b in enumerate(branches, 2):
        ws2.cell(row=i, column=1, value=b['code']).alignment = Alignment(horizontal='center')
        type_label = type_kr.get(b['type'], b['type'])
        c = ws2.cell(row=i, column=2, value=f"[{type_label}] {b['name']}")
        c.alignment = Alignment(horizontal='left')

    for i, f in enumerate(form_types, 2):
        ws2.cell(row=i, column=3, value=f['name'])
        ws2.cell(row=i, column=4, value=f['unit']).alignment = Alignment(horizontal='center')
        ws2.cell(row=i, column=5, value=f['unit_detail']).alignment = Alignment(horizontal='center')

    ws2.freeze_panes = 'A2'

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    from datetime import date
    suffix = f'_{my_branch_code}' if my_branch_code else ''
    fname  = f'입고업로드_템플릿{suffix}_{date.today().strftime("%Y%m%d")}.xlsx'
    return send_file(buf, as_attachment=True, download_name=fname,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# ── 입고 엑셀 업로드 처리 ──────────────────────────────────────────────────────

@app.route('/inbound/upload', methods=['POST'])
@login_required
def inbound_upload():
    f = request.files.get('file')
    if not f or not f.filename:
        flash('파일을 선택해주세요.', 'danger')
        return redirect(url_for('inbound'))
    if not f.filename.lower().endswith(('.xlsx', '.xls')):
        flash('xlsx 또는 xls 파일만 업로드 가능합니다.', 'danger')
        return redirect(url_for('inbound'))

    try:
        wb = openpyxl.load_workbook(f, data_only=True)
        ws = wb.active
    except Exception as e:
        flash(f'파일을 읽을 수 없습니다: {e}', 'danger')
        return redirect(url_for('inbound'))

    conn = get_db()
    branch_map = {r['code'].upper(): r['id']
                  for r in conn.execute('SELECT code, id FROM branches').fetchall()}
    form_map   = {r['name']: r['id']
                  for r in conn.execute('SELECT name, id FROM form_types').fetchall()}

    is_staff    = session.get('role') != 'admin'
    forced_bid  = None
    forced_code = ''
    if is_staff:
        if not session.get('branch_id'):
            flash('소속 지점이 없습니다. 관리자에게 문의하세요.', 'danger')
            conn.close()
            return redirect(url_for('inbound'))
        forced_bid  = session['branch_id']
        row = conn.execute('SELECT code FROM branches WHERE id=%s', (forced_bid,)).fetchone()
        forced_code = row['code'] if row else ''

    results = []
    ok_cnt  = 0

    all_rows = list(ws.iter_rows(values_only=True))
    data_rows = []
    for i, row in enumerate(all_rows):
        first = str(row[0]).strip() if row[0] is not None else ''
        if first in ('지점코드', '※ 지점코드·양식명은 드롭다운') or first.startswith('※'):
            continue
        if not any(row[:3]):
            continue
        data_rows.append((i + 1, row))

    for row_num, row in data_rows:
        raw_code = str(row[0]).strip().upper() if row[0] is not None else ''
        raw_form = str(row[1]).strip()         if row[1] is not None else ''
        raw_qty  = row[2]
        notes    = str(row[3]).strip()         if row[3] is not None else ''

        entry = {'row': row_num, 'branch': raw_code or forced_code,
                 'form': raw_form, 'qty': raw_qty}

        if is_staff:
            bid        = forced_bid
            entry['branch'] = forced_code
        else:
            if not raw_code:
                entry.update(status='skip', msg='지점코드 없음'); results.append(entry); continue
            bid = branch_map.get(raw_code)
            if not bid:
                entry.update(status='error', msg=f'지점코드 없음: {raw_code}')
                results.append(entry); continue

        if not raw_form:
            entry.update(status='skip', msg='양식명 없음'); results.append(entry); continue
        fid = form_map.get(raw_form)
        if not fid:
            entry.update(status='error', msg=f'양식명 불일치: {raw_form}')
            results.append(entry); continue

        try:
            qty = int(float(str(raw_qty)))
        except (ValueError, TypeError):
            entry.update(status='error', msg=f'수량 오류: {raw_qty}')
            results.append(entry); continue
        if qty <= 0:
            entry.update(status='skip', msg=f'수량 0 이하 ({qty})'); results.append(entry); continue

        conn.execute('''
            INSERT INTO inventory (branch_id, form_type_id, quantity, last_updated)
            VALUES (%s,%s,%s,NOW())
            ON CONFLICT(branch_id, form_type_id) DO UPDATE SET
              quantity     = inventory.quantity + EXCLUDED.quantity,
              last_updated = NOW()
        ''', (bid, fid, qty))
        conn.execute(
            "INSERT INTO transactions (type, form_type_id, to_branch_id, quantity, notes, created_by) "
            "VALUES ('IN',%s,%s,%s,%s,%s)",
            (fid, bid, qty,
             f'[엑셀업로드] {notes}' if notes else '[엑셀업로드]',
             session['username'])
        )

        entry.update(status='ok', qty=qty)
        results.append(entry)
        ok_cnt += 1

    conn.commit()
    conn.close()

    err_cnt  = sum(1 for r in results if r['status'] == 'error')
    skip_cnt = sum(1 for r in results if r['status'] == 'skip')
    return render_template('upload_result.html',
                           results=results, ok_cnt=ok_cnt,
                           err_cnt=err_cnt, skip_cnt=skip_cnt)


# ── 보유 지점 상세 API ────────────────────────────────────────────────────────

@app.route('/api/inventory/branches')
@login_required
def api_inventory_branches():
    fid = request.args.get('form_type_id')
    bf  = request.args.get('branch_id', '')
    if not fid:
        return jsonify([])

    conn = get_db()
    conditions, params = ['i.form_type_id=%s'], [fid]

    if session.get('role') != 'admin' and session.get('branch_id'):
        conditions.append('i.branch_id=%s')
        params.append(session['branch_id'])
    elif bf:
        conditions.append('i.branch_id=%s')
        params.append(bf)

    rows = conn.execute(f'''
        SELECT b.code branch_code, b.name branch_name, b.type branch_type,
               i.quantity, i.last_updated, f.min_threshold, f.name form_name
        FROM inventory i
        JOIN branches b ON i.branch_id = b.id
        JOIN form_types f ON i.form_type_id = f.id
        WHERE {" AND ".join(conditions)}
        ORDER BY b.type, b.code
    ''', params).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


# ── 최소기준 수정 ─────────────────────────────────────────────────────────────

@app.route('/api/update_threshold', methods=['POST'])
@login_required
def update_threshold():
    data = request.get_json()
    form_type_id = data.get('form_type_id')
    threshold    = data.get('threshold')
    if form_type_id is None or threshold is None or int(threshold) < 0:
        return jsonify({'ok': False, 'error': '잘못된 값'}), 400
    conn = get_db()
    conn.execute('UPDATE form_types SET min_threshold=%s WHERE id=%s', (int(threshold), form_type_id))
    conn.commit()
    conn.close()
    return jsonify({'ok': True})


# ── Excel 헬퍼 ────────────────────────────────────────────────────────────────

def _make_header_style():
    fill   = PatternFill('solid', fgColor='1A2340')
    font   = Font(color='FFFFFF', bold=True, size=10)
    align  = Alignment(horizontal='center', vertical='center')
    return fill, font, align

def _border():
    s = Side(style='thin', color='D0D7E3')
    return Border(left=s, right=s, top=s, bottom=s)

def _apply_header(ws, headers, col_widths):
    fill, font, align = _make_header_style()
    for col, (h, w) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill, cell.font, cell.alignment = fill, font, align
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.row_dimensions[1].height = 22

def _cell_border(ws, row, col):
    ws.cell(row=row, column=col).border = _border()


# ── 재고현황 Excel 다운로드 ───────────────────────────────────────────────────

@app.route('/inventory/download')
@login_required
def inventory_download():
    conn = get_db()
    bf, ff = request.args.get('branch_id',''), request.args.get('form_type_id','')
    conditions, params = ['1=1'], []

    if session.get('role') != 'admin' and session.get('branch_id'):
        conditions.append('i.branch_id=%s'); params.append(session['branch_id'])
    elif bf:
        conditions.append('i.branch_id=%s'); params.append(bf)
    if ff:
        conditions.append('i.form_type_id=%s'); params.append(ff)

    rows = conn.execute(f'''
        SELECT i.*, f.name form_name, f.unit, f.unit_detail, f.unit_price, f.min_threshold,
               b.name branch_name, b.code branch_code, b.type branch_type
        FROM inventory i
        JOIN form_types f ON i.form_type_id = f.id
        JOIN branches b ON i.branch_id = b.id
        WHERE {" AND ".join(conditions)}
        ORDER BY b.type, b.code, f.name
    ''', params).fetchall()
    conn.close()

    is_admin = session.get('role') == 'admin'

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '재고현황'

    if is_admin:
        headers    = ['구분','지점코드','지점명','양식명','단위','단위상세','현재수량','최소기준','상태','단가(원)','재고금액(원)','최종업데이트']
        col_widths = [8, 10, 14, 32, 8, 12, 10, 10, 8, 12, 14, 20]
    else:
        headers    = ['구분','지점코드','지점명','양식명','단위','단위상세','현재수량','최소기준','상태','최종업데이트']
        col_widths = [8, 10, 14, 32, 8, 12, 10, 10, 8, 20]
    _apply_header(ws, headers, col_widths)

    type_map = {'DOM':'국내','INTL':'국제','CARGO':'화물'}
    green  = PatternFill('solid', fgColor='D1FAE5')
    yellow = PatternFill('solid', fgColor='FEF9C3')
    red    = PatternFill('solid', fgColor='FEE2E2')

    for r, row in enumerate(rows, 2):
        qty = row['quantity']
        thr = row['min_threshold']
        status = '소진' if qty == 0 else ('부족' if qty <= thr else '정상')
        color  = red    if qty == 0 else (yellow if qty <= thr else green)
        if is_admin:
            vals = [
                type_map.get(row['branch_type'], row['branch_type']),
                row['branch_code'], row['branch_name'], row['form_name'],
                row['unit'], row['unit_detail'], qty, thr, status,
                row['unit_price'], qty * (row['unit_price'] or 0),
                str(row['last_updated'] or '')[:16],
            ]
        else:
            vals = [
                type_map.get(row['branch_type'], row['branch_type']),
                row['branch_code'], row['branch_name'], row['form_name'],
                row['unit'], row['unit_detail'], qty, thr, status,
                str(row['last_updated'] or '')[:16],
            ]
        status_col = 9
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border    = _border()
            cell.alignment = Alignment(vertical='center')
            if c in (7, 8):
                cell.alignment = Alignment(horizontal='center', vertical='center')
            if c == status_col:
                cell.fill = color
                cell.alignment = Alignment(horizontal='center', vertical='center')
            if is_admin and c in (10, 11):
                cell.alignment = Alignment(horizontal='right', vertical='center')
                cell.number_format = '#,##0'
        ws.row_dimensions[r].height = 16

    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = f'A1:{get_column_letter(len(headers))}1'

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    from datetime import date
    fname = f'재고현황_{date.today().strftime("%Y%m%d")}.xlsx'
    return send_file(buf, as_attachment=True, download_name=fname,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# ── 입출고이력 Excel 다운로드 ─────────────────────────────────────────────────

@app.route('/transactions/download')
@login_required
def transactions_download():
    conn = get_db()
    bf        = request.args.get('branch_id', '')
    ff        = request.args.get('form_type_id', '')
    tf        = request.args.get('type', '')
    date_from = request.args.get('date_from', '')
    date_to   = request.args.get('date_to', '')

    conditions, params = ['1=1'], []
    role, bid = session.get('role'), session.get('branch_id')

    if role != 'admin' and bid:
        conditions.append('(t.from_branch_id=%s OR t.to_branch_id=%s)'); params.extend([bid, bid])
    elif bf:
        conditions.append('(t.from_branch_id=%s OR t.to_branch_id=%s)'); params.extend([bf, bf])
    if ff:
        conditions.append('t.form_type_id=%s'); params.append(ff)
    if tf:
        conditions.append('t.type=%s'); params.append(tf)
    if date_from:
        conditions.append('t.created_at::date >= %s'); params.append(date_from)
    if date_to:
        conditions.append('t.created_at::date <= %s'); params.append(date_to)

    rows = conn.execute(f'''
        SELECT t.*, f.name form_name, f.unit,
               fb.name from_branch_name, fb.code from_branch_code,
               tb.name to_branch_name,   tb.code to_branch_code
        FROM transactions t
        JOIN form_types f ON t.form_type_id = f.id
        LEFT JOIN branches fb ON t.from_branch_id = fb.id
        LEFT JOIN branches tb ON t.to_branch_id   = tb.id
        WHERE {" AND ".join(conditions)}
        ORDER BY t.created_at DESC
    ''', params).fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '입출고이력'

    headers    = ['일시','구분','양식명','단위','출발지점코드','출발지점','도착지점코드','도착지점','수량','비고','처리자']
    col_widths = [18, 8, 32, 8, 12, 14, 12, 14, 10, 20, 12]
    _apply_header(ws, headers, col_widths)

    type_kr = {'IN':'입고', 'OUT':'출고', 'TRANSFER':'이전'}
    in_fill  = PatternFill('solid', fgColor='D1FAE5')
    out_fill = PatternFill('solid', fgColor='FEE2E2')
    tr_fill  = PatternFill('solid', fgColor='DBEAFE')

    for r, row in enumerate(rows, 2):
        t_kr = type_kr.get(row['type'], row['type'])
        fill = in_fill if row['type']=='IN' else (out_fill if row['type']=='OUT' else tr_fill)
        vals = [
            str(row['created_at'] or '')[:16], t_kr, row['form_name'], row['unit'],
            row['from_branch_code'] or '', row['from_branch_name'] or '',
            row['to_branch_code']   or '', row['to_branch_name']   or '',
            row['quantity'], row['notes'] or '', row['created_by'] or '',
        ]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border    = _border()
            cell.alignment = Alignment(vertical='center')
            if c == 2:
                cell.fill = fill
                cell.alignment = Alignment(horizontal='center', vertical='center')
            if c == 9:
                cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[r].height = 16

    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = f'A1:{get_column_letter(len(headers))}1'

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    from datetime import date
    fname = f'입출고이력_{date.today().strftime("%Y%m%d")}.xlsx'
    return send_file(buf, as_attachment=True, download_name=fname,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


try:
    init_db()
except Exception as _e:
    print(f"[init_db] {_e}")

if __name__ == '__main__':
    pass
    _port = int(os.environ.get('PORT', '5000'))
    _host = os.environ.get('FLASK_RUN_HOST', '127.0.0.1')
    print(f'\n  로컬에서 브라우저 주소창에 입력: http://127.0.0.1:{_port}/\n'
          f'  (Windows에서 5000번이 안 열리면 설정 → AirPlay 수신 끄기, 또는 PORT=5001 로 실행)\n')
    app.run(debug=True, host=_host, port=_port)
