"""
妍發科技 製程管理系統 — 資料庫層
SQLite 資料庫，包含訂單、製程步驟、出貨記錄、廠商四張表。
"""

import sqlite3
import os
from datetime import date, datetime
from contextlib import contextmanager

DB_PATH = os.path.join(os.path.dirname(__file__), "yenfa.db")


@contextmanager
def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_db():
    """建立所有資料表（若不存在）。"""
    with get_conn() as conn:
        conn.executescript("""
        -- 廠商主表
        CREATE TABLE IF NOT EXISTS vendors (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            name        TEXT NOT NULL UNIQUE,
            contact     TEXT,
            phone       TEXT,
            note        TEXT,
            created_at  TEXT DEFAULT (datetime('now','localtime'))
        );

        -- 訂單主表
        CREATE TABLE IF NOT EXISTS orders (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            order_no        TEXT,
            part_no         TEXT NOT NULL,
            customer        TEXT,
            quantity         INTEGER,
            unit_price      REAL,
            amount          REAL,
            vendor_id       INTEGER REFERENCES vendors(id),
            vendor_name     TEXT,
            order_date      TEXT,
            due_date        TEXT,
            status          TEXT DEFAULT '製程中'
                            CHECK(status IN ('製程中','已出貨','客戶暫停','待零件','已取消','延遲')),
            source_sheet    TEXT,
            note            TEXT,
            created_at      TEXT DEFAULT (datetime('now','localtime')),
            updated_at      TEXT DEFAULT (datetime('now','localtime'))
        );

        -- 製程追蹤表（每筆訂單的每一道製程）
        CREATE TABLE IF NOT EXISTS process_steps (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            order_id        INTEGER NOT NULL REFERENCES orders(id) ON DELETE CASCADE,
            part_no         TEXT NOT NULL,
            step_seq        INTEGER NOT NULL,
            process_name    TEXT NOT NULL,
            vendor_name     TEXT,
            planned_date    TEXT,
            actual_date     TEXT,
            status          TEXT DEFAULT '待處理'
                            CHECK(status IN ('待處理','進行中','完成','延遲','跳過')),
            work_date       TEXT,
            note            TEXT,
            created_at      TEXT DEFAULT (datetime('now','localtime'))
        );

        -- 出貨記錄表
        CREATE TABLE IF NOT EXISTS shipments (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            order_id        INTEGER REFERENCES orders(id),
            part_no         TEXT NOT NULL,
            ship_date       TEXT,
            ship_quantity   INTEGER,
            amount          REAL,
            note            TEXT,
            created_at      TEXT DEFAULT (datetime('now','localtime'))
        );

        -- 建立索引加速查詢
        CREATE INDEX IF NOT EXISTS idx_orders_part_no ON orders(part_no);
        CREATE INDEX IF NOT EXISTS idx_orders_status ON orders(status);
        CREATE INDEX IF NOT EXISTS idx_process_order_id ON process_steps(order_id);
        CREATE INDEX IF NOT EXISTS idx_process_part_no ON process_steps(part_no);
        CREATE INDEX IF NOT EXISTS idx_shipments_part_no ON shipments(part_no);
        """)


# ── CRUD: Vendors ──

def upsert_vendor(name, contact=None, phone=None, note=None):
    with get_conn() as conn:
        conn.execute("""
            INSERT INTO vendors (name, contact, phone, note)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(name) DO UPDATE SET
                contact = COALESCE(excluded.contact, contact),
                phone   = COALESCE(excluded.phone, phone),
                note    = COALESCE(excluded.note, note)
        """, (name, contact, phone, note))


def get_vendors():
    with get_conn() as conn:
        return conn.execute("SELECT * FROM vendors ORDER BY name").fetchall()


# ── CRUD: Orders ──

def insert_order(**kwargs):
    cols = ', '.join(kwargs.keys())
    placeholders = ', '.join(['?'] * len(kwargs))
    with get_conn() as conn:
        cur = conn.execute(
            f"INSERT INTO orders ({cols}) VALUES ({placeholders})",
            list(kwargs.values())
        )
        return cur.lastrowid


def update_order(order_id, **kwargs):
    sets = ', '.join(f"{k} = ?" for k in kwargs)
    vals = list(kwargs.values()) + [order_id]
    with get_conn() as conn:
        conn.execute(
            f"UPDATE orders SET {sets}, updated_at = datetime('now','localtime') WHERE id = ?",
            vals
        )


def get_orders(status=None, part_no=None):
    query = "SELECT * FROM orders WHERE 1=1"
    params = []
    if status:
        query += " AND status = ?"
        params.append(status)
    if part_no:
        query += " AND part_no LIKE ?"
        params.append(f"%{part_no}%")
    query += " ORDER BY due_date, part_no"
    with get_conn() as conn:
        return conn.execute(query, params).fetchall()


def get_order_by_id(order_id):
    with get_conn() as conn:
        return conn.execute("SELECT * FROM orders WHERE id = ?", (order_id,)).fetchone()


# ── CRUD: Process Steps ──

def insert_process_step(**kwargs):
    cols = ', '.join(kwargs.keys())
    placeholders = ', '.join(['?'] * len(kwargs))
    with get_conn() as conn:
        cur = conn.execute(
            f"INSERT INTO process_steps ({cols}) VALUES ({placeholders})",
            list(kwargs.values())
        )
        return cur.lastrowid


def get_process_steps(order_id=None, part_no=None):
    query = "SELECT * FROM process_steps WHERE 1=1"
    params = []
    if order_id:
        query += " AND order_id = ?"
        params.append(order_id)
    if part_no:
        query += " AND part_no LIKE ?"
        params.append(f"%{part_no}%")
    query += " ORDER BY part_no, step_seq"
    with get_conn() as conn:
        return conn.execute(query, params).fetchall()


def update_process_step(step_id, **kwargs):
    sets = ', '.join(f"{k} = ?" for k in kwargs)
    vals = list(kwargs.values()) + [step_id]
    with get_conn() as conn:
        conn.execute(f"UPDATE process_steps SET {sets} WHERE id = ?", vals)


# ── CRUD: Shipments ──

def insert_shipment(**kwargs):
    cols = ', '.join(kwargs.keys())
    placeholders = ', '.join(['?'] * len(kwargs))
    with get_conn() as conn:
        cur = conn.execute(
            f"INSERT INTO shipments ({cols}) VALUES ({placeholders})",
            list(kwargs.values())
        )
        return cur.lastrowid


def get_shipments(part_no=None, month=None):
    query = "SELECT * FROM shipments WHERE 1=1"
    params = []
    if part_no:
        query += " AND part_no LIKE ?"
        params.append(f"%{part_no}%")
    if month:
        query += " AND ship_date LIKE ?"
        params.append(f"{month}%")
    query += " ORDER BY ship_date DESC"
    with get_conn() as conn:
        return conn.execute(query, params).fetchall()


# ── 統計查詢 ──

def get_dashboard_stats():
    """取得儀表板統計數據。"""
    with get_conn() as conn:
        stats = {}
        stats['total_orders'] = conn.execute(
            "SELECT COUNT(*) FROM orders"
        ).fetchone()[0]
        stats['in_progress'] = conn.execute(
            "SELECT COUNT(*) FROM orders WHERE status = '製程中'"
        ).fetchone()[0]
        stats['shipped'] = conn.execute(
            "SELECT COUNT(*) FROM orders WHERE status = '已出貨'"
        ).fetchone()[0]
        stats['delayed'] = conn.execute(
            "SELECT COUNT(*) FROM orders WHERE status = '延遲'"
        ).fetchone()[0]
        stats['on_hold'] = conn.execute(
            "SELECT COUNT(*) FROM orders WHERE status IN ('客戶暫停','待零件')"
        ).fetchone()[0]
        stats['total_amount'] = conn.execute(
            "SELECT COALESCE(SUM(amount), 0) FROM orders"
        ).fetchone()[0]
        stats['shipped_amount'] = conn.execute(
            "SELECT COALESCE(SUM(amount), 0) FROM orders WHERE status = '已出貨'"
        ).fetchone()[0]
        return stats


def get_vendor_load():
    """取得各廠商負載統計。"""
    with get_conn() as conn:
        return conn.execute("""
            SELECT vendor_name,
                   COUNT(*) as total_steps,
                   SUM(CASE WHEN status = '完成' THEN 1 ELSE 0 END) as completed,
                   SUM(CASE WHEN status = '進行中' THEN 1 ELSE 0 END) as in_progress,
                   SUM(CASE WHEN status = '延遲' THEN 1 ELSE 0 END) as delayed
            FROM process_steps
            WHERE vendor_name IS NOT NULL AND vendor_name != ''
            GROUP BY vendor_name
            ORDER BY total_steps DESC
        """).fetchall()


def get_monthly_shipment_summary():
    """取得月出貨金額統計。"""
    with get_conn() as conn:
        return conn.execute("""
            SELECT substr(ship_date, 1, 7) as month,
                   COUNT(*) as count,
                   COALESCE(SUM(amount), 0) as total_amount
            FROM shipments
            WHERE ship_date IS NOT NULL
            GROUP BY substr(ship_date, 1, 7)
            ORDER BY month
        """).fetchall()


def get_process_progress(part_no):
    """取得特定料號的製程完成度。"""
    with get_conn() as conn:
        row = conn.execute("""
            SELECT COUNT(*) as total,
                   SUM(CASE WHEN status = '完成' THEN 1 ELSE 0 END) as completed
            FROM process_steps
            WHERE part_no = ?
        """, (part_no,)).fetchone()
        if row and row['total'] > 0:
            return row['completed'] / row['total'] * 100
        return 0.0


if __name__ == "__main__":
    init_db()
    print(f"資料庫已建立: {DB_PATH}")
