import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app, init_db

# Vercel cold start 시 DB 초기화 (테이블 없으면 생성)
try:
    init_db()
except Exception as e:
    print(f"[init_db] {e}")
