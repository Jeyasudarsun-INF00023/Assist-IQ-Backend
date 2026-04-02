"""Add price column to assets table"""
from database import engine
from sqlalchemy import text

with engine.connect() as conn:
    try:
        conn.execute(text("ALTER TABLE assets ADD COLUMN price TEXT"))
        conn.commit()
        print("✅ price column added to assets table")
    except Exception as e:
        if "duplicate column" in str(e).lower() or "already exists" in str(e).lower():
            print("ℹ️  price column already exists, skipping")
        else:
            print(f"❌ Error: {e}")
