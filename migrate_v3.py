
from sqlalchemy import create_engine, text

DATABASE_URL = "postgresql://postgres:jeya662004%40@localhost:5432/assist_iq"
engine = create_engine(DATABASE_URL)

with engine.connect() as conn:
    print("Adding last_seen column to employees table...")
    try:
        conn.execute(text("ALTER TABLE employees ADD COLUMN last_seen TIMESTAMP"))
        conn.commit()
        print("last_seen added.")
    except Exception as e:
        print(f"Error adding last_seen: {e}")

print("Migration complete.")
