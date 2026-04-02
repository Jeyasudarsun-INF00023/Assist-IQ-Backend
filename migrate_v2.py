
from sqlalchemy import create_engine, text

DATABASE_URL = "postgresql://postgres:jeya662004%40@localhost:5432/assist_iq"
engine = create_engine(DATABASE_URL)

with engine.connect() as conn:
    print("Adding columns to employees table...")
    columns = [
        ("last_app", "VARCHAR"),
        ("last_window", "VARCHAR"),
        ("is_online", "BOOLEAN DEFAULT FALSE")
    ]
    
    for col_name, col_type in columns:
        try:
            conn.execute(text(f"ALTER TABLE employees ADD COLUMN {col_name} {col_type}"))
            conn.commit()
            print(f"{col_name} added.")
        except Exception as e:
            print(f"Error adding {col_name}: {e}")

print("Migration complete.")
