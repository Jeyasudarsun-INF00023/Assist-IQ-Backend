
from sqlalchemy import create_engine, text
from datetime import datetime

DATABASE_URL = "postgresql://postgres:jeya662004%40@localhost:5432/assist_iq"
engine = create_engine(DATABASE_URL)

with engine.connect() as conn:
    print("Adding api_token to employees table...")
    try:
        conn.execute(text("ALTER TABLE employees ADD COLUMN api_token VARCHAR"))
        conn.commit()
        print("api_token added.")
    except Exception as e:
        print(f"Error adding api_token: {e}")

    print("Creating activity_logs table...")
    try:
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS activity_logs (
                id SERIAL PRIMARY KEY,
                employee_id VARCHAR,
                app VARCHAR,
                window VARCHAR,
                start_time TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                end_time TIMESTAMP
            )
        """))
        conn.commit()
        print("activity_logs table created.")
    except Exception as e:
        print(f"Error creating activity_logs: {e}")

    print("Assigning default tokens to existing employees for testing...")
    try:
        conn.execute(text("UPDATE employees SET api_token = 'SECRET123' WHERE api_token IS NULL"))
        conn.commit()
    except Exception as e:
        print(f"Error updating tokens: {e}")

print("Migration complete.")
