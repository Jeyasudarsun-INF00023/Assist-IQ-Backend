from sqlalchemy import create_engine, text
from database import DATABASE_URL

engine = create_engine(DATABASE_URL)

with engine.connect() as connection:
    try:
        connection.execute(text("ALTER TABLE employees ADD COLUMN verification_code VARCHAR;"))
        connection.execute(text("ALTER TABLE employees ADD COLUMN verification_code_expires TIMESTAMP;"))
        connection.commit()
        print("Columns for email verification added successfully.")
    except Exception as e:
        print(f"Error adding columns (they might already exist): {e}")
