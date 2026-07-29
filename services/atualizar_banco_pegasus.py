import sqlite3

conn = sqlite3.connect("phoenix_tool.db")
cursor = conn.cursor()

campos = [
    "pegasus_abertura TEXT",
    "pegasus_fechamento TEXT",
    "custo_abertura TEXT",
    "custo_fechamento TEXT"
]

for campo in campos:
    try:
        cursor.execute(
            f"ALTER TABLE solicitacoes ADD COLUMN {campo}"
        )
        print(f"{campo} criado")
    except Exception:
        print(f"{campo} já existe")

conn.commit()
conn.close()