# -*- coding: utf-8 -*-
import sqlite3

conn = sqlite3.connect('Database.db')
cursor = conn.cursor()

# Buscar todas las tablas con sus columnas
cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
tables = cursor.fetchall()

print("ESTRUCTURA COMPLETA DE DATABASE.DB:")
print("=" * 60)

for t in tables:
    table_name = t[0]
    cursor.execute(f"PRAGMA table_info({table_name})")
    columns = cursor.fetchall()
    
    # Solo mostrar tablas que puedan tener datos de objetos o animaciones
    col_names = [c[1].lower() for c in columns]
    if any(x in ' '.join(col_names) for x in ['obj', 'anim', 'shield', 'escudo', 'item', 'grh']):
        print(f"\n{table_name}:")
        for col in columns:
            print(f"  {col[1]} ({col[2]})")

# Revisar si hay alguna tabla de objetos con más campos
print("\n" + "=" * 60)
print("Buscando en tabla 'object' si existe:")
try:
    cursor.execute("SELECT * FROM object LIMIT 5")
    rows = cursor.fetchall()
    print(f"Primeras 5 filas de 'object': {rows}")
except:
    print("No hay tabla 'object' o está vacía")

# Buscar tablas con 'inventory' 
print("\n" + "=" * 60)
print("Estructura completa de inventory_item:")
cursor.execute("PRAGMA table_info(inventory_item)")
for col in cursor.fetchall():
    print(f"  {col[1]} ({col[2]})")

conn.close()
