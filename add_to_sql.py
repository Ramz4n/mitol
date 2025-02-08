import pandas as pd
import mysql.connector

# Подключение к базе данных
conn = mysql.connector.connect(
    host='10.40.1.40',
    user='root',
    password='13373a87',
    port=3306,
    database='aldi'
)

cursor = conn.cursor()

# Чтение данных из Excel-файла
df = pd.read_excel('aldi.xlsx')

# Функция для вставки данных в таблицу "Город"
def insert_city(city):
    cursor.execute("SELECT id FROM goroda WHERE Город = %s", (city,))
    result = cursor.fetchone()
    if result is None:
        cursor.execute("INSERT INTO goroda (Город) VALUES (%s)", (city,))
        conn.commit()
        return cursor.lastrowid
    else:
        return result[0]

# Функция для вставки данных в таблицу "Улица"
def insert_street(street, city_id):
    cursor.execute("SELECT id FROM street WHERE Улица = %s AND id_город = %s", (street, city_id))
    result = cursor.fetchone()
    if result is None:
        cursor.execute("INSERT INTO street (Улица, id_город) VALUES (%s, %s)", (street, city_id))
        conn.commit()
        return cursor.lastrowid
    else:
        return result[0]

# Функция для вставки данных в таблицу "Дом"
def insert_house(house_number, street_id):
    cursor.execute("SELECT id FROM doma WHERE Номер = %s AND id_улица = %s", (house_number, street_id))
    result = cursor.fetchone()
    if result is None:
        cursor.execute("INSERT INTO doma (Номер, id_улица) VALUES (%s, %s)", (house_number, street_id))
        conn.commit()
        return cursor.lastrowid
    else:
        return result[0]

# Функция для вставки данных в таблицу "Подъезд"
def insert_entrance(entrance_number):
    cursor.execute("SELECT id FROM padik WHERE Номер = %s", (entrance_number,))
    result = cursor.fetchone()
    if result is None:
        cursor.execute("INSERT INTO padik (Номер) VALUES (%s)", (entrance_number,))
        conn.commit()
        return cursor.lastrowid
    else:
        return result[0]

# Функция для вставки данных в таблицу "Лифты"
def insert_lift(house_id, entrance_id, lift_type):
    cursor.execute("INSERT INTO lifts (id_дом, id_подъезд, Тип_лифта) VALUES (%s, %s, %s)", (house_id, entrance_id, lift_type))
    conn.commit()

# Вставка данных в базу данных
for index, row in df.iterrows():
    city_id = insert_city(row['Город'])
    street_id = insert_street(row['Улица'], city_id)
    house_id = insert_house(row['Дом'], street_id)
    entrance_id = insert_entrance(row['Подъезд'])
    insert_lift(house_id, entrance_id, row['Лифт'])

# Закрытие соединения с базой данных
cursor.close()
conn.close()
