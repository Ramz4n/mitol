import pandas as pd
import mysql.connector

# Подключение к базе данных
conn = mysql.connector.connect(
    host='10.40.1.40',
    user='root',
    password='13373a87',
    port=3306,
    database='osp'
)

cursor = conn.cursor()

# Чтение данных из Excel-файла
df = pd.read_excel('osp.xlsx')

# Функция для вставки данных в таблицу "passport"
def insert_passport(row):
    cursor.execute("""
        INSERT INTO passport (город, адрес, завод_№, изготовитель, модель, год_ввода,
        год_изготовления, грузопод, этажн, кол_во_канатов, длина_канатов, диаметр_канатов,
        редуктор_леб, двигатель_леб, шкив_канат, отводной_блок, частотник, рег_привод_двер, нерег_привод_двер)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
    """, (
        row['город'], row['адрес'], row['завод_№'], row['изготовитель'], row['модель'],
        row['год_ввода'], row['год_изготовления'], row['грузопод'], row['этажн'], row['кол_во_канатов'],
        row['длина_канатов'], row['диаметр_канатов'], row['редуктор_леб'], row['двигатель_леб'],
        row['шкив_канат'], row['отводной_блок'], row['частотник'], row['рег_привод_двер'], row['нерег_привод_двер']
    ))
    conn.commit()

# Замена NaN на None
df = df.where(pd.notnull(df), None)

# Вставка данных в базу данных
for index, row in df.iterrows():
    insert_passport(row)

# Закрытие соединения с базой данных
cursor.close()
conn.close()