import sqlite3


class DB:
    db_path = "workers.db"

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    def create_db(self):
        self.cursor.execute(f"""CREATE TABLE IF NOT EXISTS workers
                                 (tabel_number text, fio text, department text)
                              """)
        self.cursor.execute(f"""CREATE TABLE IF NOT EXISTS tasks
                                         (tabel_number text, count int)
                                      """)
        self.conn.commit()

    def insert(self, data, data2):

        try:
            self.cursor.executemany("INSERT INTO workers VALUES (?, ?, ?)",
                                    (data,))  # Количество ? должно совпадать с количеством элементов во входном массиве
        except sqlite3.IntegrityError:
            pass

        try:
            self.cursor.executemany("INSERT INTO tasks VALUES (?, ?)",
                                    (data2,))  # Количество ? должно совпадать с количеством элементов во входном массиве
        except sqlite3.IntegrityError:
            pass
        self.conn.commit()

    def select(self, sql):  # Выполнение select к БД
        self.cursor.execute(sql)
        result = self.cursor.fetchall()
        return result  # Возращает список со списком записей
