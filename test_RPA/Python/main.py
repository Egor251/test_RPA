import excel
import word
import db

# Для упрощения синтаксиса создаём экземпляры класса
db = db.DB()
db.create_db()
words = word.Word()


# Простейшая функция для уникализации элементов массива (не зависимо от их типа)
def unique(list1):
    unique = []
    for number in list1:
        if number not in unique:
            unique.append(number)
    return unique


# Подсчитываем сумму выполненных задач для отделов
def get_count(tasks_list):
    tmp_list = []
    unique_list = []
    output_dict = []
    for task in tasks_list:
        tmp_list.append(task[1])
        if task[1] not in unique_list:
            unique_list.append(task[1])
    for item in unique_list:
        output_dict.append([item, tmp_list.count(item)])
    return output_dict


# Заполняем БД данными из excel-файла
def fill_db(target_list, dep_dict, tasks, last_name_pos=1, first_name_pos=2, second_name_pos=3):
    for worker in target_list:
        tmp_list = [worker[0]]
        try:
            tmp_list.append(worker[last_name_pos] + ' ' + worker[first_name_pos][0] + '. ' + worker[second_name_pos][0] + '. ')
        except TypeError:
            tmp_list.append(
                worker[last_name_pos] + ' ' + worker[first_name_pos][0] + '. ')
        tmp_list.append(dep_dict[worker[5]])
        db.cursor.executemany("INSERT INTO workers VALUES (?, ?, ?)", (tmp_list,))

        for item in tasks:
            db.cursor.executemany("INSERT INTO tasks VALUES (?, ?)", (item,))
        db.conn.commit()
    return


# Главная функция, которая вызывает все остальные функции
def main(file_name):
    departments = excel.read_excel(file_name, 'Отделы')[1:]
    dep_dict = {}
    for item in departments:
        dep_dict[item[0]] = item[1]

    tasks = get_count(excel.read_excel(file_name, 'Задачи')[1:])

    fill_db(excel.read_excel(file_name, 'Сотрудники')[1:],  dep_dict, tasks)

    sql = 'SELECT workers.fio, tasks.count, workers.department FROM workers, tasks WHERE workers.tabel_number = tasks.tabel_number ORDER BY tasks.count DESC'

    result = unique(db.select(sql))

    dep_count_tasks = []
    for item in dep_dict:
        count = 0
        for dep in result:
            if dep[2] == dep_dict[item]:
                count += int(dep[1])
        dep_count_tasks.append([dep_dict[item], count])

    dep_count_tasks.sort(key=lambda x: x[1], reverse=True)
    print(dep_count_tasks)
    final_array = []
    for item in dep_count_tasks:
        final_array.append(item)
        tmp_list = []
        for pos in result:
            if pos[2] == item[0]:
                tmp_list.append(pos)
        tmp_list.sort(key=lambda x: x[1], reverse=True)
        for fin in tmp_list:
            final_array.append([fin[0], fin[1]])
    print(final_array)

    # Создаём и сохраняем word
    words.fill_table(words.create_table(len(final_array), 2), final_array)
    words.finish("Отчёт о загрузке.docx")


if __name__ == '__main__':
    main('Data.xlsb')
