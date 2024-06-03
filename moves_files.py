import os
import shutil
def move_files():
    source_folder = "D:\\NIR\\LIST_NEW"
    destination_folder = "D:\\NIR\\LIST_NEW"

    for file in os.listdir(source_folder):
        if os.path.isfile(os.path.join(source_folder, file)):  # Проверка, является ли объект файлом
            filename, file_extension = os.path.splitext(file)
            if file_extension == ".xlsx":
                course_name = filename.split('_')[0]  # Получаем название курса из имени файла
                course_folder = os.path.join(destination_folder, course_name + "_courses")  # Папка курса
                if not os.path.exists(course_folder):
                    os.makedirs(course_folder)  # Создаем папку, если она не существует
                shutil.move(os.path.join(source_folder, file), os.path.join(course_folder, file))  # Перемещаем файл

move_files()