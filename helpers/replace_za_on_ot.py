import os


# Функция для переименования файлов
def rename_files(path):
    # Получаем список файлов и поддиректорий в директории
    files = os.listdir(path)

    # Проходим по каждому файлу и поддиректории
    for file in files:
        # Получаем полный путь к файлу или поддиректории
        full_path = os.path.join(path, file)

        # Если это директория, вызываем функцию rename_files() рекурсивно
        if os.path.isdir(full_path):
            rename_files(full_path)
        # Если это файл, проверяем его имя на наличие слова "за" и заменяем его на "от"
        elif os.path.isfile(full_path):
            if 'Счои' in file:
                new_file = file.replace('Счои', 'Сочи')
                os.rename(full_path, os.path.join(path, new_file))


# Вызываем функцию rename_files() для директории, которую нужно проверить
rename_files('C:/Users/Admin/OneDrive/Заявки Евгений/2023 год')
