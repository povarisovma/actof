import configparser
import os


def path_to_settings_file():
    """
    :return: возвращает путь к файлу настроек settings.ini
    """
    return os.path.abspath("" + "settings.ini")


def create_settings_file(path=path_to_settings_file()):
    """
    Создает файл настроек по умолчанию, если он отсутствует
    """
    if not os.path.exists(path):
        config = configparser.ConfigParser()
        config.add_section("Settings")
        config.set("Settings", "local_acts_path", "C:\\")
        config.set("Settings", "general_acts_path", "C:\\")

        with open(path, "w") as config_file:
            config.write(config_file)


def get_local_acts_path_folder(path=path_to_settings_file()):
    """
    :param path: путь до файла settings.ini, передается по умолчанию
    :return: возвращает путь до папки с локальными актами
    """
    config = configparser.ConfigParser()
    config.read(path)
    return config.get("Settings", "local_acts_path")


def get_general_acts_path_folder(path=path_to_settings_file()):
    """
    :param path: путь до файла settings.ini, передается по умолчанию
    :return: возвращает путь до папки с общими актами
    """
    config = configparser.ConfigParser()
    config.read(path)
    return config.get("Settings", "general_acts_path")


def set_local_acts_path_folder_in_settings(str_path, path=path_to_settings_file()):
    """
    записывает путь к папке локальных актов в файл настроек settings.ini
    :param str_path:
    :param path:
    """
    config = configparser.ConfigParser()
    config.read(path)
    config.set("Settings", "local_acts_path", str_path)
    with open(path, "w") as config_file:
        config.write(config_file)


def set_general_acts_path_folder_in_settings(str_path, path=path_to_settings_file()):
    """
    записывает путь к папке общих актов в файл настроек settings.ini
    :param str_path:
    :param path:
    """
    config = configparser.ConfigParser()
    config.read(path)
    config.set("Settings", "general_acts_path", str_path)
    with open(path, "w") as config_file:
        config.write(config_file)
