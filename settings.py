import configparser
import os


def path_to_prog():
    return os.path.abspath("" + "settings.ini")


def createConfig(path=path_to_prog()):
    if not os.path.exists(path):
        defaultConfig()


def defaultConfig(path=path_to_prog()):
    """
    Create a default config file
    """
    config = configparser.ConfigParser()
    config.add_section("Settings")
    config.set("Settings", "local_acts_path", "C:\\")
    config.set("Settings", "general_acts_path", "C:\\")

    with open(path, "w") as config_file:
        config.write(config_file)


def get_local_acts_path(path=path_to_prog()):
    config = configparser.ConfigParser()
    config.read(path)
    return config.get("Settings", "local_acts_path")


def get_general_acts_path(path=path_to_prog()):
    config = configparser.ConfigParser()
    config.read(path)
    return config.get("Settings", "general_acts_path")


def set_local_acts_path(str_path, path=path_to_prog()):
    config = configparser.ConfigParser()
    config.read(path)
    # str_path.replaceAll("\\", "\\")
    config.set("Settings", "local_acts_path", str_path)
    with open(path, "w") as config_file:
        config.write(config_file)


def set_general_acts_path(str_path, path=path_to_prog()):
    config = configparser.ConfigParser()
    config.read(path)
    config.set("Settings", "general_acts_path", str_path)
    with open(path, "w") as config_file:
        config.write(config_file)
