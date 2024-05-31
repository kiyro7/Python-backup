from os import system, walk, remove
from os.path import isdir, abspath, basename
from shutil import copy, copytree
from subprocess import check_output, run, PIPE
from json import loads
from datetime import datetime
from ctypes import windll

import pysnooper

FILENAME_OF_FILE_WITH_DIRS_FOR_BACKUP = "dirs_for_backup_test.txt"


def is_hidden(filepath):
    """
    Проверяет, что файл является скрытым (или нет)
    (взяты отсюда: https://stackoverflow.com/questions/284115/cross-platform-hidden-file-detection)

    :param filepath: путь до файла (в каком формате - видимо пофиг)
    :return: T/F - файл скрытый/нет
    """
    def has_hidden_attribute(filepath):
        """
        Дополнительная функция, видимо проверяет что есть атрибут "скрытый файл"

        :param filepath: путь до файла (в каком формате - видимо пофиг)
        :return: T/F - есть атрибут скрытый/нет
        """
        try:
            attrs = windll.kernel32.GetFileAttributesW(filepath)
            assert attrs != -1
            result = bool(attrs & 2)
        except (AttributeError, AssertionError):
            result = False
        return result

    name = basename(abspath(filepath))
    return name.startswith('.') or has_hidden_attribute(filepath)


def get_target_file_of_shortcut(shortcut_path):
    r"""
    Возвращает путь до файла, на который создан ярлык
    !!! в пути до ярлыка должны быть экранированы символы (обычно это \U \t - они решаются с помощью \\)
    (в этом коменте решил это с помощью r перед строкой - так тоже работает, но в либу это запихать не могу)

    :param shortcut_path: абсолютный путь до ярлыка (с экранированными символами, чтобы ничё не ломалось)
    :return: путь до файла-оригинала
    """
    # return system(f'pylnk3 p "{shortcut_path}" _link_info._path')  # это выполняет команду и мы в терминал получаем её вывод
    return check_output(["pylnk3", "p", shortcut_path, "_link_info._path"]).decode("utf-8").strip()  # а это мы получаем вывод команды, преобразуем в кириллицу и убираем возврат каретки и перенос строки


def get_all_shortcuts(path_to_dir):
    """
    Функция поиска всех файлов-ярылков в заданой папке и её подпапках (рекурсивно короче)
    !!! экранировать символы в имени папки, как в фукнции get_target_file_of_shortcut

    :param path_to_dir: абсолютный путь до папки (с экранированными символами, чтобы ничё не ломалось)
    :return: список строк - путей до файлов-ярлыков
    """
    out = []
    for subdir in walk(path_to_dir):
        subdirname = subdir[0]
        for filename in subdir[2]:
            if filename.endswith(".lnk"):
                out.append(subdirname + "\\" + filename)
    return out


def get_all_shortcuts_with_targets(path_to_dir):
    """
    Функция выявления всех файлов-ярылков в заданой папке и её подпапках (рекурсивно короче)
    и получения путей до соответствующих им файлов
    !!! экранировать символы в имени папки, как в фукнции get_target_file_of_shortcut

    :param path_to_dir: абсолютный путь до папки (с экранированными символами, чтобы ничё не ломалось)
    :return: спиоск кортежей, вида (абс путь до ярлыка, абс путь до соответствующего ему файла)
    """
    shortcuts = get_all_shortcuts(path_to_dir)
    out = []
    for item in shortcuts:
        out.append((item, get_target_file_of_shortcut(item)))
    print("\n\n", out, "\n\n")
    return out


def list_flash_drives():
    """
    Get a list of drives using WMI, excluding local
    (сурс: https://abdus.dev/posts/python-monitor-usb/)

    :return: list of letter of drives, that are not local
    """
    proc = run(args=['powershell', '-noprofile', '-command', 'Get-WmiObject -Class Win32_LogicalDisk | Select-Object deviceid,drivetype | ConvertTo-Json'], text=True, stdout=PIPE)
    if proc.returncode != 0 or not proc.stdout.strip():
        raise WindowsError("list_drives failed to get drives")
        # return []
    devices = loads(proc.stdout)
    print(devices)
    # return [d['deviceid'] for d in devices if d['drivetype'] != 3]
    return [d['deviceid'] for d in devices if d['deviceid'] not in ["C:", "D:", "E:"]]


def get_abs_path_to_backup_dir():
    """
    Функция сбора полного пути до папки, в которую будет происходить бэкап

    :return: абсолютный путь (с экранированными слешами)
    """
    dir_for_backup = datetime.now().strftime("%d.%m.%y %H-%M")
    drive_for_backup = list_flash_drives()
    print(drive_for_backup)
    if len(drive_for_backup) > 1:  # это не будет работать, если запуск по заданному сценарию)
        ind = ""
        while not ind.isnumeric() or not 0 <= int(ind) < len(drive_for_backup):
            ind = input(f"Индекс (с 0) на какой диск бэкапится ({drive_for_backup}): ")
        drive_for_backup = drive_for_backup[int(ind)]
    else:
        try:
            drive_for_backup = drive_for_backup[0]
        except IndexError:
            drive_for_backup = abspath("bac")


            # drive_for_backup = abspath(input("Не найдено флешек, напишите куда бэкапить: "))
            # while not isdir(drive_for_backup):
            #     drive_for_backup = abspath(input("Норм директорию дай: "))



    dir_for_backup = drive_for_backup + "\\Backup (" + dir_for_backup + ")"
    print(f"Будем делать бэкап в следующую директорию: {dir_for_backup}")
    return dir_for_backup


def backup_dirs(dirs_for_backup, dirpath_for_backup):
    """
    Функция копирования всех папок в целевую папку, за исключением скрытых файлов/папок

    :param dirs_for_backup: список абсолютных путей до папок, которые надо полностью скопировать (слеши должны быть экранированы)
    :param dirpath_for_backup: абс путь до папки, в которую нужно скопировать папки (слеши должны быть экранированы)
    """
    for dirpath in dirs_for_backup:

        # добавить проверку на то, что файлы/папки НЕ скрытые

        dirname = dirpath.split("\\")[-1]
        dest_dirpath = dirpath_for_backup + "\\" + dirname
        print(dirpath, dirname, dest_dirpath, sep="\t")
        system(f'backupy "{dirpath}" "{dest_dirpath}" --noprompt --nolog')


def replace_shortcuts(filepaths, dirpath_for_backup):
    """
    Заменяет все скопированные ярлыки на настоящие файлы

    :param filepaths: список кортежей вида: (абс путь до папки, возврат функции get_all_shortcuts_with_targets для неё)
    :param dirpath_for_backup: абс путь до папки, в которой лежит бэкап (слеши должны быть экранированы)
    """
    # мбмб possible_links_inside = []
    for item in filepaths:
        dirpath, shortcuts_with_targets = item
        dirname = dirpath.split("\\")[-1]
        print(dirpath, dirname, shortcuts_with_targets, sep="\t", end="\n\n")
        for jtem in shortcuts_with_targets:
            shortcut_path, target_path = jtem
            start_ind = shortcut_path.index(dirname)
            shortcut_path = shortcut_path[start_ind:]
            remove(dirpath_for_backup + "\\" + shortcut_path)
            if not isdir(target_path):
                start_ind = shortcut_path.rindex("\\")
                shortcut_path = shortcut_path[:start_ind]
                copy(target_path, dirpath_for_backup + "\\" + shortcut_path)
            else:
                # мб мб possible_links_inside.append(target_path)
                # вот тут запомнить на будущее (посмотреть в ней ярлыки и так далее)
                target_folder_name = target_path.split("\\")[-1]
                copytree(target_path, dirpath_for_backup + "\\" + dirname + "\\" + target_folder_name)


if __name__ == "__main__":



    dirpath_for_backup = get_abs_path_to_backup_dir()
    with open(FILENAME_OF_FILE_WITH_DIRS_FOR_BACKUP, mode="r", encoding="utf-8") as file:
        dirs_for_backup = list(map(str.strip, file.readlines()))  # тут по идее можем неэкранированные слеши получить, надо с этим что-то делать, реплейс не прикрутить
    print(dirs_for_backup)
    backup_dirs(dirs_for_backup, dirpath_for_backup)
    folders_with_shortcuts = [(dirpath, get_all_shortcuts_with_targets(dirpath)) for dirpath in dirs_for_backup]
    replace_shortcuts(folders_with_shortcuts, dirpath_for_backup)

