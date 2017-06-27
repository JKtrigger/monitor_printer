# coding: utf8
import ConfigParser
import codecs
import os
import re
import subprocess
import sys
import time
import win32print
import datetime

# Константы пакета

SESSION_ID_NUMBER = 'session_id_number'
SESSION_USERNAME = 'session_username'
SESSION_NAME = 'session_name'
IS_CURRENT_SESSION = 'is_current_session'
SESSION_STATUS = 'session_status'
LOGON_TIME = 'logon_time'
IDLE_TIME = 'idle_time'
SESSION_DOES_NOT_EXISTS = 'session_not_exist'
CLIENT_NAME = 'clientname'


class ExceptionMainSectionNotSet(Exception):
    """ Не задана обязательная секция """


class ExceptionNotFoundSection(ExceptionMainSectionNotSet):
    """
    Опция задана в main_section,
    но не существует секции с таким названием.
    """


class PowerShellIsNotInstalled(ExceptionNotFoundSection):
    """
    Не установлен пакет power shell

    Планировалось обойтись wmi или win32print,
    но по ряду причин с удалением принтра
    из системы эти API не справляются
    , а в случаии wmi создают массу проблемм
    """


class QueryWindowsSession(object):
    """ Класс работы с Windows терминалами(сессиями)  """
    def __init__(
            self,
            session_id_number,
            session_username,
            session_name,
            is_current_session,
            session_status,
            logon_time,
            idle_time
    ):
        """
        :param session_id_number (int): Номер сессии
        :param session_username (basestring): Имя пользователя
        :param session_name (basestring): Имя сессии
        :param session_is_current (bool): свойство является ли
                                          эта сессия текущей для процесса
        :param session_state (int): Active or Terminated
        :param logon_time (basestring): время входа
        :param idle_time (int): время простоя в минутах ( для RDP-TCP )
        """
        self.session_id_number = int(session_id_number)
        self.session_username = session_username
        self.session_name = session_name
        self.is_current_session = is_current_session
        self.session_status = session_status
        self.logon_time = logon_time
        self.idle_time = idle_time

    def __unicode__(self):
        return (
            u"SESSION.session_name = {session_name}, \r\n"
            u"SESSION.session_id_number = {session_id_number}, \r\n"
            u"SESSION.session_username = {session_username} \r\n"
            u"SESSION.is_current_session = {session_is_current}".format(
                session_id_number=self.session_id_number,
                session_name=self.session_name,
                session_username=self.session_username,
                session_is_current=self.is_current_session
            )
        )

    def __repr__(self):
        return self.__unicode__()

    def __hash__(self):
        return hash(self.__unicode__())


class PrinterAdvancedTask(object):
    """
    Утилита для автоматизации действий с принтерами по умолчанию

    Задача:
        удалять принетра с ошибками,
        удалять принтера по описанию
        выставлять принтер по умолчанию c учетом номера RDP сессий ,
            (с ограничением по маске)
    """
    # TODO: Нет реализации логирования

    # Описание вывода в консоль
    # по умолчанию с выводом в консоль
    ENABLE_STDOUT = True
    DISABLE_STDOUT = False
    PRINT_PROCESS_DETAIL = ENABLE_STDOUT

    # Параметры из файла settings.conf
    keep_printer_info = set_default_pattern = delete_printers_like = None

    # Все принетра
    ALL_PRINTERS = win32print.EnumPrinters(
        win32print.PRINTER_ENUM_LOCAL, None, 1
    )

    # Состояния печати
    ERROR_PRINTING = 8210
    STATUS_IS_PRINTING = 8208
    STATUS_IS_DELETING = 8214
    STATUS_IS_READY = 0

    # Доступные статусы
    HANDLER_STATUS = {
        ERROR_PRINTING: u'Ошибка печати',
        STATUS_IS_PRINTING: u'Печать',
        STATUS_IS_DELETING: u'Удаление',
        STATUS_IS_READY: u'Готово'
    }

    # Директория логирования
    try:
        BASE_PATH = os.path.dirname(__file__)
    except NameError:  # После применения py2exe
        import sys
        BASE_PATH = os.path.dirname(os.path.abspath(sys.argv[0]))
    LOG = 'log'
    PRINTERS = 'printers'
    PRINTERS_PATH = os.path.join(BASE_PATH, PRINTERS)
    LOG_PATH = os.path.join(BASE_PATH, LOG)
    LOG_FILE = None

    # Парсер
    parser = ConfigParser.ConfigParser()
    FILE_NAME_SETTINGS = 'settings.conf'
    MAIN_SECTION = 'main_section'
    parser.read(os.path.join(BASE_PATH, FILE_NAME_SETTINGS))

    # Команда режима
    MODE = 'MODE'
    # Режимы: просмотр, мастер, обычный
    MODE_COMMON = 'COMMON'
    MODE_MASTER = 'MASTER'
    MODE_VIEW = 'VIEW'

    # Текущая сессия
    # сессия от имени которой запущен экземляр процесса
    CURRENT_SESSION = ''
    # Все сессии
    ALL_SESSIONS = []
    user_name_session_dict = {}

    # Команды windows оболочки
    # QUERY Путь
    # Обенносить Windows
    QUERY_EXE = ''
    QUERY_EXE_1 = 'c:\\windows\\sysnative\\query.exe'
    QUERY_EXE_2 = 'c:\\WINDOWS\\system32\\query.exe'
    # SETX_PATH
    SETX_EXE = 'C:\\Windows\\System32\\setx.exe'
    # Power shell
    POWER_SHELL_EXE = (
        'C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe')

    power_shell_script = (
        # Базовый сценарий удаления в power shell
        u"""(Get-WmiObject -Query 'select * from Win32_Printer where name="{name_of_printer}"').delete()""")
    clear_all_jobs = (
        # Базовое удаление
        u"""(Get-WmiObject -Query 'select * from Win32_Printer where name="{name_of_printer}"').CancelAllJobs()""")
    # QUERY Команды
    SESSION_QUERY_COMMAND_SESSION = 'session'
    SESSION_QUERY_COMMAND_USER = 'user'

    def __init__(self, args):
        self.checks()
        self.get_sessions()
        self.init_session_keys()
        self.get_config()
        self.exec_sub_program(args)

    def get_accepts(self):
        """  Список разрещенных пользователей """
        users = []
        for option in self.parser.options(self.MAIN_SECTION):
            users.append(self.parser.get(self.MAIN_SECTION, option))
        return users

    def init_session_keys(self):
        for session in self.ALL_SESSIONS:
            self.user_name_session_dict[session.session_username] = session

    def get_config(self):
        """
        Сбор настроек для выполнения в общем режиме.

        Только для текущего пользователя
        """
        self.client_name = os.getenv(CLIENT_NAME) or 'localhost'
        self.log(u'Сбор сведений из settings.conf')
        if self.parser.has_section(self.CURRENT_SESSION.session_username):
            self.log(u'Получение списка правил для текщего пользователя')
            params = dict(
                self.parser.items(
                    self.CURRENT_SESSION.session_username)
            )

            set_default_printer_like = params.get(
                'set_default_printer_like',
                None
            )
            if set_default_printer_like:
                self.log(u'Получение описания принтера по умолчанию')
                set_default_printer_like.format(
                        client_name=self.client_name,
                        session_id_number=(
                            self.CURRENT_SESSION.session_id_number)
                    )
                self.set_default_pattern = re.compile(
                    set_default_printer_like.format(
                        client_name=self.client_name,
                        session_id_number=(
                            self.CURRENT_SESSION.session_id_number)
                    )
                )
            keep_printer_info = params.get('keep_printer_info', None)
            if keep_printer_info == 'True':
                self.log(
                    u'Получение инструкции о хранении принтера по умолчанию')
                self.keep_printer_info = True

    def checks(self):
        """ Проверка состояния пакета """
        # проверка путей
        paths = [self.LOG_PATH, self.PRINTERS_PATH]
        for path in paths:
            if not os.path.exists(path):
                os.makedirs(path)

        # Проверка секций
        if self.MAIN_SECTION not in self.parser.sections():
            sys.stdout.write(u'Не задана обязательная секция main_section\n')
            self.log(u'Нет обязательной секции main_section')
            raise ExceptionMainSectionNotSet(
                'Section "{}" not set'.format(self.MAIN_SECTION)
            )
        for main_section_key in self.parser.options(self.MAIN_SECTION):
            main_section_value = self.parser.get(
                self.MAIN_SECTION, main_section_key)
            if main_section_value not in self.parser.sections():
                self.log(u'Задано правило {} без описания'.format(
                    main_section_value)
                )
                raise ExceptionNotFoundSection(
                    'Section "{}" not found'.format(main_section_value))

        # Проверка power shell
        if not os.path.exists(self.POWER_SHELL_EXE):
            self.log(u'Не уставнолен power shell ')
            raise PowerShellIsNotInstalled("Power Shell is not found")

        # Проверка логов файла
        computer_name = os.getenv(CLIENT_NAME) or 'localhost'
        date_time = datetime.datetime.today()
        user_name = os.getenv('username')
        TXT = '.txt'
        log_file_name = '_'.join(
            [user_name, date_time.strftime("%Y_%m_%d"), computer_name]) + TXT
        full_log_path = os.path.join(self.LOG_PATH, log_file_name)

        self.LOG_FILE = full_log_path

        file_ = codecs.open(full_log_path, 'a', 'utf-8')
        file_.close()

        # Проверка доступности query.exe (отличаются для 64 и 32 Windows)
        if os.path.exists(os.path.join(self.QUERY_EXE_1)):
            self.QUERY_EXE = self.QUERY_EXE_1
        else:
            self.QUERY_EXE = self.QUERY_EXE_2
        self.log(u'Зауск приложения')

    @staticmethod
    def exec_windows_commands(*args):
        """ Выполнение команды в среде windows """
        session_query_user = subprocess.Popen(
            args,
            stdout=subprocess.PIPE,
            shell=True
        )
        out, error = session_query_user.communicate()

        if error:
            sys.stdout.write(error)
            sys.stdout.write(u'Дальнейшее выполнение не возможно')
            return
        return out

    @staticmethod
    def to_int(string_or_none):
        """
        Преобразование string_or_none в целое число

        возвращает 0 если не удалось произвести преобразование
        :param string_or_none:
        :return:
        """
        try:
            digit = int(string_or_none)
        except (ValueError, TypeError):
            return 0
        return digit

    def get_sessions(self):
        """Получение параметров сессий"""
        self.log(u'Получение паметров сессий')
        out_query_session = self.exec_windows_commands(
            self.QUERY_EXE,
            self.SESSION_QUERY_COMMAND_SESSION
        )
        compiled_pattern_for_query_session_command = re.compile(
            r'''
             (
             # Запрашиваем данные из колонок query session
             # если все колонки не пусты то они остаются в выборке
                # Захват > указателя текущей сессии
                (?P<is_current_session>\s|>)
                # Захват имени сессии
                (?P<session_name>
                    \w+(-?)?\w+\#?(\d{1,5})?|(?!\s)\w+|\s+)
                # Пропуск пробелов
                (!?\s+)
                # Захват всего имени пользователя
                (?P<session_username>
                    [a-zA-Zа-яА-Я0-9]+(\.)?([a-zA-Zа-яА-Я0-9]+)?)
                # Пропуск пробелов
                (!?\s+)
                # Захват ID сессии
                (?P<session_id_number>\d{1,5})
                # Пропуск пробелов
                (!?\s+)
                # Захват состояния
                (?P<session_status>\w+)
             )''',
            # Эранирование комметриев
            re.VERBOSE
        )
        compiled_pattern_for_query_users_command = re.compile(
            r'''
            # Запрос данных из колонок query user
            (
                # признак указателя сессии
                (?P<is_current_session>\s|>)
                # Захват всего имени пользователя
                (?P<session_username>
                    [a-zA-Zа-яА-Я0-9]+(\.)?([a-zA-Zа-яА-Я0-9]+)?)
                # Пропуск пробелов
                (!?\s+)
                # Захват имени сессии
                (?P<session_name>
                    \w+(-?)?\w+\#?(\d{1,5})?|\s+)
                # Пропуск пробелов
                (!?\s+)
                # Захват ID сессии
                (?P<session_id_number>\d{1,5})
                # Пропуск пробелов
                (!?\s+)
                # Захват состояния
                (?P<session_status>\w+)
                # Пропуск пробелов
                (!?\s+)
                # Захват времени простоя
                (?P<day>\d{1,2})?.?
                (?P<hour>\d{1,2})?.?(?P<minute>\d{1,2})?|none|\.
                # Пропуск пробелов
                (!?\s+)
                (?P<logon_time>\d{1,2}.\d{1,2}.\d{2,4}\s\d{1,2}.\d{1,2})

            )
            ''',
            re.VERBOSE
        )
        # 1) Объединить вывод в один поток
        # 2) Разбить вывод на ряд строк, используя разделить \r\n
        # 3) Отсеять шапку и следующую за ней строку [2:]
        # 4) Определить параметры сессий и сохранить их
        # Параметры будем брать из query session и query user
        params_store = {SESSION_ID_NUMBER: {}}
        for row in ''.join(out_query_session).split('\r\n')[2:]:
            found_match_session = (
                compiled_pattern_for_query_session_command.match(row))
            if found_match_session:
                session_name = found_match_session.group(SESSION_NAME)
                params_session = {
                    IS_CURRENT_SESSION: (
                        # сессия в которой выполняется текущий процесс
                        True if found_match_session.group(
                            IS_CURRENT_SESSION) != ' ' else False),
                    SESSION_NAME: session_name if len(session_name) != (
                        session_name.count(' ')) else SESSION_DOES_NOT_EXISTS,
                    SESSION_ID_NUMBER:
                        found_match_session.group(SESSION_ID_NUMBER),
                    SESSION_STATUS: found_match_session.group(SESSION_STATUS),
                    SESSION_USERNAME: found_match_session.group(
                        SESSION_USERNAME)
                }
                # Параметр SESSION_ID_NUMBER уникальный
                # и присутствует обязательно
                params_store[params_session.get(SESSION_ID_NUMBER)] = (
                    params_session)
        out_query_user = self.exec_windows_commands(
            self.QUERY_EXE,
            self.SESSION_QUERY_COMMAND_USER
        )
        # 3) Отсеять только шапку [1:]
        # 4) Включить добавать побавочные параметры из команды query session
        for row in ''.join(out_query_user).split('\r\n')[1:]:
            found_match_user = (
                compiled_pattern_for_query_users_command.match(row))
            if found_match_user:
                # TODO : повторяющий код нужно оптимизировать
                day = found_match_user.group('day')
                hour = found_match_user.group('hour')
                minute = found_match_user.group('minute')
                idle_time = sum(
                    # Сумма минут существования сессии
                    map(
                        lambda x, y:PrinterAdvancedTask.to_int(x) * y,
                        [day, hour, minute],
                        [24, 60, 1]
                    )
                )
                params_store[
                    found_match_user.group(SESSION_ID_NUMBER)
                ][LOGON_TIME] = found_match_user.group(LOGON_TIME)

                params_store[
                    found_match_user.group(SESSION_ID_NUMBER)
                ][IDLE_TIME] = idle_time
        for store_item in params_store:
            self.log(u'Обновление информации о сессиях')
            # Передаем паметры сессий
            if params_store[store_item]:
                query_windows_session = QueryWindowsSession(
                    **params_store[store_item]
                )
                if params_store[store_item].get(IS_CURRENT_SESSION):
                    # CURRENT_SESSION уже добавлен в ALL_SESSION
                    self.CURRENT_SESSION = query_windows_session

                self.ALL_SESSIONS.append(query_windows_session)
        # Проброс доаолнительных сведений в окужение пользователя
        # о текущем подключении собственной сесии
        self.exec_windows_commands(
            self.SETX_EXE,
            SESSION_ID_NUMBER,
            str(self.CURRENT_SESSION.session_id_number)
        )
        self.exec_windows_commands(
            self.SETX_EXE,
            SESSION_NAME,
            self.CURRENT_SESSION.session_name
        )
        self.exec_windows_commands(
            self.SETX_EXE,
            LOGON_TIME,
            str(self.CURRENT_SESSION.logon_time)
        )
        self.exec_windows_commands(
            self.SETX_EXE,
            IDLE_TIME,
            str(self.CURRENT_SESSION.idle_time)
        )

        sys.stdout.write(
            u'Знаения сессий обновлены,'
            u'однако они не доступны в рамсках текущего терминала (cmd).\r\n'
            u'Перезапустите терминал для доступа к новым переменным.\r\n'
            u'echo %session_id_number%\r\n'
            u'echo %session_name%\r\n'
            u'echo %idle_time%\r\n'
            u'echo %logon_time%\r\n\n'
        )
        self.log(u'Проброс дополнительных переменных в окружение пользователя')

    def exec_sub_program(self, args):
        """ Получение инструкций и их выполнение """
        self.log(u'Определение под программы ')
        if self.CURRENT_SESSION.session_username not in self.get_accepts():
            self.log(
                u'Нет зарегистриванного правила для текущего пользователя\n'
            )
            return
        self.log(u'Проверка аргуметов командной строки')
        if len(args) > 2:
            self.log(u'Слишком много аргументов')
            raise AssertionError('Is too much arguments')
        if len(args) == 1:
            self.log(u'Команда не содержит инструкции --MODE')
            raise AssertionError('should not pass')
        arg = args[1]
        argv = re.match(r'-{2}(MODE)=(\w+)', arg)
        if getattr(argv, 'group', None):
            key = argv.group(1)
            value = argv.group(2)
            if key == self.MODE:
                self.log(u'Передана инструкция --MODE')
                if value == self.MODE_COMMON:
                    # Программа должна запускаться из Автозагрузки
                    # С параметром --MODE=COMMON
                    self.log(u'Запуск в объчном режиме')
                    printer_names = []
                    for flags, desc, name, comment in self.ALL_PRINTERS:
                        printer_names.append(name)

                    if self.set_default_pattern:
                        self.log(
                            u'Есть указание на смену принтера по умолчанию'
                        )
                        # Сменить принтер по умолчанию
                        for name in printer_names:
                            new_default_printer = (
                                self.set_default_pattern.match(name))
                            if new_default_printer:
                                self.log(
                                    u'Шаблон определелил принетр по умолчанию'
                                )
                                try:
                                    win32print.SetDefaultPrinter(
                                        new_default_printer.group(0))
                                except Exception as err:
                                    self.log(err.message)
                                else:
                                    self.log(
                                        u'Установил принетр по умолчанию'
                                    )

                    if self.keep_printer_info:
                        self.log(
                            u'Получил инструкцию на '
                            u'сохрание принтера по умолчанию в файл'
                        )
                        # Сохранить имя компьютера и принтер по умолчанию
                        # в файл
                        with codecs.open(
                            os.path.join(
                                self.PRINTERS_PATH,
                                self.CURRENT_SESSION.session_username,
                            ),
                            'w',
                            'utf-8-sig'
                        ) as keep_file:
                            try:
                                keep_file.write(
                                    u'{},{}'.format(
                                        self.client_name,
                                        win32print.GetDefaultPrinter().decode(
                                            'cp1251')
                                    )
                                )
                            except RuntimeError:
                                self.log(
                                    u'Принтер по умолчанию не был задан'
                                )
                                keep_file.write(
                                    u'{},{}'.format(
                                        self.client_name,
                                        u'Нет заданного принтера'
                                    )
                                )
                            else:
                                self.log(
                                    u'Принтер по умолчанию '
                                    u'сохранен в папке printers приложения '
                                )

                elif value == self.MODE_MASTER:
                    self.log(
                        u'Получена команда удаления'
                    )
                    self.delete_printers(mode=self.MODE_MASTER)

                elif value == self.MODE_VIEW:
                    self.log(
                        u'Получена команда просмотра '
                        u'текущего состояния задачи'
                    )
                    # Подпрограмма --MODE=VIEW выполняется из профиля
                    # администратора.
                    # Программа показывает последнее имя принтера
                    # после запуска программы
                    patter_pc_name_and_printer_name = re.compile(
                        """
                        (?P<ps_name>[^,]+),(?P<printer_name>.+)
                        """
                        ,re.VERBOSE
                    )
                    sys.stdout.write(
                        u'\n\n Принтеры после применения monitor_printer \n\n')
                    head = u'{:16s}{:40s}{:16s}{:15s}\n'.format(
                        u'Имя польв.,',
                        u'Имя принтера',
                        u'Имя компьют.,',
                        u'Дата и вермя'
                    )
                    sys.stdout.write(head)
                    for file_ in os.listdir(self.PRINTERS_PATH):
                        if os.path.isfile(
                                os.path.join(self.PRINTERS_PATH,file_)):
                            username = file_
                            timestamp_modify = os.path.getmtime(
                                os.path.join(self.PRINTERS_PATH,file_)
                            )
                            line = codecs.open(
                                os.path.join(self.PRINTERS_PATH,file_),
                                'r',
                            ).readline()
                            result_of_search = (
                                patter_pc_name_and_printer_name.match(line)
                            )
                            modify_time = time.strftime(
                                "%y-%m-%d %H:%M",
                                time.gmtime(timestamp_modify)
                            )
                            if result_of_search:
                                # utf-8-sig
                                # удаление опережающих 3 байтов Windows
                                ps_name = (
                                    result_of_search.group(
                                        'ps_name').decode('utf-8-sig'))
                                printer_name = result_of_search.group(
                                    'printer_name').decode('utf-8')
                                sys.stdout.write(
                                    u'{:16s}{:40s}{:16s}{:15s}\n'.format(
                                        username,
                                        printer_name,
                                        ps_name,
                                        modify_time
                                    )
                                )

                    # Программа показывает принтера помеченные на удаление
                    self.delete_printers(mode=self.MODE_VIEW)
                self.log(u'Конец программы\n')
            else:
                self.log(u'Не была вызвана ни одна подпрограмма')

    @staticmethod
    def help():

        sys.stdout.write(
            u'1) Задать параметры в setting.conf \n'
            u'2) Нужно задать один из режимов. \n'
            u'--MODE=COMMON \n'
            u'Ставиться на logon. Пробрасывает в окружение \n'
            u'пользователя переменные об номере сессии \n'
            u'--MODE=VIEW \n'
            u'Позволяет увидеть какие принетры будут удалены,'
            u'а какие остануться после применения --MASTER '
            u'--MODE=MASTER \n'
            u'Режим для пользователя с правами администратора'
            u'Очищает очереди принтеров и затем удаляет их. \n'
            )

    def get_delete_patterns(self):
        """Сбор вырожений на удаление """
        delete_patterns = []
        for option in self.parser.options(self.MAIN_SECTION):
            rule = self.parser.get(self.MAIN_SECTION, option)
            if self.parser.has_option(rule, 'delete_printers_like'):
                pattern = self.parser.get(rule, 'delete_printers_like')
                pattern = pattern.decode('utf-8')
                # Для тех правил где указана переменная с номером сессии,
                # по умолчанию будет отправляться 0
                session = self.user_name_session_dict.get(rule)
                if hasattr(session, 'session_id_number'):
                    delete_patterns.append(
                        re.compile(
                            pattern.format(
                                session_id_number=session.session_id_number,
                            ),
                            re.UNICODE
                        )
                    )
                delete_patterns.append(
                    re.compile(
                        pattern.format(
                            session_id_number='0',
                        )
                    )
                )

        return delete_patterns

    def get_status_of_printer(self, name_of_printer):
        """Метод для получения статуса принтера"""
        try:
            printer_handler = win32print.OpenPrinter(
                name_of_printer
            )
        except:
            return
        query_printer = win32print.EnumJobs(
            printer_handler, 0, -1, 1)

        if query_printer:
            first_document = query_printer[0]
            document_status = self.HANDLER_STATUS.get(
                int(first_document['Status'])
            )
        else:
            document_status = self.HANDLER_STATUS.get(
                self.STATUS_IS_READY
            )
        return document_status

    def delete_printers(self, mode):
        """
        Команда уделения(просмотра списка удаления) принтеров.

        Удаление происходит в режиме мастера.
        В режиме просмотр выводится только список принтеров и их статус.
        """
        if mode not in [self.MODE_VIEW, self.MODE_MASTER]:
            return

        delete = False
        if mode == self.MODE_MASTER:
            delete = True
            sys.stdout.write(u'\nУдаление\n')

        if mode == self.MODE_VIEW:
            sys.stdout.write(u'\nПросмотр притеров подлежащих удалению\n\n')

        patterns_delete = self.get_delete_patterns()
        about_printer_stdout_template = u'{:60s}{:16s}\n'
        sys.stdout.write(
            about_printer_stdout_template.format(
                u"Название принтера",
                u"Состояние печати"
            )
        )

        for _, _, printer, _ in self.ALL_PRINTERS:

            for pattern in patterns_delete:

                name_printer_witch_will_be_deleted = pattern.match(
                    unicode(printer.decode('cp1251'))

                )
                if name_printer_witch_will_be_deleted:

                    status_printer = self.get_status_of_printer(
                        name_printer_witch_will_be_deleted.group()
                    )
                    # Вывод состояния принтеров
                    sys.stdout.write(
                        about_printer_stdout_template.format(
                            name_printer_witch_will_be_deleted.group(),
                            status_printer
                        )
                    )
                    if delete:

                        # Очистка очереди перед удалением
                        self.log(u'Очистка очереди на принтере')
                        clear_all_jobs = self.clear_all_jobs.format(
                            name_of_printer=unicode(printer.decode('cp1251'))
                        ).encode('cp1251')
                        self.exec_windows_commands(
                            self.POWER_SHELL_EXE,
                            clear_all_jobs
                        )
                        command = self.power_shell_script.format(
                            # Магия для удаления принтеров с кирилицей
                            name_of_printer=unicode(printer.decode('cp1251'))
                        ).encode('cp1251')
                        # Удаление
                        self.log(
                            u'Удаление принтера {}'.format(
                                unicode(
                                    printer.decode('cp1251')
                                )
                            )
                        )
                        self.exec_windows_commands(
                            self.POWER_SHELL_EXE,
                            command
                        )

    def log(self, text):
        """
        Запись лога
        """
        date_time = datetime.datetime.today()
        with codecs.open(self.LOG_FILE, 'a', 'utf-8') as file_:
            file_.write(
                u'{:20}{:40}\r\n'.format(
                    date_time.strftime("%H_%M"),
                    text)
            )

if __name__ == '__main__':
    try:
        PrinterAdvancedTask(sys.argv)
    except (
            ExceptionMainSectionNotSet,
            ExceptionNotFoundSection,
            PowerShellIsNotInstalled,
            AssertionError
    ) as err:
        sys.stdout.write(err.message)
