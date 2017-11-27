import datetime
import json
import re
import os
import time
import winreg
from comtypes import COMError
from zashel.winhttp import Requests, encode, decode, LOCALPATH
from functools import partial, wraps
from math import floor


#LOCALPATH = os.path.join(os.environ["LOCALAPPDATA"], "zashel", "gapi")

# SCOPES
class SCOPE:
    """
    Scopes for Google
    """
    BLOGGER = "https://www.googleapis.com/auth/blogger"
    BLOGGER_READONLY = "https://www.googleapis.com/auth/blogger.readonly"
    BOOKS = "https://www.googleapis.com/auth/books"
    CALENDAR = "https://www.googleapis.com/auth/calendar"
    CALENDAR_READONLY = "https://www.googleapis.com/auth/calendar"
    CONTACTS = "https://www.googleapis.com/auth/contacts"
    CONTACTS_READONLY = "https://www.googleapis.com/auth/contacts.readonly"
    DRIVE = "https://www.googleapis.com/auth/drive"
    DRIVE_APPDATA = "https://www.googleapis.com/auth/drive.appdata"
    DRIVE_FILE = "https://www.googleapis.com/auth/drive.file"
    DRIVE_METADATA = "https://www.googleapis.com/auth/drive.metadata"
    DRIVE_METADATA_READONLY = "https://www.googleapis.com/auth/drive.metadata.readonly"
    DRIVE_PHOTOS_READONLY = "https://www.googleapis.com/auth/drive.photos.readonly"
    DRIVE_READONLY = "https://www.googleapis.com/auth/drive.readonly"
    DRIVE_SCRIPTS = "https://www.googleapis.com/auth/drive.scripts"
    GMAIL = "https://mail.google.com/"
    GMAIL_COMPOSE = "https://www.googleapis.com/auth/gmail.compose"
    GMAIL_INSERT = "https://www.googleapis.com/auth/gmail.insert"
    GMAIL_LABELS = "https://www.googleapis.com/auth/gmail.labels"
    GMAIL_METADATA = "https://www.googleapis.com/auth/gmail.metadata"
    GMAIL_MODIFY = "https://www.googleapis.com/auth/gmail.modify"
    GMAIL_READONLY = "https://www.googleapis.com/auth/gmail.readonly"
    GMAIL_SEND = "https://www.googleapis.com/auth/gmail.send"
    GMAIL_SETTINGS_BASIC = "https://www.googleapis.com/auth/gmail.settings.basic"
    GMAIL_SETTINGS_SHARING = "https://www.googleapis.com/auth/gmail.settings.sharing"
    SIGN_IN_PROFILE = "profile"
    SIGN_IN_EMAIL = "email"
    SIGN_IN_OPENID = "openid"
    SPREADSHEETS = "https://www.googleapis.com/auth/spreadsheets"
    SPREADSHEETS_READONLY = "https://www.googleapis.com/auth/spreadsheets.readonly"
    PLUS_LOGIN = "https://www.googleapis.com/auth/plus.login"
    PLUS_ME = "https://www.googleapis.com/auth/plus.me"
    USER_ADDRESSES_READ = "https://www.googleapis.com/auth/user.addresses.read"
    USER_BIRTHDAY_READ = "https://www.googleapis.com/auth/user.emails.read"
    USER_PHONENUMBERS_READ = "https://www.googleapis.com/auth/user.phonenumbers.read"
    USERINFO_EMAIL = "https://www.googleapis.com/auth/userinfo.email"
    USERINFO_PROFILE = "https://www.googleapis.com/auth/userinfo.profile"


# API PATHS
DRIVE = "https://www.googleapis.com/drive/v3"
TEAMDRIVES = DRIVE + "/teamdrives"
FILESDRIVE = DRIVE + "/files"
FILEDRIVE = FILESDRIVE + "/{}"
FILEPERMISSIONS = FILEDRIVE + "/permissions"
FILEDOWNLOAD = "https://drive.google.com/open"
COPYFILE = FILESDRIVE + "/{}/copy"

SCRIPTS = "https://script.googleapis.com/v1/scripts/{}:run"

SHEETS = "https://sheets.googleapis.com/v4/spreadsheets"
SHEET_VALUES = SHEETS + "/{}/values/{}"
SHEET_APPEND = SHEET_VALUES + ":append"
SHEET_CLEAR = SHEET_VALUES + ":clear"
SHEET_BATCHUPDATE = SHEETS + "/{}:batchUpdate"


if not os.path.exists(LOCALPATH):
    os.makedirs(LOCALPATH)


QUERYTIMEOUT = 5


reghandle = winreg.CreateKey(winreg.HKEY_CURRENT_USER, "Software\\Microsoft\\Windows\\CurrentVersion\\" +\
                             "Internet Settings\\ZoneMap\\Domains\\googleapis.com\\www")
winreg.SetValueEx(reghandle, "https", 0, winreg.REG_DWORD, 0)
winreg.FlushKey(reghandle)
winreg.CloseKey(reghandle)
del(reghandle)


class DriveNotFoundError(Exception):
    pass

class TeamDriveNotFoundError(Exception):
    pass

class FileNotOpenError(Exception):
    pass

class SheetNotFoundError(Exception):
    pass

class SheetError(Exception):
    pass

class SpreadsheetNotFoundError(FileNotFoundError):
    pass

class Apps(object):
    """
    Base objects for google apps. To be inherited.
    """
    def __init__(self, gapi, name):
        """
        Initializes the app object
        :param gapi: gapi.GoogleAPI instance
        :param name: Name of the item accesed by app
        """
        self._api = gapi
        self._name = name
        self._app_name = ""

    @property
    def api(self):
        """
        Gives GoogleAPI assigned to.
        """
        return self._api

    @property
    def app_name(self):
        """
        Returns the name of the App.
        """
        return self._app_name

    @property
    def name(self):
        """
        Returns the name of the object accesed by the app.
        """
        return self._name

    def __getattribute__(self, item):
        """
        Override of object.__getattribute__ method to access data in the api directly.
        :param item: item to access to.
        :return: value of item if found.
        """
        try:
            return object.__getattribute__(self, item)
        except AttributeError:
            return self.__getattr__(item)

    def __getattr__(self, item):
        """
        __getattr__ method to get gapi.GoogleAPI methods by self.app_name+"_"+item.
        :param item: name of the item in the self.gapi to be searched for.
        :return: value or method searched.
        """
        method = "_".join((self.app_name, item))
        method = self.api.__getattribute__(method)
        return partial(method, name=self.name)


class Spreadsheets(Apps):
    """
    Class of Spreadsheet to get all "spreadsheet_" methods in gapi.GoogleAPI
    """
    class Sheet(Apps):
        """
        Sheet class to access each sheet as a list of lists.
        """
        class Row(list):
            """
            Row class to implement modification of data
            """
            def __init__(self, index, values, sheet):
                """
                Initializes Row with given values. It does not create a row remotely.
                :param index: location begining with 0 in the gieven gapi.Spreadsheet.Sheet
                :param values: list of data assigned to Row
                :param sheet: gapi.Spreadsheet.Sheet instance in which Row is defiend.
                """
                super().__init__(values)
                self._index = index
                self._sheet = sheet

            def __getitem__(self, key):
                """
                list.__getitem__ overriding to get remote data if it is locally a function
                :param key: index of key to search for.
                :return: value or method searched.
                """
                item = list.__getitem__(self, key)
                if not isinstance(item, list):
                    if list.__getitem__(self, key).startswith("="):
                        list.__setitem__(self, key,
                                         self.spreadsheet.get_range(self.spreadsheet.get_range_name(key+1,
                                                                                                    self._index+1)))
                return list.__getitem__(self, key)

            def __setitem__(self, key, value):
                """
                Overriding of list.__setitem__ to set value both loacally and remotely
                :param key: key of value to set
                :param value: new velue to set
                :return: None
                """
                self._sheet.update_range(self._sheet.get_range_name(int(key)+1, int(self._index)+1), [[value]])
                if isinstance(value, str) and value.startswith("="):
                    value = "Cargando..."
                    while value not in ("Cargando...", "Loading..."):
                        time.sleep(1)
                        value = self.spreadsheet.get_range(self.spreadsheet.get_range_name(key+1, self._index+1))
                if key >= len(self):
                    super().extend(["" for i in range(key-len(self))]+[value])
                super().__setitem__(key, value)

            def __delitem__(self, key):
                self._sheet.clear_range(self._sheet.get_range_name(key + 1, self._index + 1))
                super().__setitem__(key, "")

            @property
            def range(self):
                return self.sheet_name+"!"+self.spreadsheet.get_range_name(1, self.row_index+1) + \
                       ":"+self.spreadsheet.get_range_name(len(self)+1, self.row_index+1)

            @property
            def row_index(self):
                return self._index

            @property
            def sheet_name(self):
                return self._sheet.sheet_name

            @property
            def spreadsheet(self):
                return self._sheet.spreadsheet

            def update(self, values):
                cols, rows = self._sheet.get_sheet_dimensions()
                self._sheet.update_range(self._sheet.get_range_name(1, self._index + 1)+":"+
                                         self._sheet.get_range_name(1, cols),
                                         [values])

        def __init__(self, sheet_name, gapi, name, spreadsheet):
            Apps.__init__(self, gapi, name)
            self._app_name = "spreadsheet"
            self._sheet_name = sheet_name
            self._spreadsheet = spreadsheet
            Apps.__getattribute__(self, "api").spreadsheet_open_sheet(self.sheet_name, name=self.name)
            self._iter_index = 0

        @property
        def sheet_name(self):
            return self._sheet_name

        @property
        def spreadsheet(self):
            return self._spreadsheet

        def __getattr__(self, item):
            Apps.__getattribute__(self, "api").spreadsheet_open_sheet(self.sheet_name, name=self.name)
            return Apps.__getattr__(self, item)

        def __getitem__(self, key):
            cols, rows = Apps.__getattribute__(self, "api").spreadsheet_get_sheet_dimensions(self.sheet_name,
                                                                                             name=self.name)
            Apps.__getattribute__(self, "api").spreadsheet_open_sheet(self.sheet_name, name=self.name)
            if isinstance(key, int):
                if key < 0:
                    key = rows + key
                if key < rows and key >= 0:
                    return self.row(key,
                                    self.get_range("A" + str(key + 1) + ":" + self.get_range_name(cols, key + 1))[0])
                else:
                    raise IndexError()
            elif isinstance(key, slice):
                init = key.start
                end = key.stop
                if init is None:
                    init = 0
                if end is None or end >= rows:
                    end = rows - 1
                if init < 0:
                    init = rows + init
                if end < 0:
                    end = rows + end
                if init < end < rows and end >= init >= 0:
                    return self.row(key,
                                    self.get_range("A" + str(init + 1) + ":" + self.get_range_name(cols, end + 1)))
                else:
                    raise IndexError()


        def __setitem__(self, key, values):
            assert isinstance(values, list)
            values = [values]
            cols, rows = Apps.__getattribute__(self, "api").spreadsheet_get_sheet_dimensions(self.sheet_name,
                                                                                             name=self.name)
            Apps.__getattribute__(self, "api").spreadsheet_open_sheet(self.sheet_name, name=self.name)
            if key < 0:
                key = rows + key
            if key < rows and key >= 0:
                return self.update_range("A"+str(key+1)+":"+self.get_range_name(cols, key+1), values)
            else:
                raise IndexError()

        def __repr__(self):
            return str(self.get_sheet_values())

        def __iter__(self):
            return self.__next__()

        def __next__(self):
            data = self.get_sheet_values()
            for index, item in enumerate(data):
                yield self.row(index, item)
            raise StopIteration()

        def append_row(self, values):
            data = self.append_rows([values])
            return data #TODO Verify data
            """
            if len(data) > 0:
                return data[0]
            """

        def append_rows(self, values):
            updated_range = self.spreadsheet.append_rows("A1", values)
            return updated_range
            """ #TODO Review
            if isinstance(updated_range, str):
                sheet_name, n_range = updated_range.split("!")
                if ":" in n_range:
                    init, fin = n_range.split(":")
                else:
                    init = fin = n_range
                init = int(init.strip("ABCDEFGHIJKLMNOPQRSTUVWXYZ"))
                fin = int(fin.strip("ABCDEFGHIJKLMNOPQRSTUVWXYZ"))
                final = list()
                for row_index in range(int(init) - 1, int(fin)):
                    try:
                        row = self[row_index]
                    except KeyError:
                        return updated_range #TODO Pendiente de verificar error raro en este punto en uno de cada 10 casos
                    final.append(row)
                return final
            else:
                print(updated_range)
                return updated_range
            """

        def get_sheet_values(self):
            return self.spreadsheet.get_sheet_values(self.sheet_name, name=self.name)

        def row(self, key, range):
            return Spreadsheets.Sheet.Row(key, range, self)

        def update_rows(self, location, values):
            updated_range = self.spreadsheet.append_rows(location, values, insert_data="OVERWRITE")
            return values

    def __init__(self, gapi, name):
        super().__init__(gapi, name)
        self._app_name = "spreadsheet"

    def __getitem__(self, item):
        return self.sheet(item)

    def __setitem__(self, item, values=None):
        if values:
            if isinstance(values, Spreadsheets.Sheet):
                values = values.get_rows
            assert isinstance(values, list)
            assert all([isinstance(value, list) for value in values])
        try:
            sheet = self.sheet(item)
        except (SheetNotFoundError, SheetError):
            pass
        else:
            del(self[item])
        self.add_sheet(item)
        sheet = self.sheet(item)
        if values:
            sheet.append_rows(values)

    def __delitem__(self, item):
        try:
            sheet = self.sheet(item)
        except (SheetNotFoundError, SheetError):
            pass
        else:
            self.delete_sheet(item)

    @property
    def cells(self):
        return self.get_total_cells()

    @property
    def developer_metadata(self):
        return self.resource["developerMetadata"]

    @property
    def id(self):
        return self.resource["spreadsheetId"]

    @property
    def named_ranges(self):
        return self.resource["namedRanges"]

    @property
    def properties(self):
        return self.resource["properties"]

    @property
    def resource(self):
        _id = self.api._files_get_id_by_name(self.name)
        return self.api._opened_files[_id]

    @property
    def sheets(self):
        return self.resource["sheets"]

    @property
    def url(self):
        return self.resource["spreadsheetUrl"]

    def sheet(self, sheet):
        return Spreadsheets.Sheet(sheet, self.api, self.name, self)


class DebugRequests(Requests):
    def __init__(self, debug):
        Requests.__init__(self)
        self.debug = debug

    def request(self, method, url, *, data=None, json=None, headers=None, get=None):
        request = Requests.request(self, method, url, data=data, json=json, headers=headers, get=get)
        if self.debug:
            with open("log.txt", "a") as f:
                fp = partial(print, file=f)
                fp(method, url)
                fp(datetime.datetime.now())
                fp(f"Data: {data}")
                fp(f"JSON: {json}")
                fp(f"Headers: {headers}")
                fp(f"Get: {get}")
                fp(f"Result: {request}")
                fp()
        return request


class GoogleAPI(DebugRequests):
    def __init__(self, *, scopes, secret_file=None, secret_data=None, password=None, debug=False):
        if secret_file is not None:
            assert os.path.exists(secret_file)
        DebugRequests.__init__(self, debug)
        self.scopes = scopes
        self.secret_file = secret_file
        if password is None:
            password = " "
        self.secret_data = encode(password, json.dumps(secret_data))
        self._teamdrives = dict()
        self._drives = dict()
        self._files = None
        self._drive_id = None
        self._is_teamdrive = False
        self._lastsqueries = dict()
        self._opened_files = dict()
        self._file_id = None
        self._opened_sheet = None
        self.get = partial(self.request, "GET")
        self.post = partial(self.request, "POST")
        self.put = partial(self.request, "PUT")
        self.delete = partial(self.request, "DELETE")
        self.patch = partial(self.request, "PATCH")
        self.head = partial(self.request, "HEAD")

    @property
    def drives(self):
        pass

    @property
    def files(self):
        self.files_list()
        """if self._last_timeout("files") is True:
            self.files_list()
            self._update_timeout("files")
        """
        return self._files

    @property
    def spreadsheets(self):
        final = list()
        for item in self.files:
            if self._files[item]["mimeType"] == "application/vnd.google-apps.spreadsheet":
                final.append(item)
        return final

    @property
    def teamdrives(self):
        if self._last_timeout("teamdrives") is True:
            self._teamdrives_list()
            self._update_timeout("teamdrives")
        return list(self._teamdrives.keys())

    # TIMEOUT
    def _last_timeout(self, key):
        if key not in self._lastsqueries:
            self._lastsqueries[key] = None
        last = self._lastsqueries[key]
        if last is None or last <= datetime.datetime.now():
            return True
        else:
            return False

    def _update_timeout(self, key):
        self._lastsqueries[key] = datetime.datetime.now() + datetime.timedelta(seconds=QUERYTIMEOUT)

    # LOGIN
    def login(self, *, password=None):
        if password is None:
            password = " "
        secret_data = json.loads(decode(password, self.secret_data))
        self.oauth2(self.scopes, json_file=self.secret_file, secret_data=secret_data)

    def logout(self):
        self.oauth2_logout()

    # REQUESTING
    def request(self, method, url, *, data=None, json=None, headers=None, get=None):
        while True:
            try:
                dataX = DebugRequests.request(self, method, url, data=data, json=json, headers=headers, get=get) # Soberana CAGADA
                if not int(self.status_code) in (500, 503, 504, 429, 408):
                    break
                else:
                    time.sleep(1)
            except COMError:
                time.sleep(1)
        return dataX

    # DRIVES
    def _list_drives(self):
        pass

    # TEAMDRIVES
    def _teamdrives_list(self):
        while True:
            try:
                self.get(TEAMDRIVES)
                teamdrives = json.loads(self.text)
            except json.decoder.JSONDecodeError:
                time.sleep(1)
                continue
            else:
                self._teamdrives = dict()
                for item in teamdrives["teamDrives"]:
                    self._teamdrives[item["name"]] = item["id"]
                break

    def teamdrive_open(self, name):
        teamdrives = self.teamdrives
        if name not in teamdrives:
            raise TeamDriveNotFoundError()
        id = self._teamdrives[name]
        self._drive_id = id
        self._is_teamdrive = True
        self._lastsqueries["files"] = None #Set timeout to None to check new teamdrive contents

    # FILES
    def _files_get_id_by_name(self, name):
        if name is not None:
            if name in self.files:
                self._file_id = self._files[name]["id"]
            else:
                raise FileNotFoundError()
        return self._file_id

    def file_copy(self, origin, new_name):
        files = self.files
        if origin in files and new_name not in files:
            file_id = self._files[origin]["id"]
            self.post(COPYFILE.format(file_id), get={"supportsTeamDrives": self._is_teamdrive},
                      json={"name": new_name})
        else:
            raise FileNotFoundError

    def file_crete_premission(self, name, email, *, perm_type="user", role="writer", send_email=False,
                              body=None, where=None, is_teamdrive=False):
        """
        Sets a new permission in indicated file.
        :param name: Name of file to set a new permission.
        :param email: User email to set new permissions to.
        :param perm_type: Type of permission, user by default.
        :param role: Role to set to user. "writer" to default.
        :param send_email: Whether to send email or not. False by default.
        :param body: Body of emails to send.
        :param where: List of items to search name in. None by default.
        :param is_teamdrive: Whether where or file is teamdrive or not.
        :return: None
        """
        if where is None:
            where = self.files
        if name in where:
            file_id = self._files[name]["id"]
            self.port(FILEPERMISSIONS.format(file_id), get={"emailMessage": body,
                                                             "sendNotificationEmail": send_email,
                                                             "supportsTeamDrives": is_teamdrive},
                      data = {"role": role,
                              "type": perm_type,
                              "emailAddress": email})

    def file_download(self, name, where=None):
        if where is None:
            where = self.files
        if name in where:
            file_id = self._files[name]["id"]
            self.get(FILEDRIVE.format(file_id), get={"supportsTeamDrives": self._is_teamdrive,
                                                     "alt": "media"})
        else:
            raise FileNotFoundError()
        tempfile = os.path.join(self.tempfolder.name, name)
        with open(tempfile, "wb") as f:
            f.write(bytes(self.body))
        return tempfile

    def files_list(self, *, drive_name=None, is_teamdrive=False):
        class Files(dict):
            def __init__(self, gapi, drive_name, is_teamdrive=False):
                self.gapi = gapi
                self.drive_name = drive_name
                self.is_teamdrive = is_teamdrive
                self.last_loaded = datetime.datetime.now() - datetime.timedelta(minutes=10)
                self.load(True)
                dict.__init__(self)
            def __iter__(self):
                self.load()
                return dict.__iter__(self)
            def __getitem__(self, item):
                if item in self:
                    return dict.__getitem__(self, item)
                else:
                    raise KeyError()
            def __contains__(self, filename):
                if dict.__contains__(self, filename):
                    return True
                else:
                    self.load(True)
                    return dict.__contains__(self, filename)
            def load(self, force=False):
                if force is True or (force is False and
                                     self.last_loaded <= datetime.datetime.now() - datetime.timedelta(minutes=5)):
                    print("Loading Files")
                    if self.drive_name is not None:
                        if self.is_teamdrive is True:  # TODO
                            self.gapi._teamdrives_list()
                            drives = self.gapi._teamdrives
                        else:
                            pass
                            drives = self.gapi._drives
                        if drive_name not in drives:
                            raise is_teamdrive and TeamDriveNotFoundError() or DriveNotFoundError()
                        self.gapi._drive_id = drives[drive_name]
                        self.gapi._is_teamdrive = is_teamdrive
                    get = dict()
                    if self.gapi._is_teamdrive is True:
                        get.update({"corpora": "teamDrive",
                                    "includeTeamDriveItems": "true",
                                    "supportsTeamDrives": "true",
                                    "teamDriveId": self.gapi._drive_id})
                    get.update({"pageSize": 1000})
                    while True:
                        try:
                            self.gapi.get(FILESDRIVE, get=get)
                            data = json.loads(self.gapi.text)
                        except json.decoder.JSONDecodeError:
                            time.sleep(1)
                            continue
                        else:
                            if "files" in data:
                                for item in data["files"]:
                                    self[item["name"]] = item
                            if "nextPageToken" in data:
                                get.update({"pageToken": data["nextPageToken"]})
                                continue
                            break
                    self.last_loaded = datetime.datetime.now()
        if self._files is None or self._files.drive_name != drive_name:
            self._files = Files(self, drive_name, is_teamdrive)
        return self._files

    def _files_open(self, path, returner, name, where=None, *, args=None, kwargs=None):
        if args is None:
            args = list()
        if kwargs is None:
            kwargs = dict()
        if where is None:
            where = self.files
        if name in where:
            while True:
                try:
                    self._files_get_id_by_name(name)
                    if self._file_id not in self._opened_files:
                        self.get(path + "/" + str(self._file_id))
                        if self.status_code == 200:
                            self._opened_files[self._file_id] = json.loads(self.text)
                        else:
                            continue
                except json.decoder.JSONDecodeError:
                    time.sleep(1)
                    continue

                else:
                    return returner(self, name, *args, **kwargs)
        else:
            raise FileNotFoundError()

    # SCRIPTS
    def script(self, script_id, function, parameters, dev_mode=False):
        data = {"function": function,
                "parameters": parameters,
                "devMode": dev_mode}
        while True:
            try:
                self.post(SCRIPTS.format(script_id), json=data)
                data = json.loads(self.text)
            except json.decoder.JSONDecodeError:
                time.sleep(1)
                continue
            else:
                return data

    # SPREADSHEETS
    def spreadsheet_add_sheet(self, sheetname, *, name=None, rows=1, columns=3):
        self._files_get_id_by_name(name)
        if self._file_id is not None:
            data = {"includeSpreadsheetInResponse": "true",
                    "requests": [{"addSheet": {"properties": {"title": sheetname,
                                                              "gridProperties": {"rowCount": rows,
                                                                                 "columnCount": columns}}}},
                                 ]
                    }
            while True:
                try:
                    self.post(SHEET_BATCHUPDATE.format(self._file_id), json=data)
                    data = json.loads(self.text)
                except json.decoder.JSONDecodeError:
                    time.sleep(1)
                    continue
                else:
                    if "updatedSpreadsheet" in data:
                        self._opened_files[self._file_id].update(data["updatedSpreadsheet"])
                    break
            return self.spreadsheet_open_sheet(sheetname, name=name)
        else:
            raise FileNotOpenError()

    def spreadsheet_append_row(self, _range, values, *, name=None, input_option="USER_ENTERED", insert_data="INSERT_ROWS"):
        self._files_get_id_by_name(name)
        _range = self.spreadsheet_check_range(_range, name=name)
        return self.spreadsheet_append_rows(_range, [values], name=name, input_option=input_option, insert_data=insert_data)

    def spreadsheet_append_rows(self, _range, values, *, name=None,
                                input_option="USER_ENTERED", insert_data="INSERT_ROWS"):
        """
        Appends or replaces given rows in given range. In case of appending, it is appended to the end of the table.
        :param _range: Range in "A1" notation
        :param values: List of lists of values, each list being a new row
        :param name: name of the spreadsheet to append. Opened sheet by default
        :param input_option: how data may be processed, "USER_ENTERED" by default, "RAW" to be given if data may be
                            included as is
        :param insert_data: how data will be inserted, "INSERT_ROWS" by default, "OVERWRITE" in case it would be
                            updated
        :return: The updated range in "A1" notation
        """
        self._files_get_id_by_name(name)
        _range = self.spreadsheet_check_range(_range, name=name)
        if self._file_id is not None:
            while True:
                try:
                    self.post(SHEET_APPEND.format(self._file_id, _range), get={"valueInputOption": input_option,
                                                                               "insertDataOption": insert_data,
                                                                               "includeValuesInResponse": "true"},
                              json={"range": _range, "values": values})
                    data = json.loads(self.text)
                except json.decoder.JSONDecodeError:
                    time.sleep(1)
                    continue
                else:
                    with open("data.json", "w") as f:
                        f.write(json.dumps(data))
                    if "updates" in data:
                        updated_range = data["updates"]["updatedRange"]
                        return updated_range
                    else:
                        return data
        else:
            raise FileNotOpenError()

    def spreadsheet_clear_range(self, range, *, name=None):
        self._files_get_id_by_name(name)
        range = self.spreadsheet_check_range(range, name=name)
        if self._file_id is not None:
            self.post(SHEET_CLEAR.format(self._file_id, range))
        else:
            raise FileNotOpenError()

    def spreadsheet_delete_sheet(self, sheetname, *, name=None):
        self._files_get_id_by_name(name)
        if self._file_id is not None:
            file_data = self._opened_files[self._file_id]["sheets"]
            sheet_ids = dict([(item["properties"]["title"], item["properties"]["sheetId"]) for item in file_data])
            if sheetname in sheet_ids:
                sheet_id = sheet_ids[sheetname]
                data = {"includeSpreadsheetInResponse": "true",
                        "requests": [{"deleteSheet": {"sheetId": sheet_id}},
                                     ]
                        }
                while True:
                    try:
                        self.post(SHEET_BATCHUPDATE.format(self._file_id), json=data)
                        data = json.loads(self.text)
                    except json.decoder.JSONDecodeError:
                        time.sleep(1)
                        continue
                    else:
                        if "updatedSpreadsheet" in data:
                            self._opened_files[self._file_id].update(data["updatedSpreadsheet"])
                        break
            else:
                raise SheetNotFoundError()
        else:
            raise FileNotOpenError()

    def spreadsheet_check_range(self, range, *, name=None, autoopen=True):
        final = range
        if "!" in range:
            sheet, range = range.split("!")
            if not self._opened_sheet or autoopen is True:
                self.spreadsheet_open_sheet(sheet)
            elif self._opened_sheet != sheet and autoopen is False:
                raise SheetError()
        elif self._opened_sheet:
            final = "!".join((self._opened_sheet, range))
        return final

    def spreadsheet_get_sheet_dimensions(self, sheet_name=None, *, name=None, autoopen=True):
        self._files_get_id_by_name(name)
        self.spreadsheet_open(name=name)
        if not self._opened_sheet or autoopen is True:
            self.spreadsheet_open_sheet(sheet_name)
        elif self._opened_sheet != sheet_name and autoopen is False:
            raise SheetError()
        sheets = self._opened_files[self._file_id]["sheets"]
        for sheet in sheets:
            if sheet["properties"]["title"] == sheet_name:
                grid = sheet["properties"]["gridProperties"]
                return (grid["columnCount"], grid["rowCount"])
        raise SheetNotFoundError

    def spreadsheet_get_sheet_values(self, sheet_name=None, *, name=None, autoopen=True):
        cols, rows = self.spreadsheet_get_sheet_dimensions(sheet_name, name=name, autoopen=autoopen)
        return self.spreadsheet_get_range(sheet_name+"!A1:"+self.spreadsheet_get_range_name(cols, rows))

    def spreadsheet_get_range_name(self, column, row, **kwargs):
        letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        subcolumn = column - 1
        final = str()
        while subcolumn >= len(letters):
            final = letters[subcolumn%len(letters)]+final
            subcolumn = floor(subcolumn/len(letters)) - 1
        final = letters[subcolumn] + final
        return final + str(row)

    def spreadsheet_get_range_by_name(self, range, **kwargs):
        letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        column_name, row, void = re.split(r"([0-9]+)", range)
        column = int()
        for index, letter in enumerate(column_name):
            column += len(column_name) * (letters.index(letter) + 1)
        return column, int(row)

    def spreadsheet_get_range(self, range, *, name=None):
        self._files_get_id_by_name(name)
        range = self.spreadsheet_check_range(range, name=name)
        if self._file_id is not None:
            while True:
                try:
                    data = json.loads(self.get(SHEET_VALUES.format(self._file_id, range)))
                except json.decoder.JSONDecodeError:
                    time.sleep(1)
                    continue
                else:
                    break
            if "values" in data:
                return data["values"]
            else:
                return [[]]
        else:
            raise FileNotOpenError()

    def spreadsheet_get_total_cells(self, *, name=None):
        self._files_get_id_by_name(name)
        while True:
            try:
                sheets = self._opened_files[self._file_id]["sheets"]
            except KeyError:
                time.sleep(1) # Take a break
                pass
            else:
                time.sleep(0.5)
                break
        return sum([sheet["properties"]["gridProperties"]["columnCount"]*sheet["properties"]["gridProperties"]["rowCount"]
                    for sheet in sheets])

    def spreadsheet_open(self, name=None, **kwargs):
        if name is None and "name" in kwargs:
            name = kwargs["name"]
        elif name is None:
            raise FileNotFoundError()
        return self._files_open(SHEETS, Spreadsheets, name, self.spreadsheets)

    def spreadsheet_open_sheet(self, sheet_name, *, name=None, just_open=False):
        self._files_get_id_by_name(name)
        while True:
            try:
                sheets = self._opened_files[self._file_id]["sheets"]
            except KeyError:
                if self._file_id in self._opened_files:
                    del(self._opened_files[self._file_id])
                    just_open = True
                self.spreadsheet_open(name)
                self._files_get_id_by_name(name)
                time.sleep(1)
                pass
            else:
                time.sleep(1)
                break
        for sheet in sheets:
            if sheet["properties"]["title"] == sheet_name:
                self._opened_sheet = sheet_name
                return
        if not just_open:
            if self._file_id in self._opened_files:
                del(self._opened_files[self._file_id])
            self.spreadsheet_open(name)
            return self.spreadsheet_open_sheet(sheet_name, name=name, just_open=True)
        raise SheetNotFoundError(sheet_name)

    def spreadsheet_update_range(self, range, values, *, name=None):
        self._files_get_id_by_name(name)
        range = self.spreadsheet_check_range(range, name=name)
        if self._file_id is not None:
            while True:
                try:
                    data = json.loads(self.put(SHEET_VALUES.format(self._file_id, range),
                                               get={"valueInputOption": "USER_ENTERED"},
                                               json={"range": range, "values": values}))
                except json.decoder.JSONDecodeError:
                    time.sleep(1)
                    continue
                else:
                    break
            if "values" in data:
                return data["values"]
            else:
                return [[]]
        else:
            raise FileNotOpenError()


class SheetList(list):
    def __init__(self, sheet):
        assert isinstance(sheet, Spreadsheets.Sheet)
        data = sheet.get_sheet_values()
        if data[-1] == []:
            del(data[-1])
        if all([len(item)==1 for item in data]):
            data = [item[0] for item in data]
            cutted = True
        else:
            cutted = False
        super().__init__(data)
        self.sheet = sheet
        self.cutted = cutted

    def append(self, item):
        assert not isinstance(item, tuple)
        assert not isinstance(item, dict)
        if isinstance(item, list) and len(item)!=0 and item not in self:
            self.sheet.append_row(item)
            list.append(list(self), item)
        elif isinstance(item, list) and item not in self:
            self.sheet.append_row(item)
            if self.cutted:
                list.append(list(self), item[0])
            else:
                list.append(list(self), item)
        elif item not in self:
            self.sheet.append_row([item])
            list.append(list(self), item)

    def update(self, new_list):
        self.clear()
        self.extend(new_list)
