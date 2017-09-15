import datetime
import json
import re
import os
import time
import winreg
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

class Apps(object):
    """
    Base objects for google apps
    """
    def __init__(self, gapi, name):
        self._api = gapi
        self._name = name
        self._app_name = ""

    @property
    def api(self):
        return self._api

    @property
    def app_name(self):
        return self._app_name

    @property
    def name(self):
        return self._name

    def __getattribute__(self, item):
        try:
            return object.__getattribute__(self, item)
        except AttributeError:
            return self.__getattr__(item)

    def __getattr__(self, item):
        method = "_".join((self.app_name, item))
        method = self.api.__getattribute__(method)
        return partial(method, name=self.name)


class Spreadsheets(Apps):
    class Sheet(Apps):
        class Row(list):
            def __init__(self, index, values, sheet):
                super().__init__(values)
                self._index = index
                self._sheet = sheet

            def __getitem__(self, key):
                item = list.__getitem__(self, key)
                if not isinstance(item, list):
                    if list.__getitem__(self, key).startswith("="):
                        list.__setitem__ (self, key,
                                          self.spreadsheet.get_range(self.spreadsheet.get_range_name(key+1,
                                                                                                     self._index+1)))
                return list.__getitem__(self, key)

            def __setitem__(self, key, value):
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
                return self.sheet_name+"!"+self.spreadsheet.get_range_name(1, self.index+1) + \
                       ":"+self.spreadsheet.get_range_name(len(self)+1, self.index+1)

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
            if key < 0:
                key = rows + key
            if key < rows and key >= 0:
                return self.row(key,
                                self.get_range("A" + str(key + 1) + ":" + self.get_range_name(cols, key + 1))[0])
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
            return str(self.get_sheet_values(self.sheet_name, name=self.name))

        def __iter__(self):
            return self.__next__()

        def __next__(self):
            data = self.get_sheet_values(self.sheet_name, name=self.name)
            for index, item in enumerate(data):
                yield self.row(index, item)
            raise StopIteration()

        def append_row(self, values):
            data = self.append_rows(values)
            if len(data) > 0:
                return data[0]

        def append_rows(self, values):
            updated_range = self.spreadsheet.append_rows("A1", values)
            sheet_name, n_range = updated_range.split("!")
            init, fin = n_range.split(":")
            init = int(init.strip("ABCDEFGHIJKLMNOPQRSTUVWXYZ"))
            fin = int(fin.strip("ABCDEFGHIJKLMNOPQRSTUVWXYZ"))
            final = list()
            for row_index in range(int(init) - 1, int(fin)):
                row = self[row_index]
                final.append(row)
            return final

        def row(self, key, range):
            return Spreadsheets.Sheet.Row(key, range, self)

    def __init__(self, gapi, name):
        super().__init__(gapi, name)
        self._app_name = "spreadsheet"

    def __getitem__(self, item):
        try:
            return self.sheet(item)
        except SheetNotFoundError:
            raise KeyError

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
            sheet.append_rows("a1", values)

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

    def sheet(self, sheet):
        return Spreadsheets.Sheet(sheet, self.api, self.name, self)

#Decorators
def updatedSpreadsheet(function):
    @wraps(function)
    def wrapper(self, *args, **kwargs):
        result = function(self, *args, **kwargs)
        try:
            data = json.loads(self.text)
        except json.decoder.JSONDecodeError:
            pass
        else:
            if "updatedSpreadsheet" in data:
                self._opened_files[self._file_id].update(data["updatedSpreadsheet"])
        return result
    return wrapper

class GoogleAPI(Requests):
    def __init__(self, *, scopes, secret_file=None, secret_data=None, password=None):
        if secret_file is not None:
            assert os.path.exists(secret_file)
        Requests.__init__(self)
        self.scopes = scopes
        self.secret_file = secret_file
        if password is None:
            password = " "
        self.secret_data = encode(password, json.dumps(secret_data))
        self._teamdrives = dict()
        self._drives = dict()
        self._files = dict()
        self._drive_id = None
        self._is_teamdrive = False
        self._lastsqueries = dict()
        self._opened_files = dict()
        self._file_id = None
        self._opened_sheet = None

    @property
    def drives(self):
        pass

    @property
    def files(self):
        if self._last_timeout("files") is True:
            self.files_list()
            self._update_timeout("files")
        return list(self._files.keys())

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

    # DRIVES
    def _list_drives(self):
        pass

    # TEAMDRIVES
    def _teamdrives_list(self):
        self.get(TEAMDRIVES)
        teamdrives = json.loads(self.text)
        self._teamdrives = dict()
        for item in teamdrives["teamDrives"]:
            self._teamdrives[item["name"]] = item["id"]

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

    def file_copy(self, origin, new_name):
        files = self.files
        if origin in files and new_name not in files:
            file_id = self._files[origin]["id"]
            self.post(COPYFILE.format(file_id), get={"supportsTeamDrives": self._is_teamdrive},
                      json={"name": new_name})
        else:
            raise FileNotFoundError

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
        if drive_name is not None:
            if is_teamdrive is True:
                self._teamdrives_list()
                drives = self._teamdrives
            else:
                self._list_drives()
                drives = self._drives
            if drive_name not in drives:
                raise is_teamdrive and TeamDriveNotFoundError() or DriveNotFoundError()
            self._drive_id = drives[drive_name]
            self._is_teamdrive = is_teamdrive
        get = dict()
        if self._is_teamdrive is True:
            get.update({"corpora": "teamDrive",
                        "includeTeamDriveItems": "true",
                        "supportsTeamDrives": "true",
                        "teamDriveId": self._drive_id})
        get.update({"pageSize": 1000})
        self._files = dict()
        while True:
            self.get(FILESDRIVE, get=get)
            data = json.loads(self.text)
            if "files" in data:
                for item in data["files"]:
                    self._files[item["name"]] = item
            if "nextPageToken" in data:
                get.update({"pageToken": data["nextPageToken"]})
                continue
            break
        return self._files

    def _files_open(self, path, returner, name, where=None, *, args=None, kwargs=None):
        print(args)
        print(kwargs)
        if args is None:
            args = list()
        if kwargs is None:
            kwargs = dict()
        if where is None:
            where = self.files
        if name in where:
            self._files_get_id_by_name(name)
            self.get(path + "/" + str(self._file_id))
            self._opened_files[self._file_id] = json.loads(self.text)
            return returner(self, name, *args, **kwargs)
        else:
            raise FileNotFoundError()

    # SCRIPTS
    def script(self, script_id, function, parameters, dev_mode=False):
        data = {"function": function,
                "parameters": parameters,
                "devMode": dev_mode}
        self.post(SCRIPTS.format(script_id), json=data)
        return json.loads(self.text)

    # SPREADSHEETS
    @updatedSpreadsheet
    def spreadsheet_add_sheet(self, sheetname, *, name=None):
        self._files_get_id_by_name(name)
        if self._file_id is not None:
            data = {"includeSpreadsheetInResponse": "true",
                    "requests": [{"addSheet": {"properties": {"title": sheetname,
                                                              "gridProperties": {"rowCount": 1,
                                                                                 "columnCount": 3}}}},
                                 ]
                    }
            self.post(SHEET_BATCHUPDATE.format(self._file_id), json=data)
        else:
            raise FileNotOpenError()

    def spreadsheet_append_row(self, _range, values, *, name=None):
        self._files_get_id_by_name(name)
        _range = self.spreadsheet_check_range(_range, name=name)
        return self.spreadsheet_append_rows(_range, [values], name=name)

    def spreadsheet_append_rows(self, _range, values, *, name=None):
        self._files_get_id_by_name(name)
        _range = self.spreadsheet_check_range(_range, name=name)
        if self._file_id is not None:
            self.post(SHEET_APPEND.format(self._file_id, _range), get={"valueInputOption": "RAW",
                                                                      "insertDataOption": "INSERT_ROWS",
                                                                      "includeValuesInResponse": "true"},
                      json={"range": _range, "values": values})
            data = json.loads(self.text)
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

    @updatedSpreadsheet
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
                self.post(SHEET_BATCHUPDATE.format(self._file_id), json=data)
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
            data = json.loads(self.get(SHEET_VALUES.format(self._file_id, range)))
            if "values" in data:
                return data["values"]
            else:
                return [[]]
        else:
            raise FileNotOpenError()

    def spreadsheet_get_total_cells(self, *, name=None):
        self._files_get_id_by_name(name)
        sheets = self._opened_files[self._file_id]["sheets"]
        return sum([sheet["properties"]["gridProperties"]["columnCount"]*sheet["properties"]["gridProperties"]["rowCount"]
                    for sheet in sheets])

    def spreadsheet_open(self, name=None, **kwargs):
        if name is None and "name" in kwargs:
            name = kwargs["name"]
        elif name is None:
            raise FileNotFoundError()
        return self._files_open(SHEETS, Spreadsheets, name, self.spreadsheets)

    def spreadsheet_open_sheet(self, sheet_name, *, name=None):
        self._files_get_id_by_name(name)
        sheets = self._opened_files[self._file_id]["sheets"]
        for sheet in sheets:
            if sheet["properties"]["title"] == sheet_name:
                self._opened_sheet = sheet_name
                return
        raise SheetNotFoundError(sheet_name)

    def spreadsheet_update_range(self, range, values, *, name=None):
        self._files_get_id_by_name(name)
        range = self.spreadsheet_check_range(range, name=name)
        if self._file_id is not None:
            data = json.loads(self.put(SHEET_VALUES.format(self._file_id, range),
                                       get={"valueInputOption": "USER_ENTERED"},
                                       json={"range": range, "values": values}))
            if "values" in data:
                return data["values"]
            else:
                return [[]]
        else:
            raise FileNotOpenError()


