#!/usr/bin/env python3
__author__ = 'Derek King'

from sys import exit

try:
    from win32com.shell import shell, shellcon
    import pythoncom
except ImportError:
    shell, shellcon, pythoncom = None, None, None
    print("Missing dependency: PyWin32 version 219+ from http://www.lfd.uci.edu/~gohlke/pythonlibs/#pywin32")
    exit(2)

try:
    shell.CLSID_FileOperation
except AttributeError:
    print("Incorrect version of PyWin32")
    print("Need PyWin32 version 219+ from http://www.lfd.uci.edu/~gohlke/pythonlibs/#pywin32")
    print("run `python.exe Scripts\pywin32_postinstall.py -install` from an elevated command prompt")
    exit(2)

from os import makedirs
from os.path import exists
from datetime import datetime
from dateutil import tz
from glob import iglob
from configparser import SafeConfigParser
from ast import literal_eval


verbose_print = print if verbose else lambda *a, **k: None


def continue_running():
    response = input("Continue? [y/n] ").lower()
    while response not in ["y", "n"]:
        response = input("Please respond with 'y' or 'n': ").lower()
    if response == "n":
        return False
    else:
        return True


def get_source_folder(source_path_list):
    folder = shell.SHGetDesktopFolder()
    for item in source_path_list:
        for pidl in folder:
            if folder.GetDisplayNameOf(pidl, shellcon.SHGDN_NORMAL) == item:
                folder = folder.BindToObject(pidl, None, shell.IID_IShellFolder)
                break
        else:
            return None
    return folder


def get_date_taken(folder, photo_pidl):
    # TODO: For AAE files use date of jpg/mov file
    # Uses photo modification time to determine destination folder
    photo_stream = folder.BindToObject(photo_pidl, None, pythoncom.IID_IStream)
    date_taken = photo_stream.Stat(1)[3]  # TODO: Incorrect if file has been modified
    # verbose_print("Modified Time", date_taken)
    # Modified dates on iPhone files are stored in UTC
    date_taken = date_taken.astimezone(local_timezone)
    verbose_print("Date Taken", date_taken)

    ######
    verbose_print("Creation Time", photo_stream.Stat(1)[4])  # Returns 1601-01-01 00:00:00+00:00
    #
    # from win32com.propsys import pscon
    # import pywintypes
    # item = shell.SHCreateItemFromIDList(photo_pidl, shell.IID_IShellItem2)
    #
    # GetFileTime() not implemented for IShellItem2 just PyGShellItem2 as Gateway Implementation
    # https://github.com/Travis-Sun/pywin32/blob/master/com/win32comext/shell/src/PyIShellItem2.cpp
    # time = item.GetFileTime(pscon.PKEY_DateCreated)
    #
    # GetProperty() and GetPropertyStore() don't work because PIDL doesn't store information for special files
    # https://msdn.microsoft.com/en-us/library/bb761124
    # https://msdn.microsoft.com/en-us/library/windows/desktop/bb759748(v=vs.85).aspx
    # time = item.GetProperty(pscon.PKEY_DateCreated)
    # property_store = item.GetPropertyStore(0, pywintypes.IID("{886d8eeb-8cf2-4446-8d02-cdba1dbdcf99}"))
    #
    # SHGetDataFromIDList() not implemented and probably wouldn't work anyway because above
    # data = shell.SHGetDataFromIDList(folder, pidl, shellcon.SHGDFIL_FINDDATA)
    ######

    return date_taken


def validate_photo(dest, date_taken):
    # validate folder with exif data, if incorrect allow option to exit
    with open(dest, 'rb') as photo:
        tags = exifread.process_file(photo, details=False, stop_tag='DateTimeOriginal')

    date_taken_exif = datetime.strptime(tags["EXIF DateTimeOriginal"].values, "%Y:%m:%d %H:%M:%S")
    if date_taken.strftime(folder_format) != date_taken_exif.strftime(folder_format):
        print("Date from file info is different from date taken in exif")
        if continue_running():
            return "Validation failed and ignored"
        else:
            return False
    else:
        return "Validated"


def copy(source_pidl, dest_path_str):
    pfo = pythoncom.CoCreateInstance(shell.CLSID_FileOperation,
                                     None,
                                     pythoncom.CLSCTX_ALL,
                                     shell.IID_IFileOperation)

    pfo.SetOperationFlags(shellcon.FOF_ALLOWUNDO)

    src = shell.SHCreateItemFromIDList(source_pidl, shell.IID_IShellItem)
    
    if not exists(dest_path_str):
        # If folder does not exist then no description already on folder name
        description = input("Give description for " + dest_path_str + ":")
        if description:
            dest_path_str += " " + description + "\\"
        else:
            dest_path_str += "\\"
        makedirs(dest_path_str)
    dest = shell.SHCreateItemFromParsingName(dest_path_str, None, shell.IID_IShellItem)

    pfo.CopyItem(src, dest)

    success = pfo.PerformOperations()
    aborted = pfo.GetAnyOperationsAborted()
    return success, aborted


def main():
    photos_copied = videos_copied = other_copied = 0
    files_ignored = 0

    # Get source PyIShellFolder
    folder = get_source_folder(source_path)
    if not folder:
        print(source_path[1], "not found, exiting.")
        return 1

    # Process each file in the folder
    for pidl in folder:
        file_name = folder.GetDisplayNameOf(pidl, shellcon.SHGDN_NORMAL)
        verbose_print("\nProcessing", file_name + ":")

        date_taken = get_date_taken(folder, pidl)

        if date_range_min < date_taken <= date_range_max:
            dest_path_date = dest_path + date_taken.strftime(folder_format)
            # Allows descriptions on existing folders
            dest_path_date_glob = iglob(dest_path_date + "*")
            # dest_path_date += "\\"

            for path in dest_path_date_glob:
                dest_path_date = path
                if exists(path + "\\" + file_name):
                    verbose_print("File is already in", path)
                    files_ignored += 1
                    break
            else:
                if test:
                    if ".jpg" in file_name.lower():
                        photos_copied += 1
                    elif ".mov" in file_name.lower():
                        videos_copied += 1
                    else:
                        other_copied += 1
                    print("File would be copied to", dest_path_date)
                else:
                    # TODO: Copy doesn't work problem with pidl "The parameter is incorrect."
                    success, aborted = copy(pidl, dest_path_date)
                    if aborted:
                        print("User cancelled file operation on", file_name)
                        if not continue_running():
                            print("Exiting...")
                            return 1
                    else:
                        if success:
                            if ".jpg" in file_name.lower():
                                photos_copied += 1
                            elif ".mov" in file_name.lower():
                                videos_copied += 1
                            else:
                                other_copied += 1
                            verbose_print("File copied successfully to", dest_path_date)
                        else:
                            print("File copy failed on", file_name, "with error code:", success)
                            return 1

                    if validate and ".jpg" in file_name.lower():
                        result = validate_photo(dest_path_date + "\\" + file_name, date_taken)
                        if result:
                            verbose_print(result)
                        else:
                            print("Exiting...")
                            return 1
        else:
            files_ignored += 1
            verbose_print("Photo is outside date range, not copying")

    print("\nFinished processing photos")
    if test:
        print(photos_copied, "photos would be copied")
        print(videos_copied, "videos would be copied")
        print(other_copied, "other files would be copied")
    else:
        print(photos_copied, "photos copied")
        print(videos_copied, "videos copied")
        print(other_copied, "other files copied")
    print(files_ignored, "files already in destination or outside date range")
    return 0


def save_settings(settings_file, local_timezone, source_path, dest_path, date_range_min, date_range_max, folder_format, validate, verbose, test):
    parser = SafeConfigParser()

    parser.add_section("Device")
    parser.set("Device", "local_timezone", str(datetime.now(local_timezone).tzname()))
    parser.set("Device", "source_path", str(source_path))
    parser.set("Device", "dest_path", str(dest_path))
    parser.set("Device", "date_range_min", str(date_range_min.strftime("%Y:%m:%d%z")))
    parser.set("Device", "date_range_max", str(date_range_max.strftime("%Y:%m:%d%z")))

    parser.add_section("Options"):
    parser.set("Options", "folder_format", str(folder_format))
    parser.set("Options", "validate", str(validate))
    parser.set("Options", "verbose", str(verbose))
    parser.set("Options", "test", str(test))

    with open(settings_file, "w") as f:
        parser.write(f)

def load_settings(settings_file):
    # Use these defaults if the settings file is missing or doesn't contain the options
    local_timezone = tz.gettz("UTC")

    # "This PC" for win 10 (and possibly win 8), "Computer" for previous OSs
    source_path = ["This PC", "iPhone", "Internal Storage", "DCIM", "100APPLE"]
    # Has to use backslashes  # TODO: fix
    dest_path = "C:\\tmp\\"
    # Only copy files taken within this date range
    date_range_min = datetime.strptime("2014:09:19+1100", "%Y:%m:%d%z")
    date_range_max = datetime.now(local_timezone)  # Optional TODO: make properly optional
                                                   #                currently becomes fixed when file is saved

    # Uses format from https://docs.python.org/3.3/library/datetime.html#strftime-strptime-behavior
    folder_format = "%Y\\%m-%d"
    # Validate date taken using exif data for jpgs after copy
    validate = True
    # Print verbose info
    verbose = True
    # Print copy location instead of copying
    test = True

    parser = SafeConfigParser()
    found = parser.read(settings_file)
    valid_file = True
    
    if found
        if parser.has_section("Device"):
            if parser.has_option("Device", "local_timezone"):
                local_timezone = tz.gettz(parser.get("Device", "local_timezone"))
            else:
                valid_file = False
            if parser.has_option("Device", "source_path"):
                source_path = literal_eval(parser.get("Device", "source_path"))
            else:
                valid_file = False
            if parser.has_option("Device", "dest_path"):
                dest_path = parser.get("Device", "dest_path")
            else:
                valid_file = False
            if parser.has_option("Device", "date_range_min"):
                date_range_min = datetime.strptime(parser.get("Device", "date_range_min"))
            else:
                valid_file = False
            if parser.has_option("Device", "date_range_max"):
                date_range_max = datetime.strptime(parser.get("Device", "date_range_max"))
            else:
                date_range_max = datetime.now(local_timezone)
        else:
            valid_file = False
        if parser.has_section("Options"):
            if parser.has_option("Options", "folder_format"):
                folder_format = parser.get("Options", "folder_format")
            else:
                valid_file = False
            if parser.has_option("Options", "validate"):
                validate = parser.getboolean("Options", "validate")
            else:
                valid_file = False
            if parser.has_option("Options", "verbose"):
                verbose = parser.getboolean("Options", "verbose")
            else:
                valid_file = False
            if parser.has_option("Options", "test"):
                test = parser.getboolean("Options", "test")
            else:
                valid_file = False
        else:
            valid_file = False
    else:
        valid_file = False

    if not valid_file:
        save_settings(settings_file, local_timezone, source_path, dest_path, date_range_min, date_range_max, folder_format, validate, verbose, test)

    return local_timezone, source_path, dest_path, date_range_min, date_range_max, folder_format, validate, verbose, test


if __name__ == '__main__':
    #TODO: use a class instead of global variables
    local_timezone, source_path, dest_path, date_range_min, date_range_max, folder_format, validate, verbose, test = load_settings("settings.ini")

    if validate:
        try:
            import exifread
        except ImportError:
            exifread = None
            print("Can't validate, missing dependency: exifread")
            exit(2)

    exit(main())
