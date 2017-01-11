from glob import glob
import re
from os import getsize


def getFiles(filePath):
    files = glob(filePath)
    return files


def checkVersion(h_splited):
    if re.match("[0-9]\.", h_splited):
        return True
    else:
        return False


def getLibraryVersion(file):
    libVer = file.split("\\")[-1].replace(".jar", "")
    return libVer


def getLibrary(libVer):
    lib = ""
    for h_splited in libVer.split("-"):
        if not checkVersion(h_splited):
            lib += h_splited
    return lib


def getVersion(libVer):
    for h_splited in libVer.split("-"):
        if checkVersion(h_splited):
            return h_splited
    else:
        return "No Version"


def getFileSize(file):
    return getsize(file)


def getData(FilePath):
    data = []
    for file in getFiles(FilePath):
        libVer = getLibraryVersion(file)
        lib = getLibrary(libVer)
        ver = getVersion(libVer)
        size = getFileSize(file)
        data.append({"lib": lib, "ver": ver, "size": size, "path": file})
    return data


def makeRecord(originFilePath, targetFilePath):
    origin_data = getData(originFilePath)
    target_data = getData(targetFilePath)

    mixLibs = []
    for origin in origin_data:
        mixLibs.append(origin["lib"])
    for target in target_data:
        mixLibs.append(target["lib"])
    mixLibs = list(set(mixLibs))

    origin_record = None
    target_record = None
    for lib in mixLibs:
        for origin in origin_data:
            if lib == origin["lib"]:
                global origin_record
                origin_record = origin

        for target in target_data:
            if lib == target["lib"]:
                global target_record
                target_record = target

        if origin_record is not None and target_record is not None:
            pass
        elif origin_record is None and target_record is not None:
            pass
        elif origin_record is not None and target_record is None:
            pass
