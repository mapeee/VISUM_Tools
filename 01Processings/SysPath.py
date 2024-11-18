"""Contains helper functions to change the python system path"""

import sys
import os
import winreg
from pathlib import Path


def ChangePythonSysPath(visum_version, bit=None):
    """Inserts the python modules of the specified Visum version into sys.path

    visum_version: The number of the Visum version (e.g. 22 or 23)
    bit: obsolete"""

    try:
        key = rf"Visum.Visum-64.{visum_version}\CLSID"
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key) as hkey:
            clsid = winreg.QueryValueEx(hkey, "")[0]

        key = rf"CLSID\{clsid}\LocalServer32"
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key) as hkey:
            visum_path = winreg.QueryValueEx(hkey, "")[0]
        visum_path = Path(str(visum_path).strip('"'))
    except Exception as e:
        raise Exception("Failed to find the specified Visum version") from e

    if not visum_path.exists():
        raise Exception("The Visum installation registered as COM server does not exist anymore")

    for subfolder in ["Python", "PythonModules", "Python37Modules"]:
        path = visum_path.parent / subfolder / "Lib" / "site-packages"
        if path.exists():
            break
    else:
        raise Exception("Failed to find the python modules in the visum installation." +
                        "You might have to upgrade SysPath.py")
    path = str(path)
    if path not in sys.path:
        sys.path.insert(0, path)

    return True
