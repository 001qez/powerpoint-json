import sys
from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need fine tuning.
build_exe_options = {"packages": ["os"],
                     "excludes": ["tkinter"],
                     "include_files": ["products_details.json",
                                       "powerpoint-json-v7.py",
                                       "vba.txt"]}

# GUI applications require a different base on Windows (the default is for a
# console application).
base = 'Console'

setup(  name = "powerpoint-json",
        version = "0.7",
        description = "powerpoint-json version 7",
        options = {"build_exe": build_exe_options},
        executables = [Executable("powerpoint-json.py", base=base)])
