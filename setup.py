# coding=utf8
#launch from cmd: python setup.py build

from cx_Freeze import setup, Executable

setup(name = "IMEI Check",
      version = "2.3",
      description = "IMEI Check RC V2.3",
      executables = [Executable("imeicheck.py",base = "Win32GUI")]) # <-- base hide console!