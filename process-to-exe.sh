#!/bin/bash
pyinstaller.exe --onefile uiconfig.py
pyinstaller.exe --onefile parse_table.py
pyinstaller.exe --onefile excel_processor2.py
