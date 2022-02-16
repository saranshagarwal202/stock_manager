#!/bin/bash
pyinstaller --onefile \
    --clean \
    --log-level WARN \
    --windowed \
    -i icon.ico \
    --name CIM \
    --add-data schema.json:./ \
    --add-data approval_format.docx:./ \
    -p bin/python \
    MainWindow_backend.py