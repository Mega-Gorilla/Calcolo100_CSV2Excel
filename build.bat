python -m nuitka ^
    --standalone ^
    --enable-plugin=pyqt6 ^
    --include-qt-plugins=platforms,styles,imageformats ^
    app.py