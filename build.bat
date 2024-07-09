@echo off
chcp 65001

nuitka ^
  --msvc=latest ^
  --lto=yes ^
  --standalone ^
  --assume-yes-for-downloads ^
  --plugin-enable=upx ^
  --enable-plugin=tk-inter ^
  --windows-disable-console ^
  --output-dir=dist ^
  --remove-output ^
  --show-progress ^
  --include-data-file=start.wav=start.wav ^
  --include-data-file=stop.wav=stop.wav ^
  --include-data-file=app.ico=app.ico ^
  --windows-icon-from-ico=app.ico ^
  --windows-uac-admin ^
  --output-filename=点点侠 ^
  app.py
