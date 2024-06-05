@echo off

nuitka ^
  --lto=yes ^
  --standalone ^
  --onefile ^
  --plugin-enable=upx ^
  --enable-plugin=tk-inter ^
  --windows-disable-console ^
  --output-dir=dist ^
  --show-progress ^
  --include-data-file=start.wav=start.wav ^
  --include-data-file=stop.wav=stop.wav ^
  --include-data-file=app.ico=app.ico ^
  --windows-icon-from-ico=app.ico ^
  --windows-uac-admin ^
  app.py