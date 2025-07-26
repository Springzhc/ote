@echo off

:: 激活虚拟环境
call .venv\Scripts\activate.bat

:: 使用PyInstaller打包应用程序
pyinstaller --onefile --windowed --name "ContractManagementSystem" --add-data "data/*;data/" ContractManagementSystem.py

:: 暂停以便查看输出
pause