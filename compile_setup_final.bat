@echo off
chcp 65001 >nul

REM 检查Inno Setup是否已安装
where iscc >nul 2>nul
if %ERRORLEVEL% NEq 0 (
    echo 未找到Inno Setup编译器。请先安装Inno Setup。
    echo 下载链接: https://jrsoftware.org/isdl.php
    echo 安装完成后，请重新运行此脚本。
    pause
    exit /b 1
)

REM 创建安装程序输出目录
if not exist installer mkdir installer

REM 编译安装脚本
echo 正在编译安装脚本...
iscc setup.iss

if %ERRORLEVEL% equ 0 (
    echo 安装程序编译成功！
    echo 安装程序位于: d:\Mycode\ote\installer\ContractManagementSystem_Setup.exe
) else (
    echo 安装程序编译失败。
)

pause