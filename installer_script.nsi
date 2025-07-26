; 出口管理系统安装程序脚本
; 使用NSIS (Nullsoft Scriptable Install System) 编译
; 下载地址: https://nsis.sourceforge.io/Download

; 定义安装程序基本信息
!define PRODUCT_NAME "出口管理系统"
!define PRODUCT_VERSION "1.0"
!define PRODUCT_PUBLISHER "出口管理系统开发团队"
!define PRODUCT_WEB_SITE "https://example.com"
!define INSTALLER_NAME "出口管理系统安装程序.exe"

; 设置安装程序界面
!include "MUI2.nsh"

; 配置安装程序
Name "${PRODUCT_NAME} ${PRODUCT_VERSION}"
OutFile "${INSTALLER_NAME}"
InstallDir "$PROGRAMFILES\${PRODUCT_NAME}"
InstallDirRegKey HKLM "Software\${PRODUCT_NAME}" "InstallDir"
ShowInstDetails show
ShowUnInstDetails show

; 配置MUI
!define MUI_ABORTWARNING
!define MUI_ICON "${NSISDIR}\Contrib\Graphics\Icons\modern-install.ico"
!define MUI_UNICON "${NSISDIR}\Contrib\Graphics\Icons\modern-uninstall.ico"

; 欢迎页面
!insertmacro MUI_PAGE_WELCOME
; 安装目录选择页面
!insertmacro MUI_PAGE_DIRECTORY
; 安装进度页面
!insertmacro MUI_PAGE_INSTFILES
; 完成页面
!define MUI_FINISHPAGE_RUN "$INSTDIR\出口管理系统.exe"
!insertmacro MUI_PAGE_FINISH

; 卸载页面
!insertmacro MUI_UNPAGE_WELCOME
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_UNPAGE_FINISH

; 语言文件
!insertmacro MUI_LANGUAGE "SimpChinese"

; 安装程序部分
Section "MainSection" SEC01
  SetOutPath "$INSTDIR"
  ; 添加主程序文件
  File "dist\出口管理系统.exe"
  ; 添加数据文件夹
  SetOutPath "$INSTDIR\data"
  File /r "dist\data\*.*"

  ; 创建桌面快捷方式
  CreateShortCut "$DESKTOP\${PRODUCT_NAME}.lnk" "$INSTDIR\出口管理系统.exe" "" "$INSTDIR\出口管理系统.exe" 0

  ; 创建开始菜单快捷方式
  CreateDirectory "$SMPROGRAMS\${PRODUCT_NAME}"
  CreateShortCut "$SMPROGRAMS\${PRODUCT_NAME}\${PRODUCT_NAME}.lnk" "$INSTDIR\出口管理系统.exe" "" "$INSTDIR\出口管理系统.exe" 0
  CreateShortCut "$SMPROGRAMS\${PRODUCT_NAME}\卸载.lnk" "$INSTDIR\uninstall.exe" "" "$INSTDIR\uninstall.exe" 0

  ; 写入注册表信息
  WriteRegStr HKLM "Software\${PRODUCT_NAME}" "InstallDir" "$INSTDIR"
  WriteRegStr HKLM "Software\${PRODUCT_NAME}" "Version" "${PRODUCT_VERSION}"
  WriteUninstaller "$INSTDIR\uninstall.exe"
SectionEnd

; 卸载程序部分
Section "Uninstall"
  ; 删除文件和文件夹
  Delete "$INSTDIR\出口管理系统.exe"
  RMDir /r "$INSTDIR\data"
  Delete "$INSTDIR\uninstall.exe"

  ; 删除快捷方式
  Delete "$DESKTOP\${PRODUCT_NAME}.lnk"
  Delete "$SMPROGRAMS\${PRODUCT_NAME}\${PRODUCT_NAME}.lnk"
  Delete "$SMPROGRAMS\${PRODUCT_NAME}\卸载.lnk"
  RMDir "$SMPROGRAMS\${PRODUCT_NAME}"

  ; 删除注册表项
  DeleteRegKey HKLM "Software\${PRODUCT_NAME}"

  ; 删除安装目录
  RMDir "$INSTDIR"
SectionEnd

; 安装程序结束
Function .onInstSuccess
  MessageBox MB_OK "${PRODUCT_NAME} ${PRODUCT_VERSION} 安装成功！"
FunctionEnd

Function .onUnInstSuccess
  MessageBox MB_OK "${PRODUCT_NAME} 已成功卸载！"
FunctionEnd