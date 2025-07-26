# 出口销售合同管理系统 - 安装程序创建指南

## 前提条件

在创建Windows安装程序之前，您需要安装Inno Setup编译器。

## 步骤1：安装Inno Setup

1. 访问Inno Setup官方下载页面：https://jrsoftware.org/isdl.php
2. 下载最新版本的Inno Setup（建议下载Unicode版本）
3. 运行安装程序并按照提示完成安装

## 步骤2：编译安装脚本

1. 安装完成后，打开Inno Setup编译器
2. 在菜单栏中选择"文件" > "打开"
3. 浏览到项目目录 `d:\Mycode\ote\` 并选择 `setup.iss` 文件
4. 点击工具栏上的"编译"按钮（或按F9键）

## 步骤3：获取安装程序

1. 编译完成后，安装程序将生成在 `d:\Mycode\ote\installer\` 目录下
2. 安装程序文件名为 `ContractManagementSystem_Setup.exe`

## 安装程序功能

生成的安装程序将：
- 将应用程序文件安装到用户指定的目录
- 创建开始菜单快捷方式
- 可选地创建桌面快捷方式
- 安装完成后可以直接启动应用程序

## 故障排除

如果编译过程中遇到问题：
1. 确保Inno Setup已正确安装
2. 检查 `setup.iss` 文件中的路径是否正确
3. 确保 `d:\Mycode\ote\dist\` 目录下存在 `ContractManagementSystem.exe` 文件
4. 确保 `d:\Mycode\ote\data\` 目录下包含所有必要的数据文件

如果您需要进一步的帮助，请访问项目的GitHub页面：https://github.com/SpringZHC/ote