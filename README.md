# 出口管理系统 - GitHub仓库创建与代码上传指南

## 项目简介
这是一个出口合同管理系统，包含合同管理、进度查询、收款管理等功能。

## GitHub仓库创建与代码上传步骤

### 步骤1: 手动创建GitHub仓库
1. 打开GitHub网站 (https://github.com) 并登录您的账号
2. 点击右上角的 "+" 按钮，选择 "New repository"
3. 在 "Repository name" 字段中输入 "出口管理"
4. 选择仓库可见性 (公开或私有)
5. 勾选 "Add a README file" 选项
6. 点击 "Create repository" 按钮

### 步骤2: 安装Git并配置
1. 下载并安装Git: https://git-scm.com/downloads
2. 安装完成后，打开命令提示符(CMD)或PowerShell
3. 配置Git用户信息:
   ```
   git config --global user.name "您的GitHub用户名"
   git config --global user.email "您的GitHub邮箱"
   ```

### 步骤3: 初始化本地仓库并推送代码
1. 打开命令提示符(CMD)或PowerShell，导航到项目目录:
   ```
   cd d:\Mycode\ote
   ```
2. 初始化Git仓库:
   ```
   git init
   ```
3. 添加所有文件到暂存区:
   ```
   git add .
   ```
4. 提交更改:
   ```
   git commit -m "初始化项目"
   ```
5. 关联远程仓库(替换为您的仓库URL):
   ```
   git remote add origin https://github.com/您的用户名/出口管理.git
   ```
6. 推送代码到远程仓库:
   ```
   git push -u origin main
   ```

## 项目结构
```
├── ContractManagementSystem.py  # 主程序文件
├── data_manager.py             # 数据管理模块
├── data/                       # 数据文件夹
│   ├── contracts.xlsx          # 合同数据
│   ├── customers.xlsx          # 客户数据
│   ├── payments.xlsx           # 收款数据
│   └── salesmen.xlsx           # 销售人员数据
└── .venv/                      # 虚拟环境 (可选)
```

## 运行说明
1. 确保已安装所需依赖:
   ```
   pip install pandas openpyxl tkinter
   ```
2. 运行程序:
   ```
   python ContractManagementSystem.py
   ```