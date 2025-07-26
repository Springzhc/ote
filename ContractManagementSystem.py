import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import pandas as pd
import os
from datetime import datetime
import openpyxl
from data_manager import DataManager

# 设置中文字体支持
class ContractManagementSystem:
    def __init__(self, root):
        self.root = root
        self.root.title("出口销售合同管理系统")
        self.root.geometry("1000x600")
        self.root.minsize(800, 500)

        # 确保数据目录存在
        self.data_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
        if not os.path.exists(self.data_dir):
            os.makedirs(self.data_dir)

        # 初始化数据管理器
        from data_manager import DataManager
        self.data_manager = DataManager(self.data_dir)

        # 初始化数据文件
        self.init_data_files()

        # 创建主菜单
        self.create_main_menu()

        # 创建主界面
        self.create_main_frame()

    def init_data_files(self):
        """初始化Excel数据文件"""
        # 合同表
        self.contract_file = os.path.join(self.data_dir, "contracts.xlsx")
        if not os.path.exists(self.contract_file):
            df = pd.DataFrame({
                '合同编号': [],
                '客户名称': [],
                '业务员': [],
                '签订日期': [],
                '合同金额': [],
                '付款方式': [],
                '交货日期': [],
                '状态': [],
                '备注': []
            })
            df.to_excel(self.contract_file, index=False, engine='openpyxl')

        # 收款表
        self.payment_file = os.path.join(self.data_dir, "payments.xlsx")
        if not os.path.exists(self.payment_file):
            df = pd.DataFrame({
                '收款ID': [],
                '合同编号': [],
                '收款日期': [],
                '收款金额': [],
                '收款方式': [],
                '备注': []
            })
            df.to_excel(self.payment_file, index=False, engine='openpyxl')

        # 客户表
        self.customer_file = os.path.join(self.data_dir, "customers.xlsx")
        if not os.path.exists(self.customer_file):
            df = pd.DataFrame({
                '客户ID': [],
                '客户名称': [],
                '联系人': [],
                '联系电话': [],
                '地址': [],
                '邮箱': [],
                '备注': []
            })
            df.to_excel(self.customer_file, index=False, engine='openpyxl')

        # 业务员表
        self.salesman_file = os.path.join(self.data_dir, "salesmen.xlsx")
        if not os.path.exists(self.salesman_file):
            df = pd.DataFrame({
                '业务员ID': [],
                '姓名': [],
                '联系电话': [],
                '邮箱': [],
                '所属部门': [],
                '备注': []
            })
            df.to_excel(self.salesman_file, index=False, engine='openpyxl')

    def create_main_menu(self):
        """创建主菜单"""
        menubar = tk.Menu(self.root)

        # 文件菜单
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="打开数据库文件夹", command=self.open_data_folder)
        file_menu.add_command(label="退出", command=self.root.quit)
        menubar.add_cascade(label="文件", menu=file_menu)

        # 合同管理菜单
        contract_menu = tk.Menu(menubar, tearoff=0)
        contract_menu.add_command(label="合同管理", command=self.open_contract_management)
        contract_menu.add_command(label="合同进度查询", command=self.open_contract_progress)
        contract_menu.add_command(label="合同应收报表", command=self.open_contract_receivable_report)
        menubar.add_cascade(label="合同管理", menu=contract_menu)

        # 收款管理菜单
        payment_menu = tk.Menu(menubar, tearoff=0)
        payment_menu.add_command(label="收款管理", command=self.open_payment_management)
        menubar.add_cascade(label="收款管理", menu=payment_menu)

        # 客户管理菜单
        customer_menu = tk.Menu(menubar, tearoff=0)
        customer_menu.add_command(label="客户管理", command=self.open_customer_management)
        menubar.add_cascade(label="客户管理", menu=customer_menu)

        # 业务员管理菜单
        salesman_menu = tk.Menu(menubar, tearoff=0)
        salesman_menu.add_command(label="业务员管理", command=self.open_salesman_management)
        menubar.add_cascade(label="业务员管理", menu=salesman_menu)

        # 日志菜单
        log_menu = tk.Menu(menubar, tearoff=0)
        log_menu.add_command(label="查看日志", command=self.open_log_file)
        menubar.add_cascade(label="日志", menu=log_menu)

        # 帮助菜单
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="关于", command=self.show_about)
        menubar.add_cascade(label="帮助", menu=help_menu)

        self.root.config(menu=menubar)

        # 初始化日志文件
        self.log_file_path = os.path.join(self.data_dir, 'log.txt')
        if not os.path.exists(self.log_file_path):
            with open(self.log_file_path, 'w', encoding='utf-8') as f:
                f.write('===== 系统日志 =====\n')
                f.write(f'日志创建时间: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}\n\n')

    def create_main_frame(self):
        """创建主界面"""
        # 清空当前界面
        for widget in self.root.winfo_children():
            if isinstance(widget, tk.Frame) and widget != self.root.nametowidget('.!menu'):
                widget.destroy()

        # 创建主框架
        main_frame = tk.Frame(self.root, bg='white')
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 添加标题
        title_label = tk.Label(main_frame, text="出口销售合同管理系统", font=('SimHei', 24, 'bold'))
        title_label.pack(pady=50)

        # 创建功能按钮区域
        buttons_frame = tk.Frame(main_frame, bg='white')
        buttons_frame.pack(expand=True)

        # 合同管理按钮 - 立体效果
        contract_btn = tk.Button(buttons_frame, text="合同管理", font=('SimHei', 14, 'bold'),
                                command=self.open_contract_management, width=15, height=2,
                                bg='#f0f0f0', fg='#333333', activebackground='#e0e0e0',
                                highlightthickness=2, bd=2, relief=tk.RAISED,
                                highlightbackground='#a0a0a0', highlightcolor='#808080')
        contract_btn.grid(row=0, column=0, padx=20, pady=20)

        # 收款管理按钮 - 立体效果
        payment_btn = tk.Button(buttons_frame, text="收款管理", font=('SimHei', 14, 'bold'),
                               command=self.open_payment_management, width=15, height=2,
                               bg='#f0f0f0', fg='#333333', activebackground='#e0e0e0',
                               highlightthickness=2, bd=2, relief=tk.RAISED,
                               highlightbackground='#a0a0a0', highlightcolor='#808080')
        payment_btn.grid(row=0, column=1, padx=20, pady=20)

        # 客户管理按钮 - 立体效果
        customer_btn = tk.Button(buttons_frame, text="客户管理", font=('SimHei', 14, 'bold'),
                                command=self.open_customer_management, width=15, height=2,
                                bg='#f0f0f0', fg='#333333', activebackground='#e0e0e0',
                                highlightthickness=2, bd=2, relief=tk.RAISED,
                                highlightbackground='#a0a0a0', highlightcolor='#808080')
        customer_btn.grid(row=1, column=0, padx=20, pady=20)

        # 业务员管理按钮 - 立体效果
        salesman_btn = tk.Button(buttons_frame, text="业务员管理", font=('SimHei', 14, 'bold'),
                                command=self.open_salesman_management, width=15, height=2,
                                bg='#f0f0f0', fg='#333333', activebackground='#e0e0e0',
                                highlightthickness=2, bd=2, relief=tk.RAISED,
                                highlightbackground='#a0a0a0', highlightcolor='#808080')
        salesman_btn.grid(row=1, column=1, padx=20, pady=20)

        # 设置网格权重，确保按钮区域正确显示
        buttons_frame.grid_columnconfigure(0, weight=1)
        buttons_frame.grid_columnconfigure(1, weight=1)
        buttons_frame.grid_rowconfigure(0, weight=1)
        buttons_frame.grid_rowconfigure(1, weight=1)

        # 设置网格权重，使按钮居中
        buttons_frame.grid_columnconfigure(0, weight=1)
        buttons_frame.grid_columnconfigure(1, weight=1)
        buttons_frame.grid_rowconfigure(0, weight=1)
        buttons_frame.grid_rowconfigure(1, weight=1)

    def open_contract_progress(self):
        """打开合同进度查询界面"""
        # 清空当前界面
        for widget in self.root.winfo_children():
            if isinstance(widget, tk.Frame) and widget != self.root.nametowidget('.!menu'):
                widget.destroy()

        # 创建合同进度查询框架
        progress_frame = tk.Frame(self.root)
        progress_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 创建标题
        title_label = tk.Label(progress_frame, text="合同进度查询", font=('SimHei', 16, 'bold'))
        title_label.pack(pady=10)

        # 创建合同选择区域
        select_frame = tk.Frame(progress_frame)
        select_frame.pack(fill=tk.X, padx=5, pady=5)

        tk.Label(select_frame, text="选择合同:", font=('SimHei', 10)).pack(side=tk.LEFT, padx=5)

        # 合同列表下拉框
        contract_var = tk.StringVar()
        contract_combobox = ttk.Combobox(select_frame, textvariable=contract_var, width=30)
        contract_combobox.pack(side=tk.LEFT, padx=5)

        # 加载合同数据到下拉框
        try:
            df = self.data_manager.get_all_contracts()
            contract_list = [f"{row['合同编号']} - {row['客户名称']}" for _, row in df.iterrows()]
            contract_combobox['values'] = contract_list
            if contract_list:
                contract_combobox.current(0)
        except Exception as e:
            messagebox.showerror("错误", f"加载合同数据失败: {str(e)}")

        # 查询按钮
        def query_progress():
            """查询合同进度"""
            # 获取选中的合同
            selected_contract = contract_var.get()
            if not selected_contract:
                messagebox.showwarning("警告", "请先选择合同!")
                return

            contract_id = selected_contract.split(' - ')[0]
            contract_id = str(contract_id)  # 确保合同编号是字符串类型

            # 清空当前显示的信息
            for widget in info_frame.winfo_children():
                widget.destroy()

            for widget in payment_frame.winfo_children():
                if isinstance(widget, ttk.Treeview) or isinstance(widget, ttk.Scrollbar):
                    widget.destroy()

            # 获取合同详情
            try:
                df_contract = self.data_manager.get_all_contracts()
                # 确保合同编号列是字符串类型
                df_contract['合同编号'] = df_contract['合同编号'].astype(str)
                contract_data = df_contract[df_contract['合同编号'] == contract_id].iloc[0].to_dict()

                # 显示合同基本信息和进度
                tk.Label(info_frame, text="合同基本信息", font=('SimHei', 12, 'bold')).grid(row=0, column=0, columnspan=2, pady=5)

                tk.Label(info_frame, text="合同编号:", font=('SimHei', 10)).grid(row=1, column=0, sticky=tk.W, pady=2)
                tk.Label(info_frame, text=contract_data['合同编号'], font=('SimHei', 10)).grid(row=1, column=1, sticky=tk.W, pady=2)

                tk.Label(info_frame, text="客户名称:", font=('SimHei', 10)).grid(row=2, column=0, sticky=tk.W, pady=2)
                tk.Label(info_frame, text=contract_data['客户名称'], font=('SimHei', 10)).grid(row=2, column=1, sticky=tk.W, pady=2)

                tk.Label(info_frame, text="合同金额:", font=('SimHei', 10)).grid(row=3, column=0, sticky=tk.W, pady=2)
                tk.Label(info_frame, text=contract_data['合同金额'], font=('SimHei', 10)).grid(row=3, column=1, sticky=tk.W, pady=2)

                # 获取收款信息
                df_payment = self.data_manager.get_all_payments()
                # 确保合同编号列是字符串类型
                df_payment['合同编号'] = df_payment['合同编号'].astype(str)
                payment_data = df_payment[df_payment['合同编号'] == contract_id]

                # 计算已收款金额
                try:
                    received_amount = payment_data['收款金额'].sum() if not payment_data.empty else 0
                except Exception as e:
                    received_amount = 0
                    print(f"计算已收款金额出错: {str(e)}")

                # 计算应收款金额
                try:
                    contract_amount = float(contract_data['合同金额']) if contract_data['合同金额'] and str(contract_data['合同金额']).strip() else 0
                    receivable_amount = contract_amount - received_amount
                except ValueError:
                    contract_amount = 0
                    receivable_amount = 0
                    print(f"合同金额格式错误: {contract_data['合同金额']}")

                tk.Label(info_frame, text="已收款金额:", font=('SimHei', 10)).grid(row=4, column=0, sticky=tk.W, pady=2)
                tk.Label(info_frame, text=f"{received_amount:.2f}", font=('SimHei', 10)).grid(row=4, column=1, sticky=tk.W, pady=2)

                tk.Label(info_frame, text="应收款金额:", font=('SimHei', 10)).grid(row=5, column=0, sticky=tk.W, pady=2)
                tk.Label(info_frame, text=f"{receivable_amount:.2f}", font=('SimHei', 10)).grid(row=5, column=1, sticky=tk.W, pady=2)

                tk.Label(info_frame, text="到期时间:", font=('SimHei', 10)).grid(row=6, column=0, sticky=tk.W, pady=2)
                tk.Label(info_frame, text=contract_data['交货日期'], font=('SimHei', 10)).grid(row=6, column=1, sticky=tk.W, pady=2)

                # 检查是否超期
                today = datetime.now().date()
                delivery_date = None
                try:
                    if contract_data['交货日期'] and str(contract_data['交货日期']).strip():
                        # 尝试多种日期格式
                        date_formats = ['%Y-%m-%d', '%Y/%m/%d', '%d-%m-%Y', '%d/%m/%Y']
                        for fmt in date_formats:
                            try:
                                delivery_date = datetime.strptime(str(contract_data['交货日期']), fmt).date()
                                break
                            except ValueError:
                                continue
                except Exception as e:
                    print(f"解析交货日期出错: {str(e)}")

                if delivery_date and today > delivery_date:
                    status = "已超期"
                    status_color = "red"
                else:
                    status = "正常"
                    status_color = "black"

                tk.Label(info_frame, text="状态:", font=('SimHei', 10)).grid(row=7, column=0, sticky=tk.W, pady=2)
                tk.Label(info_frame, text=status, font=('SimHei', 10), fg=status_color).grid(row=7, column=1, sticky=tk.W, pady=2)

                # 创建收款明细表格
                columns = ("收款ID", "收款日期", "收款金额", "收款方式", "备注")
                payment_tree = ttk.Treeview(payment_frame, columns=columns, show="headings")

                # 设置列宽和标题
                for col in columns:
                    payment_tree.heading(col, text=col)
                    payment_tree.column(col, width=100, anchor=tk.CENTER)

                # 调整部分列的宽度
                payment_tree.column("收款ID", width=80)
                payment_tree.column("收款金额", width=100)
                payment_tree.column("备注", width=150)

                # 添加滚动条
                scrollbar = ttk.Scrollbar(payment_frame, orient=tk.VERTICAL, command=payment_tree.yview)
                payment_tree.configure(yscroll=scrollbar.set)
                scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
                payment_tree.pack(fill=tk.BOTH, expand=True)

                # 加载收款数据
                for index, row in payment_data.iterrows():
                    payment_tree.insert('', tk.END, values=(
                        row['收款ID'],
                        row['收款日期'],
                        row['收款金额'],
                        row['收款方式'],
                        row['备注']
                    ))

            except IndexError:
                messagebox.showerror("错误", f"未找到合同编号为 {contract_id} 的合同数据")
            except Exception as e:
                messagebox.showerror("错误", f"查询合同进度出错: {str(e)}")

        tk.Button(select_frame, text="查询", command=query_progress).pack(side=tk.LEFT, padx=5)
        tk.Button(select_frame, text="返回主界面", command=self.create_main_frame).pack(side=tk.RIGHT, padx=5)

        # 创建信息显示区域
        info_frame = tk.Frame(progress_frame, bd=1, relief=tk.SUNKEN, padx=10, pady=10)
        info_frame.pack(fill=tk.X, padx=5, pady=5)

        # 创建收款明细区域
        payment_label = tk.Label(progress_frame, text="收款明细", font=('SimHei', 12, 'bold'))
        payment_label.pack(pady=5)

        payment_frame = tk.Frame(progress_frame, bd=1, relief=tk.SUNKEN)
        payment_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    def open_contract_receivable_report(self):
        """打开合同应收报表界面"""
        # 清空当前界面
        for widget in self.root.winfo_children():
            if isinstance(widget, tk.Frame) and widget != self.root.nametowidget('.!menu'):
                widget.destroy()

        # 创建报表框架
        report_frame = tk.Frame(self.root)
        report_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 创建标题
        title_label = tk.Label(report_frame, text="合同应收报表", font=('SimHei', 16, 'bold'))
        title_label.pack(pady=10)

        # 创建搜索和过滤区域
        filter_frame = tk.Frame(report_frame)
        filter_frame.pack(fill=tk.X, padx=5, pady=5)

        # 状态过滤
        tk.Label(filter_frame, text="状态过滤:", font=('SimHei', 10)).pack(side=tk.LEFT, padx=5)
        status_var = tk.StringVar(value="全部")
        status_combo = ttk.Combobox(filter_frame, textvariable=status_var, width=10, values=["全部", "正常", "已超期"])
        status_combo.pack(side=tk.LEFT, padx=5)

        # 刷新按钮
        def refresh_report():
            """刷新报表"""
            # 清空表格
            for item in report_tree.get_children():
                report_tree.delete(item)

            try:
                # 获取所有合同数据
                df_contract = self.data_manager.get_all_contracts()
                df_payment = self.data_manager.get_all_payments()

                # 计算每个合同的已收款金额
                payment_summary = df_payment.groupby('合同编号')['收款金额'].sum().reset_index()
                payment_summary.columns = ['合同编号', '已收款金额']

                # 合并合同数据和收款数据
                merged_df = pd.merge(df_contract, payment_summary, on='合同编号', how='left')
                merged_df['已收款金额'] = merged_df['已收款金额'].fillna(0)

                # 计算未收款金额
                merged_df['未收款金额'] = merged_df['合同金额'].astype(float) - merged_df['已收款金额']

                # 获取今天的日期
                today = datetime.now().date()

                # 检查是否超期
                merged_df['是否超期'] = merged_df['交货日期'].apply(lambda x: True if x and datetime.strptime(x, '%Y-%m-%d').date() < today else False)

                # 状态过滤
                status_filter = status_var.get()
                if status_filter == "正常":
                    merged_df = merged_df[merged_df['是否超期'] == False]
                elif status_filter == "已超期":
                    merged_df = merged_df[merged_df['是否超期'] == True]

                # 按超期状态和到期日期排序
                merged_df = merged_df.sort_values(by=['是否超期', '交货日期'], ascending=[False, True])

                # 显示数据
                for index, row in merged_df.iterrows():
                    # 设置超期行的字体颜色为红色
                    tags = ('overdue',) if row['是否超期'] else ()

                    report_tree.insert('', tk.END, values=(
                        row['合同编号'],
                        row['客户名称'],
                        row['业务员'],
                        row['签订日期'],
                        row['合同金额'],
                        row['已收款金额'],
                        row['未收款金额'],
                        row['交货日期'],
                        "已超期" if row['是否超期'] else "正常"
                    ), tags=tags)

            except Exception as e:
                messagebox.showerror("错误", f"加载报表数据出错: {str(e)}")

        tk.Button(filter_frame, text="刷新", command=refresh_report).pack(side=tk.LEFT, padx=5)
        tk.Button(filter_frame, text="返回主界面", command=self.create_main_frame).pack(side=tk.RIGHT, padx=5)

        # 创建报表表格
        columns = ("合同编号", "客户名称", "业务员", "签订日期", "合同金额", "已收款金额", "未收款金额", "到期日期", "状态")
        report_tree = ttk.Treeview(report_frame, columns=columns, show="headings")

        # 设置列宽和标题
        for col in columns:
            report_tree.heading(col, text=col)
            report_tree.column(col, width=100, anchor=tk.CENTER)

        # 调整部分列的宽度
        report_tree.column("合同编号", width=80)
        report_tree.column("客户名称", width=150)
        report_tree.column("业务员", width=100)
        report_tree.column("合同金额", width=100)
        report_tree.column("已收款金额", width=100)
        report_tree.column("未收款金额", width=100)

        # 添加滚动条
        scrollbar_y = ttk.Scrollbar(report_frame, orient=tk.VERTICAL, command=report_tree.yview)
        scrollbar_x = ttk.Scrollbar(report_frame, orient=tk.HORIZONTAL, command=report_tree.xview)
        report_tree.configure(yscroll=scrollbar_y.set, xscroll=scrollbar_x.set)

        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        report_tree.pack(fill=tk.BOTH, expand=True)

        # 设置超期行的样式
        report_tree.tag_configure('overdue', foreground='red')

        # 初始加载数据
        refresh_report()

    def open_contract_management(self):
        """打开合同管理界面"""
        self.log_operation('打开合同管理界面')
        # 清空当前界面
        for widget in self.root.winfo_children():
            if isinstance(widget, tk.Frame) and widget != self.root.nametowidget('.!menu'):
                widget.destroy()

        # 创建合同管理框架
        contract_frame = tk.Frame(self.root)
        contract_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 创建搜索框
        search_frame = tk.Frame(contract_frame)
        search_frame.pack(fill=tk.X, padx=5, pady=5)

        tk.Label(search_frame, text="搜索:", font=('SimHei', 10)).pack(side=tk.LEFT, padx=5)
        search_var = tk.StringVar()
        search_entry = tk.Entry(search_frame, textvariable=search_var, width=30)
        search_entry.pack(side=tk.LEFT, padx=5)

        tk.Label(search_frame, text="搜索类型:", font=('SimHei', 10)).pack(side=tk.LEFT, padx=5)
        search_type = tk.StringVar(value="合同编号")
        search_type_combo = ttk.Combobox(search_frame, textvariable=search_type, width=10,
                                        values=["合同编号", "客户名称", "业务员", "状态"])
        search_type_combo.pack(side=tk.LEFT, padx=5)

        def search_contract():
            """搜索合同"""
            search_term = search_var.get().strip()
            search_type_val = search_type.get()

            if not search_term:
                messagebox.showwarning("警告", "请输入搜索内容!")
                return

            try:
                # 调用数据管理器进行搜索
                search_results = self.data_manager.search_contracts(search_term, search_type_val)

                # 清空表格
                for item in tree.get_children():
                    tree.delete(item)

                # 加载搜索结果
                if not search_results.empty:
                    for index, row in search_results.iterrows():
                        tree.insert('', tk.END, values=(
                            row['合同编号'],
                            row['客户名称'],
                            row['业务员'],
                            row['签订日期'],
                            row['合同金额'],
                            row['付款方式'],
                            row['交货日期'],
                            row['状态']
                        ))
                    messagebox.showinfo("成功", f"找到 {len(search_results)} 条匹配记录")
                else:
                    messagebox.showinfo("提示", "没有找到匹配的合同记录")
            except Exception as e:
                messagebox.showerror("错误", f"搜索合同出错: {str(e)}")

        tk.Button(search_frame, text="搜索", command=search_contract).pack(side=tk.LEFT, padx=5)
        tk.Button(search_frame, text="刷新", command=self.open_contract_management).pack(side=tk.LEFT, padx=5)

        # 创建按钮组
        btn_frame = tk.Frame(contract_frame)
        btn_frame.pack(fill=tk.X, padx=5, pady=5)

        def add_contract():
            """添加合同"""
            self.log_operation('打开添加合同对话框')
            # 创建添加合同对话框
            add_window = tk.Toplevel(self.root)
            add_window.title("添加合同")
            add_window.geometry("500x400")
            add_window.resizable(False, False)
            add_window.grab_set()  # 模态窗口

            # 设置字体
            font = ('SimHei', 10)

            # 创建表单框架
            form_frame = tk.Frame(add_window, padx=20, pady=20)
            form_frame.pack(fill=tk.BOTH, expand=True)

            # 客户名称
            tk.Label(form_frame, text="客户名称*:", font=font).grid(row=0, column=0, sticky=tk.W, pady=5)
            customer_var = tk.StringVar()
            tk.Entry(form_frame, textvariable=customer_var, width=25).grid(row=0, column=1, pady=5, sticky=tk.W)
            tk.Button(form_frame, text="选择", font=font, command=lambda: self.select_customer(customer_var, add_window)).grid(row=0, column=2, pady=5, padx=5)

            # 业务员
            tk.Label(form_frame, text="业务员*:", font=font).grid(row=1, column=0, sticky=tk.W, pady=5)
            salesman_var = tk.StringVar()
            tk.Entry(form_frame, textvariable=salesman_var, width=25).grid(row=1, column=1, pady=5, sticky=tk.W)
            tk.Button(form_frame, text="选择", font=font, command=lambda: self.select_salesman(salesman_var, add_window)).grid(row=1, column=2, pady=5, padx=5)

            # 签订日期
            tk.Label(form_frame, text="签订日期*:", font=font).grid(row=2, column=0, sticky=tk.W, pady=5)
            date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
            tk.Entry(form_frame, textvariable=date_var, width=30).grid(row=2, column=1, pady=5)

            # 单价
            tk.Label(form_frame, text="单价*:", font=font).grid(row=3, column=0, sticky=tk.W, pady=5)
            price_var = tk.StringVar()
            tk.Entry(form_frame, textvariable=price_var, width=30).grid(row=3, column=1, pady=5)

            # 数量
            tk.Label(form_frame, text="数量*:", font=font).grid(row=4, column=0, sticky=tk.W, pady=5)
            quantity_var = tk.StringVar()
            tk.Entry(form_frame, textvariable=quantity_var, width=30).grid(row=4, column=1, pady=5)

            # 合同金额
            tk.Label(form_frame, text="合同金额*:", font=font).grid(row=5, column=0, sticky=tk.W, pady=5)
            amount_var = tk.StringVar()
            tk.Entry(form_frame, textvariable=amount_var, width=30).grid(row=5, column=1, pady=5)
            amount_var.set("0.00")

            # 自动计算合同金额
            def calculate_amount(*args):
                try:
                    price = float(price_var.get() or 0)
                    quantity = float(quantity_var.get() or 0)
                    amount = price * quantity
                    amount_var.set(f"{amount:.2f}")
                except ValueError:
                    amount_var.set("0.00")

            # 绑定事件
            price_var.trace_add("write", calculate_amount)
            quantity_var.trace_add("write", calculate_amount)

            # 付款方式
            tk.Label(form_frame, text="付款方式*:", font=font).grid(row=6, column=0, sticky=tk.W, pady=5)
            payment_var = tk.StringVar(value="电汇")
            ttk.Combobox(form_frame, textvariable=payment_var, width=28, values=["电汇", "支票", "现金", "其他"]).grid(row=6, column=1, pady=5)

            # 交货日期
            tk.Label(form_frame, text="交货日期*:", font=font).grid(row=7, column=0, sticky=tk.W, pady=5)
            delivery_var = tk.StringVar()
            tk.Entry(form_frame, textvariable=delivery_var, width=30).grid(row=7, column=1, pady=5)

            # 状态
            tk.Label(form_frame, text="状态*:", font=font).grid(row=8, column=0, sticky=tk.W, pady=5)
            status_var = tk.StringVar(value="待执行")
            ttk.Combobox(form_frame, textvariable=status_var, width=28, values=["待执行", "执行中", "已完成", "已取消"]).grid(row=8, column=1, pady=5)

            def save_contract():
                # 数据验证
                if not customer_var.get():
                    messagebox.showerror("错误", "客户名称不能为空!")
                    return
                if not salesman_var.get():
                    messagebox.showerror("错误", "业务员不能为空!")
                    return
                if not date_var.get():
                    messagebox.showerror("错误", "签订日期不能为空!")
                    return
                if not amount_var.get():
                    messagebox.showerror("错误", "合同金额不能为空!")
                    return
                if not delivery_var.get():
                    messagebox.showerror("错误", "交货日期不能为空!")
                    return

                # 准备数据
                contract_data = {
                    "客户名称": customer_var.get(),
                    "业务员": salesman_var.get(),
                    "签订日期": date_var.get(),
                    "单价": price_var.get(),
                    "数量": quantity_var.get(),
                    "合同金额": amount_var.get(),
                    "付款方式": payment_var.get(),
                    "交货日期": delivery_var.get(),
                    "状态": status_var.get()
                }

                # 保存数据
                try:
                    if self.data_manager.add_contract(contract_data):
                        self.log_operation('添加新合同成功')
                        messagebox.showinfo("成功", "合同添加成功!")
                        add_window.destroy()
                        self.open_contract_management()  # 刷新合同列表
                    else:
                        messagebox.showerror("错误", "合同添加失败!")
                except Exception as e:
                    messagebox.showerror("错误", f"添加合同出错: {str(e)}")

            # 按钮框架
            btn_frame = tk.Frame(add_window, pady=10)
            btn_frame.pack(fill=tk.X)

            tk.Button(btn_frame, text="保存", command=save_contract).pack(side=tk.LEFT, padx=10, expand=True)
            tk.Button(btn_frame, text="取消", command=add_window.destroy).pack(side=tk.RIGHT, padx=10, expand=True)

        # 创建合同表格
        columns = ("合同编号", "客户名称", "业务员", "签订日期", "合同金额", "付款方式", "交货日期", "状态")
        tree = ttk.Treeview(contract_frame, columns=columns, show="headings")

        def edit_contract():
            """编辑合同"""
            self.log_operation('打开编辑合同对话框')
            # 获取选中的合同获取选中的行
            selected_items = tree.selection()
            if not selected_items:
                messagebox.showwarning("警告", "请先选择要编辑的合同!")
                return

            # 获取选中行的数据
            selected_item = selected_items[0]
            contract_id = tree.item(selected_item, "values")[0]

            # 获取合同详情
            try:
                df = self.data_manager.get_all_contracts()
                # 确保合同编号是相同类型
                contract_id = str(contract_id)
                df['合同编号'] = df['合同编号'].astype(str)
                
                # 查找匹配的合同
                matching_contracts = df[df['合同编号'] == contract_id]
                
                if matching_contracts.empty:
                    messagebox.showerror("错误", f"未找到合同编号为 {contract_id} 的合同!")
                    return
                
                contract_data = matching_contracts.iloc[0].to_dict()
            except Exception as e:
                messagebox.showerror("错误", f"获取合同信息失败: {str(e)}")
                return

            # 创建编辑合同对话框
            edit_window = tk.Toplevel(self.root)
            edit_window.title("编辑合同")
            edit_window.geometry("500x400")
            edit_window.resizable(False, False)
            edit_window.grab_set()  # 模态窗口

            # 设置字体
            font = ('SimHei', 10)

            # 创建表单框架
            form_frame = tk.Frame(edit_window, padx=20, pady=20)
            form_frame.pack(fill=tk.BOTH, expand=True)

            # 合同编号（不可编辑）
            tk.Label(form_frame, text="合同编号:", font=font).grid(row=0, column=0, sticky=tk.W, pady=5)
            id_var = tk.StringVar(value=contract_data['合同编号'])
            tk.Entry(form_frame, textvariable=id_var, width=30, state='disabled').grid(row=0, column=1, pady=5)

            # 客户名称
            tk.Label(form_frame, text="客户名称*:", font=font).grid(row=1, column=0, sticky=tk.W, pady=5)
            customer_var = tk.StringVar(value=contract_data['客户名称'])
            tk.Entry(form_frame, textvariable=customer_var, width=30).grid(row=1, column=1, pady=5)

            # 业务员
            tk.Label(form_frame, text="业务员*:", font=font).grid(row=2, column=0, sticky=tk.W, pady=5)
            salesman_var = tk.StringVar(value=contract_data['业务员'])
            tk.Entry(form_frame, textvariable=salesman_var, width=30).grid(row=2, column=1, pady=5)

            # 签订日期
            tk.Label(form_frame, text="签订日期*:", font=font).grid(row=3, column=0, sticky=tk.W, pady=5)
            date_var = tk.StringVar(value=contract_data['签订日期'])
            tk.Entry(form_frame, textvariable=date_var, width=30).grid(row=3, column=1, pady=5)

            # 合同金额
            tk.Label(form_frame, text="合同金额*:", font=font).grid(row=4, column=0, sticky=tk.W, pady=5)
            amount_var = tk.StringVar(value=contract_data['合同金额'])
            tk.Entry(form_frame, textvariable=amount_var, width=30).grid(row=4, column=1, pady=5)

            # 付款方式
            tk.Label(form_frame, text="付款方式*:", font=font).grid(row=5, column=0, sticky=tk.W, pady=5)
            payment_var = tk.StringVar(value=contract_data['付款方式'])
            ttk.Combobox(form_frame, textvariable=payment_var, width=28, values=["电汇", "支票", "现金", "其他"]).grid(row=5, column=1, pady=5)

            # 交货日期
            tk.Label(form_frame, text="交货日期*:", font=font).grid(row=6, column=0, sticky=tk.W, pady=5)
            delivery_var = tk.StringVar(value=contract_data['交货日期'])
            tk.Entry(form_frame, textvariable=delivery_var, width=30).grid(row=6, column=1, pady=5)

            # 状态
            tk.Label(form_frame, text="状态*:", font=font).grid(row=7, column=0, sticky=tk.W, pady=5)
            status_var = tk.StringVar(value=contract_data['状态'])
            ttk.Combobox(form_frame, textvariable=status_var, width=28, values=["待执行", "执行中", "已完成", "已取消"]).grid(row=7, column=1, pady=5)

            def save_changes():
                # 数据验证
                if not customer_var.get():
                    messagebox.showerror("错误", "客户名称不能为空!")
                    return
                if not salesman_var.get():
                    messagebox.showerror("错误", "业务员不能为空!")
                    return
                if not date_var.get():
                    messagebox.showerror("错误", "签订日期不能为空!")
                    return
                if not amount_var.get():
                    messagebox.showerror("错误", "合同金额不能为空!")
                    return
                if not delivery_var.get():
                    messagebox.showerror("错误", "交货日期不能为空!")
                    return

                # 准备更新数据
                update_data = {
                    "客户名称": customer_var.get(),
                    "业务员": salesman_var.get(),
                    "签订日期": date_var.get(),
                    "合同金额": amount_var.get(),
                    "付款方式": payment_var.get(),
                    "交货日期": delivery_var.get(),
                    "状态": status_var.get()
                }

                # 更新数据
                try:
                    if self.data_manager.update_contract(contract_id, update_data):
                        self.log_operation(f'更新合同 {contract_id} 成功')
                        messagebox.showinfo("成功", "合同更新成功!")
                        edit_window.destroy()
                        self.open_contract_management()  # 刷新合同列表
                    else:
                        messagebox.showerror("错误", "合同更新失败!")
                except Exception as e:
                    messagebox.showerror("错误", f"更新合同出错: {str(e)}")

            # 按钮框架
            btn_frame = tk.Frame(edit_window, pady=10)
            btn_frame.pack(fill=tk.X)

            tk.Button(btn_frame, text="保存", command=save_changes).pack(side=tk.LEFT, padx=10, expand=True)
            tk.Button(btn_frame, text="取消", command=edit_window.destroy).pack(side=tk.RIGHT, padx=10, expand=True)

        def delete_contract():
            """删除合同"""
            selected_item = tree.selection()
            if selected_item:
                item = tree.item(selected_item[0])
                contract_id = item['values'][0]
                self.log_operation(f'删除合同 {contract_id}')
            # 获取选中的合同获取选中的行
            selected_items = tree.selection()
            if not selected_items:
                messagebox.showwarning("警告", "请先选择要删除的合同!")
                return

            # 获取选中行的数据
            selected_item = selected_items[0]
            contract_id = tree.item(selected_item, "values")[0]
            contract_name = tree.item(selected_item, "values")[1]

            # 确认删除
            if messagebox.askyesno("确认", f"确定要删除合同编号为 {contract_id} 的 {contract_name} 合同吗?"):
                try:
                    if self.data_manager.delete_contract(contract_id):
                        messagebox.showinfo("成功", "合同删除成功!")
                        self.open_contract_management()  # 刷新合同列表
                    else:
                        messagebox.showerror("错误", "合同删除失败!")
                except Exception as e:
                    messagebox.showerror("错误", f"删除合同出错: {str(e)}")

        tk.Button(btn_frame, text="添加合同", command=add_contract).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="编辑合同", command=edit_contract).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="删除合同", command=delete_contract).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="返回主界面", command=self.create_main_frame).pack(side=tk.RIGHT, padx=5)

        # 设置列宽和标题
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=100, anchor=tk.CENTER)

        # 调整部分列的宽度
        tree.column("合同编号", width=120)
        tree.column("客户名称", width=150)
        tree.column("业务员", width=100)
        tree.column("合同金额", width=100)
        tree.column("状态", width=80)

        # 添加滚动条
        scrollbar = ttk.Scrollbar(contract_frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        tree.pack(fill=tk.BOTH, expand=True)

        # 加载合同数据
        try:
            df = pd.read_excel(self.contract_file, engine='openpyxl')
            for index, row in df.iterrows():
                tree.insert('', tk.END, values=(
                    row['合同编号'],
                    row['客户名称'],
                    row['业务员'],
                    row['签订日期'],
                    row['合同金额'],
                    row['付款方式'],
                    row['交货日期'],
                    row['状态']
                ))
        except Exception as e:
            messagebox.showerror("错误", f"加载合同数据失败: {str(e)}")

    def open_payment_management(self):
        """打开收款管理界面"""
        self.log_operation('打开收款管理界面')
        # 清空当前界面
        for widget in self.root.winfo_children():
            if isinstance(widget, tk.Frame) and widget != self.root.nametowidget('.!menu'):
                widget.destroy()

        # 创建收款管理框架
        payment_frame = tk.Frame(self.root)
        payment_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 创建搜索框
        search_frame = tk.Frame(payment_frame)
        search_frame.pack(fill=tk.X, padx=5, pady=5)

        tk.Label(search_frame, text="搜索:", font=('SimHei', 10)).pack(side=tk.LEFT, padx=5)
        search_var = tk.StringVar()
        search_entry = tk.Entry(search_frame, textvariable=search_var, width=30)
        search_entry.pack(side=tk.LEFT, padx=5)

        tk.Label(search_frame, text="搜索类型:", font=('SimHei', 10)).pack(side=tk.LEFT, padx=5)
        search_type = tk.StringVar(value="收款ID")
        search_type_combo = ttk.Combobox(search_frame, textvariable=search_type, width=10,
                                        values=["收款ID", "合同编号", "收款日期"])
        search_type_combo.pack(side=tk.LEFT, padx=5)

        def search_payment():
            # 获取搜索参数
            search_term = search_var.get().strip()
            search_type_val = search_type.get()

            if not search_term:
                messagebox.showwarning("警告", "请输入搜索内容")
                return

            # 执行搜索
            try:
                df = self.data_manager.search_payments(search_term, search_type_val)

                # 清空表格
                for item in tree.get_children():
                    tree.delete(item)

                # 加载搜索结果
                for index, row in df.iterrows():
                    tree.insert('', tk.END, values=(
                        row['收款ID'],
                        row['合同编号'],
                        row['收款日期'],
                        row['收款金额'],
                        row['收款方式']
                    ))
            except Exception as e:
                messagebox.showerror("错误", f"搜索失败: {str(e)}")

        tk.Button(search_frame, text="搜索", command=search_payment).pack(side=tk.LEFT, padx=5)
        tk.Button(search_frame, text="刷新", command=self.open_payment_management).pack(side=tk.LEFT, padx=5)

        # 创建按钮组
        btn_frame = tk.Frame(payment_frame)
        btn_frame.pack(fill=tk.X, padx=5, pady=5)

        def add_payment():
            # 创建添加收款对话框
            add_window = tk.Toplevel(self.root)
            add_window.title("添加收款")
            add_window.geometry("500x400")
            add_window.resizable(False, False)
            add_window.transient(self.root)
            add_window.grab_set()

            # 获取所有合同编号用于下拉选择
            contract_ids = []
            try:
                df_contracts = pd.read_excel(self.contract_file, engine='openpyxl')
                contract_ids = df_contracts['合同编号'].tolist()
            except Exception as e:
                messagebox.showerror("错误", f"读取合同数据失败: {str(e)}")

            # 创建表单
            form_frame = tk.Frame(add_window, padx=20, pady=20)
            form_frame.pack(fill=tk.BOTH, expand=True)

            # 合同编号
            tk.Label(form_frame, text="合同编号:", font=('SimHei', 10)).grid(row=0, column=0, sticky=tk.W, pady=5)
            contract_id_var = tk.StringVar()
            contract_id_combo = ttk.Combobox(form_frame, textvariable=contract_id_var, width=30, values=contract_ids)
            contract_id_combo.grid(row=0, column=1, pady=5)

            # 收款日期
            tk.Label(form_frame, text="收款日期:", font=('SimHei', 10)).grid(row=1, column=0, sticky=tk.W, pady=5)
            payment_date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
            tk.Entry(form_frame, textvariable=payment_date_var, width=32).grid(row=1, column=1, pady=5)

            # 收款金额
            tk.Label(form_frame, text="收款金额:", font=('SimHei', 10)).grid(row=2, column=0, sticky=tk.W, pady=5)
            amount_var = tk.StringVar()
            tk.Entry(form_frame, textvariable=amount_var, width=32).grid(row=2, column=1, pady=5)

            # 收款方式
            tk.Label(form_frame, text="收款方式:", font=('SimHei', 10)).grid(row=3, column=0, sticky=tk.W, pady=5)
            payment_method_var = tk.StringVar()
            payment_method_combo = ttk.Combobox(form_frame, textvariable=payment_method_var, width=30, 
                                               values=["银行转账", "信用证", "现金", "其他"])
            payment_method_combo.grid(row=3, column=1, pady=5)

            # 备注
            tk.Label(form_frame, text="备注:", font=('SimHei', 10)).grid(row=4, column=0, sticky=tk.NW, pady=5)
            remark_var = tk.StringVar()
            tk.Text(form_frame, height=5, width=25).grid(row=4, column=1, pady=5)

            # 保存按钮
            def save_payment():
                payment_data = {
                    '合同编号': contract_id_var.get(),
                    '收款日期': payment_date_var.get(),
                    '收款金额': amount_var.get(),
                    '收款方式': payment_method_var.get(),
                    '备注': remark_var.get()
                }

                # 验证必填字段
                if not payment_data['合同编号'] or not payment_data['收款金额']:
                    messagebox.showwarning("警告", "合同编号和收款金额不能为空")
                    return

                # 尝试转换金额为数字
                try:
                    float(payment_data['收款金额'])
                except ValueError:
                    messagebox.showwarning("警告", "收款金额必须是数字")
                    return

                # 保存数据
                if self.data_manager.add_payment(payment_data):
                    self.log_operation(f'为合同 {payment_data["合同编号"]} 添加收款 {payment_data["收款金额"]} 元')
                    messagebox.showinfo("成功", "收款添加成功")
                    add_window.destroy()
                    self.open_payment_management()
                else:
                    messagebox.showerror("错误", "收款添加失败")

            # 按钮框
            btn_frame = tk.Frame(add_window)
            btn_frame.pack(pady=10)

            tk.Button(btn_frame, text="保存", command=save_payment).pack(side=tk.LEFT, padx=10)
            tk.Button(btn_frame, text="取消", command=add_window.destroy).pack(side=tk.LEFT, padx=10)

        def edit_payment():
            # 获取选中的项
            selected_item = tree.selection()
            if not selected_item:
                messagebox.showwarning("警告", "请先选择一条收款记录")
                return

            # 获取选中项的收款ID
            payment_id = tree.item(selected_item[0], "values")[0]

            # 获取该收款的详细信息
            try:
                df = pd.read_excel(self.payment_file, engine='openpyxl')
                payment_data = df[df['收款ID'] == payment_id].iloc[0].to_dict()
            except Exception as e:
                messagebox.showerror("错误", f"读取收款数据失败: {str(e)}")
                return

            # 创建编辑对话框
            edit_window = tk.Toplevel(self.root)
            edit_window.title("编辑收款")
            edit_window.geometry("500x400")
            edit_window.resizable(False, False)
            edit_window.transient(self.root)
            edit_window.grab_set()

            # 获取所有合同编号用于下拉选择
            contract_ids = []
            try:
                df_contracts = pd.read_excel(self.contract_file, engine='openpyxl')
                contract_ids = df_contracts['合同编号'].tolist()
            except Exception as e:
                messagebox.showerror("错误", f"读取合同数据失败: {str(e)}")

            # 创建表单
            form_frame = tk.Frame(edit_window, padx=20, pady=20)
            form_frame.pack(fill=tk.BOTH, expand=True)

            # 收款ID（不可编辑）
            tk.Label(form_frame, text="收款ID:", font=('SimHei', 10)).grid(row=0, column=0, sticky=tk.W, pady=5)
            tk.Label(form_frame, text=payment_data['收款ID'], font=('SimHei', 10)).grid(row=0, column=1, sticky=tk.W, pady=5)

            # 合同编号
            tk.Label(form_frame, text="合同编号:", font=('SimHei', 10)).grid(row=1, column=0, sticky=tk.W, pady=5)
            contract_id_var = tk.StringVar(value=payment_data['合同编号'])
            contract_id_combo = ttk.Combobox(form_frame, textvariable=contract_id_var, width=30, values=contract_ids)
            contract_id_combo.grid(row=1, column=1, pady=5)

            # 收款日期
            tk.Label(form_frame, text="收款日期:", font=('SimHei', 10)).grid(row=2, column=0, sticky=tk.W, pady=5)
            payment_date_var = tk.StringVar(value=str(payment_data['收款日期']))
            tk.Entry(form_frame, textvariable=payment_date_var, width=32).grid(row=2, column=1, pady=5)

            # 收款金额
            tk.Label(form_frame, text="收款金额:", font=('SimHei', 10)).grid(row=3, column=0, sticky=tk.W, pady=5)
            amount_var = tk.StringVar(value=str(payment_data['收款金额']))
            tk.Entry(form_frame, textvariable=amount_var, width=32).grid(row=3, column=1, pady=5)

            # 收款方式
            tk.Label(form_frame, text="收款方式:", font=('SimHei', 10)).grid(row=4, column=0, sticky=tk.W, pady=5)
            payment_method_var = tk.StringVar(value=payment_data['收款方式'])
            payment_method_combo = ttk.Combobox(form_frame, textvariable=payment_method_var, width=30, 
                                               values=["银行转账", "信用证", "现金", "其他"])
            payment_method_combo.grid(row=4, column=1, pady=5)

            # 备注
            tk.Label(form_frame, text="备注:", font=('SimHei', 10)).grid(row=5, column=0, sticky=tk.NW, pady=5)
            remark_text = tk.Text(form_frame, height=5, width=25)
            remark_text.grid(row=5, column=1, pady=5)
            if '备注' in payment_data and pd.notna(payment_data['备注']):
                remark_text.insert(tk.END, payment_data['备注'])

            # 保存按钮
            def update_payment():
                update_data = {
                    '合同编号': contract_id_var.get(),
                    '收款日期': payment_date_var.get(),
                    '收款金额': amount_var.get(),
                    '收款方式': payment_method_var.get(),
                    '备注': remark_text.get(1.0, tk.END).strip()
                }

                # 验证必填字段
                if not update_data['合同编号'] or not update_data['收款金额']:
                    messagebox.showwarning("警告", "合同编号和收款金额不能为空")
                    return

                # 尝试转换金额为数字
                try:
                    float(update_data['收款金额'])
                except ValueError:
                    messagebox.showwarning("警告", "收款金额必须是数字")
                    return

                # 更新数据
                if self.data_manager.update_payment(payment_id, update_data):
                    self.log_operation(f'更新收款记录，收款ID: {payment_id}，合同编号: {update_data["合同编号"]}，金额: {update_data["收款金额"]} 元')
                    messagebox.showinfo("成功", "收款更新成功")
                    edit_window.destroy()
                    self.open_payment_management()
                else:
                    messagebox.showerror("错误", "收款更新失败")

            # 按钮框
            btn_frame = tk.Frame(edit_window)
            btn_frame.pack(pady=10)

            tk.Button(btn_frame, text="保存", command=update_payment).pack(side=tk.LEFT, padx=10)
            tk.Button(btn_frame, text="取消", command=edit_window.destroy).pack(side=tk.LEFT, padx=10)

        def delete_payment():
            # 获取选中的项
            selected_item = tree.selection()
            if not selected_item:
                messagebox.showwarning("警告", "请先选择一条收款记录")
                return

            # 获取选中项的收款ID
            payment_id = tree.item(selected_item[0], "values")[0]

            # 确认删除
            if messagebox.askyesno("确认", f"确定要删除收款ID为 {payment_id} 的记录吗?"):
                # 删除数据
                if self.data_manager.delete_payment(payment_id):
                    self.log_operation(f'删除收款记录，收款ID: {payment_id}')
                    messagebox.showinfo("成功", "收款删除成功")
                    self.open_payment_management()
                else:
                    messagebox.showerror("错误", "收款删除失败")

        tk.Button(btn_frame, text="添加收款", command=add_payment).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="编辑收款", command=edit_payment).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="删除收款", command=delete_payment).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="返回主界面", command=self.create_main_frame).pack(side=tk.RIGHT, padx=5)

        # 创建收款表格
        columns = ("收款ID", "合同编号", "收款日期", "收款金额", "收款方式")
        tree = ttk.Treeview(payment_frame, columns=columns, show="headings")

        # 设置列宽和标题
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=100, anchor=tk.CENTER)

        # 调整部分列的宽度
        tree.column("收款ID", width=100)
        tree.column("合同编号", width=120)
        tree.column("收款金额", width=100)

        # 添加滚动条
        scrollbar = ttk.Scrollbar(payment_frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        tree.pack(fill=tk.BOTH, expand=True)

        # 加载收款数据
        try:
            df = pd.read_excel(self.payment_file, engine='openpyxl')
            for index, row in df.iterrows():
                tree.insert('', tk.END, values=(
                    row['收款ID'],
                    row['合同编号'],
                    row['收款日期'],
                    row['收款金额'],
                    row['收款方式']
                ))
        except Exception as e:
            messagebox.showerror("错误", f"加载收款数据失败: {str(e)}")

    def open_customer_management(self):
        """打开客户管理界面"""
        self.log_operation('打开客户管理界面')
        # 清空当前界面
        for widget in self.root.winfo_children():
            if isinstance(widget, tk.Frame) and widget != self.root.nametowidget('.!menu'):
                widget.destroy()

        # 创建客户管理框架
        customer_frame = tk.Frame(self.root)
        customer_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 创建搜索框
        search_frame = tk.Frame(customer_frame)
        search_frame.pack(fill=tk.X, padx=5, pady=5)

        tk.Label(search_frame, text="搜索:", font=('SimHei', 10)).pack(side=tk.LEFT, padx=5)
        search_var = tk.StringVar()
        search_entry = tk.Entry(search_frame, textvariable=search_var, width=30)
        search_entry.pack(side=tk.LEFT, padx=5)

        tk.Label(search_frame, text="搜索类型:", font=('SimHei', 10)).pack(side=tk.LEFT, padx=5)
        search_type = tk.StringVar(value="客户ID")
        search_type_combo = ttk.Combobox(search_frame, textvariable=search_type, width=10,
                                        values=["客户ID", "客户名称", "联系人"])
        search_type_combo.pack(side=tk.LEFT, padx=5)

        # 添加搜索输入框
        search_entry = tk.Entry(search_frame, width=20)
        search_entry.pack(side=tk.LEFT, padx=5)

        def search_customer():
            # 获取搜索参数
            search_term = search_entry.get().strip()
            search_type_val = search_type.get()

            # 验证输入
            if not search_term:
                tk.messagebox.showwarning("警告", "请输入搜索内容！")
                return

            try:
                # 执行搜索
                df = self.data_manager.search_customers(search_term, search_type_val)

                # 清空表格
                for item in tree.get_children():
                    tree.delete(item)

                # 显示搜索结果
                if df.empty:
                    tk.messagebox.showinfo("提示", "未找到匹配的客户记录！")
                else:
                    for index, row in df.iterrows():
                        tree.insert('', tk.END, values=(
                            row['客户ID'],
                            row['客户名称'],
                            row['联系人'],
                            row['联系电话'],
                            row['地址'] if pd.notna(row['地址']) else '',
                            row['邮箱'] if pd.notna(row['邮箱']) else ''
                        ))
            except Exception as e:
                tk.messagebox.showerror("错误", f"搜索失败: {str(e)}")

        tk.Button(search_frame, text="搜索", command=search_customer).pack(side=tk.LEFT, padx=5)
        tk.Button(search_frame, text="刷新", command=self.open_customer_management).pack(side=tk.LEFT, padx=5)

        # 创建按钮组
        btn_frame = tk.Frame(customer_frame)
        btn_frame.pack(fill=tk.X, padx=5, pady=5)

        def add_customer():
            # 创建添加客户对话框
            add_window = tk.Toplevel(self.root)
            add_window.title("添加客户")
            add_window.geometry("500x350")
            add_window.resizable(False, False)
            add_window.transient(self.root)
            add_window.grab_set()

            # 创建表单框架
            form_frame = tk.Frame(add_window, padx=20, pady=20)
            form_frame.pack(fill=tk.BOTH, expand=True)

            # 客户名称
            tk.Label(form_frame, text="客户名称:*", font=('SimHei', 10)).grid(row=0, column=0, sticky=tk.W, pady=5)
            name_var = tk.StringVar()
            tk.Entry(form_frame, textvariable=name_var, width=30).grid(row=0, column=1, pady=5)

            # 联系人
            tk.Label(form_frame, text="联系人:*", font=('SimHei', 10)).grid(row=1, column=0, sticky=tk.W, pady=5)
            contact_var = tk.StringVar()
            tk.Entry(form_frame, textvariable=contact_var, width=30).grid(row=1, column=1, pady=5)

            # 联系电话
            tk.Label(form_frame, text="联系电话:*", font=('SimHei', 10)).grid(row=2, column=0, sticky=tk.W, pady=5)
            phone_var = tk.StringVar()
            tk.Entry(form_frame, textvariable=phone_var, width=30).grid(row=2, column=1, pady=5)

            # 地址
            tk.Label(form_frame, text="地址:", font=('SimHei', 10)).grid(row=3, column=0, sticky=tk.W, pady=5)
            address_var = tk.StringVar()
            tk.Entry(form_frame, textvariable=address_var, width=30).grid(row=3, column=1, pady=5)

            # 邮箱
            tk.Label(form_frame, text="邮箱:", font=('SimHei', 10)).grid(row=4, column=0, sticky=tk.W, pady=5)
            email_var = tk.StringVar()
            tk.Entry(form_frame, textvariable=email_var, width=30).grid(row=4, column=1, pady=5)

            # 保存按钮
            def save_customer():
                name = name_var.get().strip()
                contact = contact_var.get().strip()
                phone = phone_var.get().strip()
                address = address_var.get().strip()
                email = email_var.get().strip()

                # 验证必填字段
                if not name or not contact or not phone:
                    tk.messagebox.showerror("错误", "客户名称、联系人和联系电话为必填项！")
                    return

                # 准备客户数据
                customer_data = {
                    '客户名称': name,
                    '联系人': contact,
                    '联系电话': phone,
                    '地址': address,
                    '邮箱': email
                }

                # 保存数据
                if self.data_manager.add_customer(customer_data):
                    self.log_operation(f'添加新客户: {customer_data["客户名称"]}')
                    tk.messagebox.showinfo("成功", "客户添加成功！")
                    add_window.destroy()
                    self.open_customer_management()
                else:
                    tk.messagebox.showerror("错误", "客户添加失败！")

            # 按钮框架
            btn_frame = tk.Frame(add_window, pady=10)
            btn_frame.pack(fill=tk.X)

            tk.Button(btn_frame, text="保存", command=save_customer).pack(side=tk.LEFT, padx=20)
            tk.Button(btn_frame, text="取消", command=add_window.destroy).pack(side=tk.RIGHT, padx=20)

        def edit_customer():
            # 获取选中的客户
            selected_item = tree.selection()
            if not selected_item:
                tk.messagebox.showwarning("警告", "请先选择要编辑的客户！")
                return

            # 获取客户ID
            item = tree.item(selected_item[0])
            customer_id = item['values'][0]

            # 获取客户详情
            df = self.data_manager.get_all_customers()
            customer = df[df['客户ID'] == customer_id].iloc[0].to_dict()

            # 创建编辑客户对话框
            edit_window = tk.Toplevel(self.root)
            edit_window.title("编辑客户")
            edit_window.geometry("500x350")
            edit_window.resizable(False, False)
            edit_window.transient(self.root)
            edit_window.grab_set()

            # 创建表单框架
            form_frame = tk.Frame(edit_window, padx=20, pady=20)
            form_frame.pack(fill=tk.BOTH, expand=True)

            # 客户ID（不可编辑）
            tk.Label(form_frame, text="客户ID:", font=('SimHei', 10)).grid(row=0, column=0, sticky=tk.W, pady=5)
            tk.Label(form_frame, text=customer_id, font=('SimHei', 10)).grid(row=0, column=1, sticky=tk.W, pady=5)

            # 客户名称
            tk.Label(form_frame, text="客户名称:*", font=('SimHei', 10)).grid(row=1, column=0, sticky=tk.W, pady=5)
            name_var = tk.StringVar(value=customer['客户名称'])
            tk.Entry(form_frame, textvariable=name_var, width=30).grid(row=1, column=1, pady=5)

            # 联系人
            tk.Label(form_frame, text="联系人:*", font=('SimHei', 10)).grid(row=2, column=0, sticky=tk.W, pady=5)
            contact_var = tk.StringVar(value=customer['联系人'])
            tk.Entry(form_frame, textvariable=contact_var, width=30).grid(row=2, column=1, pady=5)

            # 联系电话
            tk.Label(form_frame, text="联系电话:*", font=('SimHei', 10)).grid(row=3, column=0, sticky=tk.W, pady=5)
            phone_var = tk.StringVar(value=customer['联系电话'])
            tk.Entry(form_frame, textvariable=phone_var, width=30).grid(row=3, column=1, pady=5)

            # 地址
            tk.Label(form_frame, text="地址:", font=('SimHei', 10)).grid(row=4, column=0, sticky=tk.W, pady=5)
            address_var = tk.StringVar(value=customer['地址'] if pd.notna(customer['地址']) else '')
            tk.Entry(form_frame, textvariable=address_var, width=30).grid(row=4, column=1, pady=5)

            # 邮箱
            tk.Label(form_frame, text="邮箱:", font=('SimHei', 10)).grid(row=5, column=0, sticky=tk.W, pady=5)
            email_var = tk.StringVar(value=customer['邮箱'] if pd.notna(customer['邮箱']) else '')
            tk.Entry(form_frame, textvariable=email_var, width=30).grid(row=5, column=1, pady=5)

            # 保存按钮
            def update_customer():
                name = name_var.get().strip()
                contact = contact_var.get().strip()
                phone = phone_var.get().strip()
                address = address_var.get().strip()
                email = email_var.get().strip()

                # 验证必填字段
                if not name or not contact or not phone:
                    tk.messagebox.showerror("错误", "客户名称、联系人和联系电话为必填项！")
                    return

                # 准备更新数据
                update_data = {
                    '客户名称': name,
                    '联系人': contact,
                    '联系电话': phone,
                    '地址': address,
                    '邮箱': email
                }

                # 更新数据
                if self.data_manager.update_customer(customer_id, update_data):
                    self.log_operation(f'更新客户信息，客户ID: {customer_id}，客户名称: {update_data["客户名称"]}')
                    tk.messagebox.showinfo("成功", "客户更新成功！")
                    edit_window.destroy()
                    self.open_customer_management()
                else:
                    tk.messagebox.showerror("错误", "客户更新失败！")

            # 按钮框架
            btn_frame = tk.Frame(edit_window, pady=10)
            btn_frame.pack(fill=tk.X)

            tk.Button(btn_frame, text="保存", command=update_customer).pack(side=tk.LEFT, padx=20)
            tk.Button(btn_frame, text="取消", command=edit_window.destroy).pack(side=tk.RIGHT, padx=20)

        def delete_customer():
            # 获取选中的客户
            selected_item = tree.selection()
            if not selected_item:
                tk.messagebox.showwarning("警告", "请先选择要删除的客户！")
                return

            # 确认删除
            if not tk.messagebox.askyesno("确认", "确定要删除选中的客户吗？"):
                return

            # 获取客户ID
            item = tree.item(selected_item[0])
            customer_id = item['values'][0]

            # 删除客户
            if self.data_manager.delete_customer(customer_id):
                self.log_operation(f'删除客户，客户ID: {customer_id}')
                tk.messagebox.showinfo("成功", "客户删除成功！")
                self.open_customer_management()
            else:
                tk.messagebox.showerror("错误", "客户删除失败！")

        tk.Button(btn_frame, text="添加客户", command=add_customer).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="编辑客户", command=edit_customer).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="删除客户", command=delete_customer).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="返回主界面", command=self.create_main_frame).pack(side=tk.RIGHT, padx=5)

        # 创建客户表格
        columns = ("客户ID", "客户名称", "联系人", "联系电话", "地址", "邮箱")
        tree = ttk.Treeview(customer_frame, columns=columns, show="headings")

        # 设置列宽和标题
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=100, anchor=tk.CENTER)

        # 调整部分列的宽度
        tree.column("客户ID", width=80)
        tree.column("客户名称", width=150)
        tree.column("联系人", width=100)
        tree.column("联系电话", width=120)
        tree.column("地址", width=200)
        tree.column("邮箱", width=150)

        # 添加滚动条
        scrollbar = ttk.Scrollbar(customer_frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        tree.pack(fill=tk.BOTH, expand=True)

        # 加载客户数据
        try:
            df = pd.read_excel(self.customer_file, engine='openpyxl')
            for index, row in df.iterrows():
                tree.insert('', tk.END, values=(
                    row['客户ID'],
                    row['客户名称'],
                    row['联系人'],
                    row['联系电话'],
                    row['地址'],
                    row['邮箱']
                ))
        except Exception as e:
            messagebox.showerror("错误", f"加载客户数据失败: {str(e)}")

    def open_salesman_management(self):
        """打开业务员管理界面"""
        self.log_operation('打开业务员管理界面')
        # 清空当前界面
        for widget in self.root.winfo_children():
            if isinstance(widget, tk.Frame) and widget != self.root.nametowidget('.!menu'):
                widget.destroy()

        # 创建业务员管理框架
        salesman_frame = tk.Frame(self.root)
        salesman_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 创建搜索框
        search_frame = tk.Frame(salesman_frame)
        search_frame.pack(fill=tk.X, padx=5, pady=5)

        tk.Label(search_frame, text="搜索:", font=('SimHei', 10)).pack(side=tk.LEFT, padx=5)
        search_var = tk.StringVar()
        search_entry = tk.Entry(search_frame, textvariable=search_var, width=30)
        search_entry.pack(side=tk.LEFT, padx=5)

        tk.Label(search_frame, text="搜索类型:", font=('SimHei', 10)).pack(side=tk.LEFT, padx=5)
        search_type = tk.StringVar(value="业务员ID")
        search_type_combo = ttk.Combobox(search_frame, textvariable=search_type, width=10,
                                        values=["业务员ID", "姓名", "所属部门"])
        search_type_combo.pack(side=tk.LEFT, padx=5)

        def search_salesman():
            # 获取搜索参数
            search_term = search_var.get().strip()
            search_type_val = search_type.get()

            # 验证输入
            if not search_term:
                tk.messagebox.showwarning("警告", "请输入搜索内容！")
                return

            try:
                # 执行搜索
                df = self.data_manager.search_salesmen(search_term, search_type_val)

                # 清空表格
                for item in tree.get_children():
                    tree.delete(item)

                # 显示搜索结果
                if df.empty:
                    tk.messagebox.showinfo("提示", "未找到匹配的业务员记录！")
                else:
                    for index, row in df.iterrows():
                        tree.insert('', tk.END, values=(
                            row['业务员ID'],
                            row['姓名'],
                            row['联系电话'],
                            row['邮箱'] if pd.notna(row['邮箱']) else '',
                            row['所属部门']
                        ))
            except Exception as e:
                tk.messagebox.showerror("错误", f"搜索失败: {str(e)}")

        tk.Button(search_frame, text="搜索", command=search_salesman).pack(side=tk.LEFT, padx=5)
        tk.Button(search_frame, text="刷新", command=self.open_salesman_management).pack(side=tk.LEFT, padx=5)

        # 创建按钮组
        btn_frame = tk.Frame(salesman_frame)
        btn_frame.pack(fill=tk.X, padx=5, pady=5)

        def add_salesman():
            # 创建添加业务员对话框
            add_window = tk.Toplevel(self.root)
            add_window.title("添加业务员")
            add_window.geometry("500x350")
            add_window.resizable(False, False)
            add_window.transient(self.root)
            add_window.grab_set()

            # 创建表单框架
            form_frame = tk.Frame(add_window, padx=20, pady=20)
            form_frame.pack(fill=tk.BOTH, expand=True)

            # 姓名
            tk.Label(form_frame, text="姓名:*", font=('SimHei', 10)).grid(row=0, column=0, sticky=tk.W, pady=5)
            name_var = tk.StringVar()
            tk.Entry(form_frame, textvariable=name_var, width=30).grid(row=0, column=1, pady=5)

            # 联系电话
            tk.Label(form_frame, text="联系电话:*", font=('SimHei', 10)).grid(row=1, column=0, sticky=tk.W, pady=5)
            phone_var = tk.StringVar()
            tk.Entry(form_frame, textvariable=phone_var, width=30).grid(row=1, column=1, pady=5)

            # 邮箱
            tk.Label(form_frame, text="邮箱:", font=('SimHei', 10)).grid(row=2, column=0, sticky=tk.W, pady=5)
            email_var = tk.StringVar()
            tk.Entry(form_frame, textvariable=email_var, width=30).grid(row=2, column=1, pady=5)

            # 所属部门
            tk.Label(form_frame, text="所属部门:*", font=('SimHei', 10)).grid(row=3, column=0, sticky=tk.W, pady=5)
            department_var = tk.StringVar()
            tk.Entry(form_frame, textvariable=department_var, width=30).grid(row=3, column=1, pady=5)

            # 保存按钮
            def save_salesman():
                name = name_var.get().strip()
                phone = phone_var.get().strip()
                email = email_var.get().strip()
                department = department_var.get().strip()

                # 验证必填字段
                if not name or not phone or not department:
                    tk.messagebox.showerror("错误", "姓名、联系电话和所属部门为必填项！")
                    return

                # 准备业务员数据
                salesman_data = {
                    '姓名': name,
                    '联系电话': phone,
                    '邮箱': email,
                    '所属部门': department
                }

                # 保存数据
                if self.data_manager.add_salesman(salesman_data):
                    self.log_operation(f'添加新业务员: {salesman_data["姓名"]}')
                    tk.messagebox.showinfo("成功", "业务员添加成功！")
                    add_window.destroy()
                    self.open_salesman_management()
                else:
                    tk.messagebox.showerror("错误", "业务员添加失败！")

            # 按钮框架
            btn_frame = tk.Frame(add_window, pady=10)
            btn_frame.pack(fill=tk.X)

            tk.Button(btn_frame, text="保存", command=save_salesman).pack(side=tk.LEFT, padx=20)
            tk.Button(btn_frame, text="取消", command=add_window.destroy).pack(side=tk.RIGHT, padx=20)

        def edit_salesman():
            # 获取选中的业务员
            selected_item = tree.selection()
            if not selected_item:
                tk.messagebox.showwarning("警告", "请先选择要编辑的业务员！")
                return

            # 获取业务员ID
            item = tree.item(selected_item[0])
            salesman_id = item['values'][0]

            # 获取业务员详情
            df = self.data_manager.get_all_salesmen()
            salesman = df[df['业务员ID'] == salesman_id].iloc[0].to_dict()

            # 创建编辑业务员对话框
            edit_window = tk.Toplevel(self.root)
            edit_window.title("编辑业务员")
            edit_window.geometry("500x350")
            edit_window.resizable(False, False)
            edit_window.transient(self.root)
            edit_window.grab_set()

            # 创建表单框架
            form_frame = tk.Frame(edit_window, padx=20, pady=20)
            form_frame.pack(fill=tk.BOTH, expand=True)

            # 业务员ID（不可编辑）
            tk.Label(form_frame, text="业务员ID:", font=('SimHei', 10)).grid(row=0, column=0, sticky=tk.W, pady=5)
            tk.Label(form_frame, text=salesman_id, font=('SimHei', 10)).grid(row=0, column=1, sticky=tk.W, pady=5)

            # 姓名
            tk.Label(form_frame, text="姓名:*", font=('SimHei', 10)).grid(row=1, column=0, sticky=tk.W, pady=5)
            name_var = tk.StringVar(value=salesman['姓名'])
            tk.Entry(form_frame, textvariable=name_var, width=30).grid(row=1, column=1, pady=5)

            # 联系电话
            tk.Label(form_frame, text="联系电话:*", font=('SimHei', 10)).grid(row=2, column=0, sticky=tk.W, pady=5)
            phone_var = tk.StringVar(value=salesman['联系电话'])
            tk.Entry(form_frame, textvariable=phone_var, width=30).grid(row=2, column=1, pady=5)

            # 邮箱
            tk.Label(form_frame, text="邮箱:", font=('SimHei', 10)).grid(row=3, column=0, sticky=tk.W, pady=5)
            email_var = tk.StringVar(value=salesman['邮箱'] if pd.notna(salesman['邮箱']) else '')
            tk.Entry(form_frame, textvariable=email_var, width=30).grid(row=3, column=1, pady=5)

            # 所属部门
            tk.Label(form_frame, text="所属部门:*", font=('SimHei', 10)).grid(row=4, column=0, sticky=tk.W, pady=5)
            department_var = tk.StringVar(value=salesman['所属部门'])
            tk.Entry(form_frame, textvariable=department_var, width=30).grid(row=4, column=1, pady=5)

            # 保存按钮
            def update_salesman():
                name = name_var.get().strip()
                phone = phone_var.get().strip()
                email = email_var.get().strip()
                department = department_var.get().strip()

                # 验证必填字段
                if not name or not phone or not department:
                    tk.messagebox.showerror("错误", "姓名、联系电话和所属部门为必填项！")
                    return

                # 准备更新数据
                update_data = {
                    '姓名': name,
                    '联系电话': phone,
                    '邮箱': email,
                    '所属部门': department
                }

                # 更新数据
                if self.data_manager.update_salesman(salesman_id, update_data):
                    self.log_operation(f'更新业务员信息，业务员ID: {salesman_id}，姓名: {update_data["姓名"]}')
                    tk.messagebox.showinfo("成功", "业务员更新成功！")
                    edit_window.destroy()
                    self.open_salesman_management()
                else:
                    tk.messagebox.showerror("错误", "业务员更新失败！")

            # 按钮框架
            btn_frame = tk.Frame(edit_window, pady=10)
            btn_frame.pack(fill=tk.X)

            tk.Button(btn_frame, text="保存", command=update_salesman).pack(side=tk.LEFT, padx=20)
            tk.Button(btn_frame, text="取消", command=edit_window.destroy).pack(side=tk.RIGHT, padx=20)

        def delete_salesman():
            # 获取选中的业务员
            selected_item = tree.selection()
            if not selected_item:
                tk.messagebox.showwarning("警告", "请先选择要删除的业务员！")
                return

            # 确认删除
            if not tk.messagebox.askyesno("确认", "确定要删除选中的业务员吗？"):
                return

            # 获取业务员ID
            item = tree.item(selected_item[0])
            salesman_id = item['values'][0]

            # 删除业务员
            if self.data_manager.delete_salesman(salesman_id):
                self.log_operation(f'删除业务员，业务员ID: {salesman_id}')
                tk.messagebox.showinfo("成功", "业务员删除成功！")
                self.open_salesman_management()
            else:
                tk.messagebox.showerror("错误", "业务员删除失败！")

        tk.Button(btn_frame, text="添加业务员", command=add_salesman).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="编辑业务员", command=edit_salesman).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="删除业务员", command=delete_salesman).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="返回主界面", command=self.create_main_frame).pack(side=tk.RIGHT, padx=5)

        # 创建业务员表格
        columns = ("业务员ID", "姓名", "联系电话", "邮箱", "所属部门")
        tree = ttk.Treeview(salesman_frame, columns=columns, show="headings")

        # 设置列宽和标题
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=100, anchor=tk.CENTER)

        # 调整部分列的宽度
        tree.column("业务员ID", width=80)
        tree.column("姓名", width=100)
        tree.column("联系电话", width=120)
        tree.column("邮箱", width=150)
        tree.column("所属部门", width=100)

        # 添加滚动条
        scrollbar = ttk.Scrollbar(salesman_frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        tree.pack(fill=tk.BOTH, expand=True)

        # 加载业务员数据
        try:
            df = pd.read_excel(self.salesman_file, engine='openpyxl')
            for index, row in df.iterrows():
                tree.insert('', tk.END, values=(
                    row['业务员ID'],
                    row['姓名'],
                    row['联系电话'],
                    row['邮箱'],
                    row['所属部门']
                ))
        except Exception as e:
            messagebox.showerror("错误", f"加载业务员数据失败: {str(e)}")

    def log_operation(self, operation):
        """记录用户操作到日志文件"""
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        log_entry = f'[{timestamp}] {operation}\n'
        try:
            # 确保日志目录存在
            log_dir = os.path.dirname(self.log_file_path)
            if not os.path.exists(log_dir):
                os.makedirs(log_dir)
            
            # 尝试写入日志
            with open(self.log_file_path, 'a', encoding='utf-8') as f:
                f.write(log_entry)
        except Exception as e:
            # 记录错误到控制台，但不中断程序
            print(f'写入日志失败: {str(e)}')
            # 可以选择将错误记录到备用日志或系统事件日志
            try:
                error_log = os.path.join(os.path.dirname(self.log_file_path), 'error_log.txt')
                with open(error_log, 'a', encoding='utf-8') as f:
                    f.write(f'[{timestamp}] 写入日志失败: {str(e)}\n')
            except:
                pass

    def open_log_file(self):
        """打开日志文件"""
        try:
            if os.path.exists(self.log_file_path):
                os.startfile(self.log_file_path)
            else:
                messagebox.showinfo("提示", "日志文件不存在")
        except Exception as e:
            messagebox.showerror("错误", f"打开日志文件失败: {str(e)}")

    def show_about(self):
        """显示关于信息"""
        messagebox.showinfo("关于", "出口销售合同管理系统\n版本: 1.0\n开发: Python Tkinter")

    def select_customer(self, var, add_window=None):
        """选择客户"""
        # 创建选择客户对话框
        select_window = tk.Toplevel(self.root)
        select_window.title("选择客户")
        select_window.geometry("600x400")
        select_window.transient(self.root)
        select_window.grab_set()

        # 创建搜索框
        search_frame = tk.Frame(select_window)
        search_frame.pack(fill=tk.X, padx=10, pady=10)

        tk.Label(search_frame, text="搜索:", font=('SimHei', 10)).pack(side=tk.LEFT, padx=5)
        search_var = tk.StringVar()
        search_entry = tk.Entry(search_frame, textvariable=search_var, width=30)
        search_entry.pack(side=tk.LEFT, padx=5)

        def search():
            # 清空表格
            for item in tree.get_children():
                tree.delete(item)
            
            # 搜索客户
            search_term = search_var.get().strip()
            if search_term:
                df = self.data_manager.search_customers(search_term, "客户名称")
            else:
                df = self.data_manager.get_all_customers()
            
            # 填充表格
            for index, row in df.iterrows():
                tree.insert('', tk.END, values=(row['客户ID'], row['客户名称'], row['联系人'], row['联系电话']))

        tk.Button(search_frame, text="搜索", command=search).pack(side=tk.LEFT, padx=5)

        # 创建客户表格
        columns = ("客户ID", "客户名称", "联系人", "联系电话")
        tree = ttk.Treeview(select_window, columns=columns, show="headings")

        # 设置列宽和标题
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=100, anchor=tk.CENTER)

        tree.column("客户ID", width=80)
        tree.column("客户名称", width=150)
        tree.column("联系人", width=100)
        tree.column("联系电话", width=120)

        # 添加滚动条
        scrollbar = ttk.Scrollbar(select_window, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 加载客户数据
        df = self.data_manager.get_all_customers()
        for index, row in df.iterrows():
            tree.insert('', tk.END, values=(row['客户ID'], row['客户名称'], row['联系人'], row['联系电话']))

        # 选择按钮
        def confirm_selection():
            selected_item = tree.selection()
            if selected_item:
                item = tree.item(selected_item[0])
                var.set(item['values'][1])  # 设置客户名称
                select_window.destroy()
                # 确保焦点回到新建合同窗口
                if add_window:
                    add_window.lift()
                    add_window.focus_set()

        btn_frame = tk.Frame(select_window)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)

        tk.Button(btn_frame, text="确定", command=confirm_selection).pack(side=tk.LEFT, padx=20)
        tk.Button(btn_frame, text="取消", command=select_window.destroy).pack(side=tk.RIGHT, padx=20)

    def select_salesman(self, var, add_window=None):
        """选择业务员"""
        # 创建选择业务员对话框
        select_window = tk.Toplevel(self.root)
        select_window.title("选择业务员")
        select_window.geometry("600x400")
        select_window.transient(self.root)
        select_window.grab_set()

        # 创建搜索框
        search_frame = tk.Frame(select_window)
        search_frame.pack(fill=tk.X, padx=10, pady=10)

        tk.Label(search_frame, text="搜索:", font=('SimHei', 10)).pack(side=tk.LEFT, padx=5)
        search_var = tk.StringVar()
        search_entry = tk.Entry(search_frame, textvariable=search_var, width=30)
        search_entry.pack(side=tk.LEFT, padx=5)

        def search():
            # 清空表格
            for item in tree.get_children():
                tree.delete(item)
            
            # 搜索业务员
            search_term = search_var.get().strip()
            if search_term:
                df = self.data_manager.search_salesmen(search_term, "姓名")
            else:
                df = self.data_manager.get_all_salesmen()
            
            # 填充表格
            for index, row in df.iterrows():
                tree.insert('', tk.END, values=(row['业务员ID'], row['姓名'], row['联系电话'], row['所属部门']))

        tk.Button(search_frame, text="搜索", command=search).pack(side=tk.LEFT, padx=5)

        # 创建业务员表格
        columns = ("业务员ID", "姓名", "联系电话", "所属部门")
        tree = ttk.Treeview(select_window, columns=columns, show="headings")

        # 设置列宽和标题
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=100, anchor=tk.CENTER)

        tree.column("业务员ID", width=80)
        tree.column("姓名", width=100)
        tree.column("联系电话", width=120)
        tree.column("所属部门", width=150)

        # 添加滚动条
        scrollbar = ttk.Scrollbar(select_window, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 加载业务员数据
        df = self.data_manager.get_all_salesmen()
        for index, row in df.iterrows():
            tree.insert('', tk.END, values=(row['业务员ID'], row['姓名'], row['联系电话'], row['所属部门']))

        # 选择按钮
        def confirm_selection():
            selected_item = tree.selection()
            if selected_item:
                item = tree.item(selected_item[0])
                var.set(item['values'][1])  # 设置业务员姓名
                select_window.destroy()
                # 确保焦点回到新建合同窗口
                if add_window:
                    add_window.lift()
                    add_window.focus_set()

        btn_frame = tk.Frame(select_window)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)

        tk.Button(btn_frame, text="确定", command=confirm_selection).pack(side=tk.LEFT, padx=20)
        tk.Button(btn_frame, text="取消", command=select_window.destroy).pack(side=tk.RIGHT, padx=20)

    def open_data_folder(self):
        """打开数据库文件夹"""
        try:
            os.startfile(self.data_dir)
        except Exception as e:
            messagebox.showerror("错误", f"无法打开数据库文件夹: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    # 设置中文字体
    root.option_add("*Font", "SimHei 10")
    app = ContractManagementSystem(root)
    root.mainloop()