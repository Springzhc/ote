import pandas as pd
import os
from datetime import datetime
import openpyxl

class DataManager:
    def __init__(self, data_dir):
        self.data_dir = data_dir
        
        # 初始化文件路径
        self.contract_file = os.path.join(data_dir, "contracts.xlsx")
        self.payment_file = os.path.join(data_dir, "payments.xlsx")
        self.customer_file = os.path.join(data_dir, "customers.xlsx")
        self.salesman_file = os.path.join(data_dir, "salesmen.xlsx")
        
        # 确保数据文件存在
        self._ensure_files_exist()

    def _ensure_files_exist(self):
        """确保所有数据文件存在，如果不存在则创建"""
        # 合同表
        if not os.path.exists(self.contract_file):
            df = pd.DataFrame({
                '合同编号': [],
                '客户名称': [],
                '业务员': [],
                '签订日期': [],
                '单价': [],
                '数量': [],
                '合同金额': [],
                '付款方式': [],
                '交货日期': [],
                '状态': [],
                '备注': []
            })
            df.to_excel(self.contract_file, index=False, engine='openpyxl')

        # 收款表
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

    # ------------------------------ 合同管理 ------------------------------
    def get_all_contracts(self):
        """获取所有合同数据"""
        try:
            return pd.read_excel(self.contract_file, engine='openpyxl')
        except Exception as e:
            print(f"读取合同数据失败: {str(e)}")
            return pd.DataFrame()

    def add_contract(self, contract_data):
        """添加新合同"""
        try:
            # 读取现有数据
            df = self.get_all_contracts()
            
            # 生成合同编号
            if '合同编号' not in contract_data or not contract_data['合同编号']:
                contract_data['合同编号'] = self._generate_contract_id()
            
            # 添加新数据
            new_row = pd.DataFrame([contract_data])
            df = pd.concat([df, new_row], ignore_index=True)
            
            # 保存到文件
            df.to_excel(self.contract_file, index=False, engine='openpyxl')
            return True
        except Exception as e:
            print(f"添加合同失败: {str(e)}")
            return False

    def update_contract(self, contract_id, update_data):
        """更新合同信息"""
        try:
            # 读取现有数据
            df = self.get_all_contracts()
            
            # 确保合同编号是字符串类型
            contract_id = str(contract_id)
            df['合同编号'] = df['合同编号'].astype(str)
            
            # 查找并更新合同
            index = df[df['合同编号'] == contract_id].index
            if not index.empty:
                # 检查更新数据中的列是否存在
                invalid_columns = [key for key in update_data.keys() if key not in df.columns]
                if invalid_columns:
                    print(f"更新合同失败: 无效的列名 {invalid_columns}")
                    return False
                
                for key, value in update_data.items():
                    # 尝试转换为适当的数据类型
                    try:
                        # 如果列是数字类型，尝试转换值为数字
                        if pd.api.types.is_numeric_dtype(df[key]):
                            if '.' in str(value):
                                df.at[index[0], key] = float(value)
                            else:
                                df.at[index[0], key] = int(value)
                        else:
                            df.at[index[0], key] = value
                    except ValueError:
                        # 如果转换失败，保留原始值
                        df.at[index[0], key] = value
                
                # 保存到文件
                df.to_excel(self.contract_file, index=False, engine='openpyxl')
                return True
            else:
                print(f"更新合同失败: 未找到合同编号 {contract_id}")
                return False
        except Exception as e:
            print(f"更新合同失败: {str(e)}")
            return False

    def delete_contract(self, contract_id):
        """删除合同"""
        try:
            # 读取现有数据
            df = self.get_all_contracts()
            
            # 确保合同编号是字符串类型
            contract_id = str(contract_id)
            df['合同编号'] = df['合同编号'].astype(str)
            
            # 查找并删除合同
            initial_len = len(df)
            df = df[df['合同编号'] != contract_id]
            
            # 保存到文件
            df.to_excel(self.contract_file, index=False, engine='openpyxl')
            return len(df) < initial_len
        except Exception as e:
            print(f"删除合同失败: {str(e)}")
            return False

    def search_contracts(self, search_term, search_type):
        """搜索合同"""
        try:
            df = self.get_all_contracts()
            if search_type in df.columns:
                # 对于字符串类型的列进行模糊搜索
                if df[search_type].dtype == 'object':
                    return df[df[search_type].str.contains(search_term, na=False)]
                # 对于其他类型的列进行精确匹配
                else:
                    return df[df[search_type] == search_term]
            return pd.DataFrame()
        except Exception as e:
            print(f"搜索合同失败: {str(e)}")
            return pd.DataFrame()

    def _generate_contract_id(self):
        """生成合同编号"""
        today = datetime.now().strftime('%Y%m%d')
        df = self.get_all_contracts()
        
        # 确保合同编号是字符串类型
        df['合同编号'] = df['合同编号'].astype(str)
        # 筛选今天的合同
        today_contracts = df[df['合同编号'].str.startswith(today, na=False)]
        
        # 生成序号
        if len(today_contracts) == 0:
            seq = '001'
        else:
            # 提取序号并找到最大值
            max_seq = max([int(id[-3:]) for id in today_contracts['合同编号'] if id[-3:].isdigit()])
            seq = f'{max_seq + 1:03d}'
        
        return f'{today}{seq}'

    # ------------------------------ 收款管理 ------------------------------
    def get_all_payments(self):
        """获取所有收款数据"""
        try:
            return pd.read_excel(self.payment_file, engine='openpyxl')
        except Exception as e:
            print(f"读取收款数据失败: {str(e)}")
            return pd.DataFrame()

    def add_payment(self, payment_data):
        """添加新收款"""
        try:
            # 读取现有数据
            df = self.get_all_payments()
            
            # 生成收款ID
            if '收款ID' not in payment_data or not payment_data['收款ID']:
                payment_data['收款ID'] = self._generate_payment_id()
            
            # 添加新数据
            new_row = pd.DataFrame([payment_data])
            df = pd.concat([df, new_row], ignore_index=True)
            
            # 保存到文件
            df.to_excel(self.payment_file, index=False, engine='openpyxl')
            return True
        except Exception as e:
            print(f"添加收款失败: {str(e)}")
            return False

    def update_payment(self, payment_id, update_data):
        """更新收款信息"""
        try:
            # 读取现有数据
            df = self.get_all_payments()
            
            # 查找并更新收款
            index = df[df['收款ID'] == payment_id].index
            if not index.empty:
                for key, value in update_data.items():
                    if key in df.columns:
                        df.at[index[0], key] = value
                
                # 保存到文件
                df.to_excel(self.payment_file, index=False, engine='openpyxl')
                return True
            return False
        except Exception as e:
            print(f"更新收款失败: {str(e)}")
            return False

    def delete_payment(self, payment_id):
        """删除收款"""
        try:
            # 读取现有数据
            df = self.get_all_payments()
            
            # 查找并删除收款
            initial_len = len(df)
            df = df[df['收款ID'] != payment_id]
            
            # 保存到文件
            df.to_excel(self.payment_file, index=False, engine='openpyxl')
            return len(df) < initial_len
        except Exception as e:
            print(f"删除收款失败: {str(e)}")
            return False

    def search_payments(self, search_term, search_type):
        """搜索收款"""
        try:
            df = self.get_all_payments()
            if search_type in df.columns:
                # 对于字符串类型的列进行模糊搜索
                if df[search_type].dtype == 'object':
                    return df[df[search_type].str.contains(search_term, na=False)]
                # 对于其他类型的列进行精确匹配
                else:
                    return df[df[search_type] == search_term]
            return pd.DataFrame()
        except Exception as e:
            print(f"搜索收款失败: {str(e)}")
            return pd.DataFrame()

    def _generate_payment_id(self):
        """生成收款ID"""
        today = datetime.now().strftime('%Y%m%d')
        df = self.get_all_payments()
        
        # 筛选今天的收款
        today_payments = df[df['收款ID'].str.startswith(today, na=False)]
        
        # 生成序号
        if len(today_payments) == 0:
            seq = '001'
        else:
            # 提取序号并找到最大值
            max_seq = max([int(id[-3:]) for id in today_payments['收款ID'] if id[-3:].isdigit()])
            seq = f'{max_seq + 1:03d}'
        
        return f'P{today}{seq}'

    # ------------------------------ 客户管理 ------------------------------
    def get_all_customers(self):
        """获取所有客户数据"""
        try:
            return pd.read_excel(self.customer_file, engine='openpyxl')
        except Exception as e:
            print(f"读取客户数据失败: {str(e)}")
            return pd.DataFrame()

    def add_customer(self, customer_data):
        """添加新客户"""
        try:
            # 读取现有数据
            df = self.get_all_customers()
            
            # 生成客户ID
            if '客户ID' not in customer_data or not customer_data['客户ID']:
                customer_data['客户ID'] = self._generate_customer_id()
            
            # 添加新数据
            new_row = pd.DataFrame([customer_data])
            df = pd.concat([df, new_row], ignore_index=True)
            
            # 保存到文件
            df.to_excel(self.customer_file, index=False, engine='openpyxl')
            return True
        except Exception as e:
            print(f"添加客户失败: {str(e)}")
            return False

    def update_customer(self, customer_id, update_data):
        """更新客户信息"""
        try:
            # 读取现有数据
            df = self.get_all_customers()
            
            # 查找并更新客户
            index = df[df['客户ID'] == customer_id].index
            if not index.empty:
                for key, value in update_data.items():
                    if key in df.columns:
                        df.at[index[0], key] = value
                
                # 保存到文件
                df.to_excel(self.customer_file, index=False, engine='openpyxl')
                return True
            return False
        except Exception as e:
            print(f"更新客户失败: {str(e)}")
            return False

    def delete_customer(self, customer_id):
        """删除客户"""
        try:
            # 读取现有数据
            df = self.get_all_customers()
            
            # 查找并删除客户
            initial_len = len(df)
            df = df[df['客户ID'] != customer_id]
            
            # 保存到文件
            df.to_excel(self.customer_file, index=False, engine='openpyxl')
            return len(df) < initial_len
        except Exception as e:
            print(f"删除客户失败: {str(e)}")
            return False

    def search_customers(self, search_term, search_type):
        """搜索客户"""
        try:
            df = self.get_all_customers()
            if search_type in df.columns:
                # 对于字符串类型的列进行模糊搜索
                if df[search_type].dtype == 'object':
                    return df[df[search_type].str.contains(search_term, na=False)]
                # 对于其他类型的列进行精确匹配
                else:
                    return df[df[search_type] == search_term]
            return pd.DataFrame()
        except Exception as e:
            print(f"搜索客户失败: {str(e)}")
            return pd.DataFrame()

    def _generate_customer_id(self):
        """生成客户ID"""
        df = self.get_all_customers()
        
        # 查找最大ID
        if df.empty or '客户ID' not in df.columns or df['客户ID'].isna().all():
            return 'C001'
        
        # 提取数字部分并找到最大值
        numeric_ids = []
        for id in df['客户ID']:
            if isinstance(id, str) and id.startswith('C') and id[1:].isdigit():
                numeric_ids.append(int(id[1:]))
        
        if not numeric_ids:
            return 'C001'
        
        max_id = max(numeric_ids)
        return f'C{max_id + 1:03d}'

    # ------------------------------ 业务员管理 ------------------------------
    def get_all_salesmen(self):
        """获取所有业务员数据"""
        try:
            return pd.read_excel(self.salesman_file, engine='openpyxl')
        except Exception as e:
            print(f"读取业务员数据失败: {str(e)}")
            return pd.DataFrame()

    def add_salesman(self, salesman_data):
        """添加新业务员"""
        try:
            # 读取现有数据
            df = self.get_all_salesmen()
            
            # 生成业务员ID
            if '业务员ID' not in salesman_data or not salesman_data['业务员ID']:
                salesman_data['业务员ID'] = self._generate_salesman_id()
            
            # 添加新数据
            new_row = pd.DataFrame([salesman_data])
            df = pd.concat([df, new_row], ignore_index=True)
            
            # 保存到文件
            df.to_excel(self.salesman_file, index=False, engine='openpyxl')
            return True
        except Exception as e:
            print(f"添加业务员失败: {str(e)}")
            return False

    def update_salesman(self, salesman_id, update_data):
        """更新业务员信息"""
        try:
            # 读取现有数据
            df = self.get_all_salesmen()
            
            # 查找并更新业务员
            index = df[df['业务员ID'] == salesman_id].index
            if not index.empty:
                for key, value in update_data.items():
                    if key in df.columns:
                        df.at[index[0], key] = value
                
                # 保存到文件
                df.to_excel(self.salesman_file, index=False, engine='openpyxl')
                return True
            return False
        except Exception as e:
            print(f"更新业务员失败: {str(e)}")
            return False

    def delete_salesman(self, salesman_id):
        """删除业务员"""
        try:
            # 读取现有数据
            df = self.get_all_salesmen()
            
            # 查找并删除业务员
            initial_len = len(df)
            df = df[df['业务员ID'] != salesman_id]
            
            # 保存到文件
            df.to_excel(self.salesman_file, index=False, engine='openpyxl')
            return len(df) < initial_len
        except Exception as e:
            print(f"删除业务员失败: {str(e)}")
            return False

    def search_salesmen(self, search_term, search_type):
        """搜索业务员"""
        try:
            df = self.get_all_salesmen()
            if search_type in df.columns:
                # 对于字符串类型的列进行模糊搜索
                if df[search_type].dtype == 'object':
                    return df[df[search_type].str.contains(search_term, na=False)]
                # 对于其他类型的列进行精确匹配
                else:
                    return df[df[search_type] == search_term]
            return pd.DataFrame()
        except Exception as e:
            print(f"搜索业务员失败: {str(e)}")
            return pd.DataFrame()

    def _generate_salesman_id(self):
        """生成业务员ID"""
        df = self.get_all_salesmen()
        
        # 查找最大ID
        if df.empty or '业务员ID' not in df.columns or df['业务员ID'].isna().all():
            return 'S001'
        
        # 提取数字部分并找到最大值
        numeric_ids = []
        for id in df['业务员ID']:
            if isinstance(id, str) and id.startswith('S') and id[1:].isdigit():
                numeric_ids.append(int(id[1:]))
        
        if not numeric_ids:
            return 'S001'
        
        max_id = max(numeric_ids)
        return f'S{max_id + 1:03d}'