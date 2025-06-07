import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import datetime as dt
from dateutil.relativedelta import relativedelta  # 导入relativedelta用于月份计算
from calculator import FinanceCostCalculator
from database import RecordManager
from datetime import datetime
import xlsxwriter  # 添加xlsxwriter导入
import uuid  # 添加uuid导入用于生成唯一标识

class DateEntry(ttk.Frame):
    """自定义日期输入组件，替代tkcalendar的DateEntry"""
    def __init__(self, parent, width=None, **kwargs):
        super().__init__(parent)
        
        self.date_var = tk.StringVar()
        check_date()
        # 提取并设置默认日期
        if 'date_pattern' in kwargs:
            self.date_pattern = kwargs['date_pattern']
        else:
            self.date_pattern = 'yyyy-mm-dd'
            
        # 创建日期输入框
        self.entry = ttk.Entry(self, width=width, textvariable=self.date_var)
        self.entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 初始值为当前日期
        current_date = dt.datetime.now().strftime('%Y-%m-%d')
        self.date_var.set(current_date)
        
        # 状态设置
        if 'state' in kwargs and kwargs['state'] == 'readonly':
            self.entry.configure(state='readonly')
    
    def get_date(self):
        """获取日期对象"""
        check_date()
        date_str = self.date_var.get()
        try:
            return dt.datetime.strptime(date_str, '%Y-%m-%d').date()
        except ValueError:
            return dt.datetime.now().date()
    
    def set_date(self, date):
        """设置日期值"""
        check_date()
        if isinstance(date, dt.date):
            date_str = date.strftime('%Y-%m-%d')
            self.date_var.set(date_str)
        elif isinstance(date, str):
            self.date_var.set(date)

class FinanceCostApp:
    def __init__(self, root):
        check_date()
        self.root = root
        self.root.title("企业融资成本计算工具  V1.4 @昆仑银行克拉玛依分行计划财务部 陈世恩")
        self.root.geometry("1100x900")
        
        # 使用自动模式，根据首次还款日智能选择计算方法
        self.calculator = FinanceCostCalculator(calculation_mode="auto")
        self.record_manager = RecordManager("finance_records.db")
        
        self.fees = []  # 存储费用项
        self.current_record_id = None  # 当前选中的记录ID
        
        # 定义新增字段的选项
        self.loan_channel_options = ["", "自己向银行申请", "银行自主营销", "助贷机构推荐", 
                                    "互联网平台推荐", "其他"]
        self.customer_type_options = ["", "大型企业", "中型企业", "小型企业", "微型企业", 
                                     "个体工商户", "小微企业主"]
        self.company_nature_options = ["", "国有控股", "非国有控股"]
        self.guarantee_type_options = ["", "信用", "担保", "抵质押", "其他"]
        self.loan_type_options = ["", "首贷", "无还本续贷", "借新换旧", "其他"]
        self.application_method_options = ["", "线上", "线下"]
        self.is_subsidized_options = ["否", "是"]
        
        # 存储自定义输入值
        self.custom_inputs = {}
        
        self.create_widgets()
        self.load_records()

    def create_widgets(self):
        # 创建主框架
        check_date()
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建左侧数据输入区域
        input_frame = ttk.LabelFrame(main_frame, text="贷款信息", padding="10")
        input_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 创建参数输入区域 - 基本信息
        self.create_input_area(input_frame)
        
        # 创建参数输入区域 - 附加信息
        self.create_additional_info_area(input_frame)
        
        # 创建费用输入区域
        self.create_fee_area(input_frame)
        
        # 创建右侧结果和记录显示区域
        result_frame = ttk.LabelFrame(main_frame, text="计算结果与历史记录", padding="10")
        result_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 创建结果显示区域
        self.create_result_area(result_frame)
        
        # 创建记录表格
        self.create_records_table(result_frame)
        
        # 创建底部按钮区域
        button_frame = ttk.Frame(self.root, padding="10")
        button_frame.pack(fill=tk.X)
        
        # 第一行按钮
        button_row1 = ttk.Frame(button_frame)
        button_row1.pack(fill=tk.X, pady=5)
        
        # 新建记录按钮
        ttk.Button(button_row1, text="新建记录", command=self.new_record).pack(side=tk.LEFT, padx=5)
        
        # 计算按钮
        ttk.Button(button_row1, text="计算融资成本", command=self.calculate).pack(side=tk.LEFT, padx=5)
        
        # 保存记录按钮
        ttk.Button(button_row1, text="保存记录", command=self.save_record).pack(side=tk.LEFT, padx=5)
        
        # 删除记录按钮
        ttk.Button(button_row1, text="删除选中记录", command=self.delete_record).pack(side=tk.LEFT, padx=5)
        
        # 第二行按钮 - 导入导出功能
        button_row2 = ttk.Frame(button_frame)
        button_row2.pack(fill=tk.X, pady=5)
        
        # 导入记录按钮
        ttk.Button(button_row2, text="导入记录", command=self.import_records).pack(side=tk.LEFT, padx=5)
        
        # 导出记录按钮（原功能）
        ttk.Button(button_row2, text="导出记录", command=self.export_records).pack(side=tk.LEFT, padx=5)
        
        # 导出明白纸按钮
        ttk.Button(button_row2, text="导出明白纸", command=self.export_mingbaizhi).pack(side=tk.LEFT, padx=5)
        
        # 导出明细台账按钮
        ttk.Button(button_row2, text="导出明细台账", command=self.export_detail_ledger).pack(side=tk.LEFT, padx=5)
        
        # 导出汇总表按钮
        ttk.Button(button_row2, text="导出汇总表", command=self.export_summary_table).pack(side=tk.LEFT, padx=5)
    
    def create_input_area(self, parent):
        check_date()
        input_grid = ttk.Frame(parent)
        input_grid.pack(fill=tk.BOTH, expand=False, pady=5)
        
        # 企业名称
        ttk.Label(input_grid, text="企业名称:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.company_name = ttk.Entry(input_grid, width=30)
        self.company_name.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        # 绑定企业名称变更事件
        self.company_name.bind("<KeyRelease>", self.company_name_changed)
        
        # 贷款本金
        ttk.Label(input_grid, text="贷款本金(万元):").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.loan_amount = ttk.Entry(input_grid, width=30)
        self.loan_amount.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        
        # 还款方式
        ttk.Label(input_grid, text="还款方式:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        repayment_frame = ttk.Frame(input_grid)
        repayment_frame.grid(row=2, column=1, sticky=tk.W, padx=5, pady=5)
        
        self.repayment_method = ttk.Combobox(repayment_frame, values=["等额本金", "等额本息", "一次性还本", "其他"], width=27, state="readonly")
        self.repayment_method.current(0)
        self.repayment_method.pack(side=tk.LEFT)
        self.repayment_method.bind("<<ComboboxSelected>>", lambda e: self.on_combobox_change(e, "repayment_method"))
        
        # 自定义输入框（初始隐藏）
        self.repayment_method_custom = ttk.Entry(repayment_frame, width=20)
        self.custom_inputs["repayment_method"] = self.repayment_method_custom
        
        # 贷款期限
        ttk.Label(input_grid, text="贷款期限(月):").grid(row=3, column=0, sticky=tk.W, padx=5, pady=5)
        self.loan_term = ttk.Entry(input_grid, width=30)
        self.loan_term.grid(row=3, column=1, sticky=tk.W, padx=5, pady=5)
        self.loan_term.bind("<KeyRelease>", self.update_end_date)
        
        # 付息频率
        ttk.Label(input_grid, text="付息频率:").grid(row=4, column=0, sticky=tk.W, padx=5, pady=5)
        self.interest_frequency = ttk.Combobox(input_grid, values=["日", "月", "季", "半年", "年"], width=27, state="readonly")
        self.interest_frequency.current(1)
        self.interest_frequency.grid(row=4, column=1, sticky=tk.W, padx=5, pady=5)
        
        # 贷款起始日
        ttk.Label(input_grid, text="贷款起始日:").grid(row=5, column=0, sticky=tk.W, padx=5, pady=5)
        self.start_date = DateEntry(input_grid, width=27, date_pattern="yyyy-mm-dd")
        self.start_date.grid(row=5, column=1, sticky=tk.W, padx=5, pady=5)
        self.start_date.entry.bind("<KeyRelease>", self.update_end_date)
        
        # 贷款到期日
        ttk.Label(input_grid, text="贷款到期日:").grid(row=6, column=0, sticky=tk.W, padx=5, pady=5)
        self.end_date = DateEntry(input_grid, width=27, date_pattern="yyyy-mm-dd")
        self.end_date.grid(row=6, column=1, sticky=tk.W, padx=5, pady=5)
        
        # 首次还款日
        ttk.Label(input_grid, text="首次还款日:").grid(row=7, column=0, sticky=tk.W, padx=5, pady=5)
        self.first_payment_date = DateEntry(input_grid, width=27, date_pattern="yyyy-mm-dd")
        self.first_payment_date.grid(row=7, column=1, sticky=tk.W, padx=5, pady=5)
        
        # 贷款年化率
        ttk.Label(input_grid, text="贷款年化率(%):").grid(row=8, column=0, sticky=tk.W, padx=5, pady=5)
        self.interest_rate = ttk.Entry(input_grid, width=30)
        self.interest_rate.grid(row=8, column=1, sticky=tk.W, padx=5, pady=5)
    
    def create_additional_info_area(self, parent):
        """创建附加信息输入区域"""
        additional_frame = ttk.LabelFrame(parent, text="附加信息", padding="10")
        additional_frame.pack(fill=tk.BOTH, expand=False, pady=10)
        
        # 创建两列布局
        left_frame = ttk.Frame(additional_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        right_frame = ttk.Frame(additional_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # 左侧字段
        # 获取贷款的渠道
        ttk.Label(left_frame, text="获取贷款渠道:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        loan_channel_frame = ttk.Frame(left_frame)
        loan_channel_frame.grid(row=0, column=1, columnspan=2, sticky=tk.W, padx=5, pady=5)
        
        self.loan_channel = ttk.Combobox(loan_channel_frame, values=self.loan_channel_options, width=10, state="readonly")
        self.loan_channel.pack(side=tk.LEFT)
        self.loan_channel.bind("<<ComboboxSelected>>", lambda e: self.on_combobox_change(e, "loan_channel"))
        
        self.loan_channel_custom = ttk.Entry(loan_channel_frame, width=15)
        self.custom_inputs["loan_channel"] = self.loan_channel_custom
        
        # 客户类型
        ttk.Label(left_frame, text="客户类型:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.customer_type = ttk.Combobox(left_frame, values=self.customer_type_options, width=10, state="readonly")
        self.customer_type.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        
        # 企业性质
        ttk.Label(left_frame, text="企业性质:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.company_nature = ttk.Combobox(left_frame, values=self.company_nature_options, width=10, state="readonly")
        self.company_nature.grid(row=2, column=1, sticky=tk.W, padx=5, pady=5)
        
        # 右侧字段
        # 担保方式
        ttk.Label(right_frame, text="担保方式:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        guarantee_frame = ttk.Frame(right_frame)
        guarantee_frame.grid(row=0, column=1, columnspan=2, sticky=tk.W, padx=5, pady=5)
        
        self.guarantee_type = ttk.Combobox(guarantee_frame, values=self.guarantee_type_options, width=10, state="readonly")
        self.guarantee_type.pack(side=tk.LEFT)
        self.guarantee_type.bind("<<ComboboxSelected>>", lambda e: self.on_combobox_change(e, "guarantee_type"))
        
        self.guarantee_type_custom = ttk.Entry(guarantee_frame, width=15)
        self.custom_inputs["guarantee_type"] = self.guarantee_type_custom
        
        # 贷款方式
        ttk.Label(right_frame, text="贷款方式:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        loan_type_frame = ttk.Frame(right_frame)
        loan_type_frame.grid(row=1, column=1, columnspan=2, sticky=tk.W, padx=5, pady=5)
        
        self.loan_type = ttk.Combobox(loan_type_frame, values=self.loan_type_options, width=10, state="readonly")
        self.loan_type.pack(side=tk.LEFT)
        self.loan_type.bind("<<ComboboxSelected>>", lambda e: self.on_combobox_change(e, "loan_type"))
        
        self.loan_type_custom = ttk.Entry(loan_type_frame, width=15)
        self.custom_inputs["loan_type"] = self.loan_type_custom
        
        # 贷款申请方式
        ttk.Label(right_frame, text="申请方式:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.application_method = ttk.Combobox(right_frame, values=self.application_method_options, width=10, state="readonly")
        self.application_method.grid(row=2, column=1, sticky=tk.W, padx=5, pady=5)
        
        # 是否财政贴息
        ttk.Label(right_frame, text="是否财政贴息:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=5)
        self.is_subsidized = ttk.Combobox(right_frame, values=self.is_subsidized_options, width=10, state="readonly")
        self.is_subsidized.current(0)  # 默认选择"否"
        self.is_subsidized.grid(row=3, column=1, sticky=tk.W, padx=5, pady=5)
    
    def create_fee_area(self, parent):
        check_date()
        fee_frame = ttk.LabelFrame(parent, text="费用项", padding="10")
        fee_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        fee_input_frame = ttk.Frame(fee_frame)
        fee_input_frame.pack(fill=tk.X, pady=5)
        
        # 费用名称
        ttk.Label(fee_input_frame, text="费用名称:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.fee_name = ttk.Entry(fee_input_frame, width=10)
        self.fee_name.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        # 费用金额
        ttk.Label(fee_input_frame, text="费用金额(元):").grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)
        self.fee_amount = ttk.Entry(fee_input_frame, width=10)
        self.fee_amount.grid(row=0, column=3, sticky=tk.W, padx=5, pady=5)
        
        # 费用支付频率
        ttk.Label(fee_input_frame, text="支付频率:").grid(row=0, column=4, sticky=tk.W, padx=5, pady=5)
        self.fee_frequency = ttk.Combobox(fee_input_frame, values=["年", "季", "月", "期初一次性付费"], width=12, state="readonly")
        self.fee_frequency.current(3)
        self.fee_frequency.grid(row=0, column=5, sticky=tk.W, padx=5, pady=5)
        
        # 是否银行承担
        self.is_bank_bearing = tk.BooleanVar()
        self.bank_bearing_check = ttk.Checkbutton(fee_input_frame, text="银行承担", variable=self.is_bank_bearing)
        self.bank_bearing_check.grid(row=0, column=6, sticky=tk.W, padx=5, pady=5)
        
        # 添加费用按钮
        ttk.Button(fee_input_frame, text="添加费用", command=self.add_fee).grid(row=0, column=7, padx=5, pady=5)
        
        # 费用列表
        fee_list_frame = ttk.Frame(fee_frame)
        fee_list_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        columns = ("名称", "金额", "支付频率", "银行承担")
        self.fee_tree = ttk.Treeview(fee_list_frame, columns=columns, show="headings", height=5)
        
        for col in columns:
            self.fee_tree.heading(col, text=col)
            self.fee_tree.column(col, width=100)
        
        self.fee_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(fee_list_frame, orient=tk.VERTICAL, command=self.fee_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.fee_tree.configure(yscrollcommand=scrollbar.set)
        
        # 删除费用按钮
        ttk.Button(fee_frame, text="删除选中费用", command=self.delete_fee).pack(anchor=tk.E, pady=5)
    
    def create_result_area(self, parent):
        check_date()
        result_frame = ttk.Frame(parent)
        result_frame.pack(fill=tk.X, pady=5)
        
        # 综合融资成本
        ttk.Label(result_frame, text="综合融资成本(年化):").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.total_cost_var = tk.StringVar()
        ttk.Label(result_frame, textvariable=self.total_cost_var, font=("Arial", 10, "bold")).grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        # 费用明细区域
        fee_detail_frame = ttk.LabelFrame(parent, text="费用成本明细", padding="5")
        fee_detail_frame.pack(fill=tk.X, pady=5)
        
        columns = ("费用名称", "费用金额", "年化利率", "月费率", "期间总费率")
        self.detail_tree = ttk.Treeview(fee_detail_frame, columns=columns, show="headings", height=5)
        
        for col in columns:
            self.detail_tree.heading(col, text=col)
            self.detail_tree.column(col, width=100)
        
        self.detail_tree.pack(fill=tk.X, pady=5)
    
    def create_records_table(self, parent):
        check_date()
        records_frame = ttk.LabelFrame(parent, text="历史记录", padding="5")
        records_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        columns = ("ID", "企业名称", "贷款本金(万)", "还款方式", "贷款期限(月)", 
                   "付息频率", "贷款起始日", "贷款到期日", "首次还款日", 
                   "贷款年化率(%)", "综合融资成本(%)")
        
        self.records_tree = ttk.Treeview(records_frame, columns=columns, show="headings", height=10)
        
        for col in columns:
            self.records_tree.heading(col, text=col)
            self.records_tree.column(col, width=80)
        
        self.records_tree.column("ID", width=40)
        self.records_tree.column("企业名称", width=120)
        
        self.records_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 添加垂直滚动条
        y_scrollbar = ttk.Scrollbar(records_frame, orient=tk.VERTICAL, command=self.records_tree.yview)
        y_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 添加水平滚动条
        x_scrollbar = ttk.Scrollbar(records_frame, orient=tk.HORIZONTAL, command=self.records_tree.xview)
        x_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.records_tree.configure(yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set)
        
        # 绑定选择事件
        self.records_tree.bind("<<TreeviewSelect>>", self.on_record_select)
    
    def company_name_changed(self, event=None):
        """当企业名称变更时触发，如果是新输入(不是选中记录修改)则清空其他字段"""
        # 获取当前选中的记录
        selected = self.records_tree.selection()
        if not selected:  # 如果没有选中的记录，则认为是新输入
            # 如果当前记录ID不为空，说明是在修改已有记录但刚刚清除了选择
            # 此时也应该清空表单以方便新建
            if self.current_record_id is not None:
                self.current_record_id = None
                self.clear_form_except_company_name()
            # 如果当前记录ID为空而且企业名称变更，那么可能是用户在手动新建一条记录
            # 也应该清空其他字段
            else:
                self.clear_form_except_company_name()
    
    def clear_form_except_company_name(self):
        """清空除企业名称外的所有表单字段"""
        # 清空贷款信息
        self.loan_amount.delete(0, tk.END)
        self.repayment_method.current(0)
        self.loan_term.delete(0, tk.END)
        self.interest_frequency.current(1)
        
        # 重置日期为当前日期
        today = dt.datetime.now().date()
        self.start_date.set_date(today)
        self.end_date.set_date(today)
        self.first_payment_date.set_date(today)
        
        # 清空利率
        self.interest_rate.delete(0, tk.END)
        
        # 清空附加信息 - 设置为第一个选项（空选项）
        if hasattr(self, "loan_channel"):
            self.loan_channel.current(0)
        if hasattr(self, "customer_type"):
            self.customer_type.current(0)
        if hasattr(self, "company_nature"):
            self.company_nature.current(0)
        if hasattr(self, "guarantee_type"):
            self.guarantee_type.current(0)
        if hasattr(self, "loan_type"):
            self.loan_type.current(0)
        if hasattr(self, "application_method"):
            self.application_method.current(0)
        if hasattr(self, "is_subsidized"):
            self.is_subsidized.current(0)  # 默认选择"否"
        
        # 隐藏和清空所有自定义输入框
        for field_name, custom_input in self.custom_inputs.items():
            custom_input.pack_forget()
            custom_input.delete(0, tk.END)
        
        # 清空费用列表和结果
        self.clear_fees()
        self.total_cost_var.set("")
        
        # 清空当前记录ID
        self.current_record_id = None
    
    def new_record(self):
        """创建新记录，清空所有表单字段并取消当前记录的选择"""
        # 清空所有字段
        self.company_name.delete(0, tk.END)
        self.clear_form_except_company_name()
        
        # 取消当前记录的选择
        if self.records_tree.selection():
            self.records_tree.selection_remove(self.records_tree.selection())
        
        # 重置当前记录ID
        self.current_record_id = None
    
    def clear_fees(self):
        check_date()
        """清空费用列表"""
        # 清空费用树
        for item in self.fee_tree.get_children():
            self.fee_tree.delete(item)
        
        # 清空费用列表
        self.fees = []
        
        # 清空费用明细表
        for item in self.detail_tree.get_children():
            self.detail_tree.delete(item)
    
    def add_fee(self):
        check_date()
        name = self.fee_name.get()
        amount_str = self.fee_amount.get()
        frequency = self.fee_frequency.get()
        is_bank_bearing = 1 if self.is_bank_bearing.get() else 0
        
        if not name or not amount_str:
            messagebox.showerror("错误", "费用名称和金额不能为空")
            return
        
        try:
            amount = float(amount_str)
        except ValueError:
            messagebox.showerror("错误", "费用金额必须是数字")
            return
        
        self.fee_tree.insert("", tk.END, values=(name, amount, frequency, "是" if is_bank_bearing else "否"))
        
        # 添加到费用列表
        self.fees.append({
            "name": name, 
            "amount": amount, 
            "frequency": frequency,
            "is_bank_bearing": is_bank_bearing
        })
        
        # 清空输入框
        self.fee_name.delete(0, tk.END)
        self.fee_amount.delete(0, tk.END)
        self.is_bank_bearing.set(False)
    
    def delete_fee(self):
        check_date()
        selected = self.fee_tree.selection()
        if not selected:
            messagebox.showinfo("提示", "请选择要删除的费用项")
            return
        
        for item in selected:
            index = self.fee_tree.index(item)
            self.fee_tree.delete(item)
            if 0 <= index < len(self.fees):
                self.fees.pop(index)
    
    def update_end_date(self, event=None):
        check_date()
        try:
            start_date = self.start_date.get_date()
            term_str = self.loan_term.get()
            
            if term_str:
                term = int(term_str)
                # 使用relativedelta进行月份对日计算，而不是简单加天数
                end_date = start_date + relativedelta(months=term)
                self.end_date.set_date(end_date)
        except (ValueError, TypeError):
            pass
    
    def calculate(self):
        check_date()
        try:
            # 获取输入参数
            loan_amount = float(self.loan_amount.get()) * 10000  # 转换为元
            repayment_method = self.get_field_value("repayment_method")  # 使用get_field_value获取可能的自定义值
            loan_term = int(self.loan_term.get())
            interest_frequency = self.interest_frequency.get()
            interest_rate = float(self.interest_rate.get()) / 100  # 转换为小数
            start_date = self.start_date.get_date()
            end_date = self.end_date.get_date()
            first_payment_date = self.first_payment_date.get_date()
            
            # 计算综合融资成本
            total_cost, fee_details = self.calculator.calculate_finance_cost(
                loan_amount, repayment_method, loan_term, interest_frequency,
                interest_rate, start_date, end_date, first_payment_date, self.fees
            )
            
            # 显示结果
            self.total_cost_var.set(f"{total_cost:.4f}%")
            
            # 清空并更新费用明细表
            for item in self.detail_tree.get_children():
                self.detail_tree.delete(item)
            
            # 添加基础贷款利率
            self.detail_tree.insert("", tk.END, values=("基础贷款利率", "-", f"{interest_rate*100:.4f}%", "-", "-"))
            
            # 添加其他费用
            for detail in fee_details:
                # 如果费用由银行承担，显示特殊标记
                name = detail["name"]
                if detail.get("is_bank_bearing", False):
                    name += " [银行承担]"
                
                # 计算月费率
                monthly_rate = detail['annual_rate'] / 12
                    
                self.detail_tree.insert("", tk.END, values=(
                    name, 
                    f"{detail['amount']:.2f}", 
                    f"{detail['annual_rate']*100:.4f}%",
                    f"{monthly_rate*100:.4f}%",
                    f"{detail['period_rate']*100:.4f}%"
                ))
            
            return total_cost
            
        except ValueError as e:
            messagebox.showerror("输入错误", f"请检查输入参数: {str(e)}")
            return None
    
    def save_record(self):
        check_date()
        try:
            # 首先计算结果
            total_cost = self.calculate()
            if total_cost is None:
                return
            
            # 获取所有参数
            company_name = self.company_name.get()
            loan_amount = self.loan_amount.get()
            repayment_method = self.get_field_value("repayment_method")
            loan_term = self.loan_term.get()
            interest_frequency = self.interest_frequency.get()
            start_date = self.start_date.date_var.get()
            end_date = self.end_date.date_var.get()
            first_payment_date = self.first_payment_date.date_var.get()
            interest_rate = self.interest_rate.get()
            
            # 获取附加信息 - 使用get_field_value获取可能的自定义值
            loan_channel = self.get_field_value("loan_channel") if hasattr(self, "loan_channel") else ""
            customer_type = self.customer_type.get() if hasattr(self, "customer_type") else ""
            company_nature = self.company_nature.get() if hasattr(self, "company_nature") else ""
            guarantee_type = self.get_field_value("guarantee_type") if hasattr(self, "guarantee_type") else ""
            loan_type = self.get_field_value("loan_type") if hasattr(self, "loan_type") else ""
            application_method = self.application_method.get() if hasattr(self, "application_method") else ""
            is_subsidized = 1 if self.is_subsidized.get() == "是" else 0 if hasattr(self, "is_subsidized") else 0
            
            # 保存记录
            fees_data = [{"name": fee["name"], "amount": fee["amount"], "frequency": fee["frequency"], 
                         "is_bank_bearing": fee.get("is_bank_bearing", 0)} 
                        for fee in self.fees]
            
            if self.current_record_id:
                # 更新记录
                self.record_manager.update_record(
                    self.current_record_id, company_name, loan_amount, repayment_method, loan_term,
                    interest_frequency, start_date, end_date, first_payment_date,
                    interest_rate, total_cost, fees_data, loan_channel, customer_type,
                    company_nature, guarantee_type, loan_type, application_method, is_subsidized
                )
                messagebox.showinfo("成功", "记录已更新")
            else:
                # 添加新记录
                new_id = self.record_manager.add_record(
                    company_name, loan_amount, repayment_method, loan_term,
                    interest_frequency, start_date, end_date, first_payment_date,
                    interest_rate, total_cost, fees_data, loan_channel, customer_type,
                    company_nature, guarantee_type, loan_type, application_method, is_subsidized
                )
                self.current_record_id = new_id
                messagebox.showinfo("成功", "记录已保存")
            
            # 重新加载记录
            self.load_records()
            
            # 保存后清除选择，确保下次能正确新增
            if self.records_tree.selection():
                self.records_tree.selection_remove(self.records_tree.selection())
            self.current_record_id = None
            
        except Exception as e:
            messagebox.showerror("错误", f"保存记录时发生错误: {str(e)}")
    
    def delete_record(self):
        check_date()
        selected = self.records_tree.selection()
        if not selected:
            messagebox.showinfo("提示", "请选择要删除的记录")
            return
        
        if messagebox.askyesno("确认", "确定要删除所选记录吗?"):
            for item in selected:
                record_id = self.records_tree.item(item, "values")[0]
                self.record_manager.delete_record(record_id)
            
            messagebox.showinfo("成功", "记录已删除")
            self.load_records()
            
            # 删除后清空表单
            self.new_record()
    
    def export_records(self):
        check_date()
        try:
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
            )
            
            if not file_path:
                return
            
            # 获取记录数据
            records = self.record_manager.get_all_records()
            
            # 创建DataFrame
            if not records:
                messagebox.showinfo("提示", "没有记录可以导出")
                return
                
            # 准备数据
            export_data = []
            for record in records:
                # 使用IRR方法计算费用年化率
                fees_annual_rates = {}
                fees_period_rates = {}
                if record.get("fees"):
                    loan_amount = float(record["loan_amount"]) * 10000  # 转换为元
                    loan_term = int(record["loan_term"])
                    repayment_method = record["repayment_method"]
                    
                    for fee in record["fees"]:
                        # 如果费用由银行承担，年化率为0
                        if fee.get("is_bank_bearing", 0) == 1:
                            fee_annual_rate = 0
                            period_rate = 0
                        else:
                            # 计算该费用的年化利率
                            fee_annual_rate = self.calculator.calculate_fee_annual_rate_irr(
                                fee["amount"], 
                                fee["frequency"], 
                                loan_amount, 
                                loan_term, 
                                repayment_method,
                                dt.datetime.strptime(record["start_date"], '%Y-%m-%d').date(),
                                dt.datetime.strptime(record["first_payment_date"], '%Y-%m-%d').date(),
                                record["interest_frequency"]
                            )
                            
                            # 计算周期费率
                            period_rate = fee_annual_rate * loan_term / 12
                        
                        fees_annual_rates[fee["name"]] = fee_annual_rate * 100  # 转为百分比
                        fees_period_rates[fee["name"]] = period_rate * 100  # 转为百分比
                
                # 基本记录信息
                record_data = {
                    "ID": record["id"],
                    "企业名称": record["company_name"],
                    "贷款本金(万元)": record["loan_amount"],
                    "还款方式": record["repayment_method"],
                    "贷款期限(月)": record["loan_term"],
                    "付息频率": record["interest_frequency"],
                    "贷款起始日": record["start_date"],
                    "贷款到期日": record["end_date"],
                    "首次还款日": record["first_payment_date"],
                    "贷款年化率(%)": record["interest_rate"],
                    "综合融资成本(%)": f"{record['total_cost']:.4f}",
                    "获取贷款渠道": record.get("loan_channel", ""),
                    "客户类型": record.get("customer_type", ""),
                    "企业性质": record.get("company_nature", ""),
                    "担保方式": record.get("guarantee_type", ""),
                    "贷款方式": record.get("loan_type", ""),
                    "申请方式": record.get("application_method", ""),
                    "是否财政贴息": "是" if record.get("is_subsidized", 0) == 1 else "否"
                }
                
                # 费用详细信息
                fee_detail = []
                fee_rates = []
                fee_period_rates = []
                fee_bank_bearing = []
                
                for fee in record.get("fees", []):
                    # 基本费用信息
                    fee_str = f"{fee['name']}:{fee['amount']}元({fee['frequency']})"
                    if fee.get("is_bank_bearing", 0) == 1:
                        fee_str += "[银行承担]"
                    fee_detail.append(fee_str)
                    
                    # 费用年化率信息
                    annual_rate = fees_annual_rates.get(fee['name'], 0)
                    period_rate = fees_period_rates.get(fee['name'], 0)
                    
                    fee_rates.append(f"{fee['name']}:{annual_rate:.4f}%")
                    fee_period_rates.append(f"{fee['name']}:{period_rate:.4f}%")
                    fee_bank_bearing.append("是" if fee.get("is_bank_bearing", 0) == 1 else "否")
                
                record_data["费用项"] = "; ".join(fee_detail) if fee_detail else ""
                record_data["费用年化率"] = "; ".join(fee_rates) if fee_rates else ""
                record_data["费用周期率"] = "; ".join(fee_period_rates) if fee_period_rates else ""
                record_data["银行承担"] = "; ".join(fee_bank_bearing) if fee_bank_bearing else ""
                record_data["创建时间"] = record.get("create_time", "")
                
                export_data.append(record_data)
            
            # 创建DataFrame
            df = pd.DataFrame(export_data)
            
            # 定义列顺序
            column_order = [
                "ID", "企业名称", "贷款本金(万元)", "还款方式", "贷款期限(月)",
                "付息频率", "贷款起始日", "贷款到期日", "首次还款日",
                "贷款年化率(%)", "综合融资成本(%)", "获取贷款渠道", "客户类型",
                "企业性质", "担保方式", "贷款方式", "申请方式", "是否财政贴息",
                "费用项", "费用年化率", "费用周期率", "银行承担", "创建时间"
            ]
            
            # 只保留存在的列并按顺序排列
            existing_columns = [col for col in column_order if col in df.columns]
            df = df[existing_columns]
            
            # 导出到Excel (加入自动列宽设置)
            with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name='融资成本记录', index=False)
                
                # 获取xlsxwriter对象
                workbook = writer.book
                worksheet = writer.sheets['融资成本记录']
                
                # 设置列宽
                for i, col in enumerate(df.columns):
                    # 获取列中最长字符串的长度
                    max_len = max(df[col].astype(str).apply(len).max(), len(col)) + 2
                    worksheet.set_column(i, i, max_len)
            
            messagebox.showinfo("成功", f"记录已导出到 {file_path}")
            
        except Exception as e:
            messagebox.showerror("错误", f"导出记录时发生错误: {str(e)}")
    
    def on_record_select(self, event):
        check_date()
        selected = self.records_tree.selection()
        if not selected:
            return
        
        record_id = self.records_tree.item(selected[0], "values")[0]
        record = self.record_manager.get_record(record_id)
        
        if record:
            # 设置当前记录ID
            self.current_record_id = record_id
            
            # 填充基本表单
            self.company_name.delete(0, tk.END)
            self.company_name.insert(0, record["company_name"])
            
            self.loan_amount.delete(0, tk.END)
            self.loan_amount.insert(0, record["loan_amount"])
            
            # 处理还款方式（可能有自定义值）
            repayment = record["repayment_method"]
            if repayment in ["等额本金", "等额本息", "一次性还本"]:
                self.repayment_method.set(repayment)
                # 隐藏自定义输入框
                if "repayment_method" in self.custom_inputs:
                    self.custom_inputs["repayment_method"].pack_forget()
                    self.custom_inputs["repayment_method"].delete(0, tk.END)
            else:
                # 设置为"其他"并显示自定义值
                self.repayment_method.set("其他")
                if "repayment_method" in self.custom_inputs:
                    self.custom_inputs["repayment_method"].delete(0, tk.END)
                    self.custom_inputs["repayment_method"].insert(0, repayment)
                    self.custom_inputs["repayment_method"].pack(side=tk.LEFT, padx=5)
            
            self.loan_term.delete(0, tk.END)
            self.loan_term.insert(0, record["loan_term"])
            
            self.interest_frequency.set(record["interest_frequency"])
            
            # 设置日期
            self.start_date.set_date(record["start_date"])
            self.end_date.set_date(record["end_date"])
            self.first_payment_date.set_date(record["first_payment_date"])
            
            self.interest_rate.delete(0, tk.END)
            self.interest_rate.insert(0, record["interest_rate"])
            
            # 填充附加信息 - 处理自定义值
            # 获取贷款渠道
            if hasattr(self, "loan_channel"):
                channel = record.get("loan_channel", "")
                if channel in self.loan_channel_options:
                    self.loan_channel.set(channel)
                    # 隐藏自定义输入框
                    if "loan_channel" in self.custom_inputs:
                        self.custom_inputs["loan_channel"].pack_forget()
                        self.custom_inputs["loan_channel"].delete(0, tk.END)
                elif channel:
                    # 自定义值
                    self.loan_channel.set("其他")
                    if "loan_channel" in self.custom_inputs:
                        self.custom_inputs["loan_channel"].delete(0, tk.END)
                        self.custom_inputs["loan_channel"].insert(0, channel)
                        self.custom_inputs["loan_channel"].pack(side=tk.LEFT, padx=5)
                else:
                    self.loan_channel.current(0)
            
            # 客户类型（无自定义选项）
            if hasattr(self, "customer_type"):
                ctype = record.get("customer_type", "")
                if ctype and ctype in self.customer_type_options:
                    self.customer_type.set(ctype)
                else:
                    self.customer_type.current(0)
            
            # 企业性质（无自定义选项）
            if hasattr(self, "company_nature"):
                nature = record.get("company_nature", "")
                if nature and nature in self.company_nature_options:
                    self.company_nature.set(nature)
                else:
                    self.company_nature.current(0)
            
            # 担保方式
            if hasattr(self, "guarantee_type"):
                guarantee = record.get("guarantee_type", "")
                if guarantee in self.guarantee_type_options:
                    self.guarantee_type.set(guarantee)
                    # 隐藏自定义输入框
                    if "guarantee_type" in self.custom_inputs:
                        self.custom_inputs["guarantee_type"].pack_forget()
                        self.custom_inputs["guarantee_type"].delete(0, tk.END)
                elif guarantee:
                    # 自定义值
                    self.guarantee_type.set("其他")
                    if "guarantee_type" in self.custom_inputs:
                        self.custom_inputs["guarantee_type"].delete(0, tk.END)
                        self.custom_inputs["guarantee_type"].insert(0, guarantee)
                        self.custom_inputs["guarantee_type"].pack(side=tk.LEFT, padx=5)
                else:
                    self.guarantee_type.current(0)
            
            # 贷款方式
            if hasattr(self, "loan_type"):
                loan_type = record.get("loan_type", "")
                if loan_type in self.loan_type_options:
                    self.loan_type.set(loan_type)
                    # 隐藏自定义输入框
                    if "loan_type" in self.custom_inputs:
                        self.custom_inputs["loan_type"].pack_forget()
                        self.custom_inputs["loan_type"].delete(0, tk.END)
                elif loan_type:
                    # 自定义值
                    self.loan_type.set("其他")
                    if "loan_type" in self.custom_inputs:
                        self.custom_inputs["loan_type"].delete(0, tk.END)
                        self.custom_inputs["loan_type"].insert(0, loan_type)
                        self.custom_inputs["loan_type"].pack(side=tk.LEFT, padx=5)
                else:
                    self.loan_type.current(0)
            
            # 申请方式（无自定义选项）
            if hasattr(self, "application_method"):
                method = record.get("application_method", "")
                if method and method in self.application_method_options:
                    self.application_method.set(method)
                else:
                    self.application_method.current(0)
            
            # 是否财政贴息
            if hasattr(self, "is_subsidized"):
                self.is_subsidized.set("是" if record.get("is_subsidized", 0) == 1 else "否")
            
            # 加载费用项
            self.fees = []
            for item in self.fee_tree.get_children():
                self.fee_tree.delete(item)
            
            for fee in self.record_manager.get_fees(record_id):
                is_bank_bearing = fee.get("is_bank_bearing", 0)
                self.fees.append({
                    "name": fee["name"],
                    "amount": fee["amount"],
                    "frequency": fee["frequency"],
                    "is_bank_bearing": is_bank_bearing
                })
                self.fee_tree.insert("", tk.END, values=(fee["name"], fee["amount"], fee["frequency"], 
                                                        "是" if is_bank_bearing else "否"))
    
    def load_records(self):
        check_date()
        # 清空记录表
        for item in self.records_tree.get_children():
            self.records_tree.delete(item)
        
        # 获取并显示所有记录
        records = self.record_manager.get_all_records()
        
        for record in records:
            self.records_tree.insert("", tk.END, values=(
                record["id"],
                record["company_name"],
                record["loan_amount"],
                record["repayment_method"],
                record["loan_term"],
                record["interest_frequency"],
                record["start_date"],
                record["end_date"],
                record["first_payment_date"],
                record["interest_rate"],
                f"{record['total_cost']:.4f}"
            ))

    def on_combobox_change(self, event, field_name):
        """处理下拉框选择变更事件"""
        combobox = event.widget
        selected_value = combobox.get()
        
        if selected_value == "其他" and field_name in self.custom_inputs:
            # 显示自定义输入框
            self.custom_inputs[field_name].pack(side=tk.LEFT, padx=5)
            self.custom_inputs[field_name].focus()
        elif field_name in self.custom_inputs:
            # 隐藏自定义输入框
            self.custom_inputs[field_name].pack_forget()
            self.custom_inputs[field_name].delete(0, tk.END)
    
    def get_field_value(self, field_name):
        """获取字段值，如果选择了'其他'，则返回自定义输入值"""
        combobox = getattr(self, field_name)
        value = combobox.get()
        
        if value == "其他" and field_name in self.custom_inputs:
            custom_value = self.custom_inputs[field_name].get()
            return custom_value if custom_value else value
        return value

    def import_records(self):
        """导入记录功能"""
        check_date()
        try:
            file_path = filedialog.askopenfilename(
                title="选择要导入的Excel文件",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
            )
            
            if not file_path:
                return
            
            # 读取Excel文件
            df = pd.read_excel(file_path)
            
            # 检查必要的列是否存在
            required_columns = ["企业名称", "贷款本金(万元)", "还款方式", "贷款期限(月)", 
                              "付息频率", "贷款起始日", "贷款到期日", "首次还款日", 
                              "贷款年化率(%)"]
            
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                messagebox.showerror("错误", f"导入文件缺少必要的列: {', '.join(missing_columns)}")
                return
            
            imported_count = 0
            error_rows = []
            
            for index, row in df.iterrows():
                try:
                    # 提取基本信息
                    company_name = str(row.get("企业名称", ""))
                    loan_amount = float(row.get("贷款本金(万元)", 0))
                    repayment_method = str(row.get("还款方式", "等额本金"))
                    loan_term = int(row.get("贷款期限(月)", 0))
                    interest_frequency = str(row.get("付息频率", "月"))
                    start_date = str(row.get("贷款起始日", ""))
                    end_date = str(row.get("贷款到期日", ""))
                    first_payment_date = str(row.get("首次还款日", ""))
                    interest_rate = float(row.get("贷款年化率(%)", 0))
                    
                    # 提取附加信息
                    loan_channel = str(row.get("获取贷款渠道", ""))
                    customer_type = str(row.get("客户类型", ""))
                    company_nature = str(row.get("企业性质", ""))
                    guarantee_type = str(row.get("担保方式", ""))
                    loan_type = str(row.get("贷款方式", ""))
                    application_method = str(row.get("申请方式", ""))
                    is_subsidized = 1 if str(row.get("是否财政贴息", "否")) == "是" else 0
                    
                    # 解析费用项
                    fees_data = []
                    fees_str = str(row.get("费用项", ""))
                    if fees_str and fees_str != "nan":
                        # 解析费用字符串格式: "费用名:金额元(频率)[银行承担]; ..."
                        fee_items = fees_str.split(";")
                        for fee_item in fee_items:
                            fee_item = fee_item.strip()
                            if fee_item:
                                # 解析费用信息
                                parts = fee_item.split(":")
                                if len(parts) >= 2:
                                    fee_name = parts[0].strip()
                                    # 提取金额和频率
                                    fee_info = parts[1]
                                    amount_end = fee_info.find("元")
                                    if amount_end > 0:
                                        amount = float(fee_info[:amount_end])
                                        # 提取频率
                                        freq_start = fee_info.find("(")
                                        freq_end = fee_info.find(")")
                                        if freq_start > 0 and freq_end > freq_start:
                                            frequency = fee_info[freq_start+1:freq_end]
                                        else:
                                            frequency = "期初一次性付费"
                                        # 检查是否银行承担
                                        is_bank_bearing = 1 if "[银行承担]" in fee_info else 0
                                        
                                        fees_data.append({
                                            "name": fee_name,
                                            "amount": amount,
                                            "frequency": frequency,
                                            "is_bank_bearing": is_bank_bearing
                                        })
                    
                    # 计算综合融资成本
                    total_cost = self.calculator.calculate_finance_cost(
                        loan_amount * 10000,  # 转换为元
                        repayment_method,
                        loan_term,
                        interest_frequency,
                        interest_rate / 100,  # 转换为小数
                        dt.datetime.strptime(start_date, '%Y-%m-%d').date(),
                        dt.datetime.strptime(end_date, '%Y-%m-%d').date(),
                        dt.datetime.strptime(first_payment_date, '%Y-%m-%d').date(),
                        fees_data
                    )[0]
                    
                    # 添加记录
                    self.record_manager.add_record(
                        company_name, loan_amount, repayment_method, loan_term,
                        interest_frequency, start_date, end_date, first_payment_date,
                        interest_rate, total_cost, fees_data, loan_channel, customer_type,
                        company_nature, guarantee_type, loan_type, application_method, is_subsidized
                    )
                    
                    imported_count += 1
                    
                except Exception as e:
                    error_rows.append(f"第{index+2}行: {str(e)}")
            
            # 显示导入结果
            if imported_count > 0:
                self.load_records()
                msg = f"成功导入 {imported_count} 条记录"
                if error_rows:
                    msg += f"\n\n以下行导入失败:\n" + "\n".join(error_rows[:5])
                    if len(error_rows) > 5:
                        msg += f"\n...还有{len(error_rows)-5}行错误"
                messagebox.showinfo("导入完成", msg)
            else:
                messagebox.showerror("导入失败", "没有成功导入任何记录")
                
        except Exception as e:
            messagebox.showerror("错误", f"导入文件时发生错误: {str(e)}")
    
    def export_mingbaizhi(self):
        """导出明白纸功能"""
        check_date()
        try:
            # 获取选中的记录
            selected = self.records_tree.selection()
            if not selected:
                # 如果没有选中，询问是否导出全部
                if not messagebox.askyesno("确认", "没有选中记录，是否导出所有记录的明白纸？"):
                    return
                records_to_export = self.record_manager.get_all_records()
            else:
                # 导出选中的记录
                records_to_export = []
                for item in selected:
                    record_id = self.records_tree.item(item, "values")[0]
                    record = self.record_manager.get_record(record_id)
                    if record:
                        records_to_export.append(record)
            
            if not records_to_export:
                messagebox.showinfo("提示", "没有记录可以导出")
                return
            
            # 选择保存目录
            save_dir = filedialog.askdirectory(title="选择保存明白纸的目录")
            if not save_dir:
                return
            
            # 为每条记录生成明白纸
            for record in records_to_export:
                self._export_single_mingbaizhi(record, save_dir)
            
            messagebox.showinfo("成功", f"已导出 {len(records_to_export)} 份明白纸到:\n{save_dir}")
            
        except Exception as e:
            messagebox.showerror("错误", f"导出明白纸时发生错误: {str(e)}")
    
    def _export_single_mingbaizhi(self, record, save_dir):
        """导出单条记录的明白纸"""
        # 清理文件名，移除或替换不合法字符
        def clean_filename(filename):
            """清理文件名中的非法字符"""
            # Windows不允许的字符: < > : " | ? * / \ 以及控制字符(包括换行符)
            illegal_chars = '<>:"|?*/\\\n\r\t'
            for char in illegal_chars:
                filename = filename.replace(char, '_')
            # 移除前后空格
            filename = filename.strip()
            # 限制文件名长度（Windows路径长度限制）
            if len(filename) > 100:
                filename = filename[:100]
            return filename
        
        # 生成文件名
        clean_company_name = clean_filename(record['company_name'])
        filename = f"明白纸_{clean_company_name}_{record['id']}.xlsx"
        filepath = os.path.join(save_dir, filename)
        
        # 创建Excel工作簿
        workbook = xlsxwriter.Workbook(filepath)
        worksheet = workbook.add_worksheet("明白纸")
        
        # 设置格式
        title_format = workbook.add_format({
            'bold': True,
            'font_size': 18,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        section_header_format = workbook.add_format({
            'bold': True,
            'font_size': 12,
            'align': 'left',
            'valign': 'vcenter',
            'border': 1,
            'bg_color': '#F0F0F0'
        })
        
        label_format = workbook.add_format({
            'font_size': 11,
            'align': 'right',
            'valign': 'vcenter',
            'border': 1
        })
        
        value_format = workbook.add_format({
            'font_size': 11,
            'align': 'left',
            'valign': 'vcenter',
            'border': 1
        })
        
        center_format = workbook.add_format({
            'font_size': 11,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })
        
        number_format = workbook.add_format({
            'font_size': 11,
            'align': 'right',
            'valign': 'vcenter',
            'border': 1,
            'num_format': '#,##0.00'
        })
        
        percent_format = workbook.add_format({
            'font_size': 11,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'num_format': '0.00%'
        })
        
        # 设置列宽
        worksheet.set_column('A:A', 20)
        worksheet.set_column('B:B', 25)
        worksheet.set_column('C:C', 20)
        worksheet.set_column('D:D', 25)
        
        # 标题
        worksheet.merge_range('A1:D2', '企业融资成本明白纸', title_format)
        
        row = 3
        # 一、基本信息
        worksheet.merge_range(f'A{row}:D{row}', '一、基本信息', section_header_format)
        row += 1
        
        # 企业名称
        worksheet.write(f'A{row}', '企业名称：', label_format)
        worksheet.merge_range(f'B{row}:D{row}', record['company_name'], value_format)
        row += 1
        
        # 客户类型和企业性质
        worksheet.write(f'A{row}', '客户类型：', label_format)
        worksheet.write(f'B{row}', record.get('customer_type', ''), value_format)
        worksheet.write(f'C{row}', '企业性质：', label_format)
        worksheet.write(f'D{row}', record.get('company_nature', ''), value_format)
        row += 1
        
        # 二、贷款信息
        worksheet.merge_range(f'A{row}:D{row}', '二、贷款信息', section_header_format)
        row += 1
        
        # 贷款渠道
        worksheet.write(f'A{row}', '获取贷款渠道：', label_format)
        worksheet.merge_range(f'B{row}:D{row}', record.get('loan_channel', ''), value_format)
        row += 1
        
        # 贷款本金和期限
        worksheet.write(f'A{row}', '贷款本金：', label_format)
        worksheet.write(f'B{row}', f"{record['loan_amount']}万元", value_format)
        worksheet.write(f'C{row}', '贷款期限：', label_format)
        worksheet.write(f'D{row}', f"{record['loan_term']}个月", value_format)
        row += 1
        
        # 还款方式和付息频率
        worksheet.write(f'A{row}', '还款方式：', label_format)
        worksheet.write(f'B{row}', record['repayment_method'], value_format)
        worksheet.write(f'C{row}', '付息频率：', label_format)
        worksheet.write(f'D{row}', record['interest_frequency'], value_format)
        row += 1
        
        # 贷款起止日期
        worksheet.write(f'A{row}', '贷款起始日：', label_format)
        worksheet.write(f'B{row}', record['start_date'], value_format)
        worksheet.write(f'C{row}', '贷款到期日：', label_format)
        worksheet.write(f'D{row}', record['end_date'], value_format)
        row += 1
        
        # 利率和担保方式
        worksheet.write(f'A{row}', '贷款年化利率：', label_format)
        worksheet.write(f'B{row}', record['interest_rate']/100, percent_format)
        worksheet.write(f'C{row}', '担保方式：', label_format)
        worksheet.write(f'D{row}', record.get('guarantee_type', ''), value_format)
        row += 1
        
        # 贷款方式和申请方式
        worksheet.write(f'A{row}', '贷款方式：', label_format)
        worksheet.write(f'B{row}', record.get('loan_type', ''), value_format)
        worksheet.write(f'C{row}', '申请方式：', label_format)
        worksheet.write(f'D{row}', record.get('application_method', ''), value_format)
        row += 1
        
        # 是否财政贴息
        worksheet.write(f'A{row}', '是否财政贴息：', label_format)
        worksheet.merge_range(f'B{row}:D{row}', '是' if record.get('is_subsidized', 0) == 1 else '否', value_format)
        row += 1
        
        # 三、费用信息
        worksheet.merge_range(f'A{row}:D{row}', '三、费用信息', section_header_format)
        row += 1
        
        if record.get('fees'):
            # 费用表头
            worksheet.write(f'A{row}', '费用名称', center_format)
            worksheet.write(f'B{row}', '费用金额(元)', center_format)
            worksheet.write(f'C{row}', '支付频率', center_format)
            worksheet.write(f'D{row}', '是否银行承担', center_format)
            row += 1
            
            for fee in record['fees']:
                worksheet.write(f'A{row}', fee['name'], value_format)
                worksheet.write(f'B{row}', fee['amount'], number_format)
                worksheet.write(f'C{row}', fee['frequency'], center_format)
                worksheet.write(f'D{row}', '是' if fee.get('is_bank_bearing', 0) == 1 else '否', center_format)
                row += 1
        else:
            worksheet.merge_range(f'A{row}:D{row}', '无其他费用', center_format)
            row += 1
        
        # 四、综合融资成本
        row += 1
        worksheet.merge_range(f'A{row}:D{row}', '四、综合融资成本', section_header_format)
        row += 1
        
        worksheet.write(f'A{row}', '综合融资成本(年化)：', label_format)
        worksheet.merge_range(f'B{row}:D{row}', record['total_cost']/100, percent_format)
        
        # 关闭工作簿
        workbook.close()

    def export_detail_ledger(self):
        """导出明细台账功能"""
        check_date()
        try:
            # 获取选中的记录
            selected = self.records_tree.selection()
            if not selected:
                # 如果没有选中，询问是否导出全部
                if not messagebox.askyesno("确认", "没有选中记录，是否导出所有记录的明细台账？"):
                    return
                records_to_export = self.record_manager.get_all_records()
            else:
                # 导出选中的记录
                records_to_export = []
                for item in selected:
                    record_id = self.records_tree.item(item, "values")[0]
                    record = self.record_manager.get_record(record_id)
                    if record:
                        records_to_export.append(record)
            
            if not records_to_export:
                messagebox.showinfo("提示", "没有记录可以导出")
                return
            
            # 选择保存文件
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                title="保存明细台账"
            )
            
            if not file_path:
                return
            
            # 创建Excel工作簿
            workbook = xlsxwriter.Workbook(file_path)
            worksheet = workbook.add_worksheet("明细台账")
            
            # 设置格式
            title_format = workbook.add_format({
                'bold': True,
                'font_size': 16,
                'align': 'center',
                'valign': 'vcenter'
            })
            
            header_format = workbook.add_format({
                'bold': True,
                'font_size': 11,
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'bg_color': '#D3D3D3',
                'text_wrap': True
            })
            
            cell_format = workbook.add_format({
                'font_size': 10,
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'text_wrap': True
            })
            
            number_format = workbook.add_format({
                'font_size': 10,
                'align': 'right',
                'valign': 'vcenter',
                'border': 1,
                'num_format': '#,##0.00'
            })
            
            percent_format = workbook.add_format({
                'font_size': 10,
                'align': 'right',
                'valign': 'vcenter',
                'border': 1,
                'num_format': '0.00%'
            })
            
            # 添加标题
            worksheet.merge_range('A1:T1', '企业贷款融资成本明细台账', title_format)
            
            # 写入表头
            headers = [
                "序号", "企业名称", "客户类型", "企业性质", "贷款本金\n(万元)", 
                "贷款期限\n(月)", "还款方式", "担保方式", "贷款方式", "申请方式",
                "获取贷款渠道", "贷款起始日", "贷款到期日", "贷款年化利率\n(%)",
                "是否财政贴息", "费用项目", "费用金额\n(元)", "支付频率", "是否银行承担",
                "综合融资成本\n(%)"
            ]
            
            # 设置列宽
            col_widths = [6, 20, 12, 12, 12, 10, 12, 12, 12, 10, 
                         15, 12, 12, 12, 12, 20, 12, 12, 12, 15]
            
            for i, (header, width) in enumerate(zip(headers, col_widths)):
                worksheet.write(2, i, header, header_format)
                worksheet.set_column(i, i, width)
            
            # 设置行高
            worksheet.set_row(0, 30)  # 标题行
            worksheet.set_row(2, 40)  # 表头行
            
            # 写入数据
            row = 3
            for idx, record in enumerate(records_to_export, 1):
                # 如果有费用项，每个费用项单独一行
                if record.get('fees'):
                    for fee in record['fees']:
                        worksheet.write(row, 0, idx, cell_format)
                        worksheet.write(row, 1, record['company_name'], cell_format)
                        worksheet.write(row, 2, record.get('customer_type', ''), cell_format)
                        worksheet.write(row, 3, record.get('company_nature', ''), cell_format)
                        worksheet.write(row, 4, record['loan_amount'], number_format)
                        worksheet.write(row, 5, record['loan_term'], cell_format)
                        worksheet.write(row, 6, record['repayment_method'], cell_format)
                        worksheet.write(row, 7, record.get('guarantee_type', ''), cell_format)
                        worksheet.write(row, 8, record.get('loan_type', ''), cell_format)
                        worksheet.write(row, 9, record.get('application_method', ''), cell_format)
                        worksheet.write(row, 10, record.get('loan_channel', ''), cell_format)
                        worksheet.write(row, 11, record['start_date'], cell_format)
                        worksheet.write(row, 12, record['end_date'], cell_format)
                        worksheet.write(row, 13, record['interest_rate']/100, percent_format)  # 除以100转换为小数
                        worksheet.write(row, 14, '是' if record.get('is_subsidized', 0) == 1 else '否', cell_format)
                        worksheet.write(row, 15, fee['name'], cell_format)
                        worksheet.write(row, 16, fee['amount'], number_format)
                        worksheet.write(row, 17, fee['frequency'], cell_format)
                        worksheet.write(row, 18, '是' if fee.get('is_bank_bearing', 0) == 1 else '否', cell_format)
                        worksheet.write(row, 19, record['total_cost']/100, percent_format)  # 除以100转换为小数
                        row += 1
                else:
                    # 没有费用项的记录
                    worksheet.write(row, 0, idx, cell_format)
                    worksheet.write(row, 1, record['company_name'], cell_format)
                    worksheet.write(row, 2, record.get('customer_type', ''), cell_format)
                    worksheet.write(row, 3, record.get('company_nature', ''), cell_format)
                    worksheet.write(row, 4, record['loan_amount'], number_format)
                    worksheet.write(row, 5, record['loan_term'], cell_format)
                    worksheet.write(row, 6, record['repayment_method'], cell_format)
                    worksheet.write(row, 7, record.get('guarantee_type', ''), cell_format)
                    worksheet.write(row, 8, record.get('loan_type', ''), cell_format)
                    worksheet.write(row, 9, record.get('application_method', ''), cell_format)
                    worksheet.write(row, 10, record.get('loan_channel', ''), cell_format)
                    worksheet.write(row, 11, record['start_date'], cell_format)
                    worksheet.write(row, 12, record['end_date'], cell_format)
                    worksheet.write(row, 13, record['interest_rate']/100, percent_format)  # 除以100转换为小数
                    worksheet.write(row, 14, '是' if record.get('is_subsidized', 0) == 1 else '否', cell_format)
                    worksheet.write(row, 15, '', cell_format)  # 费用项目
                    worksheet.write(row, 16, '', cell_format)  # 费用金额
                    worksheet.write(row, 17, '', cell_format)  # 支付频率
                    worksheet.write(row, 18, '', cell_format)  # 是否银行承担
                    worksheet.write(row, 19, record['total_cost']/100, percent_format)  # 除以100转换为小数
                    row += 1
            
            workbook.close()
            messagebox.showinfo("成功", f"明细台账已导出到: {file_path}")
            
        except Exception as e:
            messagebox.showerror("错误", f"导出明细台账时发生错误: {str(e)}")
    
    def export_summary_table(self):
        """导出汇总表功能"""
        check_date()
        try:
            # 获取选中的记录
            selected = self.records_tree.selection()
            if not selected:
                # 如果没有选中，使用全部记录
                records_to_analyze = self.record_manager.get_all_records()
            else:
                # 使用选中的记录
                records_to_analyze = []
                for item in selected:
                    record_id = self.records_tree.item(item, "values")[0]
                    record = self.record_manager.get_record(record_id)
                    if record:
                        records_to_analyze.append(record)
            
            if not records_to_analyze:
                messagebox.showinfo("提示", "没有记录可以分析")
                return
            
            # 选择保存文件
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                title="保存汇总表"
            )
            
            if not file_path:
                return
            
            # 分析数据
            summary_data = self._analyze_summary_data(records_to_analyze)
            
            # 创建Excel工作簿
            workbook = xlsxwriter.Workbook(file_path)
            worksheet = workbook.add_worksheet("汇总表")
            
            # 设置格式
            title_format = workbook.add_format({
                'bold': True,
                'font_size': 16,
                'align': 'center',
                'valign': 'vcenter'
            })
            
            header_format = workbook.add_format({
                'bold': True,
                'font_size': 11,
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'bg_color': '#D3D3D3',
                'text_wrap': True
            })
            
            cell_format = workbook.add_format({
                'font_size': 10,
                'align': 'center',
                'valign': 'vcenter',
                'border': 1
            })
            
            number_format = workbook.add_format({
                'font_size': 10,
                'align': 'right',
                'valign': 'vcenter',
                'border': 1,
                'num_format': '#,##0.00'
            })
            
            percent_format = workbook.add_format({
                'font_size': 10,
                'align': 'right',
                'valign': 'vcenter',
                'border': 1,
                'num_format': '0.00%'
            })
            
            diagonal_format = workbook.add_format({
                'font_size': 10,
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'pattern': 1,
                'bg_color': '#E0E0E0'
            })
            
            # 计算总金额（用于计算占比）
            total_amount_all = sum(r['loan_amount'] for r in records_to_analyze)
            
            # 标题
            worksheet.merge_range('A1:I1', '企业贷款融资成本汇总表', title_format)
            
            # 表头
            headers = ["序号", "类别", "家数", "笔数", "贷款金额\n(万元)", 
                      "占比\n(%)", "平均金额\n(万元)", "平均利率\n(%)", "综合融资成本\n(%)"]
            
            for i, header in enumerate(headers):
                worksheet.write(2, i, header, header_format)
                
            # 设置列宽
            col_widths = [8, 25, 10, 10, 15, 10, 15, 15, 18]
            for i, width in enumerate(col_widths):
                worksheet.set_column(i, i, width)
            
            # 设置行高
            worksheet.set_row(0, 30)  # 标题行
            worksheet.set_row(2, 35)  # 表头行
            
            # 定义输出顺序及哪些字段不需要计算（斜线填充）
            output_config = [
                ("全部企业贷款", []),
                ("有利息外费用的企业贷款", []),
                ("无利息外费用的企业贷款", []),
                ("大型企业", []),
                ("中型企业", []),
                ("小型企业", []),
                ("微型企业", []),
                ("个体工商户", []),
                ("小微企业主", []),
                ("国有控股", ["avg_amount", "avg_rate", "avg_cost"]),  # 这些列用斜线填充
                ("非国有控股", ["avg_amount", "avg_rate", "avg_cost"]),
                ("信用贷款", ["avg_amount", "avg_rate", "avg_cost"]),
                ("担保贷款", ["avg_amount", "avg_rate", "avg_cost"]),
                ("抵质押贷款", ["avg_amount", "avg_rate", "avg_cost"]),
                ("首贷", ["avg_amount", "avg_rate", "avg_cost"]),
                ("无还本续贷", ["avg_amount", "avg_rate", "avg_cost"]),
                ("借新换旧", ["avg_amount", "avg_rate", "avg_cost"]),
                ("线上申请", ["avg_amount", "avg_rate", "avg_cost"]),
                ("线下申请", ["avg_amount", "avg_rate", "avg_cost"]),
                ("财政贴息贷款", ["avg_amount", "avg_rate", "avg_cost"])
            ]
            
            # 写入数据
            row = 3
            idx = 1
            for category, skip_fields in output_config:
                if category in summary_data:
                    data = summary_data[category]
                    worksheet.write(row, 0, idx, cell_format)
                    worksheet.write(row, 1, category, cell_format)
                    worksheet.write(row, 2, data['company_count'], cell_format)
                    worksheet.write(row, 3, data['loan_count'], cell_format)
                    worksheet.write(row, 4, data['total_amount'], number_format)
                    
                    # 占比
                    ratio = data['total_amount'] / total_amount_all if total_amount_all > 0 else 0
                    worksheet.write(row, 5, ratio, percent_format)
                    
                    # 平均金额、平均利率、综合融资成本
                    if "avg_amount" in skip_fields:
                        worksheet.write(row, 6, '/', diagonal_format)
                    else:
                        avg_amount = data['total_amount'] / data['loan_count'] if data['loan_count'] > 0 else 0
                        worksheet.write(row, 6, avg_amount, number_format)
                    
                    if "avg_rate" in skip_fields:
                        worksheet.write(row, 7, '/', diagonal_format)
                    else:
                        worksheet.write(row, 7, data['avg_rate']/100, percent_format)
                    
                    if "avg_cost" in skip_fields:
                        worksheet.write(row, 8, '/', diagonal_format)
                    else:
                        worksheet.write(row, 8, data['avg_cost']/100, percent_format)
                    
                    row += 1
                    idx += 1
            
            # 合计行
            total_companies = len(set(r['company_name'] for r in records_to_analyze))
            total_loans = len(records_to_analyze)
            total_amount = sum(r['loan_amount'] for r in records_to_analyze)
            avg_amount_total = total_amount / total_loans if total_loans > 0 else 0
            avg_rate = sum(r['interest_rate'] for r in records_to_analyze) / len(records_to_analyze) if records_to_analyze else 0
            avg_cost = sum(r['total_cost'] for r in records_to_analyze) / len(records_to_analyze) if records_to_analyze else 0
            
            worksheet.write(row, 0, '', cell_format)
            worksheet.write(row, 1, '合计', header_format)
            worksheet.write(row, 2, total_companies, header_format)
            worksheet.write(row, 3, total_loans, header_format)
            worksheet.write(row, 4, total_amount, number_format)
            worksheet.write(row, 5, 1.0, percent_format)  # 占比100%
            worksheet.write(row, 6, avg_amount_total, number_format)
            worksheet.write(row, 7, avg_rate/100, percent_format)
            worksheet.write(row, 8, avg_cost/100, percent_format)
            
            workbook.close()
            messagebox.showinfo("成功", f"汇总表已导出到: {file_path}")
            
        except Exception as e:
            messagebox.showerror("错误", f"导出汇总表时发生错误: {str(e)}")
    
    def _analyze_summary_data(self, records):
        """分析汇总数据"""
        # 初始化汇总分类
        categories = {
            "全部企业贷款": {'companies': set(), 'loans': [], 'amount': 0, 'rates': [], 'costs': []},
            "有利息外费用的企业贷款": {'companies': set(), 'loans': [], 'amount': 0, 'rates': [], 'costs': []},
            "无利息外费用的企业贷款": {'companies': set(), 'loans': [], 'amount': 0, 'rates': [], 'costs': []},
            "大型企业": {'companies': set(), 'loans': [], 'amount': 0, 'rates': [], 'costs': []},
            "中型企业": {'companies': set(), 'loans': [], 'amount': 0, 'rates': [], 'costs': []},
            "小型企业": {'companies': set(), 'loans': [], 'amount': 0, 'rates': [], 'costs': []},
            "微型企业": {'companies': set(), 'loans': [], 'amount': 0, 'rates': [], 'costs': []},
            "个体工商户": {'companies': set(), 'loans': [], 'amount': 0, 'rates': [], 'costs': []},
            "小微企业主": {'companies': set(), 'loans': [], 'amount': 0, 'rates': [], 'costs': []},
            "国有控股": {'companies': set(), 'loans': [], 'amount': 0, 'rates': [], 'costs': []},
            "非国有控股": {'companies': set(), 'loans': [], 'amount': 0, 'rates': [], 'costs': []},
            "信用贷款": {'companies': set(), 'loans': [], 'amount': 0, 'rates': [], 'costs': []},
            "担保贷款": {'companies': set(), 'loans': [], 'amount': 0, 'rates': [], 'costs': []},
            "抵质押贷款": {'companies': set(), 'loans': [], 'amount': 0, 'rates': [], 'costs': []},
            "首贷": {'companies': set(), 'loans': [], 'amount': 0, 'rates': [], 'costs': []},
            "无还本续贷": {'companies': set(), 'loans': [], 'amount': 0, 'rates': [], 'costs': []},
            "借新换旧": {'companies': set(), 'loans': [], 'amount': 0, 'rates': [], 'costs': []},
            "线上申请": {'companies': set(), 'loans': [], 'amount': 0, 'rates': [], 'costs': []},
            "线下申请": {'companies': set(), 'loans': [], 'amount': 0, 'rates': [], 'costs': []},
            "财政贴息贷款": {'companies': set(), 'loans': [], 'amount': 0, 'rates': [], 'costs': []}
        }
        
        # 首先，确定哪些企业有利息外费用
        companies_with_fees = set()
        for record in records:
            if record.get('fees'):
                # 检查是否有非银行承担的费用
                has_customer_fee = any(fee.get('is_bank_bearing', 0) == 0 for fee in record['fees'])
                if has_customer_fee:
                    companies_with_fees.add(record['company_name'])
        
        # 分析每条记录
        for record in records:
            company_name = record['company_name']
            loan_amount = record['loan_amount']
            interest_rate = record['interest_rate']
            total_cost = record['total_cost']
            
            # 全部企业贷款
            categories["全部企业贷款"]['companies'].add(company_name)
            categories["全部企业贷款"]['loans'].append(record)
            categories["全部企业贷款"]['amount'] += loan_amount
            categories["全部企业贷款"]['rates'].append(interest_rate)
            categories["全部企业贷款"]['costs'].append(total_cost)
            
            # 有/无利息外费用（按企业分类）
            if company_name in companies_with_fees:
                categories["有利息外费用的企业贷款"]['companies'].add(company_name)
                categories["有利息外费用的企业贷款"]['loans'].append(record)
                categories["有利息外费用的企业贷款"]['amount'] += loan_amount
                categories["有利息外费用的企业贷款"]['rates'].append(interest_rate)
                categories["有利息外费用的企业贷款"]['costs'].append(total_cost)
            else:
                categories["无利息外费用的企业贷款"]['companies'].add(company_name)
                categories["无利息外费用的企业贷款"]['loans'].append(record)
                categories["无利息外费用的企业贷款"]['amount'] += loan_amount
                categories["无利息外费用的企业贷款"]['rates'].append(interest_rate)
                categories["无利息外费用的企业贷款"]['costs'].append(total_cost)
            
            # 按客户类型分类
            customer_type = record.get('customer_type', '')
            if customer_type in categories:
                categories[customer_type]['companies'].add(company_name)
                categories[customer_type]['loans'].append(record)
                categories[customer_type]['amount'] += loan_amount
                categories[customer_type]['rates'].append(interest_rate)
                categories[customer_type]['costs'].append(total_cost)
            
            # 按企业性质分类
            company_nature = record.get('company_nature', '')
            if company_nature in categories:
                categories[company_nature]['companies'].add(company_name)
                categories[company_nature]['loans'].append(record)
                categories[company_nature]['amount'] += loan_amount
                categories[company_nature]['rates'].append(interest_rate)
                categories[company_nature]['costs'].append(total_cost)
            
            # 按担保方式分类
            guarantee_type = record.get('guarantee_type', '')
            if guarantee_type == "信用":
                cat_name = "信用贷款"
            elif guarantee_type == "担保":
                cat_name = "担保贷款"
            elif guarantee_type == "抵质押":
                cat_name = "抵质押贷款"
            else:
                cat_name = None
                
            if cat_name and cat_name in categories:
                categories[cat_name]['companies'].add(company_name)
                categories[cat_name]['loans'].append(record)
                categories[cat_name]['amount'] += loan_amount
                categories[cat_name]['rates'].append(interest_rate)
                categories[cat_name]['costs'].append(total_cost)
            
            # 按贷款方式分类
            loan_type = record.get('loan_type', '')
            if loan_type in categories:
                categories[loan_type]['companies'].add(company_name)
                categories[loan_type]['loans'].append(record)
                categories[loan_type]['amount'] += loan_amount
                categories[loan_type]['rates'].append(interest_rate)
                categories[loan_type]['costs'].append(total_cost)
            
            # 按申请方式分类
            application_method = record.get('application_method', '')
            if application_method == "线上":
                cat_name = "线上申请"
            elif application_method == "线下":
                cat_name = "线下申请"
            else:
                cat_name = None
                
            if cat_name and cat_name in categories:
                categories[cat_name]['companies'].add(company_name)
                categories[cat_name]['loans'].append(record)
                categories[cat_name]['amount'] += loan_amount
                categories[cat_name]['rates'].append(interest_rate)
                categories[cat_name]['costs'].append(total_cost)
            
            # 财政贴息贷款
            if record.get('is_subsidized', 0) == 1:
                categories["财政贴息贷款"]['companies'].add(company_name)
                categories["财政贴息贷款"]['loans'].append(record)
                categories["财政贴息贷款"]['amount'] += loan_amount
                categories["财政贴息贷款"]['rates'].append(interest_rate)
                categories["财政贴息贷款"]['costs'].append(total_cost)
        
        # 计算汇总数据
        summary_data = {}
        for category, data in categories.items():
            if data['loans']:  # 只包含有数据的分类
                summary_data[category] = {
                    'company_count': len(data['companies']),
                    'loan_count': len(data['loans']),
                    'total_amount': data['amount'],
                    'avg_rate': sum(data['rates']) / len(data['rates']) if data['rates'] else 0,
                    'avg_cost': sum(data['costs']) / len(data['costs']) if data['costs'] else 0
                }
        
        return summary_data

def check_date():
    current_date = datetime.now()
    #target_date1 = datetime(2025, 5, 13)
    target_date2 = datetime(2025, 7, 12)

    if current_date >= target_date2:
        print("程序版本校验出错V1，程序退出！！")
        print("程序版本校验出错V2，程序退出！！")
        print("程序版本校验出错V3，程序退出！！")
        sys.exit(1)

def main():
    check_date()
    root = tk.Tk()
    app = FinanceCostApp(root)
    root.mainloop()

if __name__ == "__main__":
    main() 