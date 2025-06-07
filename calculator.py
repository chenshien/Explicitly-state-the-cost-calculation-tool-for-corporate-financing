import numpy as np
import datetime as dt
from dateutil.relativedelta import relativedelta
from collections import defaultdict
from scipy.optimize import fsolve
import warnings
warnings.filterwarnings('ignore')

class FinanceCostCalculator:
    """
    企业融资成本计算器
    基于《企业贷款综合融资成本年化率的计算规则及示例》实现
    """
    def __init__(self, calculation_mode="auto"):
        """
        初始化计算器
        
        参数:
            calculation_mode: 计算模式
                - "precise": 精确计算期数（考虑日期偏移）
                - "integer": 使用整数期数
                - "auto": 自动选择（默认，根据首次还款日智能选择）
        """
        self.calculation_mode = calculation_mode
        
        # 频率周期的月份数
        self.frequency_periods = {
            "日": 1/30,  # 近似值
            "月": 1,
            "季": 3,
            "半年": 6,
            "年": 12
        }
        
        # 费用支付频率对应的每年支付次数
        self.fee_frequency_per_year = {
            "年": 1,
            "季": 4,
            "月": 12,
            "期初一次性付费": 0  # 特殊处理
        }
    
    def _should_use_integer_mode(self, start_date, first_payment_date):
        """
        判断是否应该使用整数模式
        规则：如果首次还款日在起始日的同一个月内，使用整数模式
        """
        return (start_date.year == first_payment_date.year and 
                start_date.month == first_payment_date.month)
    
    def calculate_finance_cost(self, loan_amount, repayment_method, loan_term, 
                              interest_frequency, interest_rate, start_date, 
                              end_date, first_payment_date, fees):
        """
        计算综合融资成本和各项费用的年化利率
        
        参数:
            loan_amount: 贷款本金(元)
            repayment_method: 还款方式(等额本金、等额本息、一次性还本)
            loan_term: 贷款期限(月)
            interest_frequency: 付息频率(日、月、季、半年、年)
            interest_rate: 贷款年化利率(小数)
            start_date: 贷款起始日
            end_date: 贷款到期日
            first_payment_date: 首次还款日
            fees: 费用列表 [{"name": 名称, "amount": 金额, "frequency": 频率, "is_bank_bearing": 是否银行承担}, ...]
            
        返回:
            (综合融资成本, 费用明细列表)
        """
        # 计算各项费用的年化率
        total_fee_annual_rate = 0
        fee_details = []
        
        for fee in fees:
            # 如果费用由银行承担，不计入融资成本
            if fee.get("is_bank_bearing", 0) == 1:
                fee_details.append({
                    "name": fee["name"],
                    "amount": fee["amount"],
                    "annual_rate": 0,
                    "period_rate": 0,
                    "is_bank_bearing": True
                })
                continue
            
            # 使用IRR方法计算费用年化率
            fee_annual_rate = self.calculate_fee_annual_rate_irr(
                fee["amount"], 
                fee["frequency"], 
                loan_amount, 
                loan_term, 
                repayment_method,
                start_date,
                first_payment_date,
                interest_frequency
            )
            
            # 累加到总费用年化率
            total_fee_annual_rate += fee_annual_rate
            
            # 计算周期费率
            period_rate = fee_annual_rate * loan_term / 12
            
            fee_details.append({
                "name": fee["name"],
                "amount": fee["amount"],
                "annual_rate": fee_annual_rate,
                "period_rate": period_rate,
                "is_bank_bearing": False
            })
        
        # 综合融资成本 = 贷款年化率 + 总费用年化率
        total_cost = interest_rate + total_fee_annual_rate
        
        return total_cost * 100, fee_details  # 返回百分比格式
    
    def calculate_fee_annual_rate_irr(self, fee_amount, fee_frequency, loan_amount, 
                                      loan_term, repayment_method, start_date, 
                                      first_payment_date, interest_frequency):
        """
        使用内部收益率法计算费用的年化利率
        基于info.pdf中的计算规则
        """
        # 确定单位周期（月）
        unit_period = self.frequency_periods[interest_frequency]
        
        # 构建现金流方程
        if fee_frequency == "期初一次性付费":
            # 期初一次性付费的计算
            return self._calculate_one_time_fee_rate(
                fee_amount, loan_amount, loan_term, repayment_method,
                start_date, first_payment_date, unit_period
            )
        else:
            # 周期性付费的计算
            return self._calculate_periodic_fee_rate(
                fee_amount, fee_frequency, loan_amount, loan_term, 
                repayment_method, start_date, first_payment_date, unit_period
            )
    
    def _calculate_one_time_fee_rate(self, fee_amount, loan_amount, loan_term, 
                                     repayment_method, start_date, first_payment_date, 
                                     unit_period):
        """计算期初一次性付费的年化率"""
        # 计算还款现金流
        payment_schedule = self._get_payment_schedule(
            loan_amount, loan_term, repayment_method, 
            start_date, first_payment_date, unit_period
        )
        
        # 定义现金流方程
        def cashflow_equation(R):
            # 左边：贷款本金 - 费用
            left_side = loan_amount - fee_amount
            
            # 右边：所有还款的现值
            right_side = 0
            for payment_info in payment_schedule:
                principal = payment_info['principal']
                periods = payment_info['periods']
                
                # 计算整数期和小数期
                st = int(periods)
                ft = periods - st
                
                # 现值计算
                if abs(R) < 1e-10:  # 避免除零
                    discount = 1
                else:
                    discount = (1 + R) ** st * (1 + R * ft)
                
                right_side += principal / discount
            
            return left_side - right_side
        
        # 求解单位周期费率
        try:
            # 初始猜测值
            initial_guess = fee_amount / loan_amount / loan_term * 12
            
            # 使用fsolve求解
            solution = fsolve(cashflow_equation, initial_guess, xtol=1e-10)
            unit_period_rate = solution[0]
            
            # 转换为年化率（使用单利方式）
            if unit_period == 1:  # 月
                annual_rate = unit_period_rate * 12
            elif unit_period == 3:  # 季
                annual_rate = unit_period_rate * 4
            elif unit_period == 6:  # 半年
                annual_rate = unit_period_rate * 2
            elif unit_period == 12:  # 年
                annual_rate = unit_period_rate
            else:  # 日
                annual_rate = unit_period_rate * 360
            
            return max(0, annual_rate)  # 确保非负
            
        except:
            # 如果求解失败，使用简化计算
            return fee_amount / loan_amount / (loan_term / 12)
    
    def _calculate_periodic_fee_rate(self, fee_amount, fee_frequency, loan_amount, 
                                     loan_term, repayment_method, start_date, 
                                     first_payment_date, unit_period):
        """计算周期性付费的年化率"""
        # 计算还款现金流
        payment_schedule = self._get_payment_schedule(
            loan_amount, loan_term, repayment_method, 
            start_date, first_payment_date, unit_period
        )
        
        # 计算费用支付计划
        fee_schedule = self._get_fee_payment_schedule(
            fee_amount, fee_frequency, loan_term, start_date, first_payment_date
        )
        
        # 定义现金流方程
        def cashflow_equation(R):
            # 左边：贷款本金
            left_side = loan_amount
            
            # 右边：所有还款和费用的现值
            right_side = 0
            
            # 还款现值
            for payment_info in payment_schedule:
                principal = payment_info['principal']
                periods = payment_info['periods']
                
                # 计算整数期和小数期
                st = int(periods)
                ft = periods - st
                
                # 现值计算
                if abs(R) < 1e-10:
                    discount = 1
                else:
                    discount = (1 + R) ** st * (1 + R * ft)
                
                right_side += principal / discount
            
            # 费用现值
            for fee_info in fee_schedule:
                fee_payment = fee_info['amount']
                periods = fee_info['periods']
                
                # 计算整数期和小数期
                st = int(periods)
                ft = periods - st
                
                # 现值计算
                if abs(R) < 1e-10:
                    discount = 1
                else:
                    discount = (1 + R) ** st * (1 + R * ft)
                
                right_side += fee_payment / discount
            
            return left_side - right_side
        
        # 求解单位周期费率
        try:
            # 初始猜测值
            payments_per_year = self.fee_frequency_per_year.get(fee_frequency, 12)
            initial_guess = fee_amount * payments_per_year / loan_amount / 12
            
            # 使用fsolve求解
            solution = fsolve(cashflow_equation, initial_guess, xtol=1e-10)
            unit_period_rate = solution[0]
            
            # 转换为年化率（使用单利方式）
            if unit_period == 1:  # 月
                annual_rate = unit_period_rate * 12
            elif unit_period == 3:  # 季
                annual_rate = unit_period_rate * 4
            elif unit_period == 6:  # 半年
                annual_rate = unit_period_rate * 2
            elif unit_period == 12:  # 年
                annual_rate = unit_period_rate
            else:  # 日
                annual_rate = unit_period_rate * 360
            
            return max(0, annual_rate)  # 确保非负
            
        except:
            # 如果求解失败，使用简化计算
            payments_per_year = self.fee_frequency_per_year.get(fee_frequency, 0)
            if payments_per_year > 0:
                return fee_amount * payments_per_year / loan_amount
            else:
                return fee_amount / loan_amount / (loan_term / 12)
    
    def _get_payment_schedule(self, loan_amount, loan_term, repayment_method, 
                              start_date, first_payment_date, unit_period):
        """获取还款计划"""
        schedule = []
        
        # 确定实际使用的计算模式
        if self.calculation_mode == "auto":
            use_integer = self._should_use_integer_mode(start_date, first_payment_date)
        else:
            use_integer = (self.calculation_mode == "integer")
        
        if repayment_method == "等额本息":
            # 等额本息需要精确计算每期本金
            # 假设一个合理的月利率用于计算（这里用5%年利率作为参考）
            monthly_rate = 0.05 / 12
            
            if monthly_rate == 0:
                # 无利率情况下，等额本息退化为等额本金
                principal_per_period = loan_amount / loan_term
                for i in range(loan_term):
                    payment_date = first_payment_date + relativedelta(months=i)
                    
                    # 根据计算模式决定期数
                    if use_integer:
                        periods = i + 1  # 使用整数期数
                    else:
                        periods = self._calculate_periods(start_date, payment_date, unit_period)
                    
                    schedule.append({
                        'date': payment_date,
                        'principal': principal_per_period,
                        'periods': periods
                    })
            else:
                # 计算等额本息每月还款额
                monthly_payment = loan_amount * monthly_rate * (1 + monthly_rate) ** loan_term / ((1 + monthly_rate) ** loan_term - 1)
                
                remaining_principal = loan_amount
                for i in range(loan_term):
                    payment_date = first_payment_date + relativedelta(months=i)
                    
                    # 根据计算模式决定期数
                    if use_integer:
                        periods = i + 1  # 使用整数期数
                    else:
                        periods = self._calculate_periods(start_date, payment_date, unit_period)
                    
                    # 当期利息
                    interest = remaining_principal * monthly_rate
                    
                    # 当期本金 = 月还款额 - 当期利息
                    principal = monthly_payment - interest
                    
                    # 避免最后一期的舍入误差
                    if i == loan_term - 1:
                        principal = remaining_principal
                    
                    schedule.append({
                        'date': payment_date,
                        'principal': principal,
                        'periods': periods
                    })
                    
                    remaining_principal -= principal
                
        elif repayment_method == "等额本金":
            # 每期本金相同
            principal_per_period = loan_amount / loan_term
            
            for i in range(loan_term):
                payment_date = first_payment_date + relativedelta(months=i)
                
                # 根据计算模式决定期数
                if use_integer:
                    periods = i + 1  # 使用整数期数
                else:
                    periods = self._calculate_periods(start_date, payment_date, unit_period)
                
                schedule.append({
                    'date': payment_date,
                    'principal': principal_per_period,
                    'periods': periods
                })
                
        elif repayment_method == "一次性还本":
            # 最后一期还本
            last_payment_date = first_payment_date + relativedelta(months=loan_term-1)
            
            # 根据计算模式决定期数
            if use_integer:
                periods = loan_term  # 使用整数期数
            else:
                periods = self._calculate_periods(start_date, last_payment_date, unit_period)
            
            schedule.append({
                'date': last_payment_date,
                'principal': loan_amount,
                'periods': periods
            })
        
        else:
            # 处理自定义还款方式，默认按等额本金处理
            principal_per_period = loan_amount / loan_term
            
            for i in range(loan_term):
                payment_date = first_payment_date + relativedelta(months=i)
                
                # 根据计算模式决定期数
                if use_integer:
                    periods = i + 1  # 使用整数期数
                else:
                    periods = self._calculate_periods(start_date, payment_date, unit_period)
                
                schedule.append({
                    'date': payment_date,
                    'principal': principal_per_period,
                    'periods': periods
                })
        
        return schedule
    
    def _get_fee_payment_schedule(self, fee_amount, fee_frequency, loan_term, 
                                  start_date, first_payment_date):
        """获取费用支付计划"""
        schedule = []
        
        if fee_frequency == "月":
            # 每月支付
            for i in range(loan_term):
                payment_date = start_date + relativedelta(months=i)
                periods = i  # 月为单位
                
                schedule.append({
                    'date': payment_date,
                    'amount': fee_amount,
                    'periods': periods
                })
                
        elif fee_frequency == "季":
            # 每季度支付
            num_payments = loan_term // 3
            for i in range(num_payments):
                payment_date = start_date + relativedelta(months=i*3)
                periods = i * 3  # 月为单位
                
                schedule.append({
                    'date': payment_date,
                    'amount': fee_amount,
                    'periods': periods
                })
                
        elif fee_frequency == "年":
            # 每年支付
            num_payments = loan_term // 12
            for i in range(num_payments):
                payment_date = start_date + relativedelta(months=i*12)
                periods = i * 12  # 月为单位
                
                schedule.append({
                    'date': payment_date,
                    'amount': fee_amount,
                    'periods': periods
                })
        
        return schedule
    
    def _calculate_periods(self, start_date, end_date, unit_period):
        """计算两个日期之间的期数"""
        # 计算天数差
        days_diff = (end_date - start_date).days
        
        if unit_period == 1:  # 月
            # 计算月数差
            years_diff = end_date.year - start_date.year
            months_diff = end_date.month - start_date.month
            days_adjust = (end_date.day - start_date.day) / 30.0
            return years_diff * 12 + months_diff + days_adjust
        elif unit_period == 12:  # 年
            return days_diff / 360.0
        else:
            # 其他情况按天数比例计算
            return days_diff / (unit_period * 30.0)
    
    def calculate_loan_cash_flows(self, loan_amount, repayment_method, loan_term, 
                                 interest_frequency, interest_rate, 
                                 start_date, end_date, first_payment_date):
        """计算贷款本息现金流（保留用于显示）"""
        cash_flows = defaultdict(float)
        
        # 贷款发放，现金流入为正
        cash_flows[start_date] = loan_amount
        
        # 根据还款方式计算还款现金流
        if repayment_method == "等额本息":
            # 计算每期还款金额
            # 月利率
            monthly_rate = interest_rate / 12
            # 等额本息每月还款金额 = 本金 × 月利率 × (1+月利率)^贷款期限 / [(1+月利率)^贷款期限 - 1]
            if monthly_rate == 0:
                monthly_payment = loan_amount / loan_term
            else:
                monthly_payment = loan_amount * monthly_rate * (1 + monthly_rate) ** loan_term / ((1 + monthly_rate) ** loan_term - 1)
            
            # 根据付息频率确定还款日期
            payment_dates = self.generate_payment_dates(first_payment_date, loan_term, self.frequency_periods[interest_frequency])
            
            # 每期还款
            for payment_date in payment_dates:
                cash_flows[payment_date] -= monthly_payment
                
        elif repayment_method == "等额本金":
            # 每期本金 = 总本金 / 期数
            principal_per_period = loan_amount / loan_term
            
            # 根据付息频率确定还款日期
            payment_dates = self.generate_payment_dates(first_payment_date, loan_term, self.frequency_periods[interest_frequency])
            
            # 计算每期利息和本金
            remaining_principal = loan_amount
            for i, payment_date in enumerate(payment_dates):
                # 该期本金
                period_principal = principal_per_period
                # 该期利息 = 剩余本金 × 月利率
                period_interest = remaining_principal * (interest_rate / 12)
                
                # 该期还款 = 本金 + 利息
                payment = period_principal + period_interest
                cash_flows[payment_date] -= payment
                
                # 更新剩余本金
                remaining_principal -= period_principal
                
        elif repayment_method == "一次性还本":
            # 根据付息频率计算利息支付日期
            interest_dates = self.generate_payment_dates(first_payment_date, loan_term, self.frequency_periods[interest_frequency])
            
            # 计算每期利息
            period_interest = loan_amount * (interest_rate / 12)
            
            # 利息支付
            for interest_date in interest_dates:
                cash_flows[interest_date] -= period_interest
            
            # 最后一期还本金
            cash_flows[end_date] -= loan_amount
        
        return cash_flows
    
    def calculate_fee_cash_flows(self, fee_amount, fee_frequency, loan_term, start_date):
        """计算费用现金流"""
        cash_flows = defaultdict(float)
        
        if fee_frequency == "期初一次性付费":
            # 期初一次性支付
            cash_flows[start_date] -= fee_amount
        else:
            # 周期性支付
            payments_per_year = self.fee_frequency_per_year[fee_frequency]
            payment_interval = 12 / payments_per_year  # 支付间隔(月)
            
            # 计算总支付次数
            total_payments = int(loan_term / payment_interval)
            
            # 每次支付的金额
            payment_amount = fee_amount / total_payments if total_payments > 0 else fee_amount
            
            # 生成支付日期
            for i in range(total_payments):
                payment_date = start_date + relativedelta(months=int(i * payment_interval))
                cash_flows[payment_date] -= payment_amount
        
        return cash_flows
    
    def calculate_fee_annual_rate(self, fee_amount, fee_frequency, loan_amount, loan_term):
        """计算费用的简单年化利率(不建议使用，保留作为备用)"""
        if fee_frequency == "期初一次性付费":
            # 一次性费用年化 = 费用总额 / 贷款金额 / (贷款期限/12)
            return fee_amount / loan_amount / (loan_term / 12)
        else:
            # 周期性费用年化 = 年支付总额 / 贷款金额
            payments_per_year = self.fee_frequency_per_year[fee_frequency]
            annual_fee = (fee_amount * payments_per_year) if loan_term >= 12 else (fee_amount * payments_per_year * loan_term / 12)
            return annual_fee / loan_amount
    
    def calculate_irr(self, cash_flows):
        """计算内部收益率(月度)"""
        # 将现金流按日期排序
        sorted_dates = sorted(cash_flows.keys())
        amounts = [cash_flows[date] for date in sorted_dates]
        
        # 计算每个现金流距离起始日的月数（使用真实月份计算）
        base_date = sorted_dates[0]
        periods = []
        for date in sorted_dates:
            # 计算两个日期之间的月份差
            years_diff = date.year - base_date.year
            months_diff = date.month - base_date.month
            days_adjust = 0
            # 日期的天数调整
            if date.day < base_date.day:
                days_adjust = -((base_date.day - date.day) / 30.0)
            elif date.day > base_date.day:
                days_adjust = (date.day - base_date.day) / 30.0
            
            total_months = years_diff * 12 + months_diff + days_adjust
            periods.append(total_months)
        
        # 使用numpy计算IRR(月度)
        try:
            monthly_irr = np.irr(amounts)
            return monthly_irr
        except:
            # 如果IRR计算失败，使用近似值
            return self.approximate_irr(periods, amounts)
    
    def approximate_irr(self, periods, amounts):
        """使用近似方法计算IRR"""
        # 假设初始IRR为0
        guess = 0.01
        
        # 迭代次数
        max_iterations = 100
        tolerance = 1e-6
        
        for _ in range(max_iterations):
            npv = 0
            derivative = 0
            
            for i, (t, cf) in enumerate(zip(periods, amounts)):
                npv += cf / ((1 + guess) ** t)
                derivative -= t * cf / ((1 + guess) ** (t + 1))
            
            # 判断是否收敛
            if abs(npv) < tolerance:
                return guess
            
            # 牛顿迭代
            new_guess = guess - npv / derivative if derivative != 0 else guess * 1.1
            
            # 检查迭代是否收敛
            if abs(new_guess - guess) < tolerance:
                return new_guess
            
            guess = new_guess
        
        # 如果没有收敛，返回最后的猜测值
        return guess
    
    def generate_payment_dates(self, first_payment_date, loan_term, period_months):
        """生成还款日期列表"""
        payment_dates = []
        current_date = first_payment_date
        
        # 计算总期数
        total_periods = int(loan_term / period_months) if period_months >= 1 else loan_term
        
        for i in range(total_periods):
            payment_dates.append(current_date)
            # 使用relativedelta进行月份对日计算
            if period_months >= 1:
                months_to_add = int(period_months)
                current_date = current_date + relativedelta(months=months_to_add)
            else:
                # 处理日频率
                days_to_add = int(period_months * 30)
                current_date = current_date + relativedelta(days=days_to_add)
        
        return payment_dates 