import sqlite3
import json
import os
import uuid  # 添加uuid导入

class RecordManager:
    def __init__(self, db_file):
        """初始化数据库管理器"""
        self.db_file = db_file
        self.init_database()
    
    def init_database(self):
        """初始化数据库结构"""
        # 如果数据库文件不存在，创建表结构
        is_new_db = not os.path.exists(self.db_file)
        conn = sqlite3.connect(self.db_file)
        conn.execute("PRAGMA foreign_keys = ON")
        cursor = conn.cursor()
        
        if is_new_db:
            # 创建记录表 - 全新创建
            cursor.execute('''
            CREATE TABLE IF NOT EXISTS finance_records (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                uuid TEXT UNIQUE,
                company_name TEXT,
                loan_amount REAL,
                repayment_method TEXT,
                loan_term INTEGER,
                interest_frequency TEXT,
                start_date TEXT,
                end_date TEXT,
                first_payment_date TEXT,
                interest_rate REAL,
                total_cost REAL,
                loan_channel TEXT,
                customer_type TEXT,
                company_nature TEXT,
                guarantee_type TEXT,
                loan_type TEXT,
                application_method TEXT,
                is_subsidized INTEGER,
                create_time TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
            ''')
            
            # 创建费用表 - 全新创建
            cursor.execute('''
            CREATE TABLE IF NOT EXISTS finance_fees (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                record_id INTEGER,
                name TEXT,
                amount REAL,
                frequency TEXT,
                is_bank_bearing INTEGER DEFAULT 0,
                FOREIGN KEY (record_id) REFERENCES finance_records (id) ON DELETE CASCADE
            )
            ''')
        else:
            # 检查已有表结构，添加缺失的列
            # 检查finance_records表
            cursor.execute("PRAGMA table_info(finance_records)")
            columns = {row[1] for row in cursor.fetchall()}
            
            # 需要添加的新列
            new_columns = {
                "uuid": "TEXT",  # 移除UNIQUE约束，因为ALTER TABLE不能添加UNIQUE列
                "loan_channel": "TEXT",
                "customer_type": "TEXT",
                "company_nature": "TEXT", 
                "guarantee_type": "TEXT",
                "loan_type": "TEXT",
                "application_method": "TEXT",
                "is_subsidized": "INTEGER"
            }
            
            # 添加缺失的列
            for col_name, col_type in new_columns.items():
                if col_name not in columns:
                    cursor.execute(f"ALTER TABLE finance_records ADD COLUMN {col_name} {col_type}")
            
            # 为现有记录生成UUID（如果还没有的话）
            cursor.execute("SELECT id FROM finance_records WHERE uuid IS NULL OR uuid = ''")
            for row in cursor.fetchall():
                record_id = row[0]
                new_uuid = str(uuid.uuid4())
                cursor.execute("UPDATE finance_records SET uuid = ? WHERE id = ?", (new_uuid, record_id))
            
            # 检查finance_fees表
            cursor.execute("PRAGMA table_info(finance_fees)")
            fee_columns = {row[1] for row in cursor.fetchall()}
            
            # 为费用表添加是否银行承担字段
            if "is_bank_bearing" not in fee_columns:
                cursor.execute("ALTER TABLE finance_fees ADD COLUMN is_bank_bearing INTEGER DEFAULT 0")
        
        conn.commit()
        conn.close()
    
    def add_record(self, company_name, loan_amount, repayment_method, loan_term,
                  interest_frequency, start_date, end_date, first_payment_date,
                  interest_rate, total_cost, fees, loan_channel="", customer_type="",
                  company_nature="", guarantee_type="", loan_type="", 
                  application_method="", is_subsidized=0):
        """添加新记录"""
        conn = sqlite3.connect(self.db_file)
        conn.execute("PRAGMA foreign_keys = ON")
        cursor = conn.cursor()
        
        try:
            # 生成UUID
            record_uuid = str(uuid.uuid4())
            
            # 插入主记录
            cursor.execute('''
            INSERT INTO finance_records (
                uuid, company_name, loan_amount, repayment_method, loan_term,
                interest_frequency, start_date, end_date, first_payment_date,
                interest_rate, total_cost, loan_channel, customer_type,
                company_nature, guarantee_type, loan_type, application_method,
                is_subsidized
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                record_uuid, company_name, loan_amount, repayment_method, loan_term,
                interest_frequency, start_date, end_date, first_payment_date,
                interest_rate, total_cost, loan_channel, customer_type,
                company_nature, guarantee_type, loan_type, application_method,
                is_subsidized
            ))
            
            # 获取新记录的ID
            record_id = cursor.lastrowid
            
            # 插入费用记录
            for fee in fees:
                # 检查是否有银行承担字段
                is_bank_bearing = fee.get("is_bank_bearing", 0)
                
                cursor.execute('''
                INSERT INTO finance_fees (record_id, name, amount, frequency, is_bank_bearing)
                VALUES (?, ?, ?, ?, ?)
                ''', (record_id, fee["name"], fee["amount"], fee["frequency"], is_bank_bearing))
            
            conn.commit()
            return record_id
            
        except Exception as e:
            conn.rollback()
            raise e
        finally:
            conn.close()
    
    def update_record(self, record_id, company_name, loan_amount, repayment_method, loan_term,
                     interest_frequency, start_date, end_date, first_payment_date,
                     interest_rate, total_cost, fees, loan_channel="", customer_type="",
                     company_nature="", guarantee_type="", loan_type="", 
                     application_method="", is_subsidized=0):
        """更新记录"""
        conn = sqlite3.connect(self.db_file)
        conn.execute("PRAGMA foreign_keys = ON")
        cursor = conn.cursor()
        
        try:
            # 更新主记录
            cursor.execute('''
            UPDATE finance_records SET
                company_name = ?, loan_amount = ?, repayment_method = ?, loan_term = ?,
                interest_frequency = ?, start_date = ?, end_date = ?, first_payment_date = ?,
                interest_rate = ?, total_cost = ?, loan_channel = ?, customer_type = ?,
                company_nature = ?, guarantee_type = ?, loan_type = ?, application_method = ?,
                is_subsidized = ?
            WHERE id = ?
            ''', (
                company_name, loan_amount, repayment_method, loan_term,
                interest_frequency, start_date, end_date, first_payment_date,
                interest_rate, total_cost, loan_channel, customer_type,
                company_nature, guarantee_type, loan_type, application_method,
                is_subsidized, record_id
            ))
            
            # 删除旧的费用记录
            cursor.execute('DELETE FROM finance_fees WHERE record_id = ?', (record_id,))
            
            # 插入新的费用记录
            for fee in fees:
                # 检查是否有银行承担字段
                is_bank_bearing = fee.get("is_bank_bearing", 0)
                
                cursor.execute('''
                INSERT INTO finance_fees (record_id, name, amount, frequency, is_bank_bearing)
                VALUES (?, ?, ?, ?, ?)
                ''', (record_id, fee["name"], fee["amount"], fee["frequency"], is_bank_bearing))
            
            conn.commit()
            
        except Exception as e:
            conn.rollback()
            raise e
        finally:
            conn.close()
    
    def delete_record(self, record_id):
        """删除记录"""
        conn = sqlite3.connect(self.db_file)
        conn.execute("PRAGMA foreign_keys = ON")
        cursor = conn.cursor()
        
        try:
            # 删除主记录，费用记录会通过外键级联删除
            cursor.execute('DELETE FROM finance_records WHERE id = ?', (record_id,))
            conn.commit()
            
        except Exception as e:
            conn.rollback()
            raise e
        finally:
            conn.close()
    
    def get_record(self, record_id):
        """获取单条记录"""
        conn = sqlite3.connect(self.db_file)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        
        try:
            # 查询主记录
            cursor.execute('''
            SELECT * FROM finance_records WHERE id = ?
            ''', (record_id,))
            
            record = cursor.fetchone()
            
            if record:
                # 转换为字典
                record_dict = dict(record)
                
                # 查询关联的费用记录
                cursor.execute('''
                SELECT * FROM finance_fees WHERE record_id = ?
                ''', (record_id,))
                
                fees = [dict(fee) for fee in cursor.fetchall()]
                record_dict["fees"] = fees
                
                return record_dict
            else:
                return None
            
        finally:
            conn.close()
    
    def get_all_records(self):
        """获取所有记录"""
        conn = sqlite3.connect(self.db_file)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        
        try:
            # 查询所有主记录
            cursor.execute('''
            SELECT * FROM finance_records ORDER BY id DESC
            ''')
            
            records = [dict(record) for record in cursor.fetchall()]
            
            # 查询每条记录关联的费用
            for record in records:
                cursor.execute('''
                SELECT * FROM finance_fees WHERE record_id = ?
                ''', (record["id"],))
                
                fees = [dict(fee) for fee in cursor.fetchall()]
                record["fees"] = fees
            
            return records
            
        finally:
            conn.close()
    
    def get_fees(self, record_id):
        """获取指定记录的费用项"""
        conn = sqlite3.connect(self.db_file)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        
        try:
            cursor.execute('''
            SELECT * FROM finance_fees WHERE record_id = ?
            ''', (record_id,))
            
            fees = [dict(fee) for fee in cursor.fetchall()]
            return fees
            
        finally:
            conn.close() 