import pandas as pd
import re

# 自定义科目判断模块
class AccountClassifier:
    @staticmethod
    def classify(row):
        # 判断对方账户名称和附言
        counterpart_name = str(row['对方账户名称'])
        remarks = str(row['附言'])
        
        # 判断税费相关（示例规则，可扩展）
        if any(keyword in remarks for keyword in ["税", "TIPS", "代扣"]):
            return "税费"
        # 判断个人/单位
        elif "公司" in counterpart_name or "有限" in counterpart_name or "股份" in counterpart_name:
            return "其他往来-单位"
        elif "个人" in remarks or ("费用报销" in remarks and len(counterpart_name.strip()) < 4):
            return "其他往来-个人"
        else:
            return "其他往来-单位"  # 默认值

# 读取Excel并自动定位标题行
def find_header(df):
    for idx, row in df.iterrows():
        if '交易日期' in row.values and '借方发生额' in row.values:
            return idx
    return 0  # 若找不到则默认第一行为标题

# 数据清洗主函数
def clean_data(file_path):
    # 读取原始数据
    raw_df = pd.read_excel(file_path, header=None)
    header_row = find_header(raw_df)
    
    # 重新读取数据并设置标题
    df = pd.read_excel(file_path, header=header_row)
    
    # 清理列名中的特殊字符和空格
    df.columns = [col.strip() for col in df.columns.astype(str)]
    
    # 筛选指定列
    required_cols = ['交易日期', '交易时间', '对方账号', '交易账号', '单位名称', 
                    '对方账户名称', '借方发生额', '贷方发生额', '摘要', '附言']
    df = df[required_cols]
    
    # 转换金额列为数值类型（处理千分位逗号）
    df['借方发生额'] = df['借方发生额'].replace('[,-]', '', regex=True).astype(float)
    df['贷方发生额'] = df['贷方发生额'].replace('[,-]', '', regex=True).astype(float)
    
    # 新增科目列
    df['借方科目名称'] = ''
    df['贷方科目名称'] = ''
    
    # 填充科目列
    for idx, row in df.iterrows():
        if row['借方发生额'] > 0:
            df.at[idx, '借方科目名称'] = '银行存款'
            df.at[idx, '贷方科目名称'] = AccountClassifier.classify(row)
        elif row['贷方发生额'] > 0:
            df.at[idx, '贷方科目名称'] = '银行存款'
            df.at[idx, '借方科目名称'] = AccountClassifier.classify(row)
    
    # 按指定字段排序
    df = df.sort_values(by=['单位名称', '交易账号', '交易日期', '交易时间'])
    
    return df

# 执行清洗并保存
cleaned_df = clean_data(r"C:\Users\nihha\Downloads\生物医药4136.xlsx")
cleaned_df.to_excel("清洗后数据.xlsx", index=False)