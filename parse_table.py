import pandas as pd
from bs4 import BeautifulSoup
import sys
from pathlib import Path

def parse_html_table(html_file, term, week):
    # 读取HTML文件
    with open(html_file, 'r', encoding='gbk') as f:
        content = f.read()
    
    # 使用BeautifulSoup解析HTML
    soup = BeautifulSoup(content, 'html.parser')
    
    # 找到所有包含 class="main1" 的TD元素
    main_tds = soup.find_all('td', class_='main1')
    if not main_tds or len(main_tds) < 2:
        print("未找到足够的class='main1'的TD元素")
        return False
    
    # 获取第二个TD元素所在的表格
    second_td = main_tds[1]
    table = second_td.find_parent('table')
    if not table:
        print("未找到相关表格")
        return False
    
    # 解析表格数据
    table_data = []
    rows = table.find_all('tr')
    
    for row in rows:
        # 获取所有单元格
        cols = row.find_all(['td', 'th'])
        # 提取文本并去除空白
        cols = [col.text.strip() for col in cols]
        if cols:  # 如果行不为空
            # 过滤掉空列表和只包含空字符串的列表
            if any(col for col in cols if col):
                table_data.append(cols)
    
    if table_data:
        # 创建DataFrame
        df = pd.DataFrame(table_data)
        
        # 如果第一行看起来像表头，将其设置为列名
        if len(df) > 0:
            df.columns = df.iloc[0]
            df = df.iloc[1:]
            # 重置索引
            df = df.reset_index(drop=True)
            
            # 删除前两行和最后一列
            df = df.iloc[1:]  # 删除前两行
            df = df.iloc[:, :-1]  # 删除最后一列
            
            # 重置索引
            df = df.reset_index(drop=True)
        
        # 创建Excel文件名
        excel_file = f"{term}学期-{week}周.xlsx"
        
        # 保存初始数据
        df.to_excel(excel_file, index=False)
        
        # 重新打开文件并进行二次处理
        df_second = pd.read_excel(excel_file)
        
        # 删除第一行和最后一列
        df_second = df_second.iloc[:, :-2]  # 删除最后一列
        
        # 重置索引
        df_second = df_second.reset_index(drop=True)
        
        # 再次保存处理后的数据
        df_second.to_excel(excel_file, index=False)
        
        print(f"数据已保存到文件: {excel_file}")
        print(f"共处理 {len(df_second)} 行数据")
        return True
    else:
        print("表格中没有数据")
        return False

if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("使用方法: python parse_table.py <html文件> <学期> <周次>")
        sys.exit(1)
    
    html_file = sys.argv[1]
    term = sys.argv[2]
    week = sys.argv[3]
    
    if not Path(html_file).exists():
        print(f"错误: 文件 {html_file} 不存在")
        sys.exit(1)
    
    parse_html_table(html_file, term, week)