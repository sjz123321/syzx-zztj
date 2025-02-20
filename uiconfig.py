import tkinter as tk
from tkinter import ttk
import configparser
import os
import subprocess

def initialize_config():
    """初始化配置文件，如果不存在则创建"""
    config = configparser.ConfigParser()
    if not os.path.exists('config.ini'):
        config['Settings'] = {
            'semester': '1',
            'week': '1'
        }
        with open('config.ini', 'w', encoding='utf-8') as f:
            config.write(f)
    else:
        config.read('config.ini')
        if 'Settings' not in config:
            config['Settings'] = {
                'semester': '1',
                'week': '1'
            }
            with open('config.ini', 'w', encoding='utf-8') as f:
                config.write(f)
    return config['Settings'].get('semester', '1'), config['Settings'].get('week', '1')

class ConfigApp:
    def __init__(self, root, initial_semester, initial_week):
        self.root = root
        self.root.title("配置工具")
        
        # 创建主框架
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 学期输入框
        ttk.Label(main_frame, text="学期:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.semester_var = tk.StringVar(value=initial_semester)
        self.semester_entry = ttk.Entry(main_frame, textvariable=self.semester_var)
        self.semester_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        
        # 周次输入框
        ttk.Label(main_frame, text="周次:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.week_var = tk.StringVar(value=initial_week)
        self.week_entry = ttk.Entry(main_frame, textvariable=self.week_var)
        self.week_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        
        # 保存按钮
        ttk.Button(main_frame, text="保存并运行", command=self.save_and_run).grid(row=2, column=0, columnspan=2, pady=10)
            
    def save_and_run(self):
        semester = self.semester_var.get()
        week = self.week_var.get()
        
        # 保存到config.ini
        config = configparser.ConfigParser()
        config['Settings'] = {
            'semester': semester,
            'week': week
        }
        with open('config.ini', 'w', encoding='utf-8') as f:
            config.write(f)
            
        # 删除文件
        try:
            html_file = f'response-{semester}-{week}.html'
            if os.path.exists(html_file):
                os.remove(html_file)
                
            excel_file = f'{semester}学期-{week}周.xlsx'
            if os.path.exists(excel_file):
                os.remove(excel_file)
        except Exception as e:
            print(f"删除文件时出错: {e}")
            
        # 运行bat文件并传入参数
        try:
            subprocess.run(['get_html.sh', semester, week], shell=True)
            ##subprocess.run(['get_html.bat', semester, week], shell=True)
            subprocess.run(['parse_table.py', f'response{semester}-{week}.html',semester, week], shell=True)
            ##subprocess.run(['parse_table.exe', f'response{semester}-{week}.html',semester, week], shell=True)
            subprocess.run(['excel_processor2.py'], shell=True)
            ##subprocess.run(['excel_processor2.exe'], shell=True)
        except Exception as e:
            print(f"运行bat文件时出错: {e}")

def main():
    # 初始化配置并获取初始值
    initial_semester, initial_week = initialize_config()
    
    root = tk.Tk()
    app = ConfigApp(root, initial_semester, initial_week)
    root.mainloop()

if __name__ == "__main__":
    main()
