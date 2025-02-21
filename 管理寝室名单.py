import json
import os

def load_json(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        data = json.load(file)
    return data

def save_json(file_path, data):
    with open(file_path, 'w', encoding='utf-8') as file:
        json.dump(data, file, ensure_ascii=False, indent=4)

def validate_json(data):
    if not isinstance(data, dict):
        return False
    for key, value in data.items():
        if not isinstance(value, list):
            return False
    return True

def display_json(data):
    for dorm, members in data.items():
        print(f"寝室号：{dorm}\t成员：{', '.join(members)}")

def delete_node(data):
    dorm = input("请输入要删除的寝室号：")
    if dorm in data:
        del data[dorm]
        print(f"寝室号 {dorm} 已删除。")
    else:
        print(f"寝室号 {dorm} 不存在。")
    return data

def add_node(data):
    user_input = input("输入格式为 寝室号 成员名1 成员名2 ...：")
    parts = user_input.split(maxsplit=1)
    if len(parts) == 2:
        dorm = parts[0]
        members = parts[1].split()
        data[dorm] = members
        print(f"寝室号 {dorm} 已添加。")
    else:
        print("输入格式错误。")
    return data

def main():
    file_path = os.path.join(os.getcwd(), 'mingdan.json')
    
    if not os.path.exists(file_path):
        initial_data = {
            "410": ["陈", "沈"]
        }
        save_json(file_path, initial_data)
        print(f"文件 {file_path} 不存在，已创建并写入初始内容。")
    
    data = load_json(file_path)
    
    if not validate_json(data):
        print("json格式错误")
        return
    
    while True:
        print("\n当前寝室名单：")
        display_json(data)
        print("\n操作选项：\n1. 删除寝室\n2. 新增寝室\n3. 退出")
        choice = input("请选择操作：")
        
        if choice == '1':
            data = delete_node(data)
        elif choice == '2':
            data = add_node(data)
        elif choice == '3':
            save_json(file_path, data)
            print("已保存并退出。")
            break
        else:
            print("无效的选项，请重新选择。")

if __name__ == "__main__":
    main()