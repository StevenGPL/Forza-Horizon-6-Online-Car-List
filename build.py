import openpyxl
import json
import os

def build_json_from_excel(excel_path, output_json_path):
    print("开始读取 Excel 文件...")
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    sheet = wb.active

    # 获取表头作为 JSON 的 Key
    headers = [cell.value for cell in sheet[1]]
    
    cars_data = []
    
    # 从第二行开始遍历数据
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if not row[0]: # 如果 ID 为空则跳过该行
            continue
            
        car_dict = dict(zip(headers, row))
        
        # 1. 自动生成全名 (品牌 + 车名 + 年份)
        # 注意：有时候车名里已经包含了年份，可以根据你的实际情况调整拼接逻辑
        car_dict['full_name'] = f"{car_dict['brand']} {car_dict['name']}"
        
        # 2. 自动绑定图片路径 (基于 ID)
        car_dict['image'] = f"./CarImages/{car_dict['id']}.webp"
        
        # 3. 处理富文本：将特定的获取方式标红
        if car_dict['how_to_get']:
            how_to_get_str = str(car_dict['how_to_get'])
            if "季节赛" in how_to_get_str:
                car_dict['how_to_get'] = how_to_get_str.replace("季节赛", "<span class='text-red-500'>季节赛</span>")
            # 如果还有其他需要变色的关键词，可以在这里继续加 elif
            
        # 4. 清理空值并确保数字类型正确
        for k, v in car_dict.items():
            if v is None:
                car_dict[k] = ""
                
        cars_data.append(car_dict)

    # 导出为 JSON
    with open(output_json_path, 'w', encoding='utf-8') as f:
        json.dump(cars_data, f, ensure_ascii=False, indent=2)
        
    print(f"成功生成 JSON 文件！共 {len(cars_data)} 辆车。")

if __name__ == "__main__":
    build_json_from_excel("FH_Cars.xlsx", "Cars.json")