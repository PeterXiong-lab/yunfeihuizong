import os
import re
import pandas as pd
from openpyxl.styles import Alignment

def process_interactive():
    user_input = input("请输入要提取的txt文件名（多个文件请用逗号隔开）：\n> ")
    user_input = user_input.replace('，', ',')
    filenames = [f.strip() for f in user_input.split(',')]
    
    all_dataframes = []
    
    for filename in filenames:
        if not filename:
            continue
            
        if not os.path.exists(filename):
            print(f"⚠️ 找不到文件 '{filename}'，已跳过。")
            continue
            
        date_match = re.search(r'物流费用提取结果(.*?)\.txt', filename)
        if not date_match:
            print(f"⚠️ 无法从 '{filename}' 提取日期作为表头，请检查命名格式。")
            continue
            
        date_col = date_match.group(1)
        
        # 兼容不同编码
        content = ""
        for enc in ['utf-8', 'gbk', 'gb2312']:
            try:
                with open(filename, 'r', encoding=enc) as f:
                    content = f.read()
                break
            except UnicodeDecodeError:
                continue
                
        # 【修改点】以 "PO-" 作为订单的分割符，这样比按换行符分割更稳定
        orders = re.split(r'PO-', content)
        # 去除空字符串
        orders = [o for o in orders if o.strip()]
        
        data = []
        for order in orders:
            # 【修改点】把换行符全部替换为空格，消除排版差异导致的正则匹配失败
            order_text = order.replace('\n', ' ')
            
            # 正则匹配仓库名和费用（增加 \s* 以兼容可能存在的空格）
            city_match = re.search(r'天猫美团(.*?)仓', order_text)
            fees_match = re.search(r'快递\s*([\d.]+)\s*物流\s*([\d.]+)', order_text)
            
            if city_match and fees_match:
                city = city_match.group(1).strip()
                express = float(fees_match.group(1))
                logistics = float(fees_match.group(2))
                
                data.append({'仓库': city, '货运方式': '快递', date_col: express})
                data.append({'仓库': city, '货运方式': '物流', date_col: logistics})
                
        if data:
            df = pd.DataFrame(data)
            df.set_index(['仓库', '货运方式'], inplace=True)
            all_dataframes.append(df)
            print(f"✅ {filename}: 成功提取 {len(data)//2} 个仓位的数据。")
        else:
            print(f"⚠️ {filename}: 未能提取到任何数据！")
            
    if not all_dataframes:
        print("\n❌ 没有提取到任何有效数据，程序退出！")
        return

    # 横向合并所有表
    final_df = pd.concat(all_dataframes, axis=1, join='outer')
    
    # 按照列名数字大小进行升序排序，确保 5.18 在 5.19 前面
    try:
        sorted_cols = sorted(final_df.columns, key=lambda x: float(x))
        final_df = final_df[sorted_cols]
    except ValueError:
        sorted_cols = sorted(final_df.columns)
        final_df = final_df[sorted_cols]

    # 导出 Excel 
    output_excel = "动态物流费用合并结果.xlsx"
    with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
        final_df.to_excel(writer, sheet_name='费用汇总')
        worksheet = writer.sheets['费用汇总']
        
        for col in worksheet.columns:
            column = col[0].column_letter
            worksheet.column_dimensions[column].width = 14
            for cell in col:
                cell.alignment = Alignment(horizontal='center', vertical='center')
                
    print(f"\n🎉 处理完成！已按日期排好序，文件保存为: {output_excel}")

if __name__ == "__main__":
    process_interactive()
