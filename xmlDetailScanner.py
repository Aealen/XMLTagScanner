# -*- coding: utf-8 -*-
import os
import xml.etree.ElementTree as ET
import pandas as pd

def scan_xml_files(directory):
    xml_files = []
    for root_dir, _, files in os.walk(directory):
        for f in files:
            if f.endswith('.xml'):
                xml_files.append(os.path.join(root_dir, f))
    total_xml = len(xml_files)
    result = {}

    for file_path in xml_files:
        xml_file = os.path.relpath(file_path, directory)
        try:
            tree = ET.parse(file_path)
            root = tree.getroot()
            tags_count = {}
            for elem in root:
                tag = elem.tag
                tags_count[tag] = tags_count.get(tag, 0) + 1
            result[xml_file] = tags_count
        except Exception as e:
            print("解析文件 {} 时出错: {}".format(xml_file, e))

    return total_xml, result

def print_result(total_xml, result):
    print("目录下XML文件总数: {}".format(total_xml))
    # for file in sorted(result.keys()):
    #     tags = result[file]
    #     print("\n文件: {}".format(file))
    #     for tag, count in tags.items():
    #         print("  {} 标签数量: {}".format(tag, count))

def save_to_excel(result, output_path):
    # 收集所有可能出现的标签
    all_tags = set()
    for tags in result.values():
        all_tags.update(tags.keys())
    all_tags = sorted(all_tags)
    # 将结果转换为DataFrame
    data = []
    for file in sorted(result.keys()):
        tags = result[file]
        row = {'文件名': file}
        for tag in all_tags:
            row[tag] = tags.get(tag, 0)
        row['合计'] = sum(row[tag] for tag in all_tags)
        data.append(row)
    df = pd.DataFrame(data)
    # 确保文件名和合计为前两列
    cols = ['文件名', '合计'] + [col for col in df.columns if col not in ('文件名', '合计')]
    df = df[cols]
    # 使用ExcelWriter自适应列宽
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
        worksheet = writer.sheets['Sheet1']
        for i, col in enumerate(df.columns):
            max_length = max(
                df[col].astype(str).map(len).max(),
                len(str(col))
            )
            worksheet.column_dimensions[chr(65 + i)].width = max_length + 2
    print("统计结果已保存到Excel文件：{}".format(output_path))

if __name__ == "__main__":
    # directory = input("请输入要扫描的目录路径: ")
    # xml所在目录
    directory = r"C:\yourFilesDirPath"
    total_xml, result = scan_xml_files(directory)
    print_result(total_xml, result)
    save_to_excel(result, "xmlDetailResult.xlsx")
