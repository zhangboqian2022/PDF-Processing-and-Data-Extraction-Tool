
'''
THIS FILE IS PART OF Utour United Tourism Cruise Division PROJECT
Description: This program is designed to process PDF files and Excel configuration files. 
Its main functionality is to copy matching PDF files from a source folder to a destination 
folder based on specific codes found in a designated column header within the Excel files. 
Additionally, the program keeps a record of PDF files that could not be matched. Furthermore, 
the program processes room number information from the Excel files and generates statistics on room occupancy.

功能描述：该程序旨在处理PDF文件和Excel配置文件。其主要功能是根据Excel文件中指定列标题中的特定代码，
将匹配的PDF文件从源文件夹复制到目标文件夹。此外，该程序还会记录无法匹配的PDF文件。此外，
该程序还会处理来自Excel文件的房间号信息，并生成有关房间占用情况的统计信息。

Author: Zhang BoQian (Daniel)
Date: 2024-02-1 13:20:02
LastEditors: Daniel
LastEditTime: 2024-2-6 20:40:00

'''

# 导入必要的库
# Import necessary libraries
import os
import shutil
import pandas as pd
from datetime import datetime

# 定义一个函数，用于在DataFrame中查找包含特定表头的列索引
# Define a function to find the column index containing a specific header in a DataFrame
def find_column_index_with_header(df, header):
    for col in df.columns:
        if header in str(col):
            return df.columns.get_loc(col)
    return None

# 定义一个函数，用于从指定的Excel文件和工作表中提取含有特定表头的列的所有非空值
# Define a function to extract all non-empty values from columns with a specific header in a specified Excel file and worksheet
def extract_codes_from_excel(excel_file, sheet_name, header):
    try:
        # 使用pandas读取excel文件的指定工作表
        # Use pandas to read the specified worksheet from the Excel file
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
        # 查找含有特定表头的列的索引
        # Find the index of the column containing the specific header
        column_index = find_column_index_with_header(df, header)
        if column_index is not None:
            # 提取该列的所有非空值，转换为字符串列表
            # Extract all non-empty values from the column and convert them to a list of strings
            codes = df.iloc[:, column_index].dropna().astype(str).tolist()
            print(f"Extracted {len(codes)} codes from {excel_file}")
            return codes
        else:
            print(f"Header '{header}' not found in {excel_file}")
            return []
    except Exception as e:
        print(f"Error reading {excel_file}: {e}")
        return []

# 定义一个函数，用于处理房间号信息并生成房间占用的统计信息
# Define a function to process room number information and generate occupancy statistics
def process_room_numbers(df, excel_file_name, log_file):
    current_room_number = None
    room_occupancy = {}
    total_rooms = 0  # 初始化房间总数计数器

    # 遍历DataFrame的每一行
    # Iterate through each row of the DataFrame
    for index, row in df.iterrows():
        # 检查'房间'列是否有值
        # Check if the '房间' column has a value
        if not pd.isna(row['房间']):
            current_room_number = int(row['房间'])
            # 如果是新房间，则在字典中添加条目
            # If it's a new room, add an entry to the dictionary
            if current_room_number not in room_occupancy:
                room_occupancy[current_room_number] = 0
                total_rooms += 1  # 新房间：房间总数加一
        # 累计当前房间的人数
        # Accumulate the number of occupants for the current room
        if current_room_number is not None:
            room_occupancy[current_room_number] += 1

    # 根据房间占用情况统计房间类型数量
    # Count the number of room types based on room occupancy
    room_summary = {'1-person rooms': 0, '2-person rooms': 0, '3-person rooms': 0, '4-person rooms': 0}
    for occupants in room_occupancy.values():
        if occupants == 1:
            room_summary['1-person rooms'] += 1
        elif occupants == 2:
            room_summary['2-person rooms'] += 1
        elif occupants == 3:
            room_summary['3-person rooms'] += 1
        elif occupants == 4:
            room_summary['4-person rooms'] += 1

    # 将房间占用摘要写入日志文件
    # Write room occupancy summary to the log file
    log_content = f"Room Occupancy Summary for {excel_file_name}:\n"
    log_content += f"{total_rooms} rooms\n"  # 记录房间总数
    log_content += "\n".join([f"{key}: {value}" for key, value in room_summary.items()])
    log_file.write(log_content + "\n\n")

# 定义主函数，用于处理PDF文件的复制和未识别PDF文件的处理
# Define the main function to handle the copying of PDF files and processing of unrecognized PDF files
def copy_pdf_if_contains_code(source_folder, config_folder, sheet_name, header):
    # 生成日志文件名并创建日志文件
    # Generate log file name and create log file
    log_filename = datetime.now().strftime("%Y%m%d-%H%M%S-GroupList.log")
    log_filepath = os.path.join(config_folder, log_filename)
    unrecognized_pdf_path = os.path.join('c:\\zbq\\unrecognized')  # 未识别PDF文件的存储路径

    # 如果未识别PDF文件的存储路径不存在，则创建该路径
    # Create the storage path for unrecognized PDF files if it does not exist
    if not os.path.exists(unrecognized_pdf_path):
        os.makedirs(unrecognized_pdf_path)

    unrecognized_count = 0  # 初始化未识别PDF文件计数器

    with open(log_filepath, 'w', encoding='utf-8') as log_file:
        log_file.write(f"Log started on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")

        # 遍历配置文件夹中的所有Excel文件
        # Iterate through all Excel files in the configuration folder
        for excel_file in os.listdir(config_folder):
            if excel_file.startswith('~$') or not excel_file.endswith('.xlsx'):
                continue

            excel_path = os.path.join(config_folder, excel_file)
            df = pd.read_excel(excel_path)
            # 处理房间号并生成统计信息
            # Process room numbers and generate statistics
            process_room_numbers(df, excel_file, log_file)

        pdf_files = set(os.listdir(source_folder))  # 获取源文件夹中所有PDF文件的集合

        # 再次遍历配置文件夹中的所有Excel文件
        # Iterate through all Excel files in the configuration folder again
        for excel_file in os.listdir(config_folder):
            if excel_file.startswith('~$') or not excel_file.endswith('.xlsx'):
                continue

            excel_path = os.path.join(config_folder, excel_file)
            destination_folder_base = os.path.splitext(excel_file)[0]
            destination_folder = os.path.join(config_folder, destination_folder_base)
            if not os.path.exists(destination_folder):
                os.makedirs(destination_folder)

            # 提取Excel文件中的代码
            # Extract codes from the Excel file
            codes = extract_codes_from_excel(excel_path, sheet_name, header)
            copy_count = 0  # 初始化复制计数器

            # 遍历所有代码，寻找匹配的PDF文件并复制
            # Iterate through all codes, find matching PDF files, and copy them
            for code in codes:
                for file in pdf_files.copy():  # 遍历PDF文件集合的副本
                    if file.endswith('.pdf') and code in file:
                        try:
                            # 复制匹配的PDF文件到目标文件夹
                            # Copy the matching PDF file to the destination folder
                            shutil.copy2(os.path.join(source_folder, file), destination_folder)
                            log_file.write(f"Copied {file} to {destination_folder} for code {code}\n")
                            copy_count += 1
                            pdf_files.remove(file)  # 从集合中移除已处理的PDF文件
                            break
                        except Exception as e:
                            log_file.write(f"Error processing file {file}: {e}\n")

            # 如果复制了PDF文件，则重命名目标文件夹以反映复制的文件数量
            # If PDF files were copied, rename the destination folder to reflect the number of copied files
            if copy_count > 0:
                new_destination_folder = f"{destination_folder}_{copy_count}"
                os.rename(destination_folder, new_destination_folder)
                log_file.write(f"Renamed {destination_folder} to {new_destination_folder}\n")

        # 处理未识别的PDF文件
        # Process unrecognized PDF files
        for file in pdf_files:
            if file.endswith('.pdf'):
                try:
                    # 将未识别的PDF文件复制到指定路径
                    # Copy unrecognized PDF files to the specified path
                    shutil.copy2(os.path.join(source_folder, file), unrecognized_pdf_path)
                    log_file.write(f"Unrecognized PDF {file} copied to {unrecognized_pdf_path}\n")
                    unrecognized_count += 1  # 未识别PDF文件计数器加一
                except Exception as e:
                    log_file.write(f"Error processing unrecognized file {file}: {e}\n")

        # 记录未识别PDF文件的总数
        # Record the total number of unrecognized PDF files
        log_file.write(f"Total unrecognized PDF files: {unrecognized_count}\n")
        log_file.write("Processing completed.\n")

# 程序的入口点
# Main entry point of the program
if __name__ == "__main__":
    source_folder = 'C:\\zbq'  # PDF文件的源目录 / Source directory for PDF files
    config_folder = 'C:\\zbq'  # 配置文件（Excel文件）的目录 / Directory for configuration files (Excel files)

    sheet_name = 'Sheet1'  # Excel工作表名称 / Excel worksheet name
    header = '护照'  # 目标列的表头 / Target column header

    # 调用主函数开始处理
    # Call the main function to start processing
    copy_pdf_if_contains_code(source_folder, config_folder, sheet_name, header)
    print("All processing complete.")
