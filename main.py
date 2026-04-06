import os
import sys
import argparse
from config_parser import get_all_configs
from struct_to_excel import generate_excel
import config as config_module
from excel_to_struct import process_single_excel, process_target_directory

def main():
    """
    主函数
    """
    # 解析命令行参数
    parser = argparse.ArgumentParser(description='Excel 配置工具')
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument('-gen_excel', metavar='CONFIG_NAME', help='生成 Excel 配置文件，配置文件名（不含 .py 后缀），支持填 all')
    group.add_argument('-gen_struct', metavar='EXCEL_NAME', help='从 Excel 文件反向生成/更新 struct 结构文件，Excel 文件名（不含 .xlsx 后缀）')
    args = parser.parse_args()

    # 处理struct文件夹路径（支持相对路径和绝对路径）
    if os.path.isabs(config_module.STRUCT_FOLDER):
        struct_dir = config_module.STRUCT_FOLDER
    else:
        struct_dir = os.path.join(os.path.dirname(__file__), config_module.STRUCT_FOLDER)
    # 处理target文件夹路径（支持相对路径和绝对路径）
    if os.path.isabs(config_module.TARGET_FOLDER):
        target_dir = config_module.TARGET_FOLDER
    else:
        target_dir = os.path.join(os.path.dirname(__file__), config_module.TARGET_FOLDER)

    # 检查目录是否存在
    if not os.path.exists(struct_dir):
        print(f"错误: 配置文件目录 {struct_dir} 不存在")
        sys.exit(1)

    if not os.path.exists(target_dir):
        print(f"错误: 目标文件目录 {target_dir} 不存在")
        sys.exit(1)

    if args.gen_struct:
        # 从 Excel 文件反向生成/更新 struct 结构文件
        if args.gen_struct == 'all':
            process_target_directory()
        else:
            process_single_excel(args.gen_struct)
        print("\nStruct 结构文件生成/更新完成！")
        return

    # 获取所有配置文件
    all_configs = get_all_configs(struct_dir)

    if not all_configs:
        print("警告: 未找到有效的配置文件")
        sys.exit(0)

    # 根据参数筛选配置文件
    configs = []
    if args.gen_excel == 'all':
        configs = all_configs
    else:
        for config_path, config in all_configs:
            config_file = os.path.basename(config_path)
            config_name = os.path.splitext(config_file)[0]
            if config_name == args.gen_excel:
                configs.append((config_path, config))

    if not configs:
        print(f"警告: 未找到名为 {args.gen_excel} 的配置文件")
        sys.exit(0)

    # 生成 Excel 文件
    for config_path, config in configs:
        filename = config['filename']
        target_path = os.path.join(target_dir, filename)

        try:
            generate_excel(config, target_path)
        except Exception as e:
            print(f"错误: 生成 Excel 文件 {filename} 失败: {str(e)}")

    print("\nExcel 文件生成完成！")

if __name__ == '__main__':
    main()