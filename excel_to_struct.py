import os
import json
from openpyxl import load_workbook

def excel_to_struct(excel_file, struct_file):
    """
    从Excel文件反向生成/更新struct/文件夹中的描述文件

    Args:
        excel_file: Excel文件路径
        struct_file: 生成的描述文件路径
    """
    # 加载Excel文件
    workbook = load_workbook(excel_file)

    # 生成描述文件内容
    content = "from config_builder import ConfigBuilder, SheetBuilder\n\nbuilder = ConfigBuilder()\n\n"

    # 处理每个工作表
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        content += f"# 表格: {sheet_name}\n"
        content += f"sheet = builder.add_sheet(\"{sheet_name}\")\n"

        # 解析工作表内容
        erl_name = None
        lua_name = None
        includes = []
        erl_functions = []
        lua_functions = []
        fields = []
        field_names = []
        field_notes = []

        # 遍历工作表的每一行
        for row in sheet.iter_rows(min_row=1, values_only=True):
            if not row[0]:
                continue

            config_type = row[0]
            config_value = row[1]

            if config_type == 'ERL_NAME' and config_value:
                erl_name = config_value
                content += f"sheet.set_erl_name(\"{erl_name}\")\n"
            elif config_type == 'LUA_NAME' and config_value:
                lua_name = config_value
                content += f"sheet.set_lua_name(\"{lua_name}\")\n"
            elif config_type == 'ERL_INCLUDE' and config_value:
                includes.append(config_value)
                content += f"sheet.add_include(\"{config_value}\")\n"
            elif config_type == 'FIELD':
                # 提取字段名
                for i, cell in enumerate(row[1:], 1):
                    if cell:
                        field_names.append(cell)
                    else:
                        field_names.append(f"field_{i}")
            elif config_type == 'NOTE':
                # 提取字段注释
                for cell in row[1:]:
                    field_notes.append(cell if cell else "")
            elif config_type == 'ERL_FUN' and config_value:
                # 解析ERL_FUN配置
                key = []
                value = []
                if row[2]:  # params
                    try:
                        # 修复JSON字符串中的单引号问题
                        params_str = row[2].replace("'", "\"")
                        params = json.loads(params_str)
                        key = params.get('key', [])
                        value = params.get('value', [])
                    except Exception as e:
                        print(f"解析ERL_FUN params失败: {e}")
                        key = []
                        value = []

                return_type = ""
                if row[3]:  # return
                    try:
                        return_config = json.loads(row[3])
                        return_type = return_config.get('return', "")
                    except Exception as e:
                        print(f"解析ERL_FUN return失败: {e}")
                        pass

                when = ""
                if row[4]:  # when
                    try:
                        when_config = json.loads(row[4])
                        when = when_config.get('when', "")
                    except Exception as e:
                        print(f"解析ERL_FUN when失败: {e}")
                        pass

                note = ""
                if row[5]:  # note
                    note = row[5]

                # 生成ERL函数代码
                content += f"sheet.add_erl_function(\n"
                content += f"    name=\"{config_value}\",\n"
                content += f"    key={key},\n"
                content += f"    value={value},\n"
                if return_type:
                    content += f"    return_type=\"{return_type}\",\n"
                if when:
                    content += f"    when=\"{when}\",\n"
                if note:
                    content += f"    note=\"{note}\",\n"
                content += f")\n"
            elif config_type == 'LUA_FUN' and config_value:
                # 解析LUA_FUN配置
                key = []
                value = []
                if row[2]:  # params
                    try:
                        # 修复JSON字符串中的单引号问题
                        params_str = row[2].replace("'", "\"")
                        params = json.loads(params_str)
                        key = params.get('key', [])
                        value = params.get('value', [])
                    except Exception as e:
                        print(f"解析LUA_FUN params失败: {e}")
                        key = []
                        value = []

                return_type = ""
                if row[3]:  # return
                    try:
                        return_config = json.loads(row[3])
                        return_type = return_config.get('return', "")
                    except Exception as e:
                        print(f"解析LUA_FUN return失败: {e}")
                        pass

                # 生成LUA函数代码
                content += f"sheet.add_lua_function(\n"
                content += f"    name=\"{config_value}\",\n"
                content += f"    key={key},\n"
                content += f"    value={value},\n"
                if return_type:
                    content += f"    return_type=\"{return_type}\",\n"
                content += f")\n"

        # 生成字段代码
        # 确保field_notes的长度与field_names相同
        while len(field_notes) < len(field_names):
            field_notes.append("")
        # 组合字段名和注释
        if not field_names:
            # 判断field字段是否为空
            continue
        for field_name, field_note in zip(field_names, field_notes):
                content += f"sheet.add_field(\"{field_name}\", \"{field_note}\")\n"

        content += "\n"

    # 添加构建配置的代码
    content += "# 构建配置并赋值给全局变量\nconfig = builder.build()"

    # 写入文件
    os.makedirs(os.path.dirname(struct_file), exist_ok=True)
    with open(struct_file, 'w', encoding='utf-8') as f:
        f.write(content)

    print(f"成功生成/更新描述文件: {struct_file}")

def process_target_directory():
    """
    处理target目录下的所有xlsx文件，生成/更新对应的struct文件
    """
    import config as config_module

    # 处理target文件夹路径（支持相对路径和绝对路径）
    if os.path.isabs(config_module.TARGET_FOLDER):
        target_dir = config_module.TARGET_FOLDER
    else:
        target_dir = os.path.join(os.path.dirname(__file__), config_module.TARGET_FOLDER)

    struct_dir = os.path.join(os.path.dirname(__file__), 'struct')

    # 遍历target目录下的所有xlsx文件
    for filename in os.listdir(target_dir):
        if filename.endswith('.xlsx'):
            excel_file = os.path.join(target_dir, filename)
            struct_filename = os.path.splitext(filename)[0] + '.py'
            struct_file = os.path.join(struct_dir, struct_filename)

            print(f"处理文件: {excel_file}")
            excel_to_struct(excel_file, struct_file)

def process_single_excel(excel_name):
    """
    处理单个Excel文件，生成/更新对应的struct文件

    Args:
        excel_name: Excel文件名（不含 .xlsx 后缀）
    """
    import config as config_module

    # 处理target文件夹路径（支持相对路径和绝对路径）
    if os.path.isabs(config_module.TARGET_FOLDER):
        target_dir = config_module.TARGET_FOLDER
    else:
        target_dir = os.path.join(os.path.dirname(__file__), config_module.TARGET_FOLDER)

    struct_dir = os.path.join(os.path.dirname(__file__), 'struct')

    # 构建Excel文件路径和struct文件路径
    excel_file = os.path.join(target_dir, f"{excel_name}.xlsx")
    struct_file = os.path.join(struct_dir, f"{excel_name}.py")

    # 检查Excel文件是否存在
    if not os.path.exists(excel_file):
        print(f"错误: Excel文件 {excel_file} 不存在")
        return

    print(f"处理文件: {excel_file}")
    excel_to_struct(excel_file, struct_file)

if __name__ == "__main__":
    process_target_directory()
