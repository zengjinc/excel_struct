import os
import importlib.util

def parse_config(config_path):
    """
    解析配置文件
    :param config_path: 配置文件路径
    :return: 解析后的配置字典
    """
    try:
        # 导入配置模块
        spec = importlib.util.spec_from_file_location("config", config_path)
        config_module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(config_module)
        
        # 验证配置结构
        if not hasattr(config_module, 'config'):
            raise ValueError(f"配置文件 {config_path} 缺少 config 变量")
        
        config = config_module.config
        
        # 验证必要字段
        required_fields = ['filename', 'sheets']
        for field in required_fields:
            if field not in config:
                raise ValueError(f"配置文件 {config_path} 缺少 {field} 字段")
        
        # 验证 sheets 结构
        for sheet in config['sheets']:
            if 'name' not in sheet or 'config' not in sheet:
                raise ValueError(f"配置文件 {config_path} 中 sheet 缺少 name 或 config 字段")
            
            sheet_config = sheet['config']
            required_sheet_fields = ['ERL_NAME', 'LUA_NAME', 'fields']
            for field in required_sheet_fields:
                if field not in sheet_config:
                    raise ValueError(f"配置文件 {config_path} 中 sheet {sheet['name']} 缺少 {field} 字段")
        
        return config
    except Exception as e:
        raise Exception(f"解析配置文件 {config_path} 失败: {str(e)}")

def get_all_configs(struct_dir):
    """
    获取所有配置文件
    :param struct_dir: 配置文件目录
    :return: 配置文件路径列表
    """
    configs = []
    for filename in os.listdir(struct_dir):
        if filename.endswith('.py'):
            config_path = os.path.join(struct_dir, filename)
            try:
                config = parse_config(config_path)
                configs.append((config_path, config))
            except Exception as e:
                print(f"警告: {str(e)}")
    return configs