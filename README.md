# Excel 配置工具

这是一个用于处理 Excel 配置文件的工具，支持从 struct 结构文件生成 Excel 配置文件，以及从 Excel 文件反向生成/更新 struct 结构文件。

## 功能特性

- **从 struct 生成 Excel**：根据 struct 结构文件自动生成 Excel 配置文件
- **从 Excel 生成 struct**：从 Excel 文件反向生成/更新 struct 结构文件
- **支持批量处理**：可以一次性处理所有配置文件
- **多语言支持**：支持生成 Erlang 和 Lua 相关的配置
- **灵活配置**：通过配置文件自定义目标文件夹路径

## 项目结构

```
excel_struct/
├── struct/          # struct 结构文件目录
│   ├── base_test.py # 示例 struct 结构文件
│   └── ...
├── target/          # 生成的 Excel 文件目录
├── config.py        # 配置文件
├── config_builder.py # 配置构建器
├── config_parser.py  # 配置解析器
├── excel_generator.py # Excel 生成器
├── excel_to_struct.py # Excel 转 struct 工具
├── struct_to_excel.py # struct 转 Excel 工具
└── main.py          # 主入口文件
```

## 安装依赖

本项目使用了第三方库 `openpyxl` 来处理 Excel 文件，需要安装：

```bash
pip install openpyxl
```

## 使用方法

### 1. 从 struct 结构文件生成 Excel 配置文件

```bash
# 生成指定的配置文件
python main.py -gen_excel CONFIG_NAME

# 生成所有配置文件
python main.py -gen_excel all
```

### 2. 从 Excel 文件反向生成/更新 struct 结构文件

```bash
# 生成指定的 struct 文件
python main.py -gen_struct EXCEL_NAME

# 生成所有 struct 文件
python main.py -gen_struct all
```

## struct 结构文件格式

struct 结构文件使用 Python 编写，使用 `ConfigBuilder` 和 `SheetBuilder` 构建配置。以下是一个示例：

```python
from config_builder import ConfigBuilder, SheetBuilder

builder = ConfigBuilder()

# 表格: 测试数据1
sheet = builder.add_sheet("测试数据1")

sheet.add_field("id", "测试ID")
sheet.add_field("sdk_id", "SDK ID")
sheet.add_field("type", "类型")

sheet.set_erl_name("data_test.erl")
sheet.add_include("common.hrl")
sheet.add_erl_function(
    name="get",
    key=['id', 'sdk_id'],
    value=['id', 'sdk_id', 'type', 'min', 'max', 'price', 'attr1-array', 'attr2'],
    return_type="#base_test{}",
)

sheet.set_lua_name("config_test.lua")
sheet.add_lua_function(
    name="TestInfo",
    key=['id', 'sdk_id'],
    value=['id', 'sdk_id', 'type', 'min', 'max', 'price', 'reward-array4', 'attr1-array', 'attr2-array2'],
)

# 构建配置并赋值给全局变量
config = builder.build()
```

## 配置说明

在 `config.py` 文件中，可以配置目标文件夹路径：

```python
# Target folder for generated files
# Default: 'target' directory in the current working directory
# 支持相对路径和绝对路径
TARGET_FOLDER = './target'
```

## 注意事项

1. struct 结构文件必须放在 `struct` 目录下
2. 生成的 Excel 文件会保存到 `target` 目录下
3. 反向生成 struct 时，会覆盖原有文件，请确保备份重要文件
4. 支持的字段类型包括普通字段和数组字段（使用 `-array` 后缀）

## 示例

### 示例 1: 生成 Excel 文件

```bash
# 生成 base_test 配置对应的 Excel 文件
python main.py -gen_excel base_test
```

### 示例 2: 从 Excel 生成 struct

```bash
# 从 base_test.xlsx 生成 struct 结构文件
python main.py -gen_struct base_test
```

## 错误处理

如果遇到以下错误：

- "错误: 配置文件目录 struct 不存在"：请确保 struct 目录存在
- "错误: 目标文件目录 target 不存在"：请确保 target 目录存在
- "警告: 未找到有效的配置文件"：请检查 struct 目录下是否有有效的配置文件
- "警告: 未找到名为 XXX 的配置文件"：请检查配置文件名是否正确
