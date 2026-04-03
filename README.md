# Excel结构生成工具

通过编写JSON配置文件生成目标Excel文件，支持Erlang和Lua函数配置。

## 项目结构

```
excel_struct/
├── struct/              # 配置文件目录
│   ├── example_config.json
│   └── complex_example.json
├── target/             # 生成的Excel文件目录
│   ├── example_output.xlsx
│   └── complex_example.xlsx
├── config_parser.py    # 配置文件解析模块
├── excel_generator.py  # Excel生成模块
├── generate_excel.py   # 主脚本入口
├── 配置说明.md         # 详细配置说明
└── 项目需求.md         # 项目需求文档
```

## 使用方法

### 1. 创建配置文件

在 `struct/` 目录下创建JSON格式的配置文件，配置文件结构如下：

```json
{
  "output_file": "output.xlsx",
  "sheets": [
    {
      "name": "工作表名称",
      "erl_name": "data_module.erl",
      "erl_include": ["common.hrl"],
      "erl_funs": [
        {
          "name": "function_name",
          "params": {
            "key": ["field1"],
            "value": ["field2"]
          },
          "return": "",
          "when": "",
          "note": ""
        }
      ],
      "lua_name": "config_module.lua",
      "lua_funs": [
        {
          "name": "FunctionName",
          "params": {
            "key": ["field1"],
            "value": ["field2"]
          },
          "default": {
            "num": 0,
            "name": ""
          }
        }
      ],
      "fields": [
        {"name": "field1", "note": "字段说明1"},
        {"name": "field2", "note": "字段说明2"}
      ],
      "values": [
        [1, "value1"],
        [2, "value2"]
      ]
    }
  ]
}
```

### 2. 运行生成脚本

```bash
python generate_excel.py struct/example_config.json
```

### 3. 命令行参数

- `config`: 配置文件路径（必需）
- `-o, --output`: 输出文件路径（可选，默认使用配置文件中的output_file）
- `-d, --output-dir`: 输出目录（可选，默认为target目录）

示例：

```bash
# 使用默认输出路径
python generate_excel.py struct/example_config.json

# 指定输出文件
python generate_excel.py struct/example_config.json -o custom_output.xlsx

# 指定输出目录
python generate_excel.py struct/example_config.json -d custom_target
```

## 配置项说明

### 基本配置项

| 配置项 | 说明 | 示例 |
|--------|------|------|
| output_file | 输出Excel文件名 | "output.xlsx" |
| name | 工作表名称 | "测试表1" |
| erl_name | Erlang文件名 | "data_test.erl" |
| erl_include | 需要包含的头文件 | ["common.hrl"] |
| lua_name | Lua文件名 | "config_test.lua" |

### ERL_FUN 配置

| 字段 | 说明 | 示例 |
|------|------|------|
| name | 函数名（小写+下划线） | "get_name" |
| params | 参数配置 | {"key": ["id"], "value": ["name"]} |
| return | 返回类型 | "", "[]", "#base_test{}", "map", "{}", "0", "[[]]", "max" |
| when | when条件 | "" |
| note | 函数注释 | "" |
| fun_note | 字段注释 | ["field1", "field2"] |
| filter | 过滤配置 | {"field": ""} |

### LUA_FUN 配置

| 字段 | 说明 | 示例 |
|------|------|------|
| name | 函数名（双驼峰） | "GetName" |
| params | 参数配置 | {"key": ["id"], "value": ["name"]} |
| default | 默认值 | {"num": 0, "name": ""} |
| return | 返回类型 | "single", "max", "[]" |
| sub_key | 子键 | ["lv"] |
| key_split | 键分隔符 | "#" |
| filter | 过滤配置 | {"field": ""} |

### FIELD 配置

| 字段 | 说明 |
|------|------|
| name | 字段名 |
| note | 字段说明 |

### VALUE 配置

每行数据对应一个VALUE行，数据按字段顺序排列。

## 返回类型说明

### Erlang返回类型

| 返回类型 | 说明 | 示例 |
|----------|------|------|
| "" | 原样返回 | get(1)->100 |
| "[]" | 返回列表 | get_id_list()->[1, 2, 3] |
| "#base_test{}" | 返回记录 | get_event(Id)->#base_test{...} |
| "map" | 返回map | get_event(Id)->#{"field1" => "value1"} |
| "{}" | 返回N列元组 | get(1)->{1, 200, [{1001, 1000}]} |
| "0" | 返回常量 | start_time()->{12, 00} |
| "[[]]" | 返回key-list | get_id_list(1)->[1, 2, 3] |
| "max" | 返回最大值 | get_max_lv(2)->10 |

### Lua返回类型

| 返回类型 | 说明 |
|----------|------|
| "single" | 返回单个值 |
| "max" | 返回最大值 |
| "[]" | 返回列表 |

## 示例配置文件

### 简单示例

参见 [struct/example_config.json](struct/example_config.json)

### 复杂示例

参见 [struct/complex_example.json](struct/complex_example.json)

## 依赖项

- Python 3.6+
- openpyxl

安装依赖：

```bash
pip install openpyxl
```

## 验证输出

运行验证脚本查看生成的Excel文件内容：

```bash
python verify_output.py
```

## 注意事项

1. 配置文件必须使用UTF-8编码
2. 字段名和值必须与配置中的字段定义一一对应
3. 修改配置文件后重新运行脚本即可生成新的Excel文件
4. 生成的Excel文件会保留原有的VALUE行数据，仅修改结构部分
