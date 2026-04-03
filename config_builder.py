import os
import inspect
from typing import Dict, List, Optional, Any

class ConfigBuilder:
    def __init__(self) -> None:
        # 获取调用者的文件路径
        caller_frame = inspect.stack()[1]
        caller_file = caller_frame.filename
        # 提取文件名（不含扩展名）
        filename = os.path.splitext(os.path.basename(caller_file))[0] + ".xlsx"

        self.config: Dict[str, Any] = {
            "filename": filename,
            "sheets": []
        }

    def add_sheet(self, name: str, sheet_config: Optional[Dict[str, Any]] = None) -> 'SheetBuilder':
        # 默认配置
        default_config: Dict[str, Any] = {
            "ERL_NAME": f"data_{name.lower().replace(' ', '_')}.erl",
            "ERL_INCLUDE": ["common.hrl"],
            "ERL_FFUN": [],
            "LUA_NAME": f"config_{name.lower().replace(' ', '_')}.lua",
            "LUA_FUN": [],
            "fields": []
        }

        # 如果提供了自定义配置，覆盖默认值
        if sheet_config:
            default_config.update(sheet_config)

        sheet: Dict[str, Any] = {
            "name": name,
            "config": default_config
        }
        self.config["sheets"].append(sheet)
        return SheetBuilder(sheet["config"])

    def build(self) -> Dict[str, Any]:
        return self.config

class SheetBuilder:
    def __init__(self, sheet_config: Dict[str, Any]) -> None:
        self.sheet_config: Dict[str, Any] = sheet_config

    def set_erl_name(self, name: str) -> 'SheetBuilder':
        self.sheet_config["ERL_NAME"] = name
        return self

    def set_lua_name(self, name: str) -> 'SheetBuilder':
        self.sheet_config["LUA_NAME"] = name
        return self

    def add_include(self, include_file: str) -> 'SheetBuilder':
        if include_file not in self.sheet_config["ERL_INCLUDE"]:
            self.sheet_config["ERL_INCLUDE"].append(include_file)
        return self

    def add_field(self, field: str, note: str) -> 'SheetBuilder':
        self.sheet_config["fields"].append({"FIELD": field, "NOTE": note})
        return self

    def add_erl_function(self, name: str, key: List[str] = None, value: List[str] = None, return_type: str = "", when: str = "", note: str = "") -> 'SheetBuilder':
        if key is None:
            key = []
        if value is None:
            value = []

        fun_config: Dict[str, Any] = {
            "name": name,
            "params": {"key": key, "value": value},
            "return": return_type,
            "when": when,
            "note": note,
            "fun_note": key + [v for v in value if v not in key]
        }
        self.sheet_config["ERL_FFUN"].append(fun_config)
        return self

    def add_lua_function(self, name: str, key: List[str] = None, value: List[str] = None, return_type: str = "") -> 'SheetBuilder':
        if key is None:
            key = []
        if value is None:
            value = []

        fun_config: Dict[str, Any] = {
            "name": name,
            "params": {"key": key, "value": value},
            "return": return_type
        }
        self.sheet_config["LUA_FUN"].append(fun_config)
        return self