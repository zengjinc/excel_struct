import json
from typing import Dict, List, Any, Optional
from dataclasses import dataclass, field


@dataclass
class ErlFun:
    name: str
    params: Dict[str, List[str]]
    return_type: str = ""
    when_clause: str = ""
    note: str = ""
    fun_note: Optional[List[str]] = None
    filter_config: Optional[Dict[str, Any]] = None


@dataclass
class LuaFun:
    name: str
    params: Dict[str, List[str]]
    default: Optional[Dict[str, Any]] = None
    return_type: str = ""
    sub_key: Optional[List[str]] = None
    key_split: Optional[str] = None
    filter_config: Optional[Dict[str, Any]] = None


@dataclass
class PreFormat:
    key: List[str]
    value: List[str]
    server: Optional[Dict[str, str]] = None
    client: Optional[Dict[str, str]] = None


@dataclass
class Field:
    name: str
    note: str


@dataclass
class SheetConfig:
    name: str
    erl_name: str
    erl_include: List[str]
    erl_funs: List[ErlFun]
    lua_name: str
    lua_funs: List[LuaFun]
    fields: List[Field]
    values: List[List[Any]]
    pre_format: Optional[PreFormat] = None


@dataclass
class ExcelConfig:
    output_file: str
    sheets: List[SheetConfig]


class ConfigParser:
    def __init__(self):
        pass

    def parse_file(self, file_path: str) -> ExcelConfig:
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return self.parse_dict(data)

    def parse_dict(self, data: Dict[str, Any]) -> ExcelConfig:
        output_file = data.get('output_file', 'output.xlsx')
        sheets_data = data.get('sheets', [])
        
        sheets = []
        for sheet_data in sheets_data:
            sheet = self._parse_sheet(sheet_data)
            sheets.append(sheet)
        
        return ExcelConfig(output_file=output_file, sheets=sheets)

    def _parse_sheet(self, data: Dict[str, Any]) -> SheetConfig:
        name = data.get('name', 'Sheet1')
        erl_name = data.get('erl_name', '')
        erl_include = data.get('erl_include', [])
        erl_funs_data = data.get('erl_funs', [])
        lua_name = data.get('lua_name', '')
        lua_funs_data = data.get('lua_funs', [])
        fields_data = data.get('fields', [])
        values_data = data.get('values', [])
        pre_format_data = data.get('pre_format')

        erl_funs = [self._parse_erl_fun(fun_data) for fun_data in erl_funs_data]
        lua_funs = [self._parse_lua_fun(fun_data) for fun_data in lua_funs_data]
        fields = [self._parse_field(field_data) for field_data in fields_data]
        pre_format = self._parse_pre_format(pre_format_data) if pre_format_data else None

        return SheetConfig(
            name=name,
            erl_name=erl_name,
            erl_include=erl_include,
            erl_funs=erl_funs,
            lua_name=lua_name,
            lua_funs=lua_funs,
            fields=fields,
            values=values_data,
            pre_format=pre_format
        )

    def _parse_erl_fun(self, data: Dict[str, Any]) -> ErlFun:
        return ErlFun(
            name=data.get('name', ''),
            params=data.get('params', {}),
            return_type=data.get('return', ''),
            when_clause=data.get('when', ''),
            note=data.get('note', ''),
            fun_note=data.get('fun_note'),
            filter_config=data.get('filter')
        )

    def _parse_lua_fun(self, data: Dict[str, Any]) -> LuaFun:
        return LuaFun(
            name=data.get('name', ''),
            params=data.get('params', {}),
            default=data.get('default'),
            return_type=data.get('return', ''),
            sub_key=data.get('sub_key'),
            key_split=data.get('key_split'),
            filter_config=data.get('filter')
        )

    def _parse_pre_format(self, data: Dict[str, Any]) -> PreFormat:
        return PreFormat(
            key=data.get('key', []),
            value=data.get('value', []),
            server=data.get('server'),
            client=data.get('client')
        )

    def _parse_field(self, data: Dict[str, Any]) -> Field:
        return Field(
            name=data.get('name', ''),
            note=data.get('note', '')
        )

    def validate(self, config: ExcelConfig) -> List[str]:
        errors = []
        
        if not config.sheets:
            errors.append("配置文件必须包含至少一个sheet")
        
        for idx, sheet in enumerate(config.sheets):
            if not sheet.name:
                errors.append(f"Sheet {idx + 1}: 缺少名称")
            
            if not sheet.erl_name:
                errors.append(f"Sheet '{sheet.name}': 缺少erl_name")
            
            if not sheet.lua_name:
                errors.append(f"Sheet '{sheet.name}': 缺少lua_name")
            
            if not sheet.fields:
                errors.append(f"Sheet '{sheet.name}': 缺少字段定义")
            
            if not sheet.values:
                errors.append(f"Sheet '{sheet.name}': 缺少数据值")
            
            for value_row in sheet.values:
                if len(value_row) != len(sheet.fields):
                    errors.append(f"Sheet '{sheet.name}': 数据行长度与字段数不匹配")
        
        return errors
