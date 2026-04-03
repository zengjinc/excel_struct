from config_builder import ConfigBuilder, SheetBuilder

builder = ConfigBuilder()

# 表格: 测试数据1
sheet = builder.add_sheet("测试数据1")

sheet.add_field("id", "测试ID")
sheet.add_field("sdk_id", "SDK ID")
sheet.add_field("type", "类型")
sheet.add_field("min", "最小值")
sheet.add_field("max", "最大值")
sheet.add_field("price", "价格")
sheet.add_field("cost_type", "价格类型")
sheet.add_field("sys_test", "测试数据")
sheet.add_field("achieve_args", "奖励")
sheet.add_field("reward", "奖励")
sheet.add_field("attr1", "属性1数组")
sheet.add_field("attr2", "属性2")

sheet.set_erl_name("data_test.erl")
sheet.add_include("common.hrl")
sheet.add_erl_function(
    name="get",
    key=['id', 'sdk_id'],
    value=['id', 'sdk_id', 'type', 'min', 'max', 'price', 'attr1-array', 'attr2'],
    return_type="#base_test{}",
)
sheet.add_erl_function(
    name="get_id_list",
    key=[],
    value=['id'],
    return_type="[]",
)
sheet.add_erl_function(
    name="get_type_by_id",
    key=['id'],
    value=['type'],
)
sheet.add_erl_function(
    name="get_by_price",
    key=['Price'],
    value=['id', 'sdk_id', 'type', 'min', 'max', 'price', 'reward-array'],
    return_type="#base_test{}",
    when="Price>min, Price=<max",
)
sheet.add_erl_function(
    name="get_price_and_reward",
    key=['id'],
    value=['price', 'reward-array'],
    return_type="{}",
)
sheet.add_erl_function(
    name="get_id_list_by_type",
    key=['type'],
    value=['id'],
    return_type="[[]]",
)
sheet.set_lua_name("config_test.lua")
sheet.add_lua_function(
    name="TestInfo",
    key=['id', 'sdk_id'],
    value=['id', 'sdk_id', 'type', 'min', 'max', 'price', 'reward-array4', 'attr1-array', 'attr2-array2'],
)
sheet.add_lua_function(
    name="TestReward",
    key=['id'],
    value=['price', 'sys_test'],
)

# 表格: 测试数据2
sheet = builder.add_sheet("测试数据2")

sheet.add_field("open_day_lim", "开服天数限制")
sheet.add_field("start_time", "活动开启时间")
sheet.add_field("get_goods", "获得的物品列表")

sheet.set_erl_name("data_test2.erl")
sheet.add_include("common.hrl")
sheet.add_erl_function(
    name="open_day_lim",
    key=[],
    value=['open_day_lim'],
    return_type="0",
)
sheet.add_erl_function(
    name="start_time",
    key=[],
    value=['start_time'],
    return_type="0",
)
sheet.add_erl_function(
    name="get_goods",
    key=[],
    value=['get_goods-array'],
    return_type="0",
)
sheet.set_lua_name("config_test2.lua")
sheet.add_lua_function(
    name="TestInfo",
    key=[],
    value=['get_goods-array4', 'open_day_lim'],
    return_type="0",
)

# 表格: 测试数据3
sheet = builder.add_sheet("测试数据3")

sheet.add_field("id", "ID")
sheet.add_field("lv", "等级")
sheet.add_field("money1", "金钱1")
sheet.add_field("money2", "金钱2")
sheet.add_field("limit", "条件限制")
sheet.add_field("attr1", "属性1")
sheet.add_field("attr2", "属性2")
sheet.add_field("attr3", "属性3")
sheet.add_field("extra_attrs", "额外属性")
sheet.add_field("drop_goods", "掉落物品")
sheet.add_field("desc", "描述")

sheet.set_erl_name("data_test3.erl")
sheet.add_include("common.hrl")
sheet.add_erl_function(
    name="get_lv_list_by_id",
    key=['id'],
    value=['lv'],
    return_type="[[]]",
)
sheet.add_erl_function(
    name="get_lv_conf",
    key=['id'],
    value=['id', 'lv', 'attr-1,3', 'money-1,2', 'drop_goods-array', 'desc-string'],
    return_type="#base_test{}",
)
sheet.add_erl_function(
    name="get_max_lv_by_id",
    key=['id'],
    value=['lv'],
    return_type="max",
)
sheet.add_erl_function(
    name="get_max_lv",
    key=[],
    value=['lv'],
    return_type="max",
)
sheet.add_erl_function(
    name="get_test_filter",
    key=['id', 'lv'],
    value=['lv', 'extra_attrs-array'],
    return_type="{}",
)
sheet.set_lua_name("config_test3.lua")
sheet.add_lua_function(
    name="MaxLv",
    key=[],
    value=['lv'],
    return_type="max",
)

# 构建配置并赋值给全局变量
config = builder.build()