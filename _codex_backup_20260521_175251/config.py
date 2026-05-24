"""工程量清单导入配置。"""

# 飞书 API 凭证（如需写入飞书，按实际凭证填写）
APP_ID = "cli_xxxxxxxxxxxxxxxxxxxx"
APP_SECRET = "xxxxxxxxxxxxxxxxxxxxxxxx"

# 飞书多维表格
BASE_TOKEN = "ARDubM91FaL62esCg1hcnHYtnFh"
PRIMARY_TABLE_ID = "tbleZCShQmE1aRhQ"
PRIMARY_TABLE_NAME = "工程量清单"
MAX_ROWS_PER_TABLE = 20000

# lark-cli 路径
LARK_CLI = r"C:\Users\56237\AppData\Roaming\npm\lark-cli.cmd"

# Excel 源目录
EXCEL_DIR = "D:/单价库"

# 字段映射（字段名 -> field_id）
# 层级路径复用原“二级页签”字段，用 科目/二级页签 通过 / 合成。
FIELD_IDS = {
    "项目名称": "fldiVMozmY",
    "项目编号": "fldgEQEjf9",
    "项目特征描述": "fldxVGnJpn",
    "计量单位": "fldjhCzC8y",
    "工程数量": "fldPBDN6QO",
    "不含税单价": "fldbFXoOC5",
    "汇总合价": "fldGGC9sXw",
    "合同名称": "flddaIcKu7",
    "页签": "fldEvvjofm",
    "层级路径": "fldlIfsrTX",
    "备注": "fldJkEw41C",
    "定价模式": "fldpLjVfSv",
}

FIELD_TYPES = {k: "text" for k in FIELD_IDS.values()}
FIELD_TYPES.update({
    "fldbFXoOC5": "number",
    "fldGGC9sXw": "number",
    "fldPBDN6QO": "number",
})

DEDUP_KEYS = ["项目编号", "合同名称", "项目名称", "计量单位"]

SKIP_KEYWORDS = [
    "小计", "合计", "汇总", "暂列金额", "税金", "规费",
    "说明", "备注", "封皮", "编制说明", "总计",
]

TITLE_KEYWORDS = [
    "项目编号", "项目名称", "项目特征描述", "计量单位",
    "工程量", "综合单价", "基准价", "上下浮率", "投标价",
]
