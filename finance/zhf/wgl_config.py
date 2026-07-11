"""
稳岗率计算 - 配置常量
包含：档位规则、列名配置等
"""

# 档位范围
GEAR_MIN = -8
GEAR_MAX = 10

# 档位步长
GEAR_STEP = 0.05

# 项目驻场单价配置
ZC_UNIT_MIN_GEAR = -5  # 小于此档位单价为0
ZC_UNIT_MAX_GEAR = 6   # 大于此档位单价为10
ZC_UNIT_MAX = 10.0      # 封顶单价
ZC_UNIT_MIN = 0.0       # 最低单价

# 项目经理单价配置
PM_UNIT_MIN_GEAR = -5   # 小于此档位单价为0
PM_UNIT_MAX_GEAR = 6    # 大于此档位单价为2
PM_UNIT_MAX = 2.0       # 封顶单价
PM_UNIT_MIN = 0.0       # 最低单价

# 默认单价
DEFAULT_UNIT_ZC = 5.0   # 项目驻场默认单价
DEFAULT_UNIT_PM = 1.0   # 项目经理默认单价
DEFAULT_UNIT_JINTU = 0.0  # 锦途单价

# 入职补贴列名
COL_IN_SUBSIDY_ZC = "入职补贴（5元/人）"
COL_IN_SUBSIDY_PM = "入职补贴（1元/人）"

# 稳岗补贴列名
COL_STABILITY_SUBSIDY = "稳岗补贴"

# 折算后单价列名
COL_ADJUSTED_UNIT = "折算后单价5元/人 (期初)-稳岗"

# 需要排除的特殊行
EXCLUDE_NAMES = ["", "总计", "平均数", "中位数"]

# 锦途关键字
JINTU_KEYWORD = "锦途"

# 输出工作表名称
SHEET_ZC = "项目驻场"
SHEET_PM = "项目经理"
SHEET_STABILITY = "稳岗率"
SHEET_REWARD = "驻场&项目经理 - 奖励汇总"
SHEET_BASE_RATE = "基准数"

# 基准数表列数
BASE_RATE_COL_COUNT = 13

# 基准数表标题列
BASE_RATE_TITLE_COL_ZC = 2  # 项目驻场标题列
BASE_RATE_TITLE_COL_PM = 9  # 项目经理标题列

# 基准数表基准数列
BASE_RATE_COL_ZC = 1   # 项目驻场基准数列
BASE_RATE_COL_PM = 8   # 项目经理基准数列

# 基准数表档位列
BASE_RATE_GEAR_COL_ZC = 2  # 项目驻场档位列
BASE_RATE_GEAR_COL_PM = 9  # 项目经理档位列

# 基准数表稳岗率列
BASE_RATE_RATE_COL_ZC = 3  # 项目驻场稳岗率列
BASE_RATE_RATE_COL_PM = 10 # 项目经理稳岗率列

# 基准数表单价列
BASE_RATE_UNIT_COL_ZC = 4  # 项目驻场单价列
BASE_RATE_UNIT_COL_PM = 11 # 项目经理单价列

# 基准数表备注列
BASE_RATE_NOTE_COL_ZC = 5  # 项目驻场备注列
BASE_RATE_NOTE_COL_PM = 12 # 项目经理备注列

# 基准数表分隔列
BASE_RATE_SEP_COL = 6      # 分隔列
