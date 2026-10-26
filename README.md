# 浙江锦途 · Excel 表格处理工具（桌面版）

基于 **CustomTkinter** 的本地桌面工具集，把财务 / 人事日常的 Excel 批处理固化成一键操作：左侧导航切换模块，右侧功能卡片，全程本地运行、不上传云端。

> 同一套业务规则另有 Web 版 [`ExcelUtils`](https://github.com/Gitlxb/ExcelUtils)（Spring Boot + Vue）。桌面版偏单机批处理、无需部署；Web 版便于多人共用与远程访问。两端的核心算法（如手续费分类 A~N、稳岗率分段、工资表拆分）保持一致。

---

## 功能清单

### 通用功能（5 项）

| 功能 | 说明 |
|---|---|
| 拆分 · 拆成多个 Excel 文件 | 按指定列值把一个工作簿拆成多个独立文件 |
| 拆分 · 拆成多个工作表 | 按指定列值拆到同一文件的多个 Sheet |
| N 层级文件夹中只获取文件 | 从深层嵌套目录中递归提取所有文件 |
| 表格格式转换 | 文件夹内 Excel / CSV 统一转为 xlsx |
| PDF 转 Excel | 提取 PDF 表格，按列对齐输出 |

### 财务功能（9 项）

| 功能 | 说明 |
|---|---|
| 财务_收支记账_筛选 | 开关账收支记账筛选 |
| 供应商_合并表格 | 多供应商 Excel 合并 |
| 出纳_工具箱 | 含手续费处理（按运营项目 + 银行类型计费） |
| 绩效考核 | 数据匹配 / 稳岗率分析 / 合并表格 |
| 预支平账 | 「工人预支申请」与「预支人员信息查询明细表」两表匹配平账 |
| 核算汇总 | 拆分汇总与核对 |
| 银行流水_摘要修改 | 流水摘要批量改写 |
| 批量批注工具 | 按名单批量写单元格批注 |
| **合并个税** | 递归合并各簿「个税统计」表，按规则表计算手续费 |

### 人事功能（4 项）

| 功能 | 说明 |
|---|---|
| 拆分 · 拆分两次_在职离职 | 工资表按在职 / 离职两次拆分 |
| 拆分 · 指定内容放一起 | 按拆分参数把指定内容归集后拆分 |
| 工资表匹配 | 两张工资表对比并高亮差异 |
| 文档处理工具箱 | 工资表拆分（保留公式/样式）+ Word 文档智能合并 |

### 独立脚本（可直接运行，未接入主界面）

| 脚本 | 用途 |
|---|---|
| `sjqx_ghm.py` + `sjqx_ghm_gui.py` | 数据去重（核心逻辑 + GUI 分离） |
| `match_advance_tool.py` | 工人预支申请 ↔ 预支人员信息查询明细表 匹配 |
| `match_idcard_tool.py` | 身份证号码多字段组合匹配 |

---

## 目录结构

```
PythonProject/
├── main.py                 # 入口：启动画面 → 左侧导航 → Frame 切换
├── frame_general.py        # 通用功能页（卡片式入口）
├── frame_finance.py        # 财务功能页
├── frame_hr.py             # 人事功能页
├── ui_components.py        # 通用 UI 组件与主题（Theme）
├── finance/                # 财务模块实现
│   ├── hbgs_core.py        #   合并个税：核心逻辑（扫描/读取/手续费）
│   ├── hbgs_gui.py         #   合并个税：GUI
│   ├── cxm/  pmh/  xzy/  zhf/   # 分类业务模块（按经办人/业务线拆分）
│   ├── cw_gys_qzy.py       #   供应商合并
│   ├── cw_kgz_sx.py        #   开关账筛选
│   ├── cw_yhls_zyxg.py     #   银行流水摘要修改
│   └── pl_cf_pizhu.py      #   批量批注
├── general/                # 通用模块实现
│   ├── cf_biaoge_li.py     #   拆成多个工作表
│   ├── cf_duoge_wenj.py    #   拆成多个文件
│   ├── Ncengjiwenjj_fzzt.py#   N 层级文件夹提取
│   ├── biaogegesi_zh.py    #   表格格式转换
│   └── zh_pdf_Excel.py     #   PDF 转 Excel
├── hr/                     # 人事模块实现
│   ├── cf_duoge_wenj_zz.py #   拆分两次_在职离职
│   ├── cf_zdnrfyq_nrcs_wsx.py # 指定内容放一起 + 拆分参数
│   ├── rs_gzb_pipei.py     #   工资表匹配
│   └── hyp/                #   文档处理工具箱（工资表拆分 / Word 合并 / 孝心金拆分）
└── *.spec                  # PyInstaller 打包配置（每个工具一份）
```

---

## 环境要求

- **Windows**（代码用到 `ctypes` 高 DPI 感知、`Microsoft YaHei UI` 字体与 Windows 路径约定）
- **Python 3.11+**（实测 3.13）
- 依赖见 `requirements.txt`

```bash
pip install -r requirements.txt
```

主要依赖：`customtkinter`（界面）、`openpyxl`（Excel 读写）、`pandas` / `numpy`（批量数据处理）、`python-docx`（Word 合并）、`pdfplumber`（PDF 转 Excel，仅该模块需要）。

---

## 运行

```bash
python main.py
```

启动会先显示启动画面（预加载 pandas/openpyxl 等重型模块），随后进入主界面：左侧为「通用 / 财务 / 人事」导航，右侧为功能卡片。

独立脚本可直接运行，例如：

```bash
python sjqx_ghm_gui.py          # 数据去重
python match_advance_tool.py    # 预支匹配
```

---

## 打包为 exe

```bash
pyinstaller main.spec
# → dist/浙江锦途 - Excel表格处理工具.exe（单文件）
```

各子模块自带独立 `.spec`（如 `finance/cw_gys_qzy.spec`、`hr/cf_duoge_wenj_zz.spec`），需要单独分发某个工具时，进对应目录执行 `pyinstaller xxx.spec` 即可。

> `main.spec` 已声明隐藏导入（`docx`、`lxml`、`PIL`、三个 frame 模块），新增子模块时记得同步补进 `hiddenimports`，否则打包后运行会报 `ModuleNotFoundError`。

---

## 合并个税（说明）

与 Web 版逻辑一致，桌面版为「选择规则表 + 选择数据目录」两步：

1. **规则表**：需包含 `客户状态`、`运营项目`、`扣费银行（多填）`、`扣费金额` 四列，`客户状态 = 终止合作` 的行会被跳过。
2. **数据目录**：递归扫描所有 Excel，每个源簿需含 `个税统计` 表（必须有「姓名」「身份证」列）与 `员工信息-X月` 表（提供「运营项目」「开户行」用于计费，缺失则手续费留空）。
3. **输出**：19 列标准布局，白名单外的列按邻列插入，`开户行 / 发放公司 / 电话号码` 整列跳过，「采用/不采用」列跟随最后一个有效列。

手续费按规则文本精确匹配为 A~N 共 14 类，再结合开户行判断免扣；金额优先级：规则表「扣费金额」为正数时覆盖默认值，否则 A=0、I=3、其余=5。

---

## 注意事项

- **仅 Windows 可用**，跨平台需自行替换 DPI 设置与字体。
- 所有工具**只读源文件、另存输出**，不会原地覆盖；输出文件默认落在源目录或用户选择的位置。
- `__pycache__/`、`*.pyc`、`build/`、`dist/` 已在 `.gitignore` 中排除，仓库只保存源码与打包脚本。
- 部分业务规则（手续费免扣、稳岗率分段、拆分参数）属于硬编码的业务约定，改动前建议先与财务/人事确认口径。
