# 汇丰 / 渣打基金晨星分析工具

把银行销售的基金份额按 ISIN 汇总，查询晨星英国站 / 香港站的数据，并生成可供人工筛选的 Excel。当前程序主要整理历史收益、波动率和评级，不计算统一的基金推荐分数。

## 数据处理流程

| 步骤 | 脚本 | 输入与处理 | 输出 |
| --- | --- | --- | --- |
| 汇丰目录 | `hsbc_1_extract_fund_pdfs.py` | 解析 `files/hsbc/fund-nav.pdf`，从 IPFD 编号生成下载地址，下载各基金 PDF | `files/hsbc/QDII.xlsx`、`files/hsbc/funds/*.pdf` |
| 汇丰 ISIN | `hsbc_2_generate_isin.py` | 拆分 PDF 表格中的多行单元格；优先解析规范表格，再从混乱表格中提取 IPFD / ISIN | 各 PDF 对应的 Excel、`files/hsbc/ISIN_HSBC.xlsx` |
| 渣打 ISIN | `sc_1_load_funds.py` | 解析手动保存的基金列表 HTML，从 `data-row-key` 拆分 ISIN 与币种，并读取名称、标签、收益、派息和波动数据 | `files/sc/ISIN_SC.xlsx` |
| 晨星匹配 | `morningstar_1_search_id.py` | 合并两家银行的 ISIN，记录银行来源，先查 UK，返回 `NotFound` 时再查 HK | `files/morningstar/MorningStar.xlsx` |
| 页面缓存 | `morningstar_2_save_fund_pages.py` | 检查已有 HTML 和指标缓存，只补充不完整的内容；本地检查并发，网页下载顺序执行 | `files/morningstar/morningstar_pages/<ID>/<UK或HK>/*.html`、`*.xlsx` |
| 指标汇总 | `morningstar_3_update_metric.py` | 从页面对应的指标 Excel 读取数据，通过字段别名映射回统一列名 | 更新 `files/morningstar/MorningStar.xlsx` |
| Excel 后处理 | `morningstar_4_post_process.py` | 转换百分比和评级数字、设置颜色、冻结窗格、写入筛选及排序条件 | `files/morningstar/MorningStar_PostProcess.xlsx` |

相同 ISIN 只查询一次，银行来源可能为 `汇丰`、`渣打` 或 `汇丰+渣打`。不同 ISIN 的基金份额不会按名称合并。

## 核心模块

- `global_values.py`：相对路径、UK / HK 网址、页面模板、统一指标清单及字段别名。路径以运行时的工作目录为基准。
- `base_define.py`：页面模板、指标定义、单元格数值及颜色处理规则。
- `moring_star_logic.py`：浏览器会话、登录 Cookie、页面工厂、页面抓取及基金指标合并。文件名中的 `moring` 为现有拼写。
- `tools.py`：HTML 解析、缓存完整性检查、百分比转换及文件读写。
- `morningstar_isin_query.py`：搜索接口及异常重试；调用方负责 UK → HK 回退。
- `download_page.py`：针对单个基金页面的调试入口，执行时会直接触发登录和下载。

UK 当前使用 Compare、Overview、SustainabilitySAL 三类页面；HK 当前使用 HKSearch 页面，切换 Overview、短期 / 长期收益、费用、组合和风险页签，合并基金首行数据。Performance 解析类仍保留，但未加入当前 HK 模板列表。

## 指标与显示规则

| 统一字段 | 含义 | 当前显示规则 |
| --- | --- | --- |
| `Category` / `Total Assets` | 晨星分类 / 基金规模 | 保留来源文本，不统一规模的币种和单位 |
| `1 Yr` / `3 Yr` / `5 Yr` / `10Yr` | 来源页面标示的年化收益 | ≥20% 深绿；≥10% 且 <20% 浅绿 |
| `SD(3yr)` | 三年标准差 | ≤15% 深绿；>15% 且 ≤20% 浅绿 |
| `Total Return After Fees` | UK Compare 页费用后收益字段 | 与收益列使用相同百分比及颜色规则 |
| `star` / `medalist` / `sustainability` | 星级 / Medalist / 可持续性评级 | 当前共用末尾数字提取逻辑；≥5 深绿，4 浅绿 |
| `Return` | 晨星整体收益评价，当前来自 UK | High / Above Average 为深绿 / 浅绿，Low / Below Average 为深红 / 浅红 |
| `Risk` | 晨星整体风险评价，当前来自 UK | Low / Below Average 为深绿 / 浅绿，High / Above Average 为深红 / 浅红 |

百分比文本转换时会除以 100，保留四位小数，并处理 Unicode 减号。颜色阈值为代码中的固定规则，未按基金类别区分。Medalist 与可持续性评级也不能直接当作星级理解；上述表格描述的是当前实现。

最终文件冻结位置为 `D2`，即首行和前三列；设置 `5 Yr`、`10Yr`、`3 Yr` 的降序排序条件。openpyxl 只写入排序元数据，不会实际重排数据行。

## 运行方式

使用 Python 3.10 或更高版本，网页抓取需要 Chrome。建议重新创建本机虚拟环境，不复用其他电脑复制来的 `.venv`。

在项目根目录执行：

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt

.\.venv\Scripts\python.exe hsbc_1_extract_fund_pdfs.py
.\.venv\Scripts\python.exe hsbc_2_generate_isin.py
.\.venv\Scripts\python.exe sc_1_load_funds.py
.\.venv\Scripts\python.exe morningstar_1_search_id.py
.\.venv\Scripts\python.exe morningstar_2_save_fund_pages.py
.\.venv\Scripts\python.exe morningstar_3_update_metric.py
.\.venv\Scripts\python.exe morningstar_4_post_process.py
```

先准备银行原始资料。渣打输入为 `files/sc/funds_table_html_sc.txt`，需要手动更新；汇丰输入为 `files/hsbc/fund-nav.pdf`。首次抓取晨星页面时，需要在浏览器中手动登录对应站点，并在终端按回车保存 Cookie。`files/cookies/` 目录需存在；Cookie 是本机会话文件，不应新增到提交中。

## 缓存行为与已知限制

- 已存在 `QDII.xlsx` 时，不重新解析汇丰目录，只尝试处理标记为“下载异常”的记录。但失败时原下载 URL 会被该标记覆盖，当前重试分支又跳过这一标记，因此不能保证真正重试成功。
- 已存在的单基金 PDF / Excel 会跳过；晨星 ID 的有效匹配会复用，已标记 `NotFound` 的非空记录也可能被复用。更新原始资料或日期不会自动使缓存失效。
- 页面完整性主要检查 HTML 关键词和指标名称是否存在，不检查数据日期，也不保证指标值有效。需要刷新时，应先备份，再移走对应缓存后重新运行。
- 搜索接口只对 `Exception` 重试；非 200 响应会被标记为 `NotFound`。UK 查询异常不会触发 HK 回退。接口请求目前未设置超时。
- `morningstar_2_save_fund_pages.py` 有两个主入口块：批量流程结束后还会执行一次 `F00000PXGY` 的调试抓取。
- `morningstar_3_update_metric.py` 中 `write_morningstar_excel` 的进度总数引用主入口的全局 `fund_list`，直接作为库函数调用可能失败；相同晨星 ID 的汇总当前只更新第一条匹配行。
- 汇丰 PDF 未解析成功时可能返回 `(None, None)`，而调用方按表格处理；部分表格校验也在检查列数前访问第 4 列，异常输入可能中断处理。
- `morningstar_isin_query.py` 的独立演示入口缺少搜索 URL 参数，应通过正式的 `morningstar_1_search_id.py` 使用。
- 仓库中的数据文件为历史快照；日期后缀文件（如 `MorningStar_PostProcess_2025_02_15.xlsx`）为留存结果，默认脚本仍输出无日期后缀的文件名。

以上为阅读当前代码确认的行为，未验证现在线上站点的页面结构和接口可用性。
