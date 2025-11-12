# WordReportCheck

## 功能概述

- 解析统一模板的 `.docx` 实验报告，抽取固定的 18 个信息单元（学院信息、专业信息、时间、姓名、学号、班级、指导老师、课程名称、周次、实验名称、实验环境、实验内容、实验分析与体会、实验日期、备注、成绩、签名、日期）。
- 将“实验内容”中的题目与答案抽取为结构化列表 `content_items`，并可提交到 DeepSeek 进行自动评分。

## 使用准备

- 安装 Python 3.9+（建议 3.10/3.11）。
- 在项目根目录安装依赖并注册命令：

```bash
py -m pip install -e .
```

- 依赖：`python-docx`、`openai`。首次安装会自动拉取。
- Windows 终端建议使用 PowerShell；路径既支持 `\` 也支持 `/`。
- 样例文档位于 `samples/`；输出默认写入同目录或你指定的 `--out-dir`。

 

## 使用方法

### 一键执行：解析 + 评分 + 写回（auto）

无需记忆复杂参数，只需提供 DOCX 路径即可：

```bash
py -m wordreportcheck auto --doc samples/示例报告.docx
# 可选：直接传入 API Key，避免配置 .env
py -m wordreportcheck auto --doc samples/示例报告.docx --api-key sk-xxxx
# 可选：指定输出目录（解析 JSON 与评分明细会写入此目录）
py -m wordreportcheck auto --doc samples/示例报告.docx --out-dir outputs
```

执行后将自动完成以下步骤：
- 解析 DOCX 并写出同名 JSON：`samples/示例报告.json`
- 调用评分服务对“实验内容”中的题目进行批量评分
- 写出评分明细 JSON：`samples/示例报告.scores.json`
- 计算平均分并写入解析 JSON 顶层字段 `成绩`
- 将平均分写回 DOCX 的“成绩/日期”单元格（不会修改源文件，会复制到输出目录后写入）

若使用 `--out-dir`，上述两个 JSON 将改为写入指定目录：
- `<out-dir>/<stem>.json`
- `<out-dir>/<stem>.scores.json`
且 DOCX 会被复制到 `<out-dir>/<filename>.docx` 并在该复制件上写回成绩。

环境变量配置（简化使用）：
- `DEEPSEEK_API_KEY` 或 `MOONSHOT_API_KEY`：评分服务的 API Key（必需其一）
- 也可通过 `--api-key` 临时传入密钥（优先级高于环境变量）
- `WORDREPORTCHECK_PROVIDER`：评分服务提供者，`deepseek`（默认）或 `kimi`
- `WORDREPORTCHECK_MODEL`：评分模型（可选；未设置时 deepseek 用 `deepseek-chat`，kimi 用 `moonshot-v1-128k`）
- 可在项目根或当前目录放置 `.env` 文件，格式：`KEY=VALUE` 或 `export KEY=VALUE`

示例 `.env`：

```
export WORDREPORTCHECK_PROVIDER=deepseek
export DEEPSEEK_API_KEY=sk-xxx
```

运行完成后会在控制台打印简要总结，包含输入文件、输出文件、所用 provider/model、平均分等信息。

### 批量处理：遍历目录并自动评分（auto-dir）

一次性处理目录中的所有 `.docx`，为每份文档生成解析结果与评分明细，并在复制件 DOCX 中写回成绩。

```bash
# 基本用法：遍历目录并输出到指定目录
py -m wordreportcheck auto-dir --in-dir samples/学生作业2 --out-dir outputs/2 --segment-count 2

# 递归处理子目录（例如遍历 samples 下所有子目录）
py -m wordreportcheck auto-dir --in-dir samples --out-dir outputs/all --recursive

# 指定评分服务/模型与密钥（覆盖 .env 设置）
py -m wordreportcheck auto-dir --in-dir samples/学生作业2 --out-dir outputs/2 \
  --provider deepseek --model deepseek-chat --api-key <DEEPSEEK_API_KEY>

py -m wordreportcheck auto-dir --in-dir samples\1班\学生作业 --out-dir outputs\1班\学生作业 --segment-count 4 
py -m wordreportcheck auto-dir --in-dir samples\2班\学生作业 --out-dir outputs\2班\学生作业 --segment-count 4 
py -m wordreportcheck auto-dir --in-dir samples\3班\学生作业 --out-dir outputs\3班\学生作业 --segment-count 4 
py -m wordreportcheck auto-dir --in-dir samples\4班\学生作业 --out-dir outputs\4班\学生作业 --segment-count 4 
py -m wordreportcheck auto-dir --in-dir samples\5班\学生作业 --out-dir outputs\5班\学生作业 --segment-count 4 
py -m wordreportcheck auto-dir --in-dir samples\网工\2025_11_04_17_52_12 --out-dir outputs\网工\2025_11_04_17_52_12 --segment-count 4 



```

输出内容说明（写入到 `--out-dir`）：
- `<stem>.json`：解析后的报告 JSON（顶层 18 个信息单元，含 `content_items`）。
- `<stem>.scores.json`：评分明细（逐题评分的数组）。
- `<filename>.docx`：复制原始 DOCX 并在复制件上写入“成绩/日期”。
- `grades.csv`：汇总所有文件的平均分（含输入/输出路径等）。

参数与行为说明：
- 当“实验内容”中未识别到题目时，可用 `--segment-count <N>` 按本地规则切分为 N 题（示例使用 2）。
- 默认忽略临时锁文件：以 `~$` 开头的 `.docx` 不参与处理。
- 支持 `--provider deepseek|kimi`、`--model ...`、`--api-key ...` 用于覆盖 `.env` 中的默认配置。
- 单个文件失败不会中断批量任务，执行结束会打印汇总统计（`processed/failed/csv`）。

1) 解析 docx 为包含 18 个单元的 JSON（`ReportDocument`）

```bash
wordreportcheck parse --doc samples/示例报告.docx --out outputs/示例报告.json
# 如果 PATH 中找不到命令，可用模块方式运行：
py -m wordreportcheck parse --doc samples/示例报告.docx --out outputs/示例报告.json
```

输出 JSON 结构示例：

```json
{
  "学院信息": "信息工程学院",
  "专业信息": "软件工程",
  "时间": "2024年上学期",
  "姓名": "张三",
  "学号": "24064xxxxx",
  "班级": "1班",
  "指导老师": "李老师",
  "课程名称": "Java 程序设计",
  "周次": "第3-5周",
  "实验名称": "Java 基础与控制结构",
  "实验环境": "JDK 21 + Windows",
  "实验内容": {
    "raw": "实验内容原文...",
    "items": [
      {"id": "Q1", "question": "题目1：数组的基本操作", "answer": "……"},
      {"id": "Q2", "question": "题目2：条件与循环", "answer": "……"}
    ]
  },
  "实验分析与体会": "……",
  "实验日期": "2024-03-21",
  "备注": "……",
  "成绩": "……",
  "签名": "……",
  "日期": "2024-03-22"
}
```

2) 提交到评分服务进行评分（支持 DeepSeek 与 Kimi/Moonshot）

```bash
# DeepSeek：一次性提交所有题目评分（默认）
wordreportcheck score --provider deepseek --doc samples/示例报告.docx --model deepseek-chat --api-key <DEEPSEEK_API_KEY>
# 或使用已解析的 JSON：
wordreportcheck score --provider deepseek --json outputs/示例报告.json --model deepseek-chat --api-key <DEEPSEEK_API_KEY>

# DeepSeek：逐题提交评分（每题单独请求）
wordreportcheck score --provider deepseek --json outputs/示例报告.json --per-item --model deepseek-chat --api-key <DEEPSEEK_API_KEY>

# Kimi/Moonshot：一次性提交所有题目评分
wordreportcheck score --provider kimi --doc samples/示例报告.docx --model moonshot-v1-128k --api-key <MOONSHOT_API_KEY>
# 或使用已解析的 JSON：
wordreportcheck score --provider kimi --json outputs/示例报告.json --model moonshot-v1-128k --api-key <MOONSHOT_API_KEY>

# Kimi/Moonshot：逐题提交评分
wordreportcheck score --provider kimi --json outputs/示例报告.json --per-item --model moonshot-v1-128k --api-key <MOONSHOT_API_KEY>

# 写回成绩到JSON顶层“成绩”字段（取平均分）
wordreportcheck score --provider deepseek --json outputs/示例报告.json --write-back --api-key <DEEPSEEK_API_KEY>
wordreportcheck score --provider kimi     --json outputs/示例报告.json --write-back --api-key <MOONSHOT_API_KEY>

# 写回成绩到 DOCX（将平均分写入“成绩/日期”单元格）
# 方式一：评分时直接写回 DOCX（当使用 --doc 输入时）
wordreportcheck score --provider deepseek --doc samples/示例报告.docx --write-docx --api-key <DEEPSEEK_API_KEY>
wordreportcheck score --provider kimi     --doc samples/示例报告.docx --write-docx --api-key <MOONSHOT_API_KEY>

# 方式二：从已有 JSON 写回 DOCX（JSON 顶层需已有“成绩”）
wordreportcheck write-docx --doc samples/示例报告.docx --json outputs/示例报告.json
```

评分输出为一个 JSON 数组，每个元素对应一个题目：

```json
[
  {
    "id": "Q1",
    "question": "题目1：数组的基本操作",
    "answer": "……",
    "score": 0.85,
    "feedback": "答案覆盖关键点，但示例不足。"
  }
 ]
```

## 环境配置（.env）

- 在项目根目录或当前工作目录创建 `.env` 文件，即可为 CLI 提供默认参数与密钥。
- 已内置 `.env` 示例（仓库根目录），你可以直接填写其中的值：

```
# 评分服务提供者（deepseek 或 kimi）
WORDREPORTCHECK_PROVIDER=deepseek

# 默认模型（deepseek 默认 deepseek-chat；kimi 默认 moonshot-v1-128k）
WORDREPORTCHECK_MODEL=deepseek-chat

# 通用 API Key（与下面的专用 Key 二选一）
WORDREPORTCHECK_API_KEY=

# DeepSeek / Kimi 专用 API Key
DEEPSEEK_API_KEY=
MOONSHOT_API_KEY=

# 逐题评分与写回成绩
WORDREPORTCHECK_PER_ITEM=false
WORDREPORTCHECK_WRITE_BACK=false

# 默认输入（与命令行 --doc / --json 二选一）
WORDREPORTCHECK_DOC=
WORDREPORTCHECK_JSON=

# 可选：覆盖默认 Base URL（一般不需要改动）
DEEPSEEK_BASE_URL=https://api.deepseek.com/v1
MOONSHOT_BASE_URL=https://api.moonshot.cn/v1
```

### 生效与优先级
- CLI 在启动时会自动加载 `.env`（优先从当前工作目录读取，其次读取项目根目录）。
- 命令行参数 > 环境变量/`.env` 默认值。也就是说，显式传入的参数会覆盖 `.env` 中的设置。
- API Key 查找顺序：`--api-key` → 专用提供者环境变量（`DEEPSEEK_API_KEY`/`MOONSHOT_API_KEY`） → `WORDREPORTCHECK_API_KEY`。

### 示例：使用 .env 简化命令
- 在 `.env` 中设置：
  - `WORDREPORTCHECK_PROVIDER=kimi`
  - `MOONSHOT_API_KEY=sk-xxxx`
  - `WORDREPORTCHECK_PER_ITEM=true`
  - `WORDREPORTCHECK_WRITE_BACK=true`
  - `WORDREPORTCHECK_JSON=outputs/示例报告.json`
- 然后仅需执行：

```bash
wordreportcheck score --model moonshot-v1-128k
```

若未设置 `WORDREPORTCHECK_MODEL`，CLI 会根据提供者自动使用默认模型（deepseek → `deepseek-chat`，kimi → `moonshot-v1-128k`）。

## 设计方案

- 解析：读取 `.docx` 表格，识别 18 个固定信息单元；“实验内容”按统一样式聚合为原文。
- 序列化：提供 JSON 转换；`实验内容` 同时包含 `raw` 与 `items`。
- 评分：仅对 `content_items` 评分，可一次性或逐题执行；支持 DeepSeek 与 Kimi。

## 注意事项与建议

- 模板标签建议使用清晰的中文并以冒号结尾，例如“实验名称：xxx”。解析器会自动去掉冒号并进行标签归一化。
- “实验内容”中的题目建议使用“题目X：”或“问题X：”样式便于切分；若存在两列问答表格，亦可被回退逻辑识别。
- 若输出的中文在命令行显示出现乱码，通常是终端编码导致；JSON 文件本身为 UTF-8，可用编辑器或浏览器查看。

## 命令速览

- `parse`：解析 `.docx` 为 18 个信息单元的 JSON。
- `score`：对 `content_items` 评分；支持 `--per-item`、`--write-back`、`--write-docx`。
- `auto`：单文件一键解析→评分→写回。
- `auto-dir`：批量解析评分并生成 `grades.csv`。
- `write-docx`：从已有 JSON 的“成绩”写回到 DOCX。

## 写回 DOCX 的规则

- 查找包含“成绩”的行，优先写入其右侧相邻的可见值单元格；若标签位于该行最后一列，则写回前一列。
- 查找包含“日期”的行，写入其右侧值单元格；未找到时保持原值或跳过。
- 成绩写回在复制的 DOCX 上进行，源文件不修改。

## AI 分割与题目结构

- 当解析阶段未识别到题目结构，可在命令中指定 `--segment-count N`，按本地规则分割为 N 个题目。
- 分割逻辑支持标签“题目要求/题目、实验方法和步骤/方法和步骤、代码”，并忽略“运行结果”等段落；数量不一致时会补齐或截断。

## 脚本工具（scripts/）

- `scripts\copy_docs.bat`：从源目录复制 `.doc/.docx` 到目标目录，便于组织批量评分输入。
  - 用法：`scripts\copy_docs.bat <目标目录> [源目录]`，默认源目录可在脚本内修改。
- `scripts\dump_labels.py`：调试 DOCX 中“成绩/日期”所在行与值单元格位置，定位模板问题。
- `scripts\force_write_grade_date.py`：在指定 DOCX 强制写入成绩与日期，绕过模板不规范导致的写回失败。
- `scripts\validate_outputs.py`：遍历 `samples/` 解析为 JSON 并校验结构，输出汇总与问题定位。

## 常见问题（Troubleshooting）

- 未找到 API Key：通过 `--api-key` 或设置环境变量 `DEEPSEEK_API_KEY`/`MOONSHOT_API_KEY`/`WORDREPORTCHECK_API_KEY`。
- 未识别到题目：检查“实验内容”是否含“题目X：”格式，或使用 `--segment-count` 进行本地分割。
- 分割数量不一致：系统会提示并自动补齐/截断。
- 写回 DOCX 失败：模板未包含可写入的“成绩/日期”单元格或文件不可写；可用 `scripts\dump_labels.py` 定位或 `scripts\force_write_grade_date.py` 试写。

## 示例工作流

- 单份作业：
  - `py -m wordreportcheck auto --doc samples/示例报告.docx --out-dir outputs --provider deepseek`
  - 查看 `outputs/示例报告.json` 与 `outputs/示例报告.scores.json`，成绩已写回到 `outputs/示例报告.docx`。
- 批量作业（递归）：
  - `py -m wordreportcheck auto-dir --in-dir samples --out-dir outputs/all --recursive --segment-count 6`
  - 汇总成绩位于 `outputs/all/grades.csv`，每份文档均生成同名 JSON 和 `.scores.json`。