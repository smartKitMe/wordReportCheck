# WordReportCheck

## 功能概述

- 解析统一模板的 `.docx` 实验报告，抽取固定的 18 个信息单元（学院信息、专业信息、时间、姓名、学号、班级、指导老师、课程名称、周次、实验名称、实验环境、实验内容、实验分析与体会、实验日期、备注、成绩、签名、日期）。
- 将“实验内容”中的题目与答案抽取为结构化列表 `content_items`，并可提交到 DeepSeek 进行自动评分。

## 安装

```bash
py -m pip install -e .
```

依赖：`python-docx`、`openai`。

## 使用方法

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

2) 提交到 DeepSeek 进行评分

```bash
# 一次性提交所有题目评分（默认）
wordreportcheck score --doc samples/示例报告.docx --model deepseek-chat --api-key sk-xxx
# 或使用已解析的 JSON：
wordreportcheck score --json outputs/示例报告.json --model deepseek-chat --api-key sk-xxx

# 逐题提交评分（每题单独请求），适合需要逐题日志或重试的场景
wordreportcheck score --json outputs/示例报告.json --per-item --model deepseek-chat --api-key sk-xxx
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

## 设计方案

- 解析流程
  - 使用 `python-docx` 读取 `.docx` 表格。
  - 通过标签映射识别统一模板的 18 个信息单元（学院信息、姓名、课程名称、实验名称、实验内容、实验日期等），按“标签-值”方式抽取。
  - 将“实验内容”聚合为文本后，优先按“题目/问题/任务 + 编号”样式切分为题目列表；若未识别，则回退尝试在全局表格中仅采集包含“题目/问题/任务”关键词的两列问答结构，并排除 18 个固定标签行。
  - 产出 `ReportDocument`（顶层 18 个信息单元 + `content_items`）。

- 序列化
  - `schemas.py` 提供 `report_to_json` / `report_from_json` 将 `ReportDocument` 与 JSON 转换。

- 评分流程
  - `scoring/deepseek_client.py` 将 `ReportDocument.content_items` 构造成评分请求提交给 DeepSeek API，支持模型选择与 API Key 配置。

- CLI 命令
  - `parse`：解析 docx 为 18 个信息单元的报告 JSON（内含 `content_items`）。
  - `score`：解析并提交到 DeepSeek 评分（仅评分 `content_items`），或从现有 JSON 评分。

## 注意事项与建议

- 模板标签建议使用清晰的中文并以冒号结尾，例如“实验名称：xxx”。解析器会自动去掉冒号并进行标签归一化。
- “实验内容”中的题目建议使用“题目X：”或“问题X：”样式便于切分；若存在两列问答表格，亦可被回退逻辑识别。
- 若输出的中文在命令行显示出现乱码，通常是终端编码导致；JSON 文件本身为 UTF-8，可用编辑器或浏览器查看。