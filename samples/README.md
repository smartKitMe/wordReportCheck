# samples 目录

用于存放需要批改的实验报告示例（.docx）文件。你可以将待解析的 Word 文档放在此目录中，例如：

- `samples/report_example.docx`

配合命令使用示例：

- 解析：`wordreportcheck parse --doc samples/report_example.docx --out outputs/report_example.json`
- 评分：`wordreportcheck score --doc samples/report_example.docx`