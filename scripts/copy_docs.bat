@echo off
chcp 65001 >nul
setlocal EnableExtensions

REM 将源目录中的 Word 文档复制到指定目标目录。
REM 用法： copy_docs.bat <目标目录> [源目录]
REM 示例： copy_docs.bat "f:\目标\路径"
REM        copy_docs.bat "f:\目标\路径" "f:\自定义源路径"

REM 默认源目录（可按需修改）
set "DEFAULT_SRC=f:\笔记本数据备份\code\wordReportCheck\outputs\计科二班\第6周\学生作业"

if "%~1"=="" (
  echo 用法: %~nx0 ^<目标目录^> [源目录]
  echo 示例: %~nx0 "f:\目标\路径"
  echo        %~nx0 "f:\目标\路径" "f:\自定义源路径"
  exit /b 1
)

set "DEST=%~1"
set "SRC=%~2"
if "%SRC%"=="" set "SRC=%DEFAULT_SRC%"

if not exist "%SRC%" (
  echo 源目录不存在: "%SRC%"
  exit /b 2
)

if not exist "%DEST%" (
  echo 目标目录不存在，正在创建: "%DEST%"
  mkdir "%DEST%" 2>nul
)

echo 源目录: "%SRC%"
echo 目标目录: "%DEST%"

REM 使用 robocopy 复制 .doc 与 .docx 文件；返回码>=8 视为失败
robocopy "%SRC%" "%DEST%" *.doc *.docx /COPY:DAT /R:1 /W:1 /NFL /NDL /NP >nul
set "RC=%ERRORLEVEL%"

if %RC% GEQ 8 (
  echo 复制失败，错误码: %RC%
  exit /b %RC%
) else (
  echo 复制完成，返回码: %RC%
  exit /b 0
)