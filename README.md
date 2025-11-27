# Excel Diff Tool - Excel数据对比工具

一个高性能的Web版Excel文件数据对比工具，可以清晰展示两个Excel文件之间的数据差异。

## ✨ 功能特性

- **多格式支持**: 兼容 `.xls`、`.xlsx`、`.xlsm` 格式
- **多Sheet对比**: 自动对比所有工作表
- **差异高亮**: 清晰标注新增、删除、修改的单元格
- **多种视图**: 统一视图、分栏视图、仅差异视图
- **高性能**: 基于 pandas 的高效数据处理
- **现代化UI**: 深色主题，响应式设计

## 🚀 快速开始

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

### 2. 启动服务

```bash
python app.py
```

或使用 uvicorn（支持热重载）：

```bash
uvicorn app:app --reload --host 0.0.0.0 --port 8000
```

### 3. 访问工具

打开浏览器访问: http://localhost:8000

## 📁 项目结构

```
excel-diff-tool/
├── app.py                 # FastAPI 后端主程序
├── requirements.txt       # Python 依赖
├── create_test_files.py   # 测试文件生成脚本
├── static/
│   └── index.html        # 前端页面
└── test_files/           # 测试用Excel文件
    ├── original.xlsx
    └── modified.xlsx
```

## 🔧 API 接口

### POST /api/compare

对比两个Excel文件

**请求参数:**
- `original`: 原始Excel文件 (multipart/form-data)
- `compare`: 要对比的Excel文件 (multipart/form-data)

**响应示例:**

```json
{
  "file1": "original.xlsx",
  "file2": "modified.xlsx",
  "sheets": {
    "Sheet1": {
      "name": "Sheet1",
      "status": "modified",
      "diff": {
        "headers": {...},
        "rows": [...],
        "summary": {
          "total_cells": 30,
          "modified_cells": 5,
          "added_rows": 1,
          "removed_rows": 0
        }
      }
    }
  }
}
```

## 📊 差异类型说明

| 状态 | 说明 | 颜色 |
|------|------|------|
| `same` | 无变化 | 默认 |
| `modified` | 内容已修改 | 黄色 |
| `added` | 新增内容 | 绿色 |
| `removed` | 已删除内容 | 红色 |

## 🛠 技术栈

- **后端**: FastAPI + Python
- **Excel处理**: pandas + openpyxl + xlrd
- **前端**: 原生 HTML/CSS/JS
- **字体**: JetBrains Mono + Noto Sans SC

## 📝 使用示例

```python
# 使用Python代码调用对比功能
from app import compare_excel_files

with open('file1.xlsx', 'rb') as f1, open('file2.xlsx', 'rb') as f2:
    result = compare_excel_files(
        f1.read(), 'file1.xlsx',
        f2.read(), 'file2.xlsx'
    )
    
# 处理对比结果
for sheet_name, sheet_data in result['sheets'].items():
    print(f"Sheet: {sheet_name}, Status: {sheet_data['status']}")
```

## 📄 License

MIT License
