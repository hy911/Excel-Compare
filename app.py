"""
Excel文件数据对比工具 - 后端API（pandas版本）
根据新UI设计实现的API接口
"""

from io import BytesIO
from pathlib import Path
from typing import Optional

import pandas as pd
import numpy as np

from fastapi import FastAPI, UploadFile, File, HTTPException, Form, Response
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles

app = FastAPI(title="Excel Diff Tool", description="Excel文件数据对比工具")

# 配置CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


def read_excel_file(file_content: bytes, filename: str) -> dict[str, pd.DataFrame]:
    """
    读取Excel文件，支持xls和xlsx格式
    返回: {sheet_name: DataFrame}
    """
    file_ext = Path(filename).suffix.lower()

    try:
        if file_ext == ".xls":
            # 使用xlrd引擎读取旧版xls格式
            excel_file = pd.ExcelFile(BytesIO(file_content), engine="xlrd")
        elif file_ext in [".xlsx", ".xlsm"]:
            # 使用openpyxl引擎读取xlsx格式
            excel_file = pd.ExcelFile(BytesIO(file_content), engine="openpyxl")
        else:
            raise ValueError(f"不支持的文件格式: {file_ext}")

        sheets = {}
        for sheet_name in excel_file.sheet_names:
            # 统一读取为字符串，避免类型转换问题
            df = pd.read_excel(
                excel_file,
                sheet_name=sheet_name,
                dtype=str,
                na_values=[],
                keep_default_na=False,
            )
            # 填充NaN为空字符串
            df = df.fillna("")
            sheets[sheet_name] = df

        return sheets
    except Exception as e:
        raise ValueError(f"读取Excel文件失败: {str(e)}")


def safe_str(val) -> str:
    """安全地将值转换为字符串"""
    if pd.isna(val) or val is None:
        return ""
    return str(val).strip()


def compare_values(val1: str, val2: str, case_sensitive: bool = True) -> bool:
    """比较两个值是否相等"""
    if case_sensitive:
        return val1 == val2
    return val1.lower() == val2.lower()


def compare_dataframes(
    df1: pd.DataFrame,
    df2: pd.DataFrame,
    key_column: Optional[int] = None,
    case_sensitive: bool = True,
) -> dict:
    """
    比较两个DataFrame
    key_column: 关键列索引（用于行匹配），None表示按位置对比
    """
    # 获取表头
    headers1 = (
        [safe_str(c) for c in df1.columns.tolist()] if len(df1.columns) > 0 else []
    )
    headers2 = (
        [safe_str(c) for c in df2.columns.tolist()] if len(df2.columns) > 0 else []
    )

    # 计算最大列数
    max_cols = max(len(headers1), len(headers2))

    # 合并表头（优先使用新文件的表头）
    merged_headers = []
    for i in range(max_cols):
        h1 = headers1[i] if i < len(headers1) else ""
        h2 = headers2[i] if i < len(headers2) else ""
        merged_headers.append(h2 if h2 else h1)

    # 构建结果
    result = {
        "headers": merged_headers,
        "rows": [],
        "summary": {
            "total_rows": 0,
            "added": 0,
            "deleted": 0,
            "modified": 0,
            "unchanged": 0,
        },
    }

    # 辅助函数：获取行数据并补齐列数
    def get_row_values(df: pd.DataFrame, row_idx: int, cols: int) -> list[str]:
        if row_idx >= len(df):
            return [""] * cols
        row = [
            safe_str(df.iloc[row_idx, i]) if i < len(df.columns) else ""
            for i in range(cols)
        ]
        return row

    # 如果指定了关键列，使用关键列匹配
    if key_column is not None and 0 <= key_column < max_cols:
        # 构建索引 {key_value: (row_idx, row_data)}
        index1 = {}
        for idx in range(len(df1)):
            row = get_row_values(df1, idx, max_cols)
            key = row[key_column] if key_column < len(row) else ""
            if key:
                index1[key] = (idx, row)

        index2 = {}
        for idx in range(len(df2)):
            row = get_row_values(df2, idx, max_cols)
            key = row[key_column] if key_column < len(row) else ""
            if key:
                index2[key] = (idx, row)

        # 合并所有key，保持顺序
        all_keys = list(dict.fromkeys(list(index1.keys()) + list(index2.keys())))

        for key in all_keys:
            in1 = key in index1
            in2 = key in index2

            if in1 and in2:
                # 两边都有，比较差异
                row1 = index1[key][1]
                row2 = index2[key][1]

                cells = []
                has_diff = False
                for col_idx in range(max_cols):
                    v1 = row1[col_idx] if col_idx < len(row1) else ""
                    v2 = row2[col_idx] if col_idx < len(row2) else ""

                    if compare_values(v1, v2, case_sensitive):
                        cells.append({"value": v2, "old_value": None, "status": "same"})
                    else:
                        cells.append(
                            {"value": v2, "old_value": v1, "status": "modified"}
                        )
                        has_diff = True

                if has_diff:
                    result["rows"].append({"status": "modified", "cells": cells})
                    result["summary"]["modified"] += 1
                else:
                    result["rows"].append({"status": "same", "cells": cells})
                    result["summary"]["unchanged"] += 1

            elif in1:
                # 只在原文件有，已删除
                row1 = index1[key][1]
                cells = [
                    {"value": v, "old_value": None, "status": "deleted"} for v in row1
                ]
                result["rows"].append({"status": "deleted", "cells": cells})
                result["summary"]["deleted"] += 1

            else:
                # 只在新文件有，新增
                row2 = index2[key][1]
                cells = [
                    {"value": v, "old_value": None, "status": "added"} for v in row2
                ]
                result["rows"].append({"status": "added", "cells": cells})
                result["summary"]["added"] += 1

        result["summary"]["total_rows"] = len(result["rows"])

    else:
        # 按位置对比
        max_rows = max(len(df1), len(df2))

        for row_idx in range(max_rows):
            has_row1 = row_idx < len(df1)
            has_row2 = row_idx < len(df2)

            row1 = get_row_values(df1, row_idx, max_cols)
            row2 = get_row_values(df2, row_idx, max_cols)

            if not has_row1:
                # 新增行
                cells = [
                    {"value": v, "old_value": None, "status": "added"} for v in row2
                ]
                result["rows"].append({"status": "added", "cells": cells})
                result["summary"]["added"] += 1
            elif not has_row2:
                # 删除行
                cells = [
                    {"value": v, "old_value": None, "status": "deleted"} for v in row1
                ]
                result["rows"].append({"status": "deleted", "cells": cells})
                result["summary"]["deleted"] += 1
            else:
                # 比较每个单元格
                cells = []
                has_diff = False
                for col_idx in range(max_cols):
                    v1 = row1[col_idx]
                    v2 = row2[col_idx]

                    if compare_values(v1, v2, case_sensitive):
                        cells.append({"value": v2, "old_value": None, "status": "same"})
                    else:
                        cells.append(
                            {"value": v2, "old_value": v1, "status": "modified"}
                        )
                        has_diff = True

                if has_diff:
                    result["rows"].append({"status": "modified", "cells": cells})
                    result["summary"]["modified"] += 1
                else:
                    result["rows"].append({"status": "same", "cells": cells})
                    result["summary"]["unchanged"] += 1

        result["summary"]["total_rows"] = max_rows

    return result


def compare_excel_files(
    file1_content: bytes,
    file1_name: str,
    file2_content: bytes,
    file2_name: str,
    sheet_name: Optional[str] = None,
    key_column: Optional[str] = None,
    case_sensitive: bool = True,
) -> dict:
    """比较两个Excel文件"""
    # 读取文件
    sheets1 = read_excel_file(file1_content, file1_name)
    sheets2 = read_excel_file(file2_content, file2_name)

    # 解析关键列（支持 A/B/C 或 1/2/3）
    key_col_idx = None
    if key_column:
        key_column = key_column.strip().upper()
        if key_column.isdigit():
            key_col_idx = int(key_column) - 1  # 1-based to 0-based
        elif len(key_column) == 1 and key_column.isalpha():
            key_col_idx = ord(key_column) - ord("A")  # A=0, B=1, ...

    # 确定要对比的sheet
    if sheet_name and sheet_name != "all":
        sheets_to_compare = (
            [sheet_name] if sheet_name in sheets1 or sheet_name in sheets2 else []
        )
    else:
        sheets_to_compare = list(set(sheets1.keys()) | set(sheets2.keys()))

    result = {
        "file1": file1_name,
        "file2": file2_name,
        "sheets": {},
        "sheet_list": sorted(sheets_to_compare),
        "summary": {"total_differences": 0, "added": 0, "deleted": 0, "modified": 0},
    }

    for sname in sorted(sheets_to_compare):
        df1 = sheets1.get(sname, pd.DataFrame())
        df2 = sheets2.get(sname, pd.DataFrame())

        diff = compare_dataframes(df1, df2, key_col_idx, case_sensitive)
        result["sheets"][sname] = diff

        # 汇总统计
        result["summary"]["added"] += diff["summary"]["added"]
        result["summary"]["deleted"] += diff["summary"]["deleted"]
        result["summary"]["modified"] += diff["summary"]["modified"]

    result["summary"]["total_differences"] = (
        result["summary"]["added"]
        + result["summary"]["deleted"]
        + result["summary"]["modified"]
    )

    return result


@app.post("/api/compare")
async def compare_files(
    original: UploadFile = File(..., description="原始Excel文件"),
    compare: UploadFile = File(..., description="要对比的Excel文件"),
    sheet: Optional[str] = Form(None, description="指定Sheet名称，空或'all'表示全部"),
    key_column: Optional[str] = Form(None, description="关键列，如 A 或 1"),
    case_sensitive: bool = Form(True, description="是否区分大小写"),
):
    """对比两个Excel文件"""
    allowed_extensions = [".xls", ".xlsx", ".xlsm"]

    for file in [original, compare]:
        ext = Path(file.filename).suffix.lower()
        if ext not in allowed_extensions:
            raise HTTPException(
                status_code=400,
                detail=f"不支持的文件格式: {ext}，仅支持 {', '.join(allowed_extensions)}",
            )

    try:
        file1_content = await original.read()
        file2_content = await compare.read()

        result = compare_excel_files(
            file1_content,
            original.filename,
            file2_content,
            compare.filename,
            sheet_name=sheet,
            key_column=key_column,
            case_sensitive=case_sensitive,
        )

        return result

    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"对比过程中发生错误: {str(e)}")


@app.post("/api/sheets")
async def get_sheets(file: UploadFile = File(..., description="Excel文件")):
    """获取Excel文件的所有Sheet名称"""
    allowed_extensions = [".xls", ".xlsx", ".xlsm"]
    ext = Path(file.filename).suffix.lower()

    if ext not in allowed_extensions:
        raise HTTPException(
            status_code=400,
            detail=f"不支持的文件格式: {ext}",
        )

    try:
        content = await file.read()
        sheets = read_excel_file(content, file.filename)
        return {"filename": file.filename, "sheets": list(sheets.keys())}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.get("/favicon.ico")
async def favicon():
    """返回空favicon避免404"""
    return Response(status_code=204)


# 挂载静态文件
static_path = Path(__file__).parent / "static"
if static_path.exists():
    app.mount("/static", StaticFiles(directory=str(static_path)), name="static")


@app.get("/", response_class=HTMLResponse)
async def root():
    """返回前端页面"""
    html_path = Path(__file__).parent / "static" / "index.html"
    if html_path.exists():
        return html_path.read_text(encoding="utf-8")
    return HTMLResponse("<h1>Excel Diff Tool</h1><p>静态文件未找到</p>")


if __name__ == "__main__":
    import uvicorn

    print("-" * 50)
    print("[INFO] Excel Diff Tool")
    print("[INFO] 访问地址: http://127.0.0.1:8000")
    print("-" * 50)
    uvicorn.run(app, host="0.0.0.0", port=8000)
