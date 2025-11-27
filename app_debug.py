"""
Excel文件数据对比工具 - 后端API（调试版本）
支持 xls, xlsx 格式的Excel文件对比
"""

import sys

print(f"[DEBUG] Python 版本: {sys.version}")
print(f"[DEBUG] 正在导入依赖...")

try:
    from io import BytesIO

    print("[DEBUG] ✓ io")

    from pathlib import Path

    print("[DEBUG] ✓ pathlib")

    from typing import Any

    print("[DEBUG] ✓ typing")

    import pandas as pd

    print(f"[DEBUG] ✓ pandas {pd.__version__}")

    from fastapi import FastAPI, UploadFile, File, HTTPException
    import fastapi

    print(f"[DEBUG] ✓ fastapi {fastapi.__version__}")

    from fastapi.middleware.cors import CORSMiddleware
    from fastapi.responses import HTMLResponse
    from fastapi.staticfiles import StaticFiles

    print("[DEBUG] ✓ fastapi 子模块")

    import uvicorn

    print(f"[DEBUG] ✓ uvicorn {uvicorn.__version__}")

except ImportError as e:
    print(f"[ERROR] 导入失败: {e}")
    print("[提示] 请运行: pip install fastapi uvicorn pandas openpyxl xlrd python-multipart")
    sys.exit(1)

print("[DEBUG] 所有依赖导入成功！")
print("-" * 50)

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
    """读取Excel文件，支持xls和xlsx格式"""
    file_ext = Path(filename).suffix.lower()

    try:
        if file_ext == ".xls":
            excel_file = pd.ExcelFile(BytesIO(file_content), engine="xlrd")
        elif file_ext in [".xlsx", ".xlsm"]:
            excel_file = pd.ExcelFile(BytesIO(file_content), engine="openpyxl")
        else:
            raise ValueError(f"不支持的文件格式: {file_ext}")

        sheets = {}
        for sheet_name in excel_file.sheet_names:
            df = pd.read_excel(
                excel_file,
                sheet_name=sheet_name,
                dtype=str,
                na_values=[],
                keep_default_na=False,
            )
            sheets[sheet_name] = df

        return sheets
    except Exception as e:
        raise ValueError(f"读取Excel文件失败: {str(e)}")


def safe_value(val: Any) -> str:
    """安全地将值转换为字符串用于比较"""
    if pd.isna(val) or val is None:
        return ""
    return str(val).strip()


def compare_dataframes(df1: pd.DataFrame, df2: pd.DataFrame) -> dict:
    """比较两个DataFrame，返回详细的差异信息"""
    max_rows = max(len(df1), len(df2))
    max_cols = max(len(df1.columns), len(df2.columns))

    df1_cols = list(df1.columns)
    df2_cols = list(df2.columns)

    diff_result = {
        "headers": {"original": df1_cols, "compare": df2_cols, "merged": []},
        "rows": [],
        "summary": {
            "total_cells": 0,
            "modified_cells": 0,
            "added_rows": 0,
            "removed_rows": 0,
            "added_cols": 0,
            "removed_cols": 0,
        },
    }

    merged_headers = []
    for i in range(max_cols):
        h1 = df1_cols[i] if i < len(df1_cols) else None
        h2 = df2_cols[i] if i < len(df2_cols) else None

        if h1 == h2:
            merged_headers.append({"name": str(h1), "status": "same"})
        elif h1 is None:
            merged_headers.append({"name": str(h2), "status": "added"})
            diff_result["summary"]["added_cols"] += 1
        elif h2 is None:
            merged_headers.append({"name": str(h1), "status": "removed"})
            diff_result["summary"]["removed_cols"] += 1
        else:
            merged_headers.append({
                "name": f"{h1} → {h2}",
                "status": "modified",
                "original": str(h1),
                "compare": str(h2),
            })

    diff_result["headers"]["merged"] = merged_headers

    for row_idx in range(max_rows):
        row_data = {"index": row_idx, "status": "same", "cells": []}

        has_row1 = row_idx < len(df1)
        has_row2 = row_idx < len(df2)

        if not has_row1:
            row_data["status"] = "added"
            diff_result["summary"]["added_rows"] += 1
        elif not has_row2:
            row_data["status"] = "removed"
            diff_result["summary"]["removed_rows"] += 1

        row_modified = False
        for col_idx in range(max_cols):
            cell_data = {"col_index": col_idx, "status": "same", "original": "", "compare": ""}

            if has_row1 and col_idx < len(df1.columns):
                cell_data["original"] = safe_value(df1.iloc[row_idx, col_idx])

            if has_row2 and col_idx < len(df2.columns):
                cell_data["compare"] = safe_value(df2.iloc[row_idx, col_idx])

            diff_result["summary"]["total_cells"] += 1

            if not has_row1:
                cell_data["status"] = "added"
                diff_result["summary"]["modified_cells"] += 1
            elif not has_row2:
                cell_data["status"] = "removed"
                diff_result["summary"]["modified_cells"] += 1
            elif col_idx >= len(df1.columns):
                cell_data["status"] = "added"
                diff_result["summary"]["modified_cells"] += 1
            elif col_idx >= len(df2.columns):
                cell_data["status"] = "removed"
                diff_result["summary"]["modified_cells"] += 1
            elif cell_data["original"] != cell_data["compare"]:
                cell_data["status"] = "modified"
                diff_result["summary"]["modified_cells"] += 1
                row_modified = True

            row_data["cells"].append(cell_data)

        if has_row1 and has_row2 and row_modified:
            row_data["status"] = "modified"

        diff_result["rows"].append(row_data)

    return diff_result


def compare_excel_files(file1_content: bytes, file1_name: str, file2_content: bytes, file2_name: str) -> dict:
    """比较两个Excel文件的所有sheet"""
    sheets1 = read_excel_file(file1_content, file1_name)
    sheets2 = read_excel_file(file2_content, file2_name)

    all_sheets = set(sheets1.keys()) | set(sheets2.keys())

    result = {
        "file1": file1_name,
        "file2": file2_name,
        "sheets": {},
        "summary": {
            "total_sheets": len(all_sheets),
            "same_sheets": 0,
            "modified_sheets": 0,
            "added_sheets": 0,
            "removed_sheets": 0,
        },
    }

    for sheet_name in sorted(all_sheets):
        sheet_info = {"name": sheet_name, "status": "same", "diff": None}

        if sheet_name not in sheets1:
            sheet_info["status"] = "added"
            result["summary"]["added_sheets"] += 1
            empty_df = pd.DataFrame()
            sheet_info["diff"] = compare_dataframes(empty_df, sheets2[sheet_name])
        elif sheet_name not in sheets2:
            sheet_info["status"] = "removed"
            result["summary"]["removed_sheets"] += 1
            empty_df = pd.DataFrame()
            sheet_info["diff"] = compare_dataframes(sheets1[sheet_name], empty_df)
        else:
            diff = compare_dataframes(sheets1[sheet_name], sheets2[sheet_name])
            sheet_info["diff"] = diff

            if diff["summary"]["modified_cells"] > 0:
                sheet_info["status"] = "modified"
                result["summary"]["modified_sheets"] += 1
            else:
                result["summary"]["same_sheets"] += 1

        result["sheets"][sheet_name] = sheet_info

    return result


@app.post("/api/compare")
async def compare_files(
        original: UploadFile = File(..., description="原始Excel文件"),
        compare: UploadFile = File(..., description="要对比的Excel文件"),
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
            file1_content, original.filename, file2_content, compare.filename
        )

        return result

    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"对比过程中发生错误: {str(e)}")


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
    print("[DEBUG] 进入 __main__ 块")
    print(f"[DEBUG] __name__ = {__name__}")
    print(f"[DEBUG] 当前文件: {__file__}")
    print("-" * 50)
    print("[INFO] 正在启动服务器...")
    print("[INFO] 访问地址: http://127.0.0.1:8002")
    print("-" * 50)

    try:
        uvicorn.run(
            app,
            host="0.0.0.0",
            port=8002,
            log_level="info"  # 显示详细日志
        )
    except Exception as e:
        print(f"[ERROR] 服务器启动失败: {e}")
        import traceback

        traceback.print_exc()