"""
生成测试用的Excel文件
"""
import pandas as pd
from pathlib import Path

# 创建测试数据目录
test_dir = Path(__file__).parent / "test_files"
test_dir.mkdir(exist_ok=True)

# 原始数据
original_data = {
    "员工编号": ["E001", "E002", "E003", "E004", "E005"],
    "姓名": ["张三", "李四", "王五", "赵六", "钱七"],
    "部门": ["技术部", "市场部", "技术部", "人事部", "财务部"],
    "薪资": [15000, 12000, 18000, 10000, 13000],
    "入职日期": ["2020-01-15", "2019-06-20", "2021-03-10", "2018-09-01", "2022-07-15"]
}

# 修改后的数据（模拟差异）
modified_data = {
    "员工编号": ["E001", "E002", "E003", "E004", "E006"],  # E005被删除，新增E006
    "姓名": ["张三", "李小四", "王五", "赵六", "孙八"],  # 李四改为李小四
    "部门": ["技术部", "销售部", "技术部", "人事部", "技术部"],  # 市场部改为销售部
    "薪资": [16000, 12000, 18000, 11000, 14000],  # 张三和赵六涨薪
    "入职日期": ["2020-01-15", "2019-06-20", "2021-03-10", "2018-09-01", "2023-01-01"],
    "备注": ["优秀员工", "", "", "", "新员工"]  # 新增列
}

# 创建原始文件
df_original = pd.DataFrame(original_data)
df_original.to_excel(test_dir / "original.xlsx", index=False, engine='openpyxl')
print(f"✅ 创建原始文件: {test_dir / 'original.xlsx'}")

# 创建修改后的文件
df_modified = pd.DataFrame(modified_data)
df_modified.to_excel(test_dir / "modified.xlsx", index=False, engine='openpyxl')
print(f"✅ 创建对比文件: {test_dir / 'modified.xlsx'}")

# 创建多sheet的测试文件
with pd.ExcelWriter(test_dir / "multi_sheet_original.xlsx", engine='openpyxl') as writer:
    df_original.to_excel(writer, sheet_name="员工信息", index=False)
    pd.DataFrame({
        "项目": ["项目A", "项目B"],
        "预算": [100000, 200000]
    }).to_excel(writer, sheet_name="项目预算", index=False)

print(f"✅ 创建多Sheet原始文件: {test_dir / 'multi_sheet_original.xlsx'}")

with pd.ExcelWriter(test_dir / "multi_sheet_modified.xlsx", engine='openpyxl') as writer:
    df_modified.to_excel(writer, sheet_name="员工信息", index=False)
    pd.DataFrame({
        "项目": ["项目A", "项目B", "项目C"],  # 新增项目C
        "预算": [120000, 200000, 80000]  # 项目A预算增加
    }).to_excel(writer, sheet_name="项目预算", index=False)
    pd.DataFrame({
        "指标": ["完成率", "满意度"],
        "数值": ["95%", "88%"]
    }).to_excel(writer, sheet_name="绩效数据", index=False)  # 新增sheet

print(f"✅ 创建多Sheet对比文件: {test_dir / 'multi_sheet_modified.xlsx'}")

print("\n🎉 测试文件创建完成！")
print(f"📁 文件位置: {test_dir.absolute()}")
