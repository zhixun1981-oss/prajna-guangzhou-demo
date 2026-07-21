#!/usr/bin/env python3
"""Create sample input files for code-agent demos."""
import pandas as pd
from pathlib import Path

OUT = Path(__file__).resolve().parent.parent / "samples"
OUT.mkdir(exist_ok=True)

# Sales data
sales = pd.DataFrame({
    "区域": ["华南", "华南", "华北", "华北", "华东", "华东", "西南", "西南"],
    "销售人员": ["张三", "李四", "王五", "赵六", "陈七", "刘八", "周九", "吴十"],
    "销售额": [450000, 380000, 520000, 290000, 610000, 470000, 310000, 260000],
    "销售目标": [500000, 400000, 500000, 300000, 600000, 500000, 300000, 250000],
})
sales.to_excel(OUT / "sales_data.xlsx", index=False)

# Payroll data
payroll = pd.DataFrame({
    "姓名": ["张伟", "李娜", "王强", "赵敏"],
    "基本工资": [15000, 12000, 18000, 10000],
    "绩效工资": [3000, 2500, 4000, 2000],
    "社保基数": [20000, 15000, 22000, 13000],
    "公积金基数": [20000, 15000, 22000, 13000],
})
payroll.to_excel(OUT / "payroll_data.xlsx", index=False)

# Resumes
resumes = pd.DataFrame({
    "姓名": ["候选人A", "候选人B", "候选人C", "候选人D"],
    "学历": ["本科", "硕士", "大专", "本科"],
    "工作年限": [2, 4, 1, 3],
    "技能": ["电商运营, 淘宝, 数据分析", "数据分析, 用户增长, SQL", "客服, 销售", "电商运营, 天猫, Excel"],
    "项目经验": ["负责天猫店铺运营，月销百万", "搭建用户增长模型", "门店销售", "拼多多店铺操盘"],
    "自我评价": ["熟悉电商平台，数据敏感", "擅长数据驱动", "沟通能力强", "执行力强"],
})
resumes.to_excel(OUT / "resumes.xlsx", index=False)

# Inventory
inventory = pd.DataFrame({
    "SKU": ["SKU001", "SKU002", "SKU003", "SKU004", "SKU005"],
    "库存数量": [1000, 50, 300, 20, 500],
    "单价": [50, 200, 80, 500, 120],
    "月均销量": [300, 5, 100, 2, 150],
    "安全库存": [200, 30, 80, 15, 120],
})
inventory.to_excel(OUT / "inventory.xlsx", index=False)

# KPI
kpi = pd.DataFrame({
    "被考核人": ["张三", "张三", "李四", "李四"],
    "指标": ["销售额", "客户满意度", "销售额", "新客户数"],
    "权重": [0.6, 0.4, 0.7, 0.3],
    "完成率": [1.1, 0.95, 0.85, 1.2],
})
kpi.to_excel(OUT / "kpi_data.xlsx", index=False)

print(f"Samples created in {OUT}")
