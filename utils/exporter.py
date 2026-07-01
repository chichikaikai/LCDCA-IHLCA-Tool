"""
報表匯出模組
"""

import numpy as np
import pandas as pd
from pathlib import Path
from core.calculator import LCAResult


def result_to_dataframes(result: LCAResult, top_n: int = 20):
    """回傳 (summary_df, hotspot_df)"""
    rows = []
    for name, em in zip(result.material_names, result.material_emissions):
        rows.append({"類別": "輸入材料", "名稱": name, "碳排放量 (kg CO₂e)": round(float(em), 6)})
    rows.append({"類別": "產品", "名稱": result.product_name, "碳排放量 (kg CO₂e)": round(result.product_emission, 6)})
    for name, em in zip(result.sector_names, result.sector_emissions):
        rows.append({"類別": "IO 部門", "名稱": name, "碳排放量 (kg CO₂e)": round(float(em), 6)})

    summary_df = pd.DataFrame(rows)

    io_df = pd.DataFrame({
        "部門名稱": result.sector_names,
        "碳排放量 (kg CO₂e)": [round(float(e), 6) for e in result.sector_emissions],
    })
    hotspot_df = (
        io_df.assign(abs_val=io_df["碳排放量 (kg CO₂e)"].abs())
        .sort_values("abs_val", ascending=False)
        .drop(columns="abs_val")
        .head(top_n)
        .reset_index(drop=True)
    )
    return summary_df, hotspot_df


def export_to_excel(result: LCAResult, output, top_n: int = 20):
    """匯出 Excel（output 可為路徑或 BytesIO）"""
    summary_df, hotspot_df = result_to_dataframes(result, top_n)
    overview_df = pd.DataFrame({
        "項目": ["總碳排放量", "過程 LCA 小計", "IO 供應鏈小計", "部門數量"],
        "數值": [round(result.total_emission, 4), round(result.process_total, 4),
                round(result.io_total, 4), len(result.sector_names)],
        "單位": ["kg CO₂e", "kg CO₂e", "kg CO₂e", "個"],
    })
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        overview_df.to_excel(writer, sheet_name="Overview", index=False)
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        hotspot_df.to_excel(writer, sheet_name="Hotspot", index=False)
    return output
