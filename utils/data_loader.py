"""
資料載入模組
讀取 Excel，格式完全對應原 VBA 工作表結構
"""

import numpy as np
import pandas as pd
from pathlib import Path
import logging
from core.calculator import Material, Product

logger = logging.getLogger(__name__)


def load_BA_data(file) -> np.ndarray:
    """
    從單分頁 Excel 讀取 Bᴀ 碳強度向量（強化版簡化用）
      B_A 工作表（直欄格式）：
        第 1 列 = 表頭（部門編號 / 部門名稱 / 排放係數）
        第 2~N 列 = 資料，使用者填 C 欄（第 3 欄）數值
    回傳 numpy 1D array
    """
    df = pd.read_excel(file, sheet_name="B_A", header=None)
    if df.shape[0] < 2 or df.shape[1] < 3:
        raise ValueError("B_A 工作表格式不符；應為直欄表（A 編號 / B 名稱 / C 排放係數）")
    # 跳過第 1 列表頭，從第 2 列起取 C 欄（pandas iloc 0-based 為 col 2）
    BA_vec = df.iloc[1:, 2].dropna().values.astype(float)
    if BA_vec.size == 0:
        raise ValueError("B_A 的 C 欄無有效數值，請填入排放係數")
    logger.info("Bᴀ 載入完成：%d 個部門", BA_vec.size)
    return BA_vec


def load_io_data(file) -> tuple:
    """
    從 Excel 讀取 IO 資料：
      IO_A   → 第1列說明，從 A2 開始貼資料
      B_IO   → 第1列說明，從 A2 開始橫向貼資料
      Sector → 第1列說明，從 A2 開始直向貼部門名稱
    """
    df_A   = pd.read_excel(file, sheet_name="IO_A",   header=None)
    df_B   = pd.read_excel(file, sheet_name="B_IO",   header=None)
    df_sec = pd.read_excel(file, sheet_name="Sector", header=None)

    # IO_A：跳過第1列說明，從 index 1 開始
    data_A = df_A.iloc[1:].reset_index(drop=True)
    sectorCount = data_A.shape[0]
    A_matrix = data_A.iloc[:sectorCount, :sectorCount].fillna(0).values.astype(float)

    # B_IO：跳過第1列說明，從 index 1 開始
    try:
        B_IO_vector = df_B.iloc[1, :sectorCount].fillna(0).values.astype(float)
    except Exception:
        raise ValueError("B_IO 工作表讀取失敗，請確認從 A2 開始橫向貼入數值")

    # Sector：跳過第1列說明，從 index 1 開始
    sector_names = df_sec.iloc[1:sectorCount+1, 0].fillna("").astype(str).tolist()

    logger.info("IO 資料載入完成：sectorCount=%d", sectorCount)
    return A_matrix, B_IO_vector, sector_names


def parse_sector_id(sector_text: str) -> int:
    """
    從部門代碼文字取出前3碼整數
    對應 VBA：sectorID = Val(Left(sectorText, 3))
    """
    text = str(sector_text).strip()
    for length in [3, 2, 1]:
        try:
            return int(text[:length])
        except (ValueError, IndexError):
            continue
    raise ValueError(f"無法從部門代碼解析 ID：'{sector_text}'")


def parse_user_input(file) -> tuple:
    """
    從 UserInput 工作表解析投入產出資料
    格式：A=類型, B=名稱, C=對應部門, D=單位, E=數量, F=單價
    類型：產品 / 原物料 / 能資源
    第1列說明、第2列表頭，第3列起為資料

    回傳 (product, raw_materials, energy_materials)
    若 Excel 有部門欄則直接解析部門ID，否則 sector_id=0 由介面選取
    """
    df = pd.read_excel(file, sheet_name="UserInput", header=None)

    product = None
    raw_materials = []
    energy_materials = []

    # 找資料起始列（跳過說明和表頭）
    start_row = 0
    for i in range(min(4, len(df))):
        val = str(df.iloc[i, 0]).strip() if not pd.isna(df.iloc[i, 0]) else ""
        if any(kw in val for kw in ["填報", "請", "類型", "Type"]):
            start_row = i + 1
        elif val in ["產品", "原物料", "能資源"]:
            start_row = i
            break

    has_sector_col = df.shape[1] >= 6  # 6欄含部門，5欄無部門

    for idx in range(start_row, len(df)):
        row = df.iloc[idx]
        mat_type = str(row.iloc[0]).strip() if not pd.isna(row.iloc[0]) else ""
        name     = str(row.iloc[1]).strip() if not pd.isna(row.iloc[1]) else ""

        if not name or name == "nan":
            continue

        if has_sector_col:
            # 6欄：類型, 名稱, 部門, 單位, 數量, 單價
            sector_t = str(row.iloc[2]).strip() if not pd.isna(row.iloc[2]) else ""
            unit_t   = str(row.iloc[3]).strip() if not pd.isna(row.iloc[3]) else ""
            qty      = row.iloc[4]
            price_r  = row.iloc[5] if len(row) > 5 else 0.0
            # 嘗試解析部門 ID，失敗則設 0
            try:
                sector_id = parse_sector_id(sector_t) - 1  # 轉為 0-based index
            except Exception:
                sector_id = 0
        else:
            # 5欄：類型, 名稱, 單位, 數量, 單價
            sector_t  = ""
            sector_id = 0
            unit_t    = str(row.iloc[2]).strip() if not pd.isna(row.iloc[2]) else ""
            qty       = row.iloc[3]
            price_r   = row.iloc[4] if len(row) > 4 else 0.0

        if pd.isna(qty) or str(qty).strip() == "":
            qty = 0.0
        if pd.isna(price_r) or str(price_r).strip() == "":
            price_r = 0.0

        quantity = abs(float(str(qty).replace(",", "")))
        price    = abs(float(str(price_r).replace(",", "")))

        if "產品" in mat_type:
            product = {
                "name": name,
                "sector": sector_id,
                "unit": unit_t,
                "qty": quantity,
                "price": price,
            }
        elif "能資源" in mat_type or "能源" in mat_type:
            energy_materials.append(Material(
                name=name, sector_id=max(sector_id, 0),
                quantity=quantity, price=price, unit=unit_t,
            ))
        else:
            raw_materials.append(Material(
                name=name, sector_id=max(sector_id, 0),
                quantity=quantity, price=price, unit=unit_t,
            ))

    logger.info("UserInput 解析完成：產品=%s，原物料=%d，能資源=%d",
                product["name"] if product else "無",
                len(raw_materials), len(energy_materials))
    return product, raw_materials, energy_materials


def validate_io_data(A_matrix, B_IO, sector_names) -> list:
    """基本資料驗證，回傳警告訊息列表"""
    warnings = []
    sc = A_matrix.shape[0]

    if A_matrix.shape != (sc, sc):
        warnings.append(f"A 矩陣不是正方形：{A_matrix.shape}")
    if B_IO.shape[0] != sc:
        warnings.append(f"B_IO 長度 {B_IO.shape[0]} 與 A 矩陣 {sc} 不一致")
    if len(sector_names) != sc:
        warnings.append(f"部門名稱數量 {len(sector_names)} 與 A 矩陣 {sc} 不一致")
    if int(np.sum(B_IO < 0)) > 0:
        warnings.append(f"B_IO 中有 {int(np.sum(B_IO < 0))} 個負值，請確認資料")

    return warnings
