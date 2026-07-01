"""
Integrated Hybrid LCA 計算核心
完整對應最新版 VBA（動態 sectorCount，不寫死 163）

矩陣維度：
  processCount = n_materials + 1
  totalSize    = processCount + sectorCount
  大矩陣 = [ T    |  0  ]
           [ -Cu  | I-A ]
"""

import numpy as np
import logging

logger = logging.getLogger(__name__)


class Material:
    def __init__(self, name: str, sector_id: int, quantity: float, price: float, unit: str = ""):
        self.name = name
        self.sector_id = sector_id
        self.quantity = quantity
        self.price = price
        self.unit = unit


class Product:
    def __init__(self, name: str, sector_id: int, price: float):
        self.name = name
        self.sector_id = sector_id
        self.price = price


class LCAResult:
    def __init__(self, material_names, material_emissions,
                 product_name, product_emission,
                 sector_names, sector_emissions):
        self.material_names = material_names
        self.material_emissions = np.array(material_emissions)
        self.product_name = product_name
        self.product_emission = float(product_emission)
        self.sector_names = sector_names
        self.sector_emissions = np.array(sector_emissions)

    @property
    def total_emission(self):
        return float(np.sum(self.material_emissions)) + self.product_emission + float(np.sum(self.sector_emissions))

    @property
    def process_total(self):
        return float(np.sum(self.material_emissions)) + self.product_emission

    @property
    def io_total(self):
        return float(np.sum(self.sector_emissions))


class HybridLCACalculator:
    """
    對應 VBA 各 Sub：
      Build_T          → _build_T
      Build_Bpb        → _build_Bpb
      Build_Ccon       → _build_Ccon
      Build_ACcon      → _build_ACcon
      Build_ACconP     → _build_ACconP
      Build_Ta         → _build_Ta
      Build_Cu         → _build_Cu
      Build_Integrated → _build_integrated
      Build_Inverse    → _build_inverse
      Build_Result     → calculate
      GetSectorCount   → 由 A_matrix.shape[0] 動態取得
    """

    def __init__(self, A_matrix: np.ndarray, B_IO: np.ndarray, sector_names: list):
        """
        A_matrix    : IO 技術矩陣，shape (sectorCount, sectorCount)
        B_IO        : 碳強度向量，shape (sectorCount,)  ← 對應 B_IO 第1列
        sector_names: 部門名稱列表，長度 sectorCount
        """
        self.A = A_matrix.astype(float)
        self.B_IO = B_IO.astype(float)
        self.sector_names = sector_names
        self.sectorCount = A_matrix.shape[0]  # 動態，不寫死

        assert self.A.shape == (self.sectorCount, self.sectorCount)
        assert self.B_IO.shape == (self.sectorCount,)
        assert len(sector_names) == self.sectorCount

    # ── Phase 1 ──────────────────────────────

    def _build_T(self, materials):
        """對應 Build_T：製程需求矩陣，對角線=1，最後欄填數量
        使用者輸入正數，此處自動轉為負數（投入材料為負產出）"""
        n = len(materials)
        pc = n + 1
        T = np.eye(pc, dtype=float)
        for i, mat in enumerate(materials):
            T[i, pc - 1] = -abs(mat.quantity)  # 自動轉負數
        return T

    def _build_Bpb(self, materials, product):
        """對應 Build_Bpb：製程碳強度 = price × B_IO[sector]"""
        n = len(materials)
        pc = n + 1
        Bpb = np.zeros(pc, dtype=float)
        for i, mat in enumerate(materials):
            Bpb[i] = mat.price * self.B_IO[mat.sector_id - 1]
        Bpb[pc - 1] = product.price * self.B_IO[product.sector_id - 1]
        return Bpb

    def _build_Ccon(self, materials, product):
        """對應 Build_Ccon：部門對應矩陣，shape (sectorCount, processCount)"""
        n = len(materials)
        pc = n + 1
        Ccon = np.zeros((self.sectorCount, pc), dtype=float)
        for i, mat in enumerate(materials):
            Ccon[mat.sector_id - 1, i] = 1.0
        Ccon[product.sector_id - 1, pc - 1] = 1.0
        return Ccon

    def _build_ACcon(self, Ccon):
        """對應 Build_ACcon：ACcon = A × Ccon"""
        return self.A @ Ccon

    def _build_ACconP(self, ACcon, materials, product):
        """對應 Build_ACconP：ACconP[:,i] = ACcon[:,i] × price_i"""
        n = len(materials)
        pc = n + 1
        ACconP = ACcon.copy()
        for i, mat in enumerate(materials):
            ACconP[:, i] *= mat.price
        ACconP[:, pc - 1] *= product.price
        return ACconP

    def _build_Ta(self, Ccon, T):
        """對應 Build_Ta：Ta = Ccon × T"""
        return Ccon @ T

    def _build_Cu(self, Ta, ACconP):
        """
        對應 Build_Cu（截斷修正，Hybrid LCA 核心）：
          Ta == 0 → 保留 ACconP（IO 未被製程覆蓋）
          Ta != 0 → 設為 0（已被製程 LCA 覆蓋，避免重複計算）
        """
        return np.where(Ta == 0, ACconP, 0.0)

    # ── Phase 2 ──────────────────────────────

    def _build_integrated(self, T, Cu):
        """
        對應 Build_Integrated：
        [ T    |  0  ]
        [ -Cu  | I-A ]
        """
        pc = T.shape[0]
        total = pc + self.sectorCount
        M = np.zeros((total, total), dtype=float)
        M[:pc, :pc] = T                              # 左上：T
        # 右上：0（預設）
        M[pc:, :pc] = -Cu                            # 左下：-Cu
        M[pc:, pc:] = np.eye(self.sectorCount) - self.A  # 右下：I-A
        return M

    def _build_inverse(self, M):
        """對應 Build_Inverse：求大矩陣逆矩陣（對應 MINVERSE）"""
        try:
            return np.linalg.inv(M)
        except np.linalg.LinAlgError as e:
            raise ValueError(f"矩陣奇異，無法求逆。請確認輸入資料。\n錯誤：{e}") from e

    # ── Phase 3 ──────────────────────────────

    def calculate(self, materials: list, product: Product) -> LCAResult:
        """
        對應 Build_Result（含完整計算流程）：
        y = [0...1...0]（第 processCount 個設為 1）
        x = Inverse × y
        碳排 = Bpb × x（製程）+ B_IO × x（IO 部門）
        """
        n = len(materials)
        pc = n + 1
        total = pc + self.sectorCount

        logger.info("開始計算：%s，材料數=%d，sectorCount=%d", product.name, n, self.sectorCount)

        # Phase 1
        T      = self._build_T(materials)
        Bpb    = self._build_Bpb(materials, product)
        Ccon   = self._build_Ccon(materials, product)
        ACcon  = self._build_ACcon(Ccon)
        ACconP = self._build_ACconP(ACcon, materials, product)
        Ta     = self._build_Ta(Ccon, T)
        Cu     = self._build_Cu(Ta, ACconP)

        # Phase 2
        M     = self._build_integrated(T, Cu)
        M_inv = self._build_inverse(M)

        # Phase 3：y 向量
        y = np.zeros(total, dtype=float)
        y[pc - 1] = 1.0
        x = M_inv @ y

        # 碳排放結果（對應 Build_Result 的 4、5、6 段）
        material_emissions = np.array([Bpb[i] * x[i] for i in range(n)])
        product_emission   = Bpb[pc - 1] * x[pc - 1]
        sector_emissions   = np.array([self.B_IO[j] * x[pc + j] for j in range(self.sectorCount)])

        logger.info("計算完成。總碳排=%.4f", float(np.sum(material_emissions)) + product_emission + float(np.sum(sector_emissions)))

        return LCAResult(
            material_names=[m.name for m in materials],
            material_emissions=material_emissions,
            product_name=product.name,
            product_emission=product_emission,
            sector_names=self.sector_names,
            sector_emissions=sector_emissions,
        )

    def calculate_with_custom_bpb(
        self, materials: list, product: Product, custom_Bpb: np.ndarray
    ) -> LCAResult:
        """
        使用使用者修改後的 Bpb 進行計算
        custom_Bpb：shape (processCount,)，順序為 [材料1, 材料2, ..., 產品]
        其餘計算流程與 calculate 完全相同
        """
        n = len(materials)
        pc = n + 1
        total = pc + self.sectorCount

        assert len(custom_Bpb) == pc, \
            f"custom_Bpb 長度 {len(custom_Bpb)} 與 processCount {pc} 不一致"

        T      = self._build_T(materials)
        Ccon   = self._build_Ccon(materials, product)
        ACcon  = self._build_ACcon(Ccon)
        ACconP = self._build_ACconP(ACcon, materials, product)
        Ta     = self._build_Ta(Ccon, T)
        Cu     = self._build_Cu(Ta, ACconP)
        M      = self._build_integrated(T, Cu)
        M_inv  = self._build_inverse(M)

        y = np.zeros(total, dtype=float)
        y[pc - 1] = 1.0
        x = M_inv @ y

        # 用使用者修改後的 Bpb 計算排放
        material_emissions = np.array([custom_Bpb[i] * x[i] for i in range(n)])
        product_emission   = custom_Bpb[pc - 1] * x[pc - 1]
        sector_emissions   = np.array([self.B_IO[j] * x[pc + j] for j in range(self.sectorCount)])

        logger.info("自訂Bpb計算完成。總碳排=%.4f",
                    float(np.sum(material_emissions)) + product_emission + float(np.sum(sector_emissions)))

        return LCAResult(
            material_names=[m.name for m in materials],
            material_emissions=material_emissions,
            product_name=product.name,
            product_emission=product_emission,
            sector_names=self.sector_names,
            sector_emissions=sector_emissions,
        )
