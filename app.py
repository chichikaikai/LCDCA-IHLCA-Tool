"""
工程材料碳足跡計算工具
Integrated Hybrid LCA — Streamlit 主程式
"""

import streamlit as st
import numpy as np
import pandas as pd
import plotly.express as px
import io
import logging

from core.calculator import HybridLCACalculator, Material, Product
from utils.data_loader import load_io_data, validate_io_data
from utils.exporter import export_to_excel, result_to_dataframes
from utils.i18n import T

logging.basicConfig(level=logging.INFO)

st.set_page_config(
    page_title="碳足跡計算工具 / Carbon Footprint Calculator",
    page_icon="🌿",
    layout="wide",
)

# ── 語言切換（右上角）──────────────────────────
if "lang" not in st.session_state:
    st.session_state.lang = "zh"

col_title, col_lang = st.columns([10, 1])
with col_title:
    st.title(T("app_title", st.session_state.lang))
    st.caption(T("app_caption", st.session_state.lang))
with col_lang:
    st.write("")
    if st.session_state.lang == "zh":
        if st.button("EN", use_container_width=True):
            st.session_state.lang = "en"
            st.rerun()
    else:
        if st.button("中文", use_container_width=True):
            st.session_state.lang = "zh"
            st.rerun()

lang = st.session_state.lang

EMPTY_ROW = lambda: {"name": "", "unit": "", "sector": 0, "qty": None, "price": None}

if "calculator" not in st.session_state:
    st.session_state.calculator = None
if "result" not in st.session_state:
    st.session_state.result = None
if "sector_options" not in st.session_state:
    st.session_state.sector_options = [f"{i+1:03d}" for i in range(163)]

# 強化版（V2）— 啟動只載入內建 A 矩陣 + 部門名稱（B_A 必須使用者上傳）
if "BA_loaded" not in st.session_state:
    st.session_state.BA_loaded = False
if st.session_state.calculator is None:
    try:
        import numpy as _np_init, json as _json_init
        from pathlib import Path as _Path_init
        _data_dir = _Path_init(__file__).parent / "data"
        _npz_init = _np_init.load(_data_dir / "tw110_io.npz")
        with open(_data_dir / "tw110_sectors.json", encoding="utf-8") as _f_init:
            _sectors_init = _json_init.load(_f_init)
        _A = _npz_init["A"].astype(float)
        # B 留空（zeros），待使用者上傳 B_A 才能計算 — 資訊保護考量
        _B_placeholder = _np_init.zeros(_A.shape[0], dtype=float)
        st.session_state.calculator     = HybridLCACalculator(_A, _B_placeholder, _sectors_init)
        st.session_state.sector_options = _sectors_init
    except Exception as _e_init:
        st.error(f"內建 A 矩陣載入失敗：{_e_init}")
if "raw_rows" not in st.session_state:
    st.session_state.raw_rows = [EMPTY_ROW()]
if "energy_rows" not in st.session_state:
    st.session_state.energy_rows = [EMPTY_ROW()]
if "bpb_data" not in st.session_state:
    st.session_state.bpb_data = None
if "widget_version" not in st.session_state:
    st.session_state.widget_version = 0
# ── 強化版新增 session keys ──
if "nas_path" not in st.session_state:
    st.session_state.nas_path = ""
if "quality_raw_rows" not in st.session_state:
    st.session_state.quality_raw_rows = []
if "quality_energy_rows" not in st.session_state:
    st.session_state.quality_energy_rows = []
if "db_meta" not in st.session_state:
    st.session_state.db_meta = {
        "pcces": "", "product_en": "", "location": "台灣",
        "boundary": "搖籃到大門",
        "exclusions": "無",   # 排除項目（新增；預設「無」）
        "period": "2021/01/01~2021/12/31",
        "process_desc": "",
        # LCA 方法學 — 預設完整 ISO 說明
        "lca_method": (
            "本係數依據 ISO 14067:2018 第 6.3.4 條完整性要求，"
            "採用 Integrated Hybrid LCA 建置。"
            "前景系統以製程 LCA 資料建模，"
            "背景系統以臺灣產業關聯表補足截止門檻外之間接排放；"
            "對於前景製程投入資料不可得之產品，前景系統邊界收斂至功能單位層級，"
            "係數完整由背景 IO 矩陣估算，"
            "仍符合 Hybrid LCA 之系統邊界完整性要求。"
        ),
        "activity_src": "財政部製造業原物料耗用通常水準調查報告、主計總處110年產業關聯統計",
        "emission_src": "能源署110年能源平衡表、環境部溫室氣體排放係數管理表、環境部2023年國家溫室氣體排放清冊",
        "gwp": "IPCC AR6", "org": "LCDCA低碳數位營建聯盟",
        "audit_note": "此係數之建置方法學預計於 2026 年完成第三方之 AUP 查證。",
        "version": "V1.0",
        # 固定為 "IH"（Integrated Hybrid）；不再讓使用者選
        "calc_method": "IH",
        "up_src": "工程會公共工程價格資料庫",
        "quality_reli": 2,   # 數據品質等級可靠性（預設 2）
        "quality_comp": 2,   # 數據品質等級完整性（預設 2）
    }

with st.sidebar:
    st.header(T("sidebar_header", lang))
    from utils.template_generator import create_BA_template, create_userinput_template
    from utils.data_loader import load_BA_data

    # 單一狀態列：依 B_A 是否上傳顯示不同顏色
    if st.session_state.get("BA_loaded", False):
        st.success(T("sidebar_status_BA_ready", lang))
    else:
        st.warning(T("sidebar_status_need_BA", lang))

    # 下載範本（無小標、緊接著放）
    st.download_button(
        label=T("sidebar_dl_BA", lang),
        data=create_BA_template(),
        file_name="B_A填報範本.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    # 上傳區
    ba_file = st.file_uploader(T("sidebar_upload_BA_label", lang),
                               type=["xlsx", "xls"], key="ba_upload",
                               label_visibility="collapsed")
    if ba_file:
        try:
            with st.spinner("讀取中..."):
                BA_vec = load_BA_data(ba_file)
            calc = st.session_state.calculator
            if calc is None:
                st.error("內建 A 矩陣尚未載入，無法套用 B_A")
            elif len(BA_vec) != calc.sectorCount:
                st.warning(T("sidebar_BA_size_warn", lang,
                             n_new=len(BA_vec), n_base=calc.sectorCount))
            else:
                calc.B_IO = BA_vec.astype(float)
                st.session_state.calculator = calc
                st.session_state.BA_loaded = True
                st.success(T("sidebar_BA_upload_ok", lang, n=len(BA_vec)))
        except Exception as e:
            st.error(T("sidebar_BA_upload_fail", lang, e=e))

    # 底部小字註記（內建 A + 資訊保護理由併在一起）
    st.caption(T("sidebar_BA_cap", lang))


def render_material_table(rows_key, prefix):
    rows = st.session_state[rows_key]
    v = st.session_state.widget_version

    col_add, col_clr, _ = st.columns([1, 1, 4])
    with col_add:
        if st.button(T("btn_add", lang), use_container_width=True, key=f"{prefix}_add"):
            rows.append(EMPTY_ROW())
            st.rerun()
    with col_clr:
        if st.button(T("btn_clear", lang), use_container_width=True, key=f"{prefix}_clr"):
            st.session_state[rows_key] = [EMPTY_ROW()]
            st.rerun()

    h1, h2, h3, h4, h5, h6 = st.columns([2, 3, 1, 1.2, 1.5, 0.5])
    h1.markdown(f"**{T('col_name', lang)}**")
    h2.markdown(f"**{T('col_sector', lang)}**")
    h3.markdown(f"**{T('col_unit', lang)}**")
    h4.markdown(f"**{T('col_qty', lang)}**")
    h5.markdown(f"**{T('col_price', lang)}**")
    h6.markdown(f"**{T('col_delete', lang)}**")

    to_delete = []
    for idx, row in enumerate(rows):
        c1, c2, c3, c4, c5, c6 = st.columns([2, 3, 1, 1.2, 1.5, 0.5])
        with c1:
            row["name"] = st.text_input(T("col_name", lang), value=row["name"],
                label_visibility="collapsed", key=f"{prefix}_n{v}_{idx}")
        with c2:
            row["sector"] = st.selectbox(T("col_sector", lang),
                range(len(st.session_state.sector_options)),
                format_func=lambda x: st.session_state.sector_options[x],
                index=min(row["sector"], len(st.session_state.sector_options) - 1),
                label_visibility="collapsed", key=f"{prefix}_s{v}_{idx}")
        with c3:
            row["unit"] = st.text_input(T("col_unit", lang), value=row.get("unit", ""),
                label_visibility="collapsed", placeholder=T("placeholder_unit", lang),
                key=f"{prefix}_u{v}_{idx}")
        with c4:
            row["qty"] = st.number_input(T("col_qty", lang), value=row["qty"],
                label_visibility="collapsed", placeholder=T("placeholder_qty", lang),
                key=f"{prefix}_q{v}_{idx}")
        with c5:
            row["price"] = st.number_input(T("col_price", lang), value=row["price"], min_value=0.0,
                label_visibility="collapsed", placeholder=T("placeholder_price", lang),
                key=f"{prefix}_p{v}_{idx}")
        with c6:
            if st.button("✕", key=f"{prefix}_d{v}_{idx}", use_container_width=True):
                to_delete.append(idx)

    if to_delete:
        st.session_state[rows_key] = [
            r for i, r in enumerate(rows) if i not in to_delete
        ]
        st.rerun()


# ═══════════════════════════════════════════════════════
# 強化版：資料品質評分（Pedigree Matrix）
# ═══════════════════════════════════════════════════════
def render_quality_scores(rows_key, prefix):
    """
    對應 rows_key（raw_rows / energy_rows）所有有名稱的材料，
    顯示 5 維度（Re/Co/Ti/Ge/Te）× 2 資料類型（活動數據 / 排放係數）的評分編輯器
    分數寫回 session_state[f"quality_{rows_key}"]
    """
    rows = st.session_state[rows_key]
    materials = [r for r in rows if (r.get("name") or "").strip()]
    if not materials:
        return

    qkey = f"quality_{rows_key}"
    existing = {q.get("material", ""): q for q in st.session_state.get(qkey, [])}
    new_quality = []
    for m in materials:
        prev = existing.get(m["name"], {})
        new_quality.append({
            "material": m["name"],
            "Re_act": int(prev.get("Re_act", 2) or 2),
            "Co_act": int(prev.get("Co_act", 2) or 2),
            "Ti_act": int(prev.get("Ti_act", 2) or 2),
            "Ge_act": int(prev.get("Ge_act", 2) or 2),
            "Te_act": int(prev.get("Te_act", 2) or 2),
            "Re_em":  int(prev.get("Re_em",  2) or 2),
            "Co_em":  int(prev.get("Co_em",  2) or 2),
            "Ti_em":  int(prev.get("Ti_em",  2) or 2),
            "Ge_em":  int(prev.get("Ge_em",  2) or 2),
            "Te_em":  int(prev.get("Te_em",  2) or 2),
        })
    st.session_state[qkey] = new_quality

    with st.expander(T("quality_expander", lang), expanded=False):
        st.caption(T("quality_caption", lang))
        df = pd.DataFrame(new_quality)
        col_cfg = {
            "material": st.column_config.TextColumn(T("quality_col_material", lang), disabled=True),
            "Re_act": st.column_config.NumberColumn("Re_活動", min_value=1, max_value=5, step=1, help="活動數據-可靠性"),
            "Co_act": st.column_config.NumberColumn("Co_活動", min_value=1, max_value=5, step=1, help="活動數據-完整性"),
            "Ti_act": st.column_config.NumberColumn("Ti_活動", min_value=1, max_value=5, step=1, help="活動數據-時間相關性"),
            "Ge_act": st.column_config.NumberColumn("Ge_活動", min_value=1, max_value=5, step=1, help="活動數據-地理相關性"),
            "Te_act": st.column_config.NumberColumn("Te_活動", min_value=1, max_value=5, step=1, help="活動數據-技術相關性"),
            "Re_em":  st.column_config.NumberColumn("Re_係數", min_value=1, max_value=5, step=1, help="排放係數-可靠性"),
            "Co_em":  st.column_config.NumberColumn("Co_係數", min_value=1, max_value=5, step=1, help="排放係數-完整性"),
            "Ti_em":  st.column_config.NumberColumn("Ti_係數", min_value=1, max_value=5, step=1, help="排放係數-時間相關性"),
            "Ge_em":  st.column_config.NumberColumn("Ge_係數", min_value=1, max_value=5, step=1, help="排放係數-地理相關性"),
            "Te_em":  st.column_config.NumberColumn("Te_係數", min_value=1, max_value=5, step=1, help="排放係數-技術相關性"),
        }
        edited = st.data_editor(
            df, use_container_width=True, hide_index=True,
            column_config=col_cfg, key=f"{prefix}_quality_editor",
        )
        st.session_state[qkey] = edited.to_dict("records")


tab1, tab2, tab3, tab4 = st.tabs([T("tab_input", lang), T("tab_result", lang),
                                  T("tab_detail", lang), T("tab_db", lang)])

with tab1:

    with st.expander(T("import_expander", lang), expanded=False):
        dl_col, up_col = st.columns(2)
        with dl_col:
            st.download_button(
                label=T("import_dl_btn", lang),
                data=create_userinput_template(),
                file_name="投入產出填報範本.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with up_col:
            upload_ui = st.file_uploader(T("import_upload", lang), type=["xlsx"], key="ui_upload")

        if upload_ui:
            try:
                from utils.data_loader import parse_user_input
                prod_imp, raw_imp, energy_imp = parse_user_input(upload_ui)
                n_prod = 1 if prod_imp else 0
                st.success(T("import_success", lang, p=n_prod, r=len(raw_imp), e=len(energy_imp)))
                if st.button(T("import_apply", lang)):
                    if prod_imp:
                        st.session_state["imp_product"] = prod_imp
                    st.session_state.raw_rows = [
                        {"name": m.name, "unit": m.unit, "sector": m.sector_id,
                         "qty": m.quantity, "price": m.price}
                        for m in raw_imp
                    ] or [EMPTY_ROW()]
                    st.session_state.energy_rows = [
                        {"name": m.name, "unit": m.unit, "sector": m.sector_id,
                         "qty": m.quantity, "price": m.price}
                        for m in energy_imp
                    ] or [EMPTY_ROW()]
                    st.session_state.widget_version += 1
                    st.rerun()
            except Exception as e:
                st.error(T("import_fail", lang, e=e))

    st.divider()
    st.subheader(T("prod_header", lang))
    st.caption(T("prod_note", lang))

    imp_prod = st.session_state.get("imp_product", None)
    default_prod_name   = imp_prod["name"]   if imp_prod else T("prod_default", lang)
    default_prod_sector = imp_prod["sector"] if imp_prod else 0
    default_prod_unit   = imp_prod["unit"]   if imp_prod else ""
    default_prod_qty    = imp_prod["qty"]    if imp_prod else 1.0
    default_prod_price  = imp_prod["price"]  if imp_prod else None

    v = st.session_state.widget_version
    c1, c2, c3, c4, c5 = st.columns([2, 3, 1, 1.2, 1.5])
    with c1:
        product_name = st.text_input(T("prod_name", lang), value=default_prod_name, key=f"prod_name_{v}")
    with c2:
        product_sector = st.selectbox(
            T("prod_sector", lang),
            options=list(range(len(st.session_state.sector_options))),
            format_func=lambda x: st.session_state.sector_options[x],
            index=min(default_prod_sector, len(st.session_state.sector_options)-1),
            key=f"prod_sector_{v}",
        )
    with c3:
        product_unit = st.text_input(T("prod_unit", lang), value=default_prod_unit,
                                     placeholder=T("prod_unit", lang), key=f"prod_unit_{v}")
    with c4:
        product_qty = st.number_input(T("prod_qty", lang), min_value=0.0, max_value=1.0,
                                      value=1.0, key=f"prod_qty_{v}")
        if product_qty != 1.0:
            st.warning(T("prod_qty_warn", lang))
    with c5:
        product_price = st.number_input(T("prod_price", lang), min_value=0.0,
                                        value=default_prod_price,
                                        placeholder=T("placeholder_price", lang),
                                        step=1.0, key=f"prod_price_{v}")

    st.divider()
    st.subheader(T("raw_header", lang))
    st.caption(T("raw_note", lang))
    render_material_table("raw_rows", "raw")
    # 強化版：每張材料表後接資料品質評分
    render_quality_scores("raw_rows", "raw")

    st.divider()
    st.subheader(T("energy_header", lang))
    st.caption(T("energy_note", lang))
    render_material_table("energy_rows", "energy")
    render_quality_scores("energy_rows", "energy")

    st.divider()

    if st.button(T("bpb_btn", lang), use_container_width=True):
        if not st.session_state.calculator:
            st.error(T("no_io", lang))
        elif not st.session_state.get("BA_loaded", False):
            st.error(T("no_BA_err", lang))
        else:
            all_rows = st.session_state.raw_rows + st.session_state.energy_rows
            valid = [r for r in all_rows if r["name"] and (r["qty"] or 0) >= 0 and (r["price"] or 0) >= 0]
            if not valid:
                st.error(T("no_mat", lang))
            else:
                calc = st.session_state.calculator
                prod_sid = product_sector + 1
                bpb_rows = []
                for r in valid:
                    sid = r["sector"] + 1
                    price = float(r["price"] or 0)
                    bio = float(calc.B_IO[sid - 1]) if sid - 1 < len(calc.B_IO) else 0.0
                    bpb_rows.append({T("bpb_col_name", lang): r["name"], "BT": round(price * bio, 8)})
                prod_price = float(product_price or 0)
                prod_bio   = float(calc.B_IO[prod_sid - 1]) if prod_sid - 1 < len(calc.B_IO) else 0.0
                bpb_rows.append({T("bpb_col_name", lang): product_name, "BT": round(prod_price * prod_bio, 8)})
                st.session_state.bpb_data = bpb_rows
                st.session_state.result = None

    if st.session_state.bpb_data:
        st.divider()
        st.subheader(T("bpb_header", lang))
        st.caption(T("bpb_caption", lang))
        name_col = T("bpb_col_name", lang)
        edited = st.data_editor(
            pd.DataFrame(st.session_state.bpb_data),
            use_container_width=True,
            hide_index=True,
            disabled=[name_col],
            column_config={"BT": st.column_config.NumberColumn("Bᴛ", format="%.8f")},
            key="bpb_editor",
        )
        st.session_state.bpb_data = edited.to_dict("records")

        st.divider()
        if st.button(T("calc_btn", lang), type="primary", use_container_width=True):
            try:
                with st.spinner("..."):
                    calc = st.session_state.calculator
                    all_rows = st.session_state.raw_rows + st.session_state.energy_rows
                    valid = [r for r in all_rows if r["name"] and (r["qty"] or 0) >= 0 and (r["price"] or 0) >= 0]
                    # ── 強化版：IHLCA 只算 Integrated → 必須至少有一筆 T 資料（qty > 0）──
                    has_T = any((r.get("qty") or 0) > 0 for r in valid)
                    if not has_T:
                        st.error(T("no_T_err", lang))
                        st.stop()
                    product = Product(product_name, product_sector + 1, float(product_price or 0))
                    materials = [
                        Material(name=r["name"], sector_id=r["sector"] + 1,
                                 quantity=float(r["qty"] or 0), price=float(r["price"] or 0),
                                 unit=r.get("unit", ""))
                        for r in valid
                    ]
                    bpb_values = [row["BT"] for row in st.session_state.bpb_data]
                    import numpy as np_bpb
                    custom_Bpb = np_bpb.array(bpb_values, dtype=float)
                    result = calc.calculate_with_custom_bpb(materials, product, custom_Bpb)
                    result.n_raw = len([r for r in st.session_state.raw_rows if r["name"]])
                    st.session_state.result = result
                st.success(T("calc_ok", lang))
            except Exception as e:
                st.error(T("calc_fail", lang, e=e))

with tab2:
    if not st.session_state.result:
        st.info(T("res_no_calc", lang))
    else:
        r = st.session_state.result
        n_raw = getattr(r, "n_raw", len(r.material_names))

        prod_unit_display = st.session_state.get(f"prod_unit_{st.session_state.widget_version}", "")
        prod_qty_display  = st.session_state.get(f"prod_qty_{st.session_state.widget_version}", 1.0)

        info_col1, info_col2, info_col3, info_col4 = st.columns(4)
        info_col1.metric(T("res_prod_name", lang), r.product_name)
        info_col2.metric(T("res_unit", lang), prod_unit_display or "—")
        info_col3.metric(T("res_qty", lang), f"{prod_qty_display}")
        info_col4.metric(T("res_boundary", lang), T("res_boundary_val", lang))

        st.divider()

        k1, k2, k3 = st.columns(3)
        k1.metric(T("res_total", lang), f"{r.total_emission:,.4f} kg CO₂e")
        k2.metric(T("res_pblca", lang), f"{r.process_total:,.4f} kg CO₂e")
        k3.metric(T("res_iolca", lang), f"{r.io_total:,.4f} kg CO₂e")

        st.divider()

        ELEC_SECTOR_ID = 107
        raw_names = r.material_names[:n_raw]
        raw_em    = r.material_emissions[:n_raw]
        eng_names = r.material_names[n_raw:]
        eng_em    = r.material_emissions[n_raw:]

        energy_rows = st.session_state.get("energy_rows", [])
        scope2_em = 0.0
        scope3_eng_em = 0.0
        for i, (name, em) in enumerate(zip(eng_names, eng_em)):
            sid = (energy_rows[i]["sector"] + 1) if i < len(energy_rows) else 0
            if sid == ELEC_SECTOR_ID:
                scope2_em += float(em)
            else:
                scope3_eng_em += float(em)

        scope1 = r.product_emission
        scope2 = scope2_em
        scope3 = sum(float(e) for e in raw_em) + scope3_eng_em + float(r.io_total)

        em_col = T("res_em_col", lang)
        scope_df = pd.DataFrame({
            T("res_scope_col", lang): ["Scope 1", "Scope 2", "Scope 3"],
            T("res_desc_col", lang):  [T("scope1_desc", lang), T("scope2_desc", lang), T("scope3_desc", lang)],
            em_col: [round(scope1, 4), round(scope2, 4), round(scope3, 4)],
        })
        st.subheader(T("res_scope_title", lang))
        st.dataframe(scope_df, use_container_width=True, hide_index=True,
                     column_config={em_col: st.column_config.NumberColumn(format="%.4f")})

        st.divider()

        TRANSPORT_IDX = list(range(120, 125))
        transport_em   = sum(float(r.sector_emissions[i]) for i in TRANSPORT_IDX if i < len(r.sector_emissions))
        manufacture_em = r.product_emission + scope2
        non_transport_io = sum(float(e) for i, e in enumerate(r.sector_emissions) if i not in TRANSPORT_IDX)
        raw_material_em  = sum(float(e) for e in raw_em) + scope3_eng_em + non_transport_io

        lifecycle_df = pd.DataFrame({
            T("lca_stage_col", lang): [T("lca_raw", lang), T("lca_transport", lang), T("lca_manufacture", lang)],
            em_col: [round(raw_material_em, 4), round(transport_em, 4), round(manufacture_em, 4)],
        })
        st.subheader(T("lca_title", lang))
        st.dataframe(lifecycle_df, use_container_width=True, hide_index=True,
                     column_config={em_col: st.column_config.NumberColumn(format="%.4f")})

        st.divider()

        mfg_suffix = "製造" if lang == "zh" else " (Mfg)"
        proc_names = [f"{n}{mfg_suffix}" for n in r.material_names] + [f"{r.product_name}{mfg_suffix}"]
        proc_em    = [float(e) for e in r.material_emissions] + [r.product_emission]

        all_df = pd.DataFrame({
            T("col_name", lang): proc_names + list(r.sector_names),
            em_col: [round(e, 6) for e in proc_em + [float(e) for e in r.sector_emissions]],
        }).sort_values(em_col, ascending=False).reset_index(drop=True)
        all_df.index += 1

        st.subheader(T("hotspot_title", lang))
        top20 = all_df.head(20)
        fig_bar = px.bar(
            top20.sort_values(em_col),
            x=em_col, y=T("col_name", lang), orientation="h",
            color=em_col, color_continuous_scale="Blues",
        )
        fig_bar.update_layout(yaxis_title="", coloraxis_showscale=False, height=520)
        st.plotly_chart(fig_bar, use_container_width=True)
        st.dataframe(all_df.head(20), use_container_width=True,
                     column_config={em_col: st.column_config.NumberColumn(format="%.6f")})

with tab3:
    if not st.session_state.result:
        st.info(T("detail_no_calc", lang))
    else:
        r = st.session_state.result
        n_raw = getattr(r, "n_raw", len(r.material_names))

        raw_names = r.material_names[:n_raw]
        raw_em    = r.material_emissions[:n_raw]
        eng_names = r.material_names[n_raw:]
        eng_em    = r.material_emissions[n_raw:]

        # 建立明細 DataFrame，依照使用者分類
        rows_detail = []
        for name, em in zip(raw_names, raw_em):
            mfg_suffix = "製造" if lang == "zh" else " (Mfg)"
        cat_raw    = T("cat_raw",     lang)
        cat_energy = T("cat_energy",  lang)
        cat_prod   = T("cat_product", lang)
        cat_io     = T("cat_io",      lang)
        em_col     = T("res_em_col",  lang)
        name_col   = T("col_name",    lang)
        cat_col    = T("res_scope_col", lang) if lang == "en" else "類別"

        rows_detail = []
        for name, em in zip(raw_names, raw_em):
            rows_detail.append({cat_col: cat_raw,    name_col: f"{name}{mfg_suffix}", em_col: round(float(em), 6)})
        for name, em in zip(eng_names, eng_em):
            rows_detail.append({cat_col: cat_energy, name_col: f"{name}{mfg_suffix}", em_col: round(float(em), 6)})
        rows_detail.append({cat_col: cat_prod, name_col: f"{r.product_name}{mfg_suffix}", em_col: round(r.product_emission, 6)})
        for name, em in zip(r.sector_names, r.sector_emissions):
            rows_detail.append({cat_col: cat_io, name_col: name, em_col: round(float(em), 6)})

        summary = pd.DataFrame(rows_detail)

        f1, f2, _ = st.columns([1, 1, 2])
        with f1:
            cats = st.multiselect(T("detail_filter", lang),
                                  [cat_raw, cat_energy, cat_prod, cat_io],
                                  default=[cat_raw, cat_energy, cat_prod, cat_io])
        with f2:
            sort_by = st.radio(T("detail_sort", lang),
                               [T("detail_sort_orig", lang), T("detail_sort_desc", lang)],
                               horizontal=True)

        df = summary[summary[cat_col].isin(cats)].copy()
        if sort_by == T("detail_sort_desc", lang):
            df = df.sort_values(em_col, ascending=False)

        st.caption(T("detail_count", lang, n=len(df), nz=int((df[em_col] != 0).sum())))
        st.dataframe(df, use_container_width=True, hide_index=True,
                     column_config={em_col: st.column_config.NumberColumn(format="%.6f")})


# ═════════════════════════════════════════════════════════
# Tab 4 — 強化版：產出 DB 上傳格式 Excel
# ═════════════════════════════════════════════════════════
with tab4:
    if not st.session_state.result:
        st.info(T("res_no_calc", lang))
    else:
        r = st.session_state.result
        # 從 widget_version 取得使用者填的單位
        prod_unit_display = st.session_state.get(
            f"prod_unit_{st.session_state.widget_version}", ""
        )

        st.subheader(T("db_section", lang))
        st.caption(T("db_caption", lang))

        meta = st.session_state.db_meta
        # 從 Tab 1 抓「產品單價」自動帶入
        prod_price_display = st.session_state.get(
            f"prod_price_{st.session_state.widget_version}", 0.0
        ) or 0.0

        with st.expander(T("db_meta_header", lang), expanded=True):
            m_col1, m_col2 = st.columns(2)
            with m_col1:
                meta["pcces"]      = st.text_input(T("db_pcces", lang),    value=meta["pcces"],     key="meta_pcces")
                meta["version"]    = st.text_input(T("db_version", lang),  value=meta["version"],   key="meta_version")
                meta["product_en"] = st.text_input(T("db_prod_en", lang),  value=meta["product_en"],key="meta_product_en")
                meta["location"]   = st.text_input(T("db_location", lang), value=meta["location"],  key="meta_location")
                meta["boundary"]   = st.text_input(T("db_boundary", lang), value=meta["boundary"],  key="meta_boundary")
                meta["period"]     = st.text_input(T("db_period", lang),   value=meta["period"],    key="meta_period")
                # 產品單價（自動帶入、可覆寫）
                st.text_input(T("db_product_price", lang),
                              value=f"{float(prod_price_display):g}",
                              key="meta_product_price_display", disabled=True)
            with m_col2:
                meta["gwp"]         = st.text_input(T("db_gwp", lang),    value=meta["gwp"],         key="meta_gwp")
                meta["org"]         = st.text_input(T("db_org", lang),    value=meta["org"],         key="meta_org")
                meta["activity_src"]= st.text_area(T("db_act_src", lang), value=meta["activity_src"],height=70, key="meta_activity_src")
                meta["emission_src"]= st.text_area(T("db_em_src", lang),  value=meta["emission_src"],height=70, key="meta_emission_src")
                meta["audit_note"]  = st.text_area(T("db_audit", lang),   value=meta["audit_note"],  height=70, key="meta_audit_note")
            # 數據品質等級（可靠性、完整性）
            q_col1, q_col2 = st.columns(2)
            with q_col1:
                meta["quality_reli"] = st.number_input(T("db_q_reli", lang),
                                                       min_value=1, max_value=5, step=1,
                                                       value=int(meta.get("quality_reli", 2) or 2),
                                                       key="meta_quality_reli")
            with q_col2:
                meta["quality_comp"] = st.number_input(T("db_q_comp", lang),
                                                       min_value=1, max_value=5, step=1,
                                                       value=int(meta.get("quality_comp", 2) or 2),
                                                       key="meta_quality_comp")
            # LCA 方法學 — 統合含計算方法的完整說明
            meta["lca_method"]   = st.text_area(T("db_method", lang),   value=meta["lca_method"],   height=140, key="meta_lca_method")
            # 製程描述
            meta["process_desc"] = st.text_area(T("db_process", lang), value=meta["process_desc"], key="meta_process_desc")
            # 排除項目
            meta["exclusions"]   = st.text_area(T("db_exclusions", lang),
                                                value=meta.get("exclusions", "無") or "無",
                                                height=70, key="meta_exclusions",
                                                help="盤查中排除的項目與理由（例：辦公室消耗、非重要副產品）")
            st.session_state.db_meta = meta
            # 內部固定為 IH（不再讓使用者選 IO/Integrated）
            meta["calc_method"] = "IH"

        if st.button(T("db_btn", lang), type="primary", use_container_width=True, key="db_export_btn"):
            try:
                from utils.db_exporter import export_db_format, build_coefficient_id
                meta_payload = dict(meta)
                meta_payload["unit"] = prod_unit_display
                meta_payload["product_price"] = float(prod_price_display or 0)
                qs = (st.session_state.get("quality_raw_rows", []) +
                      st.session_state.get("quality_energy_rows", []))
                buf = export_db_format(
                    result=r,
                    metadata=meta_payload,
                    raw_rows=st.session_state.raw_rows,
                    energy_rows=st.session_state.energy_rows,
                    quality_scores=qs,
                    calculator=st.session_state.calculator,
                )
                cid = build_coefficient_id(meta["pcces"], meta["calc_method"], meta["version"])
                st.success(T("db_ok", lang, cid=cid))
                st.download_button(
                    label=T("db_dl", lang),
                    data=buf.getvalue(),
                    file_name=f"{cid}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="db_download_btn",
                )
            except Exception as e:
                st.error(T("db_fail", lang, e=e))
