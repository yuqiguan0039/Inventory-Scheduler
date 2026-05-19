import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="Smart Inventory Allocation System V8.6", layout="wide")

st.title("📦 Smart Inventory Allocation System V8.6")

# --- 1. Sidebar Configuration ---
st.sidebar.header("⚙️ Warehouse Allocation Configuration")

default_order = ["TX", "PHX", "ID", "IL"]
user_priority = []

for i in range(1, 5):
    target_wh = st.sidebar.selectbox(
        f"Priority {i}",
        ["Default"] + default_order,
        key=f"p_{i}"
    )

    if target_wh != "Default" and target_wh not in user_priority:
        user_priority.append(target_wh)

final_priority = user_priority + [w for w in default_order if w not in user_priority]

st.sidebar.write("Manual Allocation Order:", " -> ".join(final_priority))

enable_variant_match = st.sidebar.checkbox("Enable Variant Matching (YN/MG)", value=True)
variant_prefixes = ["YN", "MG"]


# --- 2. Demand Input Section ---
st.subheader("Step 1: Input Order Demand")

input_method = st.radio(
    "Input Method",
    ["Extract from Text", "Manual Table Input", "Upload File"],
    horizontal=True
)

df_demand = pd.DataFrame(columns=["SKU", "Quantity"])
original_sku_order = []  # 记录原始输入的 SKU 顺序

if input_method == "Extract from Text":
    raw_text = st.text_area("Paste Raw Text Here:", height=200)

    if raw_text:
        matches = re.finditer(r"SKU:\s*(?P<sku>[\w\-+]+)", raw_text)
        extracted = []
        all_skus = list(matches)

        for i, m in enumerate(all_skus):
            sku = m.group("sku")
            end_pos = m.end()
            next_start = all_skus[i + 1].start() if i + 1 < len(all_skus) else len(raw_text)

            sub_text = raw_text[end_pos:next_start]
            qty_m = re.search(r"(?<!\$)(?P<qty>\d+\.?\d*)", sub_text)

            if qty_m:
                sku_clean = str(sku).strip()
                extracted.append({
                    "SKU": sku_clean,
                    "Quantity": float(qty_m.group("qty"))
                })
                if sku_clean not in original_sku_order:
                    original_sku_order.append(sku_clean)

        if extracted:
            df_demand = pd.DataFrame(extracted)
            st.success(f"✅ Successfully extracted {len(df_demand)} demand records")
            st.dataframe(df_demand, use_container_width=True)

elif input_method == "Manual Table Input":
    df_demand = st.data_editor(
        pd.DataFrame([{"SKU": "", "Quantity": 0.0}] * 10),
        num_rows="dynamic",
        use_container_width=True
    )
    # 动态抓取非空的原始顺序
    if not df_demand.empty:
        for s in df_demand["SKU"].dropna():
            s_clean = str(s).strip()
            if s_clean != "" and s_clean not in original_sku_order:
                original_sku_order.append(s_clean)

elif input_method == "Upload File":
    f = st.file_uploader("Upload Demand File", type=["xlsx", "csv"])

    if f:
        df_demand = pd.read_csv(f) if f.name.endswith(".csv") else pd.read_excel(f)
        df_demand.columns = [str(c).strip() for c in df_demand.columns]

        s_col = st.selectbox("Select SKU Column", df_demand.columns)
        q_col = st.selectbox("Select Quantity Column", df_demand.columns)

        df_demand = df_demand.rename(columns={
            s_col: "SKU",
            q_col: "Quantity"
        })
        if not df_demand.empty:
            for s in df_demand["SKU"].dropna():
                s_clean = str(s).strip()
                if s_clean not in original_sku_order:
                    original_sku_order.append(s_clean)


# --- Helper Functions ---
def get_sku_detail(sku):
    sku = str(sku).strip().upper() if pd.notnull(sku) else ""

    if enable_variant_match:
        for idx, v in enumerate(variant_prefixes, start=1):
            for p in ["C-", "B-"]:
                if sku.startswith(p + v):
                    return p + sku[len(p) + len(v):], idx, sku

    return sku, 0, sku


def safe_display_df(df):
    if df is None or df.empty:
        return df

    df_show = df.copy()
    for c in df_show.columns:
        if df_show[c].dtype == "object":
            df_show[c] = df_show[c].apply(
                lambda x: "" if pd.isna(x) else str(x)
            )
    return df_show


# 新增：给触发变体（SKU不一致）的行高亮标红
def style_variant_rows(df):
    def make_style(row):
        # 如果原始需求SKU（通过Match_SKU间接对应）和实际出货SKU不一致，则标红
        # 这里为了稳妥，我们在分配指南里保存了“原始需求 SKU”作为辅助判断
        if "Original Demand SKU" in row.index and "Order SKU" in row.index:
            if str(row["Original Demand SKU"]).strip().upper() != str(row["Order SKU"]).strip().upper():
                return ['background-color: #FFECEC; color: #CC0000; font-weight: bold;'] * len(row)
        return [''] * len(row)
    
    if df.empty:
        return df
    return df.style.apply(make_style, axis=1)


# --- 3. Allocation Logic ---
st.divider()

inv_f = st.file_uploader("Upload Inventory File (Kingdee)", type=["xlsx", "csv"])

if inv_f and not df_demand.empty:
    try:
        # --- Load Inventory File ---
        df_inv = pd.read_csv(inv_f) if inv_f.name.endswith(".csv") else pd.read_excel(inv_f)
        df_inv.columns = [str(c).strip() for c in df_inv.columns]

        # --- 需求 1：智能/手动列名映射 ---
        st.subheader("🔍 Inventory Column Mapping")
        
        # 尝试自动寻找默认列名
        default_wh_col = "Warehouse Name" if "Warehouse Name" in df_inv.columns else df_inv.columns[0]
        default_sku_col = "SKU" if "SKU" in df_inv.columns else df_inv.columns[0]
        
        if "Available For Sale" in df_inv.columns:
            default_qty_col = "Available For Sale"
        elif "Accounting Available For Sale" in df_inv.columns:
            default_qty_col = "Accounting Available For Sale"
        else:
            default_qty_col = df_inv.columns[0]

        # 创建三个可选的映射项（让用户在表格格式改变时可以手动选）
        col_wh = st.selectbox("Warehouse Column:", df_inv.columns, index=list(df_inv.columns).index(default_wh_col))
        col_sku = st.selectbox("SKU Column:", df_inv.columns, index=list(df_inv.columns).index(default_sku_col))
        col_qty = st.selectbox("Inventory Quantity Column:", df_inv.columns, index=list(df_inv.columns).index(default_qty_col))

        # 统一重命名，确保不破坏后面所有的处理逻辑流程
        df_inv = df_inv.rename(columns={
            col_wh: "Warehouse Name",
            col_sku: "SKU",
            col_qty: "Available For Sale"
        })
        col = "Available For Sale" # 保持后续逻辑一致

        # --- Data Cleaning ---
        df_inv["SKU"] = df_inv["SKU"].astype(str).str.strip()
        df_inv["Warehouse Name"] = df_inv["Warehouse Name"].astype(str).str.strip()

        df_inv[col] = pd.to_numeric(
            df_inv[col],
            errors="coerce"
        ).fillna(0)

        df_demand["SKU"] = df_demand["SKU"].astype(str).str.strip()
        df_demand["Quantity"] = pd.to_numeric(
            df_demand["Quantity"],
            errors="coerce"
        ).fillna(0)

        df_demand = df_demand[
            (df_demand["SKU"] != "") &
            (df_demand["SKU"].str.lower() != "nan") &
            (df_demand["Quantity"] > 0)
        ].copy()

        if df_demand.empty:
            st.warning("⚠️ Demand file is empty or contains no valid SKU / Quantity.")
            st.stop()

        # 重新捕获最新最准的输入顺序
        original_sku_order = list(df_demand["SKU"].unique())

        # 建立 原始需求 SKU -> 标准 SKU 的映射字典，用于后面追溯原始需求
        sku_to_match_dict = {s: get_sku_detail(s)[0] for s in original_sku_order}

        # --- Standardize Demand SKU ---
        demand_grouped = df_demand.groupby(
            df_demand.apply(lambda r: get_sku_detail(r["SKU"])[0], axis=1)
        )["Quantity"].sum().to_dict()

        # --- Standardize Inventory SKU ---
        df_inv[["Match_SKU", "Ver_Priority", "Clean_SKU"]] = df_inv.apply(
            lambda r: pd.Series(get_sku_detail(r["SKU"])),
            axis=1
        )

        df_inv["Match_SKU"] = df_inv["Match_SKU"].astype(str)
        df_inv["Clean_SKU"] = df_inv["Clean_SKU"].astype(str)
        df_inv["Ver_Priority"] = pd.to_numeric(
            df_inv["Ver_Priority"],
            errors="coerce"
        ).fillna(0).astype(int)

        # --- Core Allocation Logic ---
        def run_allocation(rem_demand, wh_list):
            guide = []
            logic = []
            rem = rem_demand.copy()

            for wh_tag in wh_list:
                wh_names = [
                    n for n in df_inv["Warehouse Name"].unique()
                    if str(wh_tag).upper() in str(n).upper()
                ]

                for wh in wh_names:
                    wh_stock = df_inv[df_inv["Warehouse Name"] == wh].copy()

                    if wh_stock.empty:
                        continue

                    wh_summary = (
                        wh_stock
                        .groupby(["Match_SKU", "Ver_Priority"])
                        .agg({
                            col: "sum",
                            "Clean_SKU": lambda x: "/".join(
                                pd.Series(x).dropna().astype(str).unique()
                            )
                        })
                        .unstack()
                    )

                    if col in wh_summary.columns.get_level_values(0):
                        wh_summary[col] = wh_summary[col].fillna(0)

                    if "Clean_SKU" in wh_summary.columns.get_level_values(0):
                        wh_summary["Clean_SKU"] = wh_summary["Clean_SKU"].fillna("")

                    for m_sku in list(rem.keys()):
                        if rem[m_sku] <= 0:
                            continue

                        sku_options = (
                            wh_stock[wh_stock["Match_SKU"] == m_sku]
                            .sort_values(by="Ver_Priority")
                        )

                        if sku_options.empty:
                            continue

                        def get_val(p, field):
                            if m_sku in wh_summary.index and (field, p) in wh_summary.columns:
                                val = wh_summary.loc[m_sku, (field, p)]
                                if pd.isna(val):
                                    return 0 if field == col else ""
                                if field == col:
                                    return float(val)
                                return str(val)
                            return 0 if field == col else ""

                        n_avail = get_val(0, col)
                        n_name = get_val(0, "Clean_SKU")
                        v_avail = get_val(1, col) + get_val(2, col)
                        v_name = "/".join(filter(None, [
                            get_val(1, "Clean_SKU"),
                            get_val(2, "Clean_SKU")
                        ]))

                        take_total = min(rem[m_sku], n_avail + v_avail)

                        if take_total > 0:
                            logic.append({
                                "Standard SKU": str(m_sku),
                                "Allocated Warehouse": str(wh),
                                "Allocated Qty": float(take_total),
                                "Regular Version Reference":
                                    f"{n_name} (Remain:{n_avail})" if n_avail > 0 else "-",
                                "YN/MG Version Reference":
                                    f"{v_name} (Remain:{v_avail})" if v_avail > 0 else "-"
                            })

                            for _, row in sku_options.iterrows():
                                if rem[m_sku] <= 0:
                                    break

                                available_qty = float(row[col])
                                item_take = min(rem[m_sku], available_qty)

                                if item_take > 0:
                                    # 尝试找回原始需求对应的SKU全称，用于排序和比对高亮
                                    orig_sku_matches = [orig for orig, match in sku_to_match_dict.items() if match == m_sku]
                                    orig_sku_label = orig_sku_matches[0] if orig_sku_matches else str(m_sku)

                                    guide.append({
                                        "Warehouse": str(wh),
                                        "Original Demand SKU": orig_sku_label, # 新增隐藏辅助列：用于精准排序与高亮
                                        "Order SKU": str(row["Clean_SKU"]),
                                        "Quantity": float(item_take)
                                    })

                                    rem[m_sku] -= item_take

            return pd.DataFrame(guide), pd.DataFrame(logic), rem

        # --- Plan A/B ---
        df_guide_ab, df_logic_ab, rem_ab = run_allocation(demand_grouped, final_priority)

        # --- Plan C ---
        wh_scores = []
        for wh in df_inv["Warehouse Name"].unique():
            score = 0
            wh_inv_part = df_inv[df_inv["Warehouse Name"] == wh].groupby("Match_SKU")[col].sum()

            for s, q in demand_grouped.items():
                available = wh_inv_part.get(s, 0)
                if available >= q:
                    score += 10
                elif available > 0:
                    score += 1

            wh_scores.append({"name": wh, "score": score})

        auto_priority = [x["name"] for x in sorted(wh_scores, key=lambda x: x["score"], reverse=True)]
        df_guide_c, df_logic_c, rem_c = run_allocation(demand_grouped, auto_priority)

        # --- 需求 2：执行严格按照输入 SKU 顺序排序逻辑 ---
        def sort_by_original_order(df):
            if df.empty or "Original Demand SKU" not in df.columns:
                return df
            # 将原始输入的序列转为 Categorical 类型以确保按输入顺序排序
            df["Original Demand SKU"] = pd.Categorical(df["Original Demand SKU"], categories=original_sku_order, ordered=True)
            df = df.sort_values(by=["Original Demand SKU", "Warehouse"]).reset_index(drop=True)
            return df

        df_guide_ab = sort_by_original_order(df_guide_ab)
        df_guide_c = sort_by_original_order(df_guide_c)

        # 在展示前，由于后台导出的 Excel 不需要保留 "Original Demand SKU" 辅助列，我们在前端展示或导出前灵活处理
        # 为了配合高亮，我们把 Original Demand SKU 放在前端展示中，确认完后一目了然

        # --- Display Results ---
        tab_ab, tab_c = st.tabs([
            "🚀 Plan A/B: Manual Priority Allocation",
            "📦 Plan C: Logistics Optimized"
        ])

        with tab_ab:
            if not df_guide_ab.empty:
                # 重新调整下前端展示列顺序，并将 st.table 升级为 st.dataframe 以渲染高亮
                show_cols = ["Warehouse", "Original Demand SKU", "Order SKU", "Quantity"]
                df_disp_ab = safe_display_df(df_guide_ab[show_cols])
                st.dataframe(style_variant_rows(df_disp_ab), use_container_width=True, height=400)
            else:
                st.write("No allocation result.")

            with st.expander("View Detailed Allocation Logic"):
                st.dataframe(safe_display_df(df_logic_ab), use_container_width=True)

        with tab_c:
            if not df_guide_c.empty:
                show_cols = ["Warehouse", "Original Demand SKU", "Order SKU", "Quantity"]
                df_disp_c = safe_display_df(df_guide_c[show_cols])
                st.dataframe(style_variant_rows(df_disp_c), use_container_width=True, height=400)
            else:
                st.write("No allocation result.")

            with st.expander("View Detailed Allocation Logic"):
                st.dataframe(safe_display_df(df_logic_c), use_container_width=True)

        st.divider()
        st.subheader("📥 Export & Shortage Report")

        df_short = pd.DataFrame([
            {
                "SKU": str(k),
                "Unfulfilled Quantity": float(v)
            }
            for k, v in rem_ab.items()
            if v > 0
        ])

        if not df_short.empty:
            st.error("The following SKUs have insufficient inventory:")
            st.dataframe(safe_display_df(df_short), use_container_width=True)
        else:
            st.success("✨ All demands have been fulfilled!")

        # --- Export Excel (清洗掉辅助列，保持干净的用户导出格式) ---
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as wr:
            if not df_guide_ab.empty:
                # 导出时去掉技术辅助列，只保留最干净的 3 列结果
                export_cols = ["Warehouse", "Order SKU", "Quantity"]
                df_guide_ab[export_cols].to_excel(wr, sheet_name="PlanAB_OrderGuide", index=False)

            if not df_guide_c.empty:
                export_cols = ["Warehouse", "Order SKU", "Quantity"]
                df_guide_c[export_cols].to_excel(wr, sheet_name="PlanC_LogisticsOptimized", index=False)

            if not df_short.empty:
                df_short.to_excel(wr, sheet_name="ShortageReport", index=False)

            if not df_logic_ab.empty:
                df_logic_ab.to_excel(wr, sheet_name="DetailedReference", index=False)

        out.seek(0)
        st.download_button(
            label="📥 Download Full Excel Report",
            data=out.getvalue(),
            file_name="Inventory_Allocation_V8.6_Updated.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Analysis Error: {e}")
