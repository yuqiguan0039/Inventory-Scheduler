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
                extracted.append({
                    "SKU": sku,
                    "Quantity": float(qty_m.group("qty"))
                })

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
    """
    修复 Streamlit Cloud / Arrow 对混合类型列的报错。
    只影响前端展示，不改变计算和导出结果。
    """
    if df is None or df.empty:
        return df

    df_show = df.copy()

    for c in df_show.columns:
        if df_show[c].dtype == "object":
            df_show[c] = df_show[c].apply(
                lambda x: "" if pd.isna(x) else str(x)
            )

    return df_show


# --- 3. Allocation Logic ---
st.divider()

inv_f = st.file_uploader("Upload Inventory File (Kingdee)", type=["xlsx", "csv"])

if inv_f and not df_demand.empty:
    try:
        # --- Load Inventory File ---
        df_inv = pd.read_csv(inv_f) if inv_f.name.endswith(".csv") else pd.read_excel(inv_f)

        df_inv.columns = [str(c).strip() for c in df_inv.columns]

        # --- Required Column Validation ---
        if "Warehouse Name" not in df_inv.columns:
            st.error(
                f"❌ Error: 'Warehouse Name' column not found. Current columns: {list(df_inv.columns)}"
            )
            st.stop()

        if "SKU" not in df_inv.columns:
            st.error(
                f"❌ Error: 'SKU' column not found. Current columns: {list(df_inv.columns)}"
            )
            st.stop()

        if "Available For Sale" in df_inv.columns:
            col = "Available For Sale"

        elif "Accounting Available For Sale" in df_inv.columns:
            col = "Accounting Available For Sale"

        else:
            st.error(
                "❌ Error: Neither 'Available For Sale' nor 'Accounting Available For Sale' column was found."
            )
            st.stop()

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

                    # 关键修复：
                    # 不能使用 unstack(fill_value=0)，否则 Clean_SKU 文字字段也会被填成数字 0
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

                    # 数量字段缺失填 0
                    if col in wh_summary.columns.get_level_values(0):
                        wh_summary[col] = wh_summary[col].fillna(0)

                    # 文字字段缺失填空字符串
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

                                item_take = min(
                                    rem[m_sku],
                                    available_qty
                                )

                                if item_take > 0:
                                    guide.append({
                                        "Warehouse": str(wh),
                                        "Order SKU": str(row["Clean_SKU"]),
                                        "Quantity": float(item_take)
                                    })

                                    rem[m_sku] -= item_take

            return pd.DataFrame(guide), pd.DataFrame(logic), rem

        # --- Plan A/B ---
        df_guide_ab, df_logic_ab, rem_ab = run_allocation(
            demand_grouped,
            final_priority
        )

        # --- Plan C ---
        wh_scores = []

        for wh in df_inv["Warehouse Name"].unique():
            score = 0

            wh_inv_part = (
                df_inv[df_inv["Warehouse Name"] == wh]
                .groupby("Match_SKU")[col]
                .sum()
            )

            for s, q in demand_grouped.items():
                available = wh_inv_part.get(s, 0)

                if available >= q:
                    score += 10

                elif available > 0:
                    score += 1

            wh_scores.append({
                "name": wh,
                "score": score
            })

        auto_priority = [
            x["name"] for x in sorted(
                wh_scores,
                key=lambda x: x["score"],
                reverse=True
            )
        ]

        df_guide_c, df_logic_c, rem_c = run_allocation(
            demand_grouped,
            auto_priority
        )

        # --- Display Results ---
        tab_ab, tab_c = st.tabs([
            "🚀 Plan A/B: Manual Priority Allocation",
            "📦 Plan C: Logistics Optimized"
        ])

        with tab_ab:
            st.table(safe_display_df(df_guide_ab))

            with st.expander("View Detailed Allocation Logic"):
                st.table(safe_display_df(df_logic_ab))

        with tab_c:
            st.table(safe_display_df(df_guide_c))

            with st.expander("View Detailed Allocation Logic"):
                st.table(safe_display_df(df_logic_c))

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
            st.table(safe_display_df(df_short))

        else:
            st.success("✨ All demands have been fulfilled!")

        # --- Export Excel ---
        out = io.BytesIO()

        with pd.ExcelWriter(out, engine="openpyxl") as wr:

            if not df_guide_ab.empty:
                df_guide_ab.to_excel(
                    wr,
                    sheet_name="PlanAB_OrderGuide",
                    index=False
                )

            if not df_guide_c.empty:
                df_guide_c.to_excel(
                    wr,
                    sheet_name="PlanC_LogisticsOptimized",
                    index=False
                )

            if not df_short.empty:
                df_short.to_excel(
                    wr,
                    sheet_name="ShortageReport",
                    index=False
                )

            if not df_logic_ab.empty:
                df_logic_ab.to_excel(
                    wr,
                    sheet_name="DetailedReference",
                    index=False
                )

        out.seek(0)

        st.download_button(
            label="📥 Download Full Excel Report",
            data=out.getvalue(),
            file_name="Inventory_Allocation_V8.6.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Analysis Error: {e}")
