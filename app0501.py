import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="智能库存分配调度系统 V8.6", layout="wide")

st.title("📦 智能库存分配调度系统 V8.6")

# --- 1. 侧边栏：仓库调度配置中心 ---
st.sidebar.header("⚙️ 仓库调度配置中心")

default_order = ["TX", "PHX", "ID", "IL"]
user_priority = []

for i in range(1, 5):
    target_wh = st.sidebar.selectbox(
        f"优先级 {i}",
        ["默认"] + default_order,
        key=f"p_{i}"
    )
    if target_wh != "默认" and target_wh not in user_priority:
        user_priority.append(target_wh)

final_priority = user_priority + [w for w in default_order if w not in user_priority]

st.sidebar.write("手动执行顺序：", " -> ".join(final_priority))

enable_variant_match = st.sidebar.checkbox("开启多版本兼容 (YN/MG)", value=True)
variant_prefixes = ["YN", "MG"]


# --- 2. 需求录入板块 ---
st.subheader("第一步：录入订单需求")

input_method = st.radio(
    "录入方式",
    ["自由文本提取", "手动输入/编辑表格", "上传文件"],
    horizontal=True
)

df_demand = pd.DataFrame(columns=["SKU", "Quantity"])

if input_method == "自由文本提取":
    raw_text = st.text_area("在此粘贴原始文本：", height=200)

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
            st.success(f"✅ 成功提取出 {len(df_demand)} 条需求")
            st.dataframe(df_demand, use_container_width=True)

elif input_method == "手动输入/编辑表格":
    df_demand = st.data_editor(
        pd.DataFrame([{"SKU": "", "Quantity": 0.0}] * 10),
        num_rows="dynamic",
        use_container_width=True
    )

elif input_method == "上传文件":
    f = st.file_uploader("上传需求表", type=["xlsx", "csv"])

    if f:
        df_demand = pd.read_csv(f) if f.name.endswith(".csv") else pd.read_excel(f)
        df_demand.columns = [str(c).strip() for c in df_demand.columns]

        s_col = st.selectbox("SKU列", df_demand.columns)
        q_col = st.selectbox("数量列", df_demand.columns)

        df_demand = df_demand.rename(columns={
            s_col: "SKU",
            q_col: "Quantity"
        })


# --- 辅助函数 ---
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
            df_show[c] = df_show[c].apply(lambda x: "" if pd.isna(x) else str(x))

    return df_show


# --- 3. 执行调度逻辑 ---
st.divider()

inv_f = st.file_uploader("上传库存表 (Kingdee)", type=["xlsx", "csv"])

if inv_f and not df_demand.empty:
    try:
        # --- 读取库存表 ---
        df_inv = pd.read_csv(inv_f) if inv_f.name.endswith(".csv") else pd.read_excel(inv_f)
        df_inv.columns = [str(c).strip() for c in df_inv.columns]

        # --- 必要字段检查 ---
        if "Warehouse Name" not in df_inv.columns:
            st.error(f"❌ 错误：库存表中找不到 'Warehouse Name' 列。当前的列名有：{list(df_inv.columns)}")
            st.stop()

        if "SKU" not in df_inv.columns:
            st.error(f"❌ 错误：库存表中找不到 'SKU' 列。当前的列名有：{list(df_inv.columns)}")
            st.stop()

        if "Available For Sale" in df_inv.columns:
            col = "Available For Sale"
        elif "Accounting Available For Sale" in df_inv.columns:
            col = "Accounting Available For Sale"
        else:
            st.error("❌ 错误：库存表中找不到 'Available For Sale' 或 'Accounting Available For Sale' 列。")
            st.stop()

        # --- 数据清洗 ---
        df_inv["SKU"] = df_inv["SKU"].astype(str).str.strip()
        df_inv["Warehouse Name"] = df_inv["Warehouse Name"].astype(str).str.strip()
        df_inv[col] = pd.to_numeric(df_inv[col], errors="coerce").fillna(0)

        df_demand["SKU"] = df_demand["SKU"].astype(str).str.strip()
        df_demand["Quantity"] = pd.to_numeric(df_demand["Quantity"], errors="coerce").fillna(0)

        df_demand = df_demand[
            (df_demand["SKU"] != "") &
            (df_demand["SKU"].str.lower() != "nan") &
            (df_demand["Quantity"] > 0)
        ].copy()

        if df_demand.empty:
            st.warning("⚠️ 需求表为空，或没有有效的 SKU / 数量。")
            st.stop()

        # --- 标准化需求 SKU ---
        demand_grouped = df_demand.groupby(
            df_demand.apply(lambda r: get_sku_detail(r["SKU"])[0], axis=1)
        )["Quantity"].sum().to_dict()

        # --- 标准化库存 SKU ---
        df_inv[["Match_SKU", "Ver_Priority", "Clean_SKU"]] = df_inv.apply(
            lambda r: pd.Series(get_sku_detail(r["SKU"])),
            axis=1
        )

        df_inv["Match_SKU"] = df_inv["Match_SKU"].astype(str)
        df_inv["Clean_SKU"] = df_inv["Clean_SKU"].astype(str)
        df_inv["Ver_Priority"] = pd.to_numeric(df_inv["Ver_Priority"], errors="coerce").fillna(0).astype(int)

        # --- 核心分配逻辑 ---
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
                                "标准SKU": str(m_sku),
                                "分配仓库": str(wh),
                                "下单总量": float(take_total),
                                "普通版库存参考": f"{n_name} (余:{n_avail})" if n_avail > 0 else "-",
                                "YN/MG版库存参考": f"{v_name} (余:{v_avail})" if v_avail > 0 else "-"
                            })

                            for _, row in sku_options.iterrows():
                                if rem[m_sku] <= 0:
                                    break

                                available_qty = float(row[col])
                                item_take = min(rem[m_sku], available_qty)

                                if item_take > 0:
                                    guide.append({
                                        "下单仓库": str(wh),
                                        "具体下单 SKU": str(row["Clean_SKU"]),
                                        "数量": float(item_take)
                                    })

                                    rem[m_sku] -= item_take

            return pd.DataFrame(guide), pd.DataFrame(logic), rem

        # --- 方案 A/B：指定顺序优先 ---
        df_guide_ab, df_logic_ab, rem_ab = run_allocation(
            demand_grouped,
            final_priority
        )

        # --- 方案 C：物流最优 ---
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

        # --- 界面展示 ---
        tab_ab, tab_c = st.tabs([
            "🚀 方案 A/B：指定顺序优先",
            "📦 方案 C：物流最优 (最少仓+不拆箱)"
        ])

        with tab_ab:
            st.table(safe_display_df(df_guide_ab))

            with st.expander("查看详细分配逻辑"):
                st.table(safe_display_df(df_logic_ab))

        with tab_c:
            st.table(safe_display_df(df_guide_c))

            with st.expander("查看详细分配逻辑"):
                st.table(safe_display_df(df_logic_c))

        st.divider()
        st.subheader("📥 方案导出与缺口明细")

        df_short = pd.DataFrame([
            {
                "SKU": str(k),
                "未满足数量": float(v)
            }
            for k, v in rem_ab.items()
            if v > 0
        ])

        if not df_short.empty:
            st.error("以下产品库存不足：")
            st.table(safe_display_df(df_short))
        else:
            st.success("✨ 所有需求均已满足！")

        # --- 下载功能 ---
        out = io.BytesIO()

        with pd.ExcelWriter(out, engine="openpyxl") as wr:
            if not df_guide_ab.empty:
                df_guide_ab.to_excel(wr, sheet_name="方案AB-傻瓜指南", index=False)

            if not df_guide_c.empty:
                df_guide_c.to_excel(wr, sheet_name="方案C-物流最优", index=False)

            if not df_short.empty:
                df_short.to_excel(wr, sheet_name="缺口明细", index=False)

            if not df_logic_ab.empty:
                df_logic_ab.to_excel(wr, sheet_name="详细参考(供核对)", index=False)

        out.seek(0)

        st.download_button(
            label="📥 下载 Excel 完整结果报告",
            data=out.getvalue(),
            file_name="Inventory_Allocation_V8.6.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"分析出错：{e}")
