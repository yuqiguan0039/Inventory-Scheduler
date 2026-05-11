import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="智能库存分配调度系统 V8.8", layout="wide")

st.title("📦 智能库存分配调度系统 V8.8")

# --- 1. 侧边栏：仓库调度配置中心 ---
st.sidebar.header("⚙️ 仓库调度配置中心")
default_order = ["TX", "PHX", "ID", "IL"]
user_priority = []
for i in range(1, 5):
    target_wh = st.sidebar.selectbox(f"优先级 {i}", ["默认"] + default_order, key=f"p_{i}")
    if target_wh != "默认" and target_wh not in user_priority:
        user_priority.append(target_wh)
final_priority = user_priority + [w for w in default_order if w not in user_priority]

st.sidebar.write("手动执行顺序：", " -> ".join(final_priority))

enable_variant_match = st.sidebar.checkbox("开启多版本兼容 (YN/MG)", value=True)
variant_prefixes = ["YN", "MG"] 

# --- 2. 需求录入板块 ---
st.subheader("第一步：录入订单需求")
input_method = st.radio("录入方式", ["自由文本提取", "手动输入/编辑表格", "上传文件"], horizontal=True)

df_demand = pd.DataFrame(columns=['SKU', 'Quantity'])
if input_method == "自由文本提取":
    raw_text = st.text_area("在此粘贴原始文本：", height=200)
    if raw_text:
        matches = re.finditer(r"SKU:\s*(?P<sku>[\w\-+]+)", raw_text)
        extracted = []
        all_skus = list(matches)
        for i, m in enumerate(all_skus):
            sku, end_pos = m.group('sku'), m.end()
            next_start = all_skus[i+1].start() if i+1 < len(all_skus) else len(raw_text)
            sub_text = raw_text[end_pos:next_start]
            qty_m = re.search(r"(?<!\$)(?P<qty>\d+\.?\d*)", sub_text)
            if qty_m: extracted.append({"SKU": sku, "Quantity": float(qty_m.group('qty'))})
        if extracted:
            df_demand = pd.DataFrame(extracted)
            st.success(f"✅ 成功提取出 {len(df_demand)} 条需求")
            st.dataframe(df_demand, use_container_width=True)

elif input_method == "手动输入/编辑表格":
    df_demand = st.data_editor(pd.DataFrame([{"SKU": "", "Quantity": 0.0}] * 10), num_rows="dynamic", use_container_width=True)

elif input_method == "上传文件":
    f = st.file_uploader("上传需求表", type=["xlsx", "csv"], key="d_up")
    if f:
        df_demand = pd.read_csv(f) if f.name.endswith('.csv') else pd.read_excel(f)
        df_demand.columns = [str(c).strip() for c in df_demand.columns]
        s_col = st.selectbox("需求表：选择 SKU 列", df_demand.columns)
        q_col = st.selectbox("需求表：选择 数量 列", df_demand.columns)
        df_demand = df_demand.rename(columns={s_col: 'SKU', q_col: 'Quantity'})

# --- 辅助函数 ---
def get_sku_detail(sku):
    sku = str(sku).strip().upper()
    for idx, v in enumerate(variant_prefixes, start=1):
        for p in ["C-", "B-"]:
            if sku.startswith(p + v): return p + sku[len(p)+len(v):], idx, sku
    return sku, 0, sku

# --- 3. 执行调度逻辑 ---
st.divider()
st.subheader("第二步：上传库存表并匹配列名")
inv_f = st.file_uploader("上传库存表 (Kingdee/Vape)", type=["xlsx", "csv"], key="i_up")

if inv_f and not df_demand.empty:
    try:
        df_inv_raw = pd.read_csv(inv_f) if inv_f.name.endswith('.csv') else pd.read_excel(inv_f)
        
        # 核心修复：清理表头空格，防止 KeyError
        df_inv_raw.columns = [str(c).strip() for c in df_inv_raw.columns]
        
        # 列名自选功能
        col1, col2 = st.columns(2)
        with col1:
            default_sku_idx = list(df_inv_raw.columns).index('SKU') if 'SKU' in df_inv_raw.columns else 0
            inv_sku_col = st.selectbox("库存表：选择 SKU 对应列", df_inv_raw.columns, index=default_sku_idx)
        with col2:
            possible_qty = ['Available For Sale', 'Accounting Available For Sale', 'Quantity']
            default_qty_idx = 0
            for pq in possible_qty:
                if pq in df_inv_raw.columns:
                    default_qty_idx = list(df_inv_raw.columns).index(pq)
                    break
            inv_qty_col = st.selectbox("库存表：选择 库存量 对应列", df_inv_raw.columns, index=default_qty_idx)

        # 检查 Warehouse Name
        if 'Warehouse Name' not in df_inv_raw.columns:
            st.error(f"❌ 找不到 'Warehouse Name' 列。现有列名为: {list(df_inv_raw.columns)}")
            st.stop()

        df_inv = df_inv_raw.copy().rename(columns={inv_sku_col: 'SKU', inv_qty_col: 'Calc_Qty'})
        
        df_demand['Quantity'] = pd.to_numeric(df_demand['Quantity'], errors='coerce').fillna(0)
        demand_grouped = df_demand.groupby(df_demand.apply(lambda r: get_sku_detail(r['SKU'])[0], axis=1))['Quantity'].sum().to_dict()
        df_inv[['Match_SKU', 'Ver_Priority', 'Clean_SKU']] = df_inv.apply(lambda r: pd.Series(get_sku_detail(r['SKU'])), axis=1)

        def run_allocation(rem_demand, wh_list):
            guide = []; logic = []; rem = rem_demand.copy()
            for wh_tag in wh_list:
                wh_names = [n for n in df_inv['Warehouse Name'].unique() if wh_tag in str(n).upper()]
                for wh in wh_names:
                    wh_stock = df_inv[df_inv['Warehouse Name'] == wh].copy()
                    # 聚合库存
                    wh_summary = wh_stock.groupby(['Match_SKU', 'Ver_Priority']).agg({'Calc_Qty': 'sum', 'Clean_SKU': lambda x: "/".join(x.unique())}).unstack(fill_value=0)
                    
                    for m_sku in list(rem.keys()):
                        if rem[m_sku] <= 0: continue
                        sku_options = wh_stock[wh_stock['Match_SKU'] == m_sku].sort_values(by='Ver_Priority')
                        if sku_options.empty: continue
                        
                        def get_val(p, field): 
                            return wh_summary.loc[m_sku, (field, p)] if (m_sku in wh_summary.index and (field, p) in wh_summary.columns) else (0 if field=='Calc_Qty' else "")
                        
                        n_avail, n_name = get_val(0, 'Calc_Qty'), get_val(0, 'Clean_SKU')
                        v_avail = get_val(1, 'Calc_Qty') + get_val(2, 'Calc_Qty')
                        v_name = "/".join(filter(None, [get_val(1, 'Clean_SKU'), get_val(2, 'Clean_SKU')]))
                        
                        take_total = min(rem[m_sku], n_avail + v_avail)
                        if take_total > 0:
                            logic.append({"标准SKU": m_sku, "分配仓库": wh, "下单总量": take_total, "普通版参考": f"{n_name} (余:{n_avail})" if n_avail > 0 else "-", "YN/MG版参考": f"{v_name} (余:{v_avail})" if v_avail > 0 else "-"})
                            for _, row in sku_options.iterrows():
                                if rem[m_sku] <= 0: break
                                item_take = min(rem[m_sku], row['Calc_Qty'])
                                if item_take > 0:
                                    guide.append({"下单仓库": wh, "具体下单 SKU": row['Clean_SKU'], "数量": item_take})
                                    rem[m_sku] -= item_take
            return pd.DataFrame(guide), pd.DataFrame(logic), rem

        # 运行方案
        df_guide_ab, df_logic_ab, rem_ab = run_allocation(demand_grouped, final_priority)
        
        # 方案 C 评分逻辑
        wh_scores = []
        for wh in df_inv['Warehouse Name'].unique():
            score = 0
            wh_inv_s = df_inv[df_inv['Warehouse Name'] == wh].groupby('Match_SKU')['Calc_Qty'].sum()
            for s, q in demand_grouped.items():
                if wh_inv_s.get(s, 0) >= q: score += 10
                elif wh_inv_s.get(s, 0) > 0: score += 1
            wh_scores.append({'name': wh, 'score': score})
        auto_priority = [x['name'] for x in sorted(wh_scores, key=lambda x: x['score'], reverse=True)]
        df_guide_c, df_logic_c, rem_c = run_allocation(demand_grouped, auto_priority)

        # 展示
        tab_ab, tab_c = st.tabs(["🚀 方案 A/B (手动顺位)", "📦 方案 C (物流最优)"])
        with tab_ab:
            st.table(df_guide_ab)
            with st.expander("逻辑参考"): st.table(df_logic_ab)
        with tab_c:
            st.table(df_guide_c)
            with st.expander("逻辑参考"): st.table(df_logic_c)

        # 下载与缺口
        st.divider()
        st.subheader("📥 方案导出与缺口明细")
        df_short = pd.DataFrame([{"SKU": k, "未满足数量": v} for k, v in rem_ab.items() if v > 0])
        if not df_short.empty:
            st.error("库存缺口明细：")
            st.table(df_short)
        else:
            st.success("✨ 所有需求均已满足！")

        # 加固下载逻辑
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as wr:
            if not df_guide_ab.empty: df_guide_ab.to_excel(wr, sheet_name='方案AB-指南', index=False)
            if not df_guide_c.empty: df_guide_c.to_excel(wr, sheet_name='方案C-最优', index=False)
            if not df_short.empty: df_short.to_excel(wr, sheet_name='缺口表', index=False)
        
        buffer.seek(0) # 关键：重置指针
        st.download_button(
            label="📥 下载 Excel 完整结果报告",
            data=buffer,
            file_name="Dispatch_Report_V8.8.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"分析出错：{e}")
