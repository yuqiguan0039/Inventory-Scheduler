import pandas as pd

# 1. 加载数据 (请确保文件名与你实际的文件名一致)
demand_file = 'HS需求.xlsx - Sheet1.csv'
inventory_file = 'Vape_Inventory_1777395253091.xlsx - sheet1.csv'

df_demand = pd.read_csv(demand_file)
df_inventory = pd.read_csv(inventory_file)

# 数据清洗：去除SKU两端空格
df_demand['SKU'] = df_demand['SKU'].str.strip()
df_inventory['SKU'] = df_inventory['SKU'].str.strip()

# 2. 统计每个仓库对总需求的覆盖能力
warehouses = df_inventory['Warehouse Name'].unique()
coverage_report = []

for wh in warehouses:
    wh_stock = df_inventory[df_inventory['Warehouse Name'] == wh].groupby('SKU')['Accounting Available For Sale'].sum()
    
    match_count = 0
    total_needed = 0
    total_satisfied = 0
    
    for _, row in df_demand.iterrows():
        sku = row['SKU']
        qty_needed = row['Quantity']
        total_needed += qty_needed
        
        if sku in wh_stock:
            can_provide = min(qty_needed, wh_stock[sku])
            total_satisfied += can_provide
            if can_provide > 0:
                match_count += 1
                
    coverage_report.append({
        'Warehouse': wh,
        'Matched SKU Types': match_count,
        'Satisfied Quantity': total_satisfied,
        'Satisfaction Rate (%)': round((total_satisfied / total_needed) * 100, 2) if total_needed > 0 else 0
    })

# 找出满足率最高的仓库
df_report = pd.DataFrame(coverage_report).sort_values(by='Satisfaction Rate (%)', ascending=False)
best_wh = df_report.iloc[0]['Warehouse']

print("--- 仓库满足率排名 ---")
print(df_report)
print(f"\n推荐主发货仓: {best_wh}")

# 3. 生成详细下单方案
order_plan = []
remaining_demand = df_demand.set_index('SKU')['Quantity'].to_dict()

# 第一轮：优先从最优仓划扣
wh_stock_best = df_inventory[df_inventory['Warehouse Name'] == best_wh].groupby('SKU')['Accounting Available For Sale'].sum().to_dict()
for sku in list(remaining_demand.keys()):
    needed = remaining_demand[sku]
    available = wh_stock_best.get(sku, 0)
    take = min(needed, available)
    if take > 0:
        order_plan.append({'SKU': sku, 'Warehouse': best_wh, 'Take Quantity': take})
        remaining_demand[sku] -= take

# 第二轮：剩余部分从其他仓补足（按满足率顺序）
other_whs = df_report.iloc[1:]['Warehouse'].tolist()
for wh in other_whs:
    wh_stock = df_inventory[df_inventory['Warehouse Name'] == wh].groupby('SKU')['Accounting Available For Sale'].sum().to_dict()
    for sku in list(remaining_demand.keys()):
        needed = remaining_demand[sku]
        if needed <= 0: continue
        
        available = wh_stock.get(sku, 0)
        take = min(needed, available)
        if take > 0:
            order_plan.append({'SKU': sku, 'Warehouse': wh, 'Take Quantity': take})
            remaining_demand[sku] -= take

# 4. 输出结果
df_plan = pd.DataFrame(order_plan)
df_shortage = pd.DataFrame([{'SKU': k, 'Shortage': v} for k, v in remaining_demand.items() if v > 0])

# 导出到 Excel
with pd.ExcelWriter('下单方案结果.xlsx') as writer:
    df_report.to_excel(writer, sheet_name='仓库满足率排名', index=False)
    df_plan.to_excel(writer, sheet_name='下单方案', index=False)
    df_shortage.to_excel(writer, sheet_name='仍有缺口SKU', index=False)

print("\n方案生成完毕！请查看 '下单方案结果.xlsx'")