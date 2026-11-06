# 智能库存分配调度系统
这是一个基于 Streamlit 的物流调度工具，支持多仓库优先级匹配和 YN/MG 版本自动兼容。

## Brand 简称映射

产品标题简称由 `brand_aliases.xlsx` 管理，不再写在 `app0501.py` 中。

- 第一列必须为 `BRAND`，其内容必须与 `Product Catalog.xlsx` 的 `Brand` 完全一致。
- 后续列填写可能出现的简称，例如 `MT15`、`MT15K`、`MT 15K`。
- 程序匹配简称时不区分大小写，并忽略空格和符号。
- 同一个简称可以对应多个 Brand。程序会优先使用输入中明确出现的 YN、MG、0%、2%、World Cup、Lemonade、Turbo、Thermal 等版本信息；没有明确版本时优先普通版本，再结合口味与后续库存分配判断。
- 如果映射表中的 Brand 与产品目录不完全一致，页面会显示需要修正的 Brand 清单。

## Order Detail 历史记录

每次在 Paste Text 表单点击 `Start Matching` 时，非空的 Order Detail 会追加保存到 `data/order_history.jsonl`。每条记录包含 UTC 时间、唯一记录 ID、是否启用 AI 以及原始粘贴内容。该文件已排除在 Git 之外，避免意外提交客户订单。

本地文件适合电脑上的日常测试。部署到可能重置本地磁盘的云端服务时，应改用持久化数据库或受控的共享存储。
