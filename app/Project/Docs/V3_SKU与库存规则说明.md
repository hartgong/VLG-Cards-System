# V3 SKU与库存规则说明

## 1. SKU定义

V3 中，一个 SKU 就是一个可以独立销售、独立计价、独立扣库存的单位。

如果同一个商品在不同平台的销售单位不同，应建立不同 SKU：

| 场景 | SKU示例 | 销售单位规格 |
|---|---|---|
| TikTok卖单包 | TCG00001 | 包 |
| eBay卖整盒 | TCG00002 | 盒 |
| 批发卖整箱 | TCG00003 | 箱 |

## 2. Seller Note扣库规则

Seller Note 示例：

```text
#VLG
#Cammie
#TCG00001 2
#TCG00002 1
```

解析结果：

| 字段 | 结果 |
|---|---|
| Company_Code | VLG |
| Host_Name | Cammie |
| SKU_ID | TCG00001，数量2 |
| SKU_ID | TCG00002，数量1 |

## 3. 库存计算

库存不从 SKU_Master 或 Purchase_Inbound 直接读取，而是从 Inventory_Ledger 汇总。

```text
Current_Qty = SUM(Inventory_Ledger[Qty_Change] where SKU_ID = 当前SKU)
```

## 4. 手工采购流程

采购数据量较少，所以采购不需要 Python 导入。只需要：

1. 在 Purchase_Inbound 填采购数据。
2. 运行 refresh_inventory.py。
3. 脚本自动把 Purchase_Inbound 转成 Inventory_Ledger 的 Purchase In 行。
4. Current_Inventory 自动刷新。

## 库存刷新规则

- 手工只维护 `Purchase_Inbound` 的采购入库数据。
- 订单扣库来自 `Order_Line`。`Order_Line` 通常由 TikTok 订单导入脚本根据 `Seller Note` 解析生成。
- `Inventory_Ledger` 是库存流水账，会由 `refresh_inventory.py` 重建自动流水。
- `Current_Inventory` 是库存结果表，会由 `Inventory_Ledger` 自动汇总。
- 运行 `python Scripts/refresh_inventory.py Data/card_data_warehouse_V3.xlsx` 后，脚本直接修改该 Excel 文件，不再输出一个新文件。
