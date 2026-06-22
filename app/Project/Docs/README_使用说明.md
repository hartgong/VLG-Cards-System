# VLG TCG System V3 使用说明

本版本为 V3 覆盖版，已去掉旧版本中的 `00_`、`01_`、`02_` 等编号文件夹，目录名称改为更直观的英文名称。

## 一、目录结构

```text
VLG_TCG_System_V3/
├─ Raw_Files/                  原始文件，不建议修改
│  ├─ TikTok_Order/             TikTok订单CSV样本
│  ├─ SKU_Product_Info/         商品信息CSV样本
│  └─ Purchase/                 采购CSV样本
├─ Templates/                  标准数据仓库模板
│  └─ card_data_warehouse_template_V3.xlsx
├─ Data/                       真实 SKU 主数据仓库
│  └─ card_data_warehouse_V3.xlsx
├─ Scripts/                    Python脚本
│  ├─ refresh_inventory.py      库存刷新脚本
│  ├─ run_order_demo.py         订单导入演示脚本
│  ├─ common/                   通用解析脚本
│  ├─ order_import/             订单导入模块
│  ├─ logistics_import/         物流导入模块模板
│  ├─ settlement_import/        结算导入模块模板
│  └─ live_import/              直播导入模块模板
├─ Docs/                       说明文档
└─ requirements.txt            运行依赖
```

## 二、V3 商品和库存口径

V3 已取消复杂的箱/盒/包/单卡自动换算逻辑。

核心规则：

```text
一个 SKU = 一个可销售单位 = 一个库存单位
```

例如：

- TikTok 上按“包”卖：建立 SKU `TCG00001`，销售单位规格为“包”。
- eBay 上按“盒”卖：建立另一个 SKU `TCG00002`，销售单位规格为“盒”。
- 两者不再自动互相换算，库存各自独立管理。

订单里的 Seller Note 直接按 SKU 数量扣库存：

```text
#VLG
#Cammie
#TCG00001 2
#TCG00002 1
```

表示：

- 公司/团队：VLG
- 主播：Cammie
- 销售 `TCG00001` 2个
- 销售 `TCG00002` 1个

## 三、三张库存相关表的关系

### 1. Purchase_Inbound

采购入库表。你只需要在这里填写采购数据。

- 采购数量直接等于 SKU 数量。
- 所有金额均为 USD。
- 不做箱/盒/包之间的自动换算。

### 2. Inventory_Ledger

库存流水表。库存变化都要沉淀到这里。

常见类型：

| Txn_Type | 含义 | Qty_Change |
|---|---|---:|
| Purchase In | 采购入库 | 正数 |
| Sales Out | 销售扣库 | 负数 |
| Return In | 退货入库 | 正数 |
| Adjustment | 手工调整 | 正数或负数 |

### 3. Current_Inventory

当前库存表。它不是手工录入表，而是由 `Inventory_Ledger` 汇总出来。

```text
当前库存 = 该 SKU 所有 Qty_Change 合计
```

## 四、刷新库存脚本

脚本位置：

```text
Scripts/refresh_inventory.py
```

用途：

1. 读取 `Purchase_Inbound`，生成采购入库流水 `Purchase In`。
2. 读取 `Order_Line`，生成销售扣库流水 `Sales Out`。
3. 保留原有 `Inventory_Ledger` 中的 `Adjustment`、`Return In` 等手工调整行。
4. 重建 `Inventory_Ledger` 和 `Current_Inventory`。

运行示例：

```bash
cd VLG_TCG_System_V3
python Scripts/refresh_inventory.py
```

也可以显式指定主数据仓库：

```bash
python Scripts/refresh_inventory.py Data/card_data_warehouse_V3.xlsx
```

脚本会直接修改 `Data/card_data_warehouse_V3.xlsx` 本身，默认不会复制或另存新文件。

## 五、日常操作顺序

### 采购入库

1. 打开数据仓库 Excel。
2. 在 `Purchase_Inbound` 填写采购记录。
3. 运行 `Scripts/refresh_inventory.py`。
4. 查看 `Current_Inventory`。

### 订单扣库

1. 订单导入程序会从 TikTok 订单 CSV 中读取 `Seller Note`。
2. 生成 `Order_Master` 和 `Order_Line`。
3. 运行 `Scripts/refresh_inventory.py`。
4. 查看 `Current_Inventory`。

## 六、V3 使用重点

- `SKU_Master` 只维护 SKU 名称、销售单位规格、美元售价、KG、LB 等基础信息。
- `Purchase_Inbound` 只维护采购记录，所有金额为 USD。
- `Inventory_Ledger` 是唯一可信的库存流水来源。
- `Current_Inventory` 是自动汇总结果，不建议手工修改。
- Seller Note 中的 SKU 必须与 `SKU_Master` 中的 SKU_ID 一致。
