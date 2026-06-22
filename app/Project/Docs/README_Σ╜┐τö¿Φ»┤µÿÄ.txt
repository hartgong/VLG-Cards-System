VLG-TCG System V3 使用说明

重点更新：
1. 库存刷新脚本现在必须明确刷新你传入的同一个Excel文件；不再默认生成输出文件。
2. 你在 Purchase_Inbound 填入库信息后，运行 refresh_inventory.py，会直接重建该文件里的 Inventory_Ledger 和 Current_Inventory。
3. 请先关闭Excel正在打开的文件，否则Windows会阻止保存。

推荐运行：
python Scripts/refresh_inventory.py Data/card_data_warehouse_V3.xlsx

可选备份：
python Scripts/refresh_inventory.py Data/card_data_warehouse_V3.xlsx --backup

测试增加少量库存：
python Scripts/test_add_small_purchase.py Data/card_data_warehouse_V3.xlsx --sku TCG00001 --qty 1 --unit-cost 0.10

如果运行后没有变化，请检查：
1. 你修改的Excel文件路径，是否就是命令里传入的文件路径。
2. Excel文件是否仍处于打开状态。
3. Purchase_Inbound 是否填写了 SKU_ID 和 Purchase_Qty。
4. Purchase_Qty 是否大于0。
5. Status 是否不是 Canceled/Cancelled/Void/Invalid/取消/作废。

库存逻辑：
Purchase_Inbound -> Inventory_Ledger 中 Purchase In 正数
Order_Line -> Inventory_Ledger 中 Sales Out 负数
Current_Inventory = Inventory_Ledger 汇总结果
