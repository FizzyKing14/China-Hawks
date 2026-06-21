# KIA 销售提成计算表 🚗

KIA 销售算自己每单能拿多少提成的表格。

## 文件
- `KIA提成表.xlsx` —— 直接用 Excel / WPS / 手机表格打开就能用
- `build_excel.py` —— 重新生成表格的脚本（`python3 build_excel.py`，需要 `pip install openpyxl`）

## 提成公式
```
我的提成 = 提成比例 × ( LUX Care份数 × 单份利润
                      + Greenway Advantage
                      + 车架利润
                      + Trade 差价 )

Trade 差价 = Trade 实际价值 − 折抵给客户的价
```

默认参数（橙色格子，可改）：
- LUX Care 单份利润 = **$50**
- Greenway Advantage = **$2,995**
- 提成比例 = **25%**

例：1 份 LUX + Greenway + 车架 $2000 + Trade 差价(18000−16000=2000)
→ 总利润 50+2995+2000+2000 = **$7,045**，×25% = **$1,761.25**

## 怎么用
1. **橙色格子**先确认参数对不对（提成比例、Greenway 价格等），不对就改
2. 每成交一单填一行（**黄色格子**）：客户/车型、LUX Care 份数、Greenway 是/否（下拉选）、车架利润；有置换就填 Trade 实际价值 + 折抵价，没有就留空
3. **绿色格子**自动算：Trade 差价、店总利润、我的提成
4. 最下面**汇总区**：成交单数、LUX 总份数、Greenway 卖出数、店总利润、我的总提成、平均每单提成

默认 60 行，不够照公式往下拖，或改 `build_excel.py` 里的 `DATA_ROWS` 重新生成。

## 参考
- Greenway Advantage 套餐：https://www.greenwaykias.com/advantage-pkg/
- LuxCare XT 外观保护：https://www.timbrookkia.com/luxcare-xt.htm
