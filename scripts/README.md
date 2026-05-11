# 1688 抓价脚本

独立脚本：`fetch_1688_prices.py`

## 用法

```powershell
python scripts\fetch_1688_prices.py 店小秘产品表.xlsx -o 1688-cost-table.csv
```

测试前几条：

```powershell
python scripts\fetch_1688_prices.py 店小秘产品表.csv -o 1688-cost-table.csv --limit 5
```

强制重抓，不使用已有成功结果断点：

```powershell
python scripts\fetch_1688_prices.py 店小秘产品表.csv -o 1688-cost-table.csv --force
```

## 输出字段

```csv
来源Url,offerId,1688商品标题,1688采购价,1688重量(g),1688运费,抓取状态,备注
```

脚本会按 1688 offer ID 去重，并跳过输出文件中已经成功抓取过的链接。

注意：1688 可能要求登录、验证码或触发反爬。此时脚本会把该行标记为“失败”，不会中断整个任务。
