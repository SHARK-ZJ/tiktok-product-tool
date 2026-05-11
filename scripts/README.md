# 1688 抓价脚本

这里有两个独立脚本，和前端网页分开。

## 1. 普通抓取脚本

适合 1688 页面不要求登录时使用：

```powershell
python scripts\fetch_1688_prices.py "店小秘产品表.xlsx" -o "1688-cost-table.csv"
```

先测试前 5 条：

```powershell
python scripts\fetch_1688_prices.py "店小秘产品表.xlsx" -o "1688-cost-table.csv" --limit 5
```

## 2. 浏览器登录态脚本

如果普通脚本一直提示“页面需要登录或验证”，用这个：

```powershell
pip install playwright
python -m playwright install chromium
python scripts\fetch_1688_prices_browser.py "店小秘产品表.xlsx" -o "1688-cost-table.csv"
```

运行后会打开浏览器。你需要：

1. 在打开的浏览器里登录 1688。
2. 如果出现验证码/验证，手动完成。
3. 回到命令行窗口按 Enter。
4. 脚本继续批量打开 1688 商品页抓价。

浏览器登录态会保存在：

```text
.1688-browser-profile
```

下次运行可以复用登录状态。

## 输出字段

```csv
来源Url,offerId,1688商品标题,1688采购价,1688重量(g),1688运费,抓取状态,备注
```

脚本会按 1688 offer ID 去重，并跳过输出文件中已经成功抓取过的链接。

注意：脚本不会破解验证码，也不会绕过 1688 登录限制。如果页面要求验证，需要你在浏览器里手动完成。
