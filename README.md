# BUFF Inventory Exporter

Python exporter for BUFF CS inventory pages. It reads BUFF cookies for multiple accounts, fetches the inventory API, and exports an Excel workbook.

## Setup

```powershell
python -m venv .venv
.\.venv\Scripts\pip install -r requirements.txt
Copy-Item config.example.json config.json
```

Edit `config.json` and fill in multiple accounts:
List more accounts in the following format
```
{
  "name": "account1",
  "steamid": "...",
  "cookie": "..."
}
```
- `name`: Excel sheet name.
- `steamid`: SteamID used in the BUFF inventory URL.
- `cookie`: Copy the BUFF request `Cookie` header from the browser.

Do not commit or share `config.json`. The cookie is equivalent to a login session.

## How to Get Cookie

1. Open BUFF in the browser and log in.
2. Open DevTools, go to `Network`.
3. Refresh the BUFF inventory page.
4. Filter by `steam_inventory` or `api`.
5. Click the inventory API request.
6. Copy `Request Headers -> Cookie`.
7. Paste it into the matching account in `config.json`.
<img width="1572" height="730" alt="image" src="https://github.com/user-attachments/assets/799b318b-f71f-45e1-b8ee-4fc3e691d26f" />

The cookie format can include or omit the `Cookie:` prefix.

## Export

```powershell
.\.venv\Scripts\python export_inventory.py
```

The workbook is written to `exports/buff_inventory_YYYY-MM-DD.xlsx`. Each account gets one sheet with these columns:

```text
饰品, 磨损, 买入价, 当前售价, 冷却时间, 利润, 利润率
```

`利润` and `利润率` are Excel formulas, not static values.

Export details:

- 饰品名称优先使用 BUFF 网页显示的中文名。
- 金额列写入数字格式，不带 `¥` 符号。
- 饰品顺序按网页展示倒序写入，网页第一个饰品会在 Excel 最下面。

If fields do not parse correctly after BUFF changes its response shape, run:

```powershell
.\.venv\Scripts\python export_inventory.py --debug-json
```

Then inspect `logs/*_first_page.json`.

## Linux Cron Example

```cron
0 12 * * 0 cd /opt/buff-inventory && .venv/bin/python export_inventory.py >> logs/cron.log 2>&1
```
