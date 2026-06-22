# Cell 期刊 issue 爬虫

把要爬的 Cell issue URL 放进 `urls.txt`，程序用模拟浏览器逐个打开并抓取
**标题 / DOI / 作者 / 首条作者单位注释**，写入 Excel。每个 issue 一个 sheet。
支持 Cloudflare 人工放行、模拟人类节律避免 IP 封禁。

## 安装

```bash
pip install -r requirements.txt
playwright install chromium
```

## 运行

1. 编辑 `urls.txt`，每行一个 issue URL（`#` 开头为注释）：
   ```
   https://www.sciencedirect.com/journal/cell/vol/189/issue/10
   https://www.sciencedirect.com/journal/cell/vol/189/issue/11
   ```
2. 关闭可能占用输出文件的程序（Excel 等）
3. 运行：
   ```bash
   python scraper.py
   # 输出默认按日期命名：cell_2026-06-22.xlsx
   # 也可指定路径
   python scraper.py --out result.xlsx
   ```

## 断点续爬（按 PII 缓存）

每篇成功抓取的文章以 `cache/<PII>.json` 存盘。重跑时：

- **缓存命中**：直接从磁盘读，**不访问网络、不开浏览器**，但仍写入 Excel → 输出永远完整
- **缓存未命中**：正常抓取并写缓存
- **CF 未通过 / 数据残缺**：不写缓存，下次重试

控制缓存：

```bash
# 全部重抓（等价 --fresh）
rm -rf cache

# 只重抓某篇（删除其缓存文件）
rm cache/S0092867426003946.json

# 命令行强制忽略缓存
python scraper.py --fresh
```

注意：如果你修改了 `scraper.py` 的字段提取逻辑，旧缓存不会自动失效——
要重新应用新逻辑请 `rm -rf cache` 或 `--fresh`。

## Cloudflare 验证流程

首次访问 ScienceDirect 几乎一定会触发 Cloudflare。程序会：

1. 检测到挑战后暂停，终端打印：
   ```
   [Cloudflare] 检测到人机验证挑战（第 1/3 次）。
   [Cloudflare] 目标 URL: https://www.sciencedirect.com/...
   [Cloudflare] 当前 URL: ...
   [Cloudflare] 请在浏览器窗口中：
     1) 完成人机验证（勾选复选框或等自动放行）
     2) 必要时手动把地址栏改成目标 URL 并回车
     3) 确认浏览器停在目标文章页（非 CF 挑战页）
   >>> 完成后回到此终端按 Enter（第 1/3 次）:
   ```
2. 你切到浏览器窗口，完成验证或手动导航到目标文章页。
3. 回到终端按 Enter，程序重新扫描页面内容判断是否真的通过。
4. 最多提示 3 次；3 次仍未通过则该篇记为 `[CF BLOCKED]` 并继续下一篇。

`./browser_profile` 持久化浏览器 cookie，CF 通过后的 cf_clearance cookie 通常会被
记住，后续同域名访问可能免挑战。

## 输出格式

默认输出文件按运行日期命名：`cell_YYYY-MM-DD.xlsx`（同日重跑会覆盖）。
**每个 issue 一个 sheet**（按 `v189-i10` 形式命名）。
每个 sheet 内部布局：

| 行/列 | A | B | C | D | E |
|---|---|---|---|---|---|
| 1 | issue URL | | | | |
| 2 | Articles（分类名独占一行） | | | | |
| 3 | 文章 URL | 标题 | DOI | 作者(分号拼接) | 首条 affiliation |
| 4 | 文章 URL | 标题 | DOI | 作者 | 首条 affiliation |
| 5 | Short Articles | | | | |
| 6 | ... | | | | |

只抓 `Articles / Short Articles / Resources` 三个 section
（跳过 Leading Edge / Previews / Review / Corrections）。

## 反 IP 封禁策略

- 文章间随机停顿 6-15s，每 5-8 篇插入 45-120s 长歇
- 进文章前 1.5-4s 随机停顿 + 鼠标抖动
- 文章页做 4-8 段慢速滚动模拟阅读
- 点击 show-more 前鼠标抖动
- 所有原本固定的 wait 全部带随机抖动
- 鼠标在视口内做不规则移动
- `navigator.webdriver` 隐藏

## 输出文件被占用？

若输出文件（如 `cell_2026-06-22.xlsx`）被 Excel 打开导致写入失败，
程序会自动写到 `cell_2026-06-22.<时间戳>.xlsx`，终端会打印实际路径。

## 关键文件

- `scraper.py`  主程序（CLI 入口、Playwright、CF 处理、字段提取）
- `excel_writer.py`  Excel 输出（多 sheet）
- `urls.txt`  待抓 issue URL 列表
- `requirements.txt`  依赖
- `browser_profile/`  持久化浏览器配置（运行后生成）

## 字段提取策略

文章页数据已在首次加载的 HTML 内嵌 JSON 数据模型中（无需点击 Show more）：

- 标题：`<meta name="citation_title">`
- DOI：`<meta name="citation_doi">` 或第一个 `10.1016/j.cell.xxx`
- 作者：按 `author-id` 去重，配对 `{"#name":"given-name","_":"X"}` 与 `{"#name":"surname","_":"Y"}`
- 首条 affiliation：定位第一个 `"id":"aff1"` 之后的 textfn
  （跳过前部 "Preview" 等文章类型脚注）

若 affiliation 缺失，回退 JS 点击 `#show-more-btn` 再试；仍取不到则 DOM 选择器兜底。
