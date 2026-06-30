# CNS 期刊爬虫

统一的 Cell / Nature / Science 期刊 issue 爬虫：从一期 issue 的 URL 抓取**每篇论文的标题、DOI、作者、一作单位、一作国别、是否中国**，写入 Excel（每个 issue 一个 sheet）。

主要用途：**统计一期 CNS 期刊里有多少篇论文的一作在中国机构**。

**支持的 issue URL 模式**（三个不一样，容易混）：
- Cell：`https://www.sciencedirect.com/journal/cell/vol/<v>/issue/<i>`
- Nature：`https://www.nature.com/nature/volumes/<v>/issues/<i>`
- Science：`https://www.science.org/toc/science/<v>/<i>`

特性：
- 统一 Web 控制台，**Cell / Nature / Science 三个 source 完全隔离，可同时并行运行**
- Cloudflare / Cookie 同意墙人工放行（暂停 → 浏览器里过验证 → 网页继续）
- 反 IP 封禁：随机停顿、鼠标抖动、慢速滚动、长歇
- 按文章 ID 缓存（断点续爬），缓存命中不访问网络
- 国别识别：60+ 国家别名表，一作 Affiliation 末段匹配

## 迁移到新设备

```bash
git clone https://github.com/wangxz01/cnsfindchina.git
cd cnsfindchina
pip install -r requirements.txt
playwright install chromium
python src/unified_web.py
# 浏览器打开 http://127.0.0.1:8000/
```

前提：已装 **Python 3.10+**（Windows 安装时勾选 "Add to PATH"）。

目录布局：
```
cnsfindchina/
├── src/                  # 所有源码
│   ├── unified_web.py    # 统一入口
│   ├── scraper.py / nature_scraper.py / science_scraper.py
│   ├── scraper_common.py / excel_writer.py
│   └── static/           # 前端
├── data/                 # 运行时产物（git 不入库；urls*.txt 入库）
│   ├── urls.txt / urls_nature.txt / urls_science.txt
│   ├── browser_profile{,_nature,_science}/   # 运行后生成
│   ├── cache{,_nature,_science}/             # 运行后生成
│   └── *_*.xlsx                               # 输出
├── README.md
└── requirements.txt
```

迁移不带的（已在 `.gitignore`）：
- `data/browser_profile*/`（含 Cookie，新设备会重新过验证）
- `data/cache*/`（已抓缓存）
- `data/*.xlsx`（输出）

## 运行：统一 Web 控制台（推荐）

```bash
python src/unified_web.py            # 默认 http://127.0.0.1:8000/
python src/unified_web.py --port 8080
```

顶部 **Cell / Nature / Science** 按钮切换 source。每个 source 独立线程、独立浏览器、独立状态，**可同时启动三个并行跑**。正在运行的 source 按钮上有绿色脉动小圆点。

操作流程：
1. 选 source，编辑对应 `urls_*.txt`，点 **💾 保存到文件**
2. 点 **▶ 开始抓取**，Playwright 窗口弹出
3. CF / Cookie 墙触发时网页显眼提示；在 Playwright 窗口过验证后点 **✓ 我已通过验证**
4. 抓取中表格实时填充，**中国一作的行有粉色背景**
5. 完成后点 **⬇ 下载 Excel**

## 运行：CLI 命令行

三个独立 CLI（不依赖 Web 服务）：

```bash
python src/scraper.py            # Cell    → data/cell_YYYY-MM-DD.xlsx
python src/nature_scraper.py     # Nature  → data/nature_YYYY-MM-DD.xlsx
python src/science_scraper.py    # Science → data/science_YYYY-MM-DD.xlsx
```

通用参数：`--out PATH` / `--fresh`（忽略缓存重抓）/ `--headless`（不推荐，过 CF 需可见）

> Cell/Nature/Science 各自的 `web.py` / `nature_web.py` / `science_web.py` 已合并到 `unified_web.py`，单独的 web 入口已删除。

URL 输入文件（在 `data/` 下）：
| Source | 文件 | URL 模式 |
|---|---|---|
| Cell | `data/urls.txt` | `sciencedirect.com/journal/cell/vol/<v>/issue/<i>` |
| Nature | `data/urls_nature.txt` | `nature.com/nature/volumes/<v>/issues/<i>` |
| Science | `data/urls_science.txt` | `science.org/toc/science/<v>/<i>` |

每行一个 URL，`#` 开头为注释。

## 断点续爬（按文章 ID 缓存）

每篇成功抓取的文章以 `cache[_nature|_science]/<id>.json` 存盘。重跑时：

- **缓存命中**：直接读盘，**不开浏览器、不访问网络**，但仍写入 Excel
- **缓存未命中**：正常抓取并写缓存
- **CF 未通过 / 数据残缺**：不写缓存，下次重试

```bash
# 全部重抓
rm -rf data/cache data/cache_nature data/cache_science
# 或加 --fresh
python src/scraper.py --fresh

# 只重抓某篇
rm data/cache/S0092867426003946.json
```

修改了 scraper 的字段提取逻辑后，旧缓存不会自动失效——请手动清缓存或加 `--fresh`。

## Cloudflare / Cookie 同意墙流程

首次访问 Cell 几乎一定触发 Cloudflare；Nature/Science 有 Cookie 同意横幅。程序检测到拦截页时：

1. 网页顶部出现醒目黄色提示，含目标 URL 和当前 URL
2. 你切到 Playwright 浏览器窗口：
   - Cloudflare：勾选复选框或等自动放行
   - Cookie 墙：点 "Accept All" / "Manage Preferences" 关闭横幅
   - 必要时手动把地址栏改回目标 URL 并回车
3. 确认浏览器停在目标文章页（不是挑战页）
4. 回网页点 **✓ 我已通过 Cloudflare 验证**，程序重新扫描判断

最多提示 3 次；3 次未过则该篇记为 `[CF BLOCKED]` 继续下一篇。

`browser_profile*/` 持久化 Cookie，通过后的 `cf_clearance` 通常会被记住，同域名后续访问可能免挑战。

## 输出格式

默认输出按运行日期命名：`<source>_YYYY-MM-DD.xlsx`（同日重跑会覆盖；若文件被 Excel 占用，自动写到 `<source>_YYYY-MM-DD.<时间戳>.xlsx` 并在日志/网页提示）。

每个 issue 一个 sheet，命名 `v<v>-i<i>`（如 `v189-i10`）。每 sheet 内布局：

| 行/列 | A | B | C | D | E | F | G | H | I |
|---|---|---|---|---|---|---|---|---|---|
| 1 | issue URL | | | | | | | | |
| 2 | Articles / Research Articles / Perspective（分类名独占一行） | | | | | | | | |
| 3+ | 文章 URL | 标题 | DOI | 类型 | 一作 | 一作单位 | 一作国家 | 是否中国 | 作者列表 |

## 反 IP 封禁策略

- 文章间随机停顿 6-25s（Science 拉长到 12-25s）
- 每 4-5 篇插入 45-120s 长歇模拟阅读间歇
- 进文章前 1.5-6s 随机停顿 + 鼠标抖动
- 文章页 4-8 段慢速滚动模拟阅读
- 点击 show-more / Expand All 前鼠标抖动
- 所有固定 wait 全部带随机抖动
- `navigator.webdriver` 隐藏
- 三个 source 各自独立 profile，cookie 互不污染

## 关键文件

| 文件 | 作用 |
|---|---|
| `src/unified_web.py` | **统一 Web 服务（推荐入口）** |
| `src/static/unified_index.html` | 前端单页（Vue 3 CDN，无构建） |
| `src/scraper.py` | Cell 抓取逻辑（CLI 入口 + 三套共用的 ScraperCallbacks / CF 检测 / 随机化辅助） |
| `src/nature_scraper.py` | Nature 抓取逻辑 + 国别判定函数（`parse_country`） |
| `src/science_scraper.py` | Science 抓取逻辑 |
| `src/scraper_common.py` | 共用工具（缓存、URL 加载、DATA_DIR、safe_goto） |
| `src/excel_writer.py` | Excel 输出（多 sheet + 自定义列 schema） |
| `data/urls.txt` / `data/urls_nature.txt` / `data/urls_science.txt` | 各 source 的 issue URL 列表 |
| `requirements.txt` | 依赖 |
| `data/browser_profile*/` | Playwright 持久化浏览器配置（运行后生成，不入库） |
| `data/cache*/` | 按 PII/DOI 的文章缓存（运行后生成，不入库） |

## 字段提取策略

| 字段 | Cell | Nature | Science |
|---|---|---|---|
| 标题 | `citation_title` meta | `dc.title` meta | `dc.Title` meta |
| DOI | `citation_doi` meta | `prism.doi` meta | URL 里直接含 |
| 作者 | 内嵌 JSON `#name:author`（按 author-id 去重） | `dc.creator` meta | `dc.Creator` meta |
| 一作单位 | JSON 内 `"id":"aff1"` 后首个 `textfn` | `<li id="Aff1">` 内 address | `#con1_content` 内首个 affiliation `<span property="name">` |
| 一作国家 | `parse_country(aff)` —— Nature/Science/Cell 同口径 | 同左 | 同左 |

Cell 文章页 JSON 数据已在首次加载 HTML 内嵌，无需点击 Show more；Science affiliation 通常 display:none 但已在 DOM，无需点 Expand All；只在直取失败时才作兜底点击。
