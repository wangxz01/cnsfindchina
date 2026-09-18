# CNS 期刊爬虫

抓取 Cell / Nature / Science 的 issue 目录、文章信息、作者单位及**文章发布日期**，输出 Excel，并统计署名第一作者的中国单位情况。

## 日期与统计口径

| 来源 | Excel 日期列 | 取值规则 |
|---|---|---|
| Cell | `Available online`、`Version of Record` | 两个独立字段；读取页面明确标签或同名日期元数据；需要时点击 Show more，再打开文章历史入口 |
| Nature | `发布日期` | 文章的 online / publication 元数据，或 Published 标记 |
| Science | `发布日期` | 文章的 online / publication 元数据，或 Published 标记 |

- 日期以真正的 Excel 日期单元格写入，显示 `yyyy-mm-dd`，可排序、筛选和计算。
- Cell 的两个日期不互相代替，也不用 Received、Accepted、期刊期次日期或抓取当天日期补齐。
- 缺失日期在 Excel 留空，在网页显示“待补全”；“待补字段”指出具体缺项，“日期来源”保存提取依据。
- 统计对象是**署名第一作者**，不自动把共同一作纳入。检查其所有明确关联单位，任一国家为 China 则计“是”。
- 作者与单位通过引用关系或单位作者名单对应；不会仅因单位排在第一条就当作一作单位。
- “是否中国”有“是 / 否 / 待核实”三种状态。缺失单位、对应关系不明、国家无法识别，均不会计入“否”。
- 国家名称沿用页面地址和别名表；当前 China 判定不自动合并单独标为 Hong Kong、Taiwan 等的地址。

## 安装与启动

Windows PowerShell：

```powershell
git clone https://github.com/wangxz01/cnsfindchina.git
cd cnsfindchina
# 已安装 uv 可跳过
irm https://astral.sh/uv/install.ps1 | iex
uv sync --locked
uv run playwright install chromium
uv run python src/unified_web.py
```

macOS / Linux：

```bash
git clone https://github.com/wangxz01/cnsfindchina.git
cd cnsfindchina
curl -LsSf https://astral.sh/uv/install.sh | sh
uv sync --locked
uv run playwright install chromium
uv run python src/unified_web.py
```

打开 `http://127.0.0.1:8000/`。项目使用 `.python-version` 指定的 Python 3.12 和 `uv.lock` 锁定依赖。

```bash
uv run python src/unified_web.py --port 8080
```

## 使用网页

1. 选择 Cell / Nature / Science，填写 issue URL，每行一个；支持空行与 `#` 注释。
2. 点击“开始抓取”，程序保存配置后运行。三个来源可并行，同一来源只允许一个任务。
3. 遇到人机验证，在爬虫浏览器中手动完成，然后点“我已通过验证”；也可跳过当前文章。
4. 网页显示日期、国家判断及数据完整性；中国记录标粉色。
5. 抓取结束后下载 Excel。“部分完成”表示仍有缺失字段、失败文章或数量校验差异。
6. 点击停止后显示“正在停止”，到检查点保存已有记录，再显示“已取消”。导航请求自身有超时，停止不承诺立即中断正在进行的网络请求。

刷新网页或 SSE 断线重连后，会从后端快照恢复已抓文章、日志与校验结果，不会因重复事件重复计数。
这些网页快照保存在当前服务进程内；重启服务后不恢复旧任务界面，磁盘缓存与 Excel 保留。

“立即导出”在任务停止后，导出该来源缓存目录中的**所有历史记录**，不局限于文本框中的 issue；会按缓存记载的 issue 分组。旧缓存字段未验证时标“待核实”。运行中禁止清缓存和缓存导出，避免与抓取写入互相覆盖。

## URL 格式与 CLI

| 来源 | 配置文件 | issue URL 模式 |
|---|---|---|
| Cell | `data/urls.txt` | `https://www.sciencedirect.com/journal/cell/vol/<v>/issue/<i>` |
| Nature | `data/urls_nature.txt` | `https://www.nature.com/nature/volumes/<v>/issues/<i>` |
| Science | `data/urls_science.txt` | `https://www.science.org/toc/science/<v>/<i>` |

```bash
uv run python src/scraper.py
uv run python src/nature_scraper.py
uv run python src/science_scraper.py
# 可选参数：--out PATH / --fresh / --headless
```

`--fresh` 忽略已有缓存。遇到人工验证时建议使用默认的可见浏览器。

## Excel 输出

默认 `data/<source>_YYYY-MM-DD.xlsx`，文件名日期是运行日期，和表格内的文章发布日期不同。
同日重跑会替换同名文件；文件被 Excel 占用时改写到唯一后缀的备用文件，并返回实际路径。
Excel 和 JSON 缓存均通过临时文件原子替换，降低中途异常损坏原文件的风险。

- 第一个 sheet 为“汇总”：计划篇数、已取得记录、成功访问、失败、中国、非中国、国家待核实、数据待补全、数量校验与异常说明。“成功访问”不等于所有字段均已补齐。
- 后续每个 issue 一个 sheet，命名如 `v189-i10`；第一行为 issue URL，第二行为表头，文章按 section 分组。
- 基础字段：URL、标题、DOI、类型、第一作者、单位、国家、中国判断、作者列表。
- Cell 额外两列日期；Nature / Science 额外一列日期；最后为提取状态、待补字段、日期来源。
- 标题等外部文字按文本写入，不解释成 Excel 公式；冻结顶部表头，中国记录标粉色。

## 缓存与补抓

每篇按文章 ID 存入 `data/cache*/*.json`。缓存包含版本、来源、完整性、日期证据与作者单位关联信息。

- 当前版本且字段完整：跳过该文章页的网络访问。
- 缺日期、缺单位、国别待核实等：保留已得字段供导出；重跑会重新访问该文章，合并此前验证过的字段。
- 旧版本缓存：自动视为未命中，不用旧判断跳过日期补抓；无需手动清空全部缓存。
- 人机验证未通过或导航失败：保留失败记录在当次结果中，不保存为成功缓存。
- 即使所有文章缓存命中，程序仍启动浏览器并读取 issue 目录，以获得本期列表和校验依据。

`data/browser_profile*` 保留浏览器 Cookie。缓存、Cookie、Excel 均不入 Git。

## 数量校验

文章列表通过浏览器 DOM 提取；独立路径使用 Python HTMLParser 遍历文章链接。两条路径遵循同一纳入范围，并比较**文章 ID 集合**，不只比较总数。

- Cell：本刊 PII 链接，排除 Author / Publisher Correction。
- Nature：各 section 的 `s41586-` 链接，排除 Author / Publisher Correction；不表示包括所有新闻等其他 DOI 前缀。
- Science：各 section 的 `10.1126/science.*` 链接。
- 排除 header / footer / nav / aside 与 `.card-related` 的链接。
- 不一致时报告漏项和多项的 ID；`0/0` 不显示为校验成功。

数量一致仅说明两条提取路径得到相同文章集合，不代表日期、作者单位完整，也不能排除页面尚未加载的共同遗漏。

## 验证与代码结构

```bash
uv run python -m unittest discover -s tests -v
# 用模拟数据手动检查 UI；不访问期刊网站，不修改正式 data 文件
uv run python tests/preview_server.py
# 打开 http://127.0.0.1:8765/
# 真实文章单篇检查：打开本地浏览器，手动过验证后在终端按 Enter
uv run python tests/live_check.py cell
uv run python tests/live_check.py science
# 可替换文章；已有验证 Cookie 时可自动开始
uv run python tests/live_check.py cell --url https://www.sciencedirect.com/science/article/pii/S0092867426003946 --auto
```

`tests/fixtures` 为自行构造的 HTML 样例，覆盖双日期、作者单位关系、勘误等情形。离线测试通过不代表出版商当前页面一定可访问；真实页面仍可能遇到验证、权限或结构变化。

真实文章检查把提取前后 HTML、字段 JSON 和 Excel 留在 `data/live_checks/<source>/`，仅供本地诊断，不入 Git；同一来源再次检查会更新这些文件。Cell 新版页面的 `dates` 数据中，`Available online` 和 `Version of Record` 按各自的标签读取，值可以相同，也可以不同。

| 文件 | 用途 |
|---|---|
| `src/scraper.py` / `nature_scraper.py` / `science_scraper.py` | 各期刊页面提取与 CLI |
| `src/article_metadata.py` | 日期、作者单位关联、完整性与集合校验 |
| `src/countries.py` | 共用国家名解析 |
| `src/scraper_common.py` | 缓存、取消检查与统一任务生命周期 |
| `src/excel_writer.py` | 日期单元格、汇总页与原子输出 |
| `src/unified_web.py` | 三来源任务管理、结果快照与 SSE |
| `src/static/unified_index.html` | 网页控制台 |
| `tests/test_regressions.py` | 离线回归验证 |
