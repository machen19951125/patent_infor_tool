# patent_infor_tool
用cas号查询相关动态的工具
# 询单产品证据工作台

工作台将流程分为两段：证据采集脚本负责采集、归档、去重和压缩；总结脚本只将已入包的证据逐产品交给大模型。两段独立运行、独立计量，总结阶段不重新搜索网页。

## 数据源

- PubChem：身份、结构、同义词、药物注释和专利标识候选
- Google Patents：从 PubChem 专利索引中选取代表性页面，提取标题、摘要及 CAS/名称附近的正文上下文
- ChEMBL：药物/生物活性数据库身份与最高临床阶段
- Europe PMC、PubMed：论文和摘要
- ClinicalTrials.gov：临床试验候选
- Tavily Search（默认启用、需要 key）：供应商目录、原始专利网页及 CAS 直接用途线索
- MedChemExpress / MCE（可选、需要 Tavily key）：定向收集 MCE 人工整理的靶点、机制、通路、IC50 及实验背景
- Brave Search（可选备用）：另一套开放网页检索接口

没有 Tavily/Brave key 时，其余结构化数据源仍可正常工作。

## 安装

```powershell
python -m pip install -r requirements.txt
```

## 最小用法

```powershell
python scripts\collect_inquiry_evidence.py "询单范例.xlsx"
```

默认读取第一个工作表中的 `BD`、`CAS` 两列，结果写入带时间戳的 `outputs/evidence-*` 目录。

## Streamlit 图形界面

安装依赖后，双击工作区根目录中的 `启动询单证据采集器.bat`。启动器会向 Windows 自动申请一个当前空闲的随机端口，并且只监听本机 `127.0.0.1`；没有固定端口，也不会向局域网开放服务。

也可以从 PowerShell 启动：

```powershell
python scripts\run_inquiry_evidence_app.py
```

界面支持：

- 使用工作区示例或上传新的 Excel/CSV
- 切换到“突发单个 CAS”，只输入 CAS 号即可直接采集，无需临时创建 Excel
- 自动恢复上次的输入模式、文件路径、数据源、并发参数、模型设置和 API key
- API key 按用户授权明文保存在 `.cache/gui_settings.json`，不写入任务输出或命令行
- 勾选数据源及设置并发数、网页结果数和证据包大小
- 将 MedChemExpress (MCE) 作为独立可选数据源，单独统计其 Tavily credits
- 先测试 1 个产品，再处理全部产品
- 实时查看脚本日志和停止任务
- 查看 Tavily credits、缓存命中、候选/入选证据数和证据包 Token 粗估
- 分别下载结果文件或下载完整 ZIP
- 选择历史证据包，先离线导出大模型请求
- 在 OpenAI Responses API 与 DeepSeek Chat Completions API 之间切换
- 根据提供商动态显示模型、推理、思考模式、温度和 API 地址等参数
- 逐产品调用模型，输出严格校验的 JSON、CSV 和 Markdown 总结
- 记录 API 返回的每个产品输入、输出、推理和总 token

指定输出目录：

```powershell
python scripts\collect_inquiry_evidence.py "询单范例.xlsx" `
  --output-dir "outputs\本次询单证据"
```

只处理前两个产品做联网测试：

```powershell
python scripts\collect_inquiry_evidence.py "询单范例.xlsx" --limit 2
```

突发单产品可以不经过文件直接查询：

```powershell
python scripts\collect_inquiry_evidence.py --cas 50-78-2
```

`--cas` 会检查 CAS 格式及校验位，并在 `run_metadata.json` 中记录 `input_mode=single_cas`。

## Tavily 开放网页检索

Tavily 能补齐商业目录、原始专利和 CAS 直接用途线索。脚本固定使用 `basic` 检索，关闭 Tavily 的生成式回答，只保存搜索结果和查询相关片段。密钥通过环境变量传入，不会写入输出文件或 HTTP 缓存：

```powershell
$env:TAVILY_API_KEY = "你的密钥"
python scripts\collect_inquiry_evidence.py "询单范例.xlsx"
```

每个产品默认发出 2 个查询：精确 CAS，以及“CAS 或 PubChem 规范英文名” + 专利/中间体/杂质/合成用途。因此英文名扩展不会增加默认查询次数。包含目标 CAS 的结果作为直接证据；不含 CAS 时，只有在原始专利/论文权威域名上精确出现足够独特的完整英文名，才可进入默认证据包。供应商页面或宽泛短名称的无 CAS 结果仍标为 `indirect`。Tavily 返回的 credits、本次实际联网消耗和缓存命中数会写入原始记录及 `run_metadata.json`。重复运行命中本地缓存时不会再次调用 Tavily。

如果确实需要收集名称相关的近期事件，可在界面中手动勾选，或在命令行增加：

```powershell
--include-recent-events
```

这类结果默认按间接候选归档，不进入大模型证据包。

先处理一个产品验证密钥和结果：

```powershell
$env:TAVILY_API_KEY = "你的密钥"
python scripts\collect_inquiry_evidence.py "询单范例.xlsx" `
  --limit 1 `
  --output-dir "outputs\tavily-smoke-test"
```

Brave 仍作为可选备用源保留；需要同时启用时显式指定：

```powershell
$env:BRAVE_SEARCH_API_KEY = "你的密钥"
python scripts\collect_inquiry_evidence.py "询单范例.xlsx" `
  --sources pubchem,googlepatents,chembl,europepmc,pubmed,clinicaltrials,tavily,brave
```

## MedChemExpress (MCE) 定向资料

MCE 在 GUI 的“选择数据源”中单独列出，默认不勾选。启用后，每个产品额外执行 1 次 Tavily `basic` 查询，并通过 `include_domains=["medchemexpress.com"]` 将结果限定在 MCE 域名。脚本不直接并发抓取 MCE 网页。

为防止大批量询单引起不必要的服务器压力，实施了以下保护：

- 默认不截断产品数；200 或更多产品会自动排队完成，不需要手工拆批。
- Tavily 主机请求全局串行节流，两次联网请求至少间隔 0.5 秒，并继续尊重 `429/Retry-After`。
- 14 天内的同一查询优先读取本地缓存，不重复联网。
- GUI 保留可选软上限：`0` 表示处理全部；如果只想试跑前 50 个，才主动填写 `50`。超过软上限的产品会显式标记 `skipped`。

MCE 页面按“人工整理的二级药理证据”处理：

- 页面文本精确包含目标 CAS 时，可进入默认证据包，用于支持靶点、机制、通路和实验背景。
- 只命中英文名称但未出现 CAS 时，仍作为 `indirect` 候选保存。
- MCE 资料不自动取代其引用的原始论文或专利，也不单独证明近期询单放量。
- MCE 定向查询的实际消耗和缓存命中分别写入 `mce_tavily_credits_this_run` 和 `mce_tavily_cache_hits`。

命令行启用示例：

```powershell
$env:TAVILY_API_KEY = "你的密钥"
python scripts\collect_inquiry_evidence.py --cas 50-78-2 `
  --sources pubchem,mce
```

如果某次只想试跑前 20 个 MCE 查询：

```powershell
--mce-max-products 20
```

也可以设置 NCBI API key 和真实联系邮箱，以获得更稳定的 PubMed 批量访问：

```powershell
$env:NCBI_API_KEY = "你的密钥"
$env:NCBI_EMAIL = "your.name@example.com"
```

## 输出文件

| 文件 | 用途 |
|---|---|
| `llm_evidence_pack.md` | 已压缩、分产品排列，可直接交给大模型总结 |
| `llm_token_estimates.json` | 证据包总体及逐产品的粗略输入 token 估计 |
| `evidence_candidates.csv` | 一条证据一行，含匹配依据、直接 CAS 标记和 direct/indirect 范围 |
| `products.jsonl` | 每个产品的身份、数据源状态和原始文件位置 |
| `raw/*.json` | 每个产品的完整标准化采集结果，供审计和重新生成证据包 |
| `run_metadata.json` | 运行参数、来源状态和输出清单 |
| `resume_state.json` | 断点续跑进度、原产品列表和非密钥采集参数 |
| `checkpoints/*.json` | 每完成一个分子就原子写入的产品级检查点 |

HTTP 原始响应会缓存在 `.cache/inquiry_evidence`。默认 14 天内复用，因此重复运行同一批 CAS 会更快，也不会重复请求公共接口。

## 断点续跑

采集器 0.4.1 起具备产品级断点。每完成一个分子，立即写入 `checkpoints/` 并更新 `resume_state.json`；不需要等全批产品完成才落盘。

- GUI 会自动扫描未完成任务，显示“已完成/总数”，点击“从检查点继续”即可。
- 续跑直接跳过已完成分子，只调度剩余产品。
- 中断时正在处理、尚未写成产品检查点的分子会重新调度；它们已成功返回的 HTTP/Tavily 请求仍会优先命中 14 天缓存。
- 原产品列表、数据源和会影响结果的参数从续跑状态恢复；API Key 不写入状态，使用当前 GUI/环境变量中的值。
- 所有产品完成后，脚本按原输入顺序合并检查点，再生成正常的 CSV/JSONL/Markdown 结果。

命令行续跑：

```powershell
python scripts\collect_inquiry_evidence.py `
  --resume-dir "outputs\streamlit-原任务时间戳"
```

0.4.1 之前已中断的旧任务没有产品检查点，无法还原当时只存在内存中的完整产品记录；但重新发起同一批 CAS 时，仍会复用已保存的 HTTP 缓存。

## 大模型总结

总结器读取某个采集输出目录中的 `llm_evidence_pack.md`。它不配置任何网页搜索工具，也不调用 Tavily、PubChem 或其他采集源。每个产品使用独立请求，可以将 API token 用量精确归到 BD/CAS。

总结结果中的化合物名称只保留 `chemical_name_en`，优先使用 PubChem preferred name，必要时使用 IUPAC English name。不生成或展示中文翻译，以保留立体化学、位点和标点的原始表达，并便于后续英文检索。

| 提供商 | 官方根地址 | 接口 | 默认模型/选项 | 结构化方式 |
|---|---|---|---|---|
| OpenAI | `https://api.openai.com/v1` | `POST /responses` | `gpt-5.6-luna/terra/sol` 或自定义 | strict `json_schema` |
| DeepSeek | `https://api.deepseek.com` | `POST /chat/completions` | `deepseek-v4-flash/pro` 或自定义 | `json_object` + 本地业务 Schema 校验 |

DeepSeek JSON 输出仍可能出现空内容或业务字段不合格。总结器因此保留本地严格校验和可配置重试；重试产生的所有 token 会累加到对应产品。

先做不联网的完整拆包与 token 预算测试：

```powershell
python scripts\summarize_inquiry_evidence.py `
  "outputs\streamlit-20260819-182739-154" `
  --dry-run `
  --output-dir "outputs\summary-dry-run"
```

使用 OpenAI 真实调用：

```powershell
$env:OPENAI_API_KEY = "你的模型 API 密钥"
python scripts\summarize_inquiry_evidence.py `
  "outputs\streamlit-20260819-182739-154" `
  --model "你的 API 账户可用的模型名称" `
  --reasoning-effort medium
```

使用 DeepSeek 真实调用：

```powershell
$env:DEEPSEEK_API_KEY = "你的 DeepSeek API 密钥"
python scripts\summarize_inquiry_evidence.py `
  "outputs\streamlit-20260819-182739-154" `
  --provider deepseek `
  --model deepseek-v4-flash `
  --thinking-mode enabled `
  --reasoning-effort high `
  --temperature 0.2
```

真实总结输出：

| 文件 | 用途 |
|---|---|
| `summary_results.jsonl` | 每个产品一行的严格 JSON 结果；含 E→证据详情和专利号数组 |
| `summary_results.csv` | 可直接打开或并回 Excel 的总结表；含 E→标题/来源/URL、已核验专利号、候选专利号 |
| `summary_report.md` | 分产品的可读报告；正文后附引用证据索引 |
| `summary_token_usage.json` | API 返回的总计及逐产品 token |
| `raw_responses/*.json` | 去除密钥后的原始模型响应，供审计 |
| `summary_run_metadata.json` | 模型、接口、成功/失败产品和错误记录 |

离线 `--dry-run` 只产生 `prompts/*.request.json`、`prompt_manifest.json` 和运行元数据，不会连接模型服务。

E01/E02 等是“当前产品证据包中的局部序号”，不是专利号。总结脚本在模型输出校验后，会从证据包确定性地回填引用的证据元数据；这一步不调用模型或搜索 API。`verified_original_patent` 的发布号进入 `cited_verified_patent_numbers`；`patent_index` 中的号码只进入 `cited_candidate_patent_numbers`，不会被误标为已核验专利。

已有的历史总结可以完全离线回填，默认新建同级 `-evidence-mapped` 目录，不覆盖原结果：

```powershell
python scripts\backfill_summary_citations.py `
  "outputs\summary-streamlit-20260820-160010-069"
```

## 常用参数

```text
--sheet Sheet1
--bd-column BD
--cas-column CAS
--sources pubchem,chembl,europepmc,pubmed,clinicaltrials,tavily
--max-workers 4
--patent-pages 4
--cache-ttl-days 14
--offline
--pack-items 14
--snippet-chars 700
```

`--offline` 只读取已有缓存，适合在采集完成后重复生成输出。完整参数请运行：

```powershell
python scripts\collect_inquiry_evidence.py --help
```

## 证据边界

- CAS 命中 PubChem 或供应商目录，只证明身份候选，不自动证明药物路线。
- PubChem 中出现专利标识，只是专利候选；必须再检查专利正文是否直接使用该 CAS、名称或结构。
- 名称回退检索可能带来同名或宽泛结果，完整 CSV/JSON 会保留检索式、匹配范围和相关性分数供复核。
- 名称回退、无身份匹配的专利候选和名称相关临床试验只进入完整候选档案，默认不进入大模型证据包。
- 证据包始终执行来源、证据类型和网页域名配额，即使总候选数少于 `--pack-items` 也不会让单一搜索源占满。
- `llm_token_estimates.json` 是本地字符法估计，不是 Codex/ChatGPT 账户的实际扣量。

