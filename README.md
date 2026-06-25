# 产品网评爬虫

从 **Best Buy (BBY)**、Amazon、Walmart、Target 等站点爬取商品评论，**按机型分类**导出到 **Excel**。

## 项目路径

```
E:\Work @ US R&D\产品网评&call log\爬虫
```

## 功能

| 功能 | 说明 |
|------|------|
| 多机型 | 在 `config/sites.yaml` 中配置多个机型 |
| 多站点 | 每机型可配置 bestbuy / amazon / walmart / target |
| URL 预留 | `url: ""` 留空，你手动填入商品或评论页地址 |
| Excel 导出 | `data/output/` 下生成 xlsx：总表 + 每个机型一个 sheet |

## 快速开始

```powershell
cd "E:\Work @ US R&D\产品网评&call log\爬虫"

python -m venv .venv
.venv\Scripts\activate
pip install -e .
python -m playwright install chromium
```

### Amazon 评论（需要先保存登录）

Amazon 在非登录环境下常跳到 **Sign-In**，因此需在本机跑一次登录保存脚本：

```powershell
python scripts/save_amazon_state.py
```

浏览器打开后在 Amazon 登录或完成验证 → 回到终端 **按回车** → 生成 `data/amazon_state.json`。再执行：

```powershell
python scripts/run_scraper.py
# 或
review-scraper run -o amazon_B0GR9NR9XV.xlsx
```

Amazon 已按星级拆分抓取，以减少 `all_stars` 页面漏抓：

```yaml
amazon_5:  # five_star
amazon_4:  # four_star
amazon_3:  # three_star
amazon_2:  # two_star
amazon_1:  # one_star
```

只跑某一机型的全部 Amazon 星级：

```powershell
review-scraper run --model S7S --site amazon
```

只跑某个星级：

```powershell
review-scraper run --model S7S --site amazon_5
```

若人机验证较难识别，可把 `config/sites.yaml` 中 `amazon_playwright_headless` 改为 `false`，用有界面浏览器操作。

### 1. 填写 URL

编辑 **`config/sites.yaml`**：

```yaml
models:
  - model_id: "pixel-9"
    display_name: "Google Pixel 9"
    sites:
      bestbuy:
        enabled: true
        url: "https://www.bestbuy.com/site/..."   # ← 填这里
        max_pages: 5
```

- `enabled: true` 且 `url` 非空才会爬取
- `max_pages`：Best Buy 分页时最多爬几页

### 2. 查看配置

```powershell
review-scraper list
```

### 3. 执行爬取

```powershell
review-scraper run
```

只爬某一机型或站点：

```powershell
review-scraper run --model "Google Pixel 9"
review-scraper run --site bestbuy
```

输出示例：`data/output/reviews_20260515_143022.xlsx`

### 3.1 一键爬取 + AI 分析

如果已经配置好 API Key，并希望执行后自动完成“网上抓评论 → 1-3 星低分语义分类 → 生成最终报告”，可直接运行：

```powershell
python scripts/run_full_workflow.py -o data/output/review_issue_analysis_latest.xlsx
```

只跑某个机型或站点：

```powershell
python scripts/run_full_workflow.py --model S7S -o data/output/review_issue_analysis_latest.xlsx
python scripts/run_full_workflow.py --site amazon -o data/output/review_issue_analysis_latest.xlsx
```

该流程会先在 `data/output/` 下生成新的 `reviews_YYYYMMDD_HHMMSS.xlsx`，再自动把这份文件传给 `analyze_reviews.py` 做 API 语义分析。

### 4. 分析低分评论

爬取完成后，可按机型分析 3 星及以下评论，并输出问题分类占比报告：

```powershell
python scripts/analyze_reviews.py
```

默认读取 `data/output/` 下最新的 `reviews_*.xlsx`，输出：

```text
data/output/review_issue_analysis_YYYYMMDD_HHMMSS.xlsx
```

问题分类默认读取 `data/output/Call log分类.xlsx`，按该文件的列名作为分类名称、列内内容作为匹配规则。

也可以指定输入和输出文件：

```powershell
python scripts/analyze_reviews.py -i data/output/reviews_20260611_140954.xlsx -o data/output/my_analysis.xlsx
```

如需指定另一份分类表：

```powershell
python scripts/analyze_reviews.py -c data/output/Call log分类.xlsx
```

如果没有 OpenAI API Key，但希望用 ChatGPT 网页版做语义分类，可先导出待分类模板：

```powershell
python scripts/analyze_reviews.py --classification-mode export-template --template-output data/output/review_classification_template_latest.xlsx
```

把生成的 `review_classification_template_latest.xlsx` 上传到 ChatGPT，让 ChatGPT 阅读完整 review 语义后填写 `一级分类_请填写` 和 `二级分类_请填写`。

ChatGPT 网页版分类完成并保存为 `review_classification_template_latest_completed.xlsx` 后，生成最终汇总报告：

```powershell
python scripts/summarize_classified_reviews.py -o data/output/review_classification_summary_latest.xlsx
```

报告会包含一个 `总表`，按型号、一级分类统计问题数和占比；同时每个机型单独一个 sheet。

各机型 sheet 的明细区域会保留完整的原始评论内容，不输出 `ReviewID`、“标题”和“问题摘要”列。长评论单元格会自动换行并扩展行高，便于直接检查完整网评。

后续如果有可用 OpenAI API Key，可在 `.env` 中添加：

```env
OPENAI_API_KEY=你的key
```

然后直接运行默认语义分类：

```powershell
python scripts/analyze_reviews.py --classification-mode semantic -o data/output/review_issue_analysis_latest.xlsx
```

也可以明确指定输入文件：

```powershell
python scripts/analyze_reviews.py -i data/output/reviews_YYYYMMDD_HHMMSS.xlsx --classification-mode semantic -o data/output/review_issue_analysis_latest.xlsx
```

API 模式会直接生成最终分析报告，无需再执行“导出模板 → 上传网页版 ChatGPT → 下载 → 汇总”的人工流程。程序会把 `Call log分类.xlsx` 作为固定分类体系传给模型，并在本地校验一级、二级分类；无法确定时回退为 `Other / Please add comments`。

语义分类内置以下严格边界：

- `Power_on_Hardware / Cannot turn on`：只有评论明确表示无法开机、按电源无反应、无法从待机唤醒或设备无法启动时才能使用；不能仅凭 `stopped working`、`died`、`broken`、`failed` 判断。
- `OTA_failure`：只有升级、固件或 OTA 更新过程本身检测、下载、安装、完成失败或卡住时才能使用；更新完成后出现的其他问题应按实际故障分类。

API 返回结果还会写入简短的 `分类理由`，方便人工复核。

## Excel 列说明

| 列名 | 说明 |
|------|------|
| 机型ID / 机型名称 | 来自配置 |
| 来源站点 | bestbuy / amazon 等 |
| 来源URL | 实际爬取的页面 |
| 评分 / 标题 / 评论内容 | 解析结果 |
| 评论日期 / 用户 | 若有 |
| 抓取时间 | ISO 时间戳 |

## 与 OpenClaw / 其他助手

可将 `review-scraper run` 注册为 Shell Skill，由助手定时或按需触发。

## 注意事项

1. **合规**：仅爬取你有权访问的公开页面，遵守网站 ToS。
2. **Amazon**：必须用 `scripts/save_amazon_state.py` 保存 Cookie 后方可稳定抓取评论。
3. **反爬**：部分站点为动态加载，若抓不到数据，可用 F12 检查 DOM，并修改 `scrapers/` 下选择器。
4. Best Buy：`scrapers/bestbuy.py`。请求间隔：`config/sites.yaml` → `defaults.request_delay_seconds`。

## 项目结构

```
爬虫/
├── config/sites.yaml      # 机型 + URL（手动填写）
├── src/review_scraper/
│   ├── cli.py
│   ├── pipeline.py
│   ├── scrapers/          # bestbuy / amazon / generic
│   └── export/excel.py
└── data/output/           # 导出的 xlsx
```
