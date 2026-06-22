# 中国各省旅游收入交互式可视化看板

这是一个基于 `pyecharts` 与 ECharts 的中国旅游收入地图交互式看板项目。项目读取 Excel 中的 2023-2025 年各省旅游收入数据，并结合各省代表旅游城市、热门景点、地方美食和图片资源，生成一个可离线打开的 HTML 数据看板。

最终产物是：

```text
china_tourism_income_2023_2025.html
```

打开后可以通过年份按钮、左右切换按钮、地图悬停和 Top10 排行图查看不同年份各省旅游收入及配套旅游信息。

## 项目结构

```text
中国旅游收入地图/
├─ main.py                              # 主程序：读取数据并生成交互式 HTML
├─ tourism_income.xlsx                  # 2023-2025 年各省旅游收入数据
├─ province_info.xlsx                   # 各省旅游城市、景点、美食信息
├─ china_tourism_income_2023_2025.html  # 已生成的最终看板页面
├─ 视频讲解.mp4                         # 项目讲解视频
├─ assets/                              # 各省旅游图片资源
│  ├─ beijing.jpg
│  ├─ shanghai.jpg
│  ├─ sichuan.jpg
│  └─ ...
├─ assets_js/                           # 本地 JS 依赖，支持离线打开
│  ├─ echarts.min.js
│  └─ maps/
│     └─ china.js
└─ README.md
```

## 数据文件

### `tourism_income.xlsx`

旅游收入主数据，共 31 行、4 列：

| 字段 | 说明 |
|---|---|
| `省份` | 省级行政区名称 |
| `2023年旅游总收入（亿元）` | 2023 年旅游总收入 |
| `2024年旅游总收入（亿元）` | 2024 年旅游总收入 |
| `2025年旅游总收入（亿元）` | 2025 年旅游总收入 |

数据单位为“亿元”。程序会自动识别列名中的年份，因此后续如果增加 `2026年旅游总收入（亿元）` 这类列，也可以沿用同一套年份提取逻辑。

### `province_info.xlsx`

省份补充信息数据，共 31 行、4 列：

| 字段 | 说明 |
|---|---|
| `省份` | 省级行政区名称 |
| `著名旅游景点` | 代表景区或景点 |
| `地方美食` | 代表地方美食 |
| `著名旅游城市` | 代表旅游城市 |

这些信息会显示在页面右侧的信息卡片中。

### `assets/`

保存各省对应图片。`main.py` 中通过省份名称映射到图片文件名，例如：

| 省份 | 图片 |
|---|---|
| 北京市 | `assets/beijing.jpg` |
| 四川省 | `assets/sichuan.jpg` |
| 云南省 | `assets/yunnan.jpg` |
| 新疆维吾尔自治区 | `assets/xinjiang.jpg` |

当前图片覆盖 31 个省级行政区。

## 功能说明

看板主要包含四个交互模块：

| 模块 | 功能 |
|---|---|
| 中国地图 | 按省份展示当前年份旅游收入，颜色越深表示收入越高 |
| 年份切换 | 支持 2023、2024、2025 三个年份切换，并带有数值过渡动画 |
| 省份信息卡 | 鼠标悬停省份时展示旅游收入、强旅游城市、热门景点、特色美食和省份图片 |
| Top10 排行图 | 展示当前年份旅游收入最高的 10 个省份 |

页面采用本地 `echarts.min.js` 和 `china.js`，生成后的 HTML 不依赖外网 CDN。

## 代码框架

`main.py` 的主要流程如下：

| 函数 | 作用 |
|---|---|
| `normalize_province_name()` | 将省份简称统一为 ECharts 中国地图使用的完整名称 |
| `province_to_asset_key()` | 将省份名称映射到 `assets/` 下的图片文件名 |
| `load_tourism_dataframe()` | 读取并整理旅游收入 Excel |
| `load_province_info()` | 读取省份旅游补充信息，并关联图片资源 |
| `extract_year_columns()` | 从收入表列名中提取年份列 |
| `build_year_payloads()` | 将不同年份的数据整理成前端可直接使用的 JSON 结构 |
| `get_top10_pairs()` | 计算某一年旅游收入 Top10 省份 |
| `build_base_map()` | 使用 pyecharts 构建中国地图基础配置 |
| `build_base_bar()` | 使用 pyecharts 构建 Top10 横向柱状图基础配置 |
| `build_html()` | 拼接完整 HTML、CSS 和 JavaScript 交互逻辑 |
| `main()` | 串联读取数据、生成图表配置、写出 HTML |

整体架构可以概括为：

```text
Excel 数据
   ↓
pandas 清洗与结构化
   ↓
pyecharts 生成地图/柱状图基础 option
   ↓
自定义 HTML + CSS + JavaScript
   ↓
交互式离线可视化看板
```

## 环境依赖

建议使用 Python 3.10 或更高版本。

需要安装：

```bash
pip install pandas openpyxl pyecharts
```

其中：

- `pandas`：读取和处理 Excel 数据；
- `openpyxl`：支持 `.xlsx` 文件读取；
- `pyecharts`：生成 ECharts 图表配置。

## 运行方式

进入项目目录：

```bash
cd D:\CODE\中国旅游收入地图
```

运行主程序：

```bash
python main.py
```

运行后会生成或覆盖：

```text
china_tourism_income_2023_2025.html
china_tourism_income_2023.html
```

其中 `china_tourism_income_2023.html` 是兼容旧命名的输出文件，当前主要使用 `china_tourism_income_2023_2025.html`。

## 查看页面

直接用浏览器打开：

```text
D:\CODE\中国旅游收入地图\china_tourism_income_2023_2025.html
```

如果浏览器对本地资源加载有限制，也可以在项目目录启动一个简单静态服务器：

```bash
python -m http.server 8000
```

然后访问：

```text
http://localhost:8000/china_tourism_income_2023_2025.html
```

## 当前数据概览

当前数据覆盖 31 个省级行政区和 2023-2025 三个年份。以 2025 年数据为例，旅游收入 Top10 省份为：

| 排名 | 省份 | 2025 年旅游总收入（亿元） |
|---:|---|---:|
| 1 | 四川 | 18000.0 |
| 2 | 江苏 | 14800.0 |
| 3 | 云南 | 12700.0 |
| 4 | 浙江 | 12179.5 |
| 5 | 广东 | 11800.0 |
| 6 | 湖南 | 11576.0 |
| 7 | 广西 | 11500.0 |
| 8 | 贵州 | 10800.0 |
| 9 | 山东 | 10800.0 |
| 10 | 江西 | 10500.0 |

## 页面设计特点

- 使用暖色系视觉映射旅游收入强弱；
- 地图、信息卡和 Top10 排行图联动；
- 年份切换时使用平滑过渡动画；
- 静态资源本地化，便于课程展示和离线演示；
- 页面右侧保留省份图片、城市、景点、美食信息，使数据看板不只是数值展示，也具有旅游主题表达。

## 扩展方式

如需扩展项目，可以从以下方向入手：

- 在 `tourism_income.xlsx` 中增加新的年份列；
- 在 `province_info.xlsx` 中补充更多省份介绍字段；
- 为港澳台或更多地区补充地图数据和图片资源；
- 增加总收入趋势折线图、年度增速图、区域分组统计等模块；
- 将静态 HTML 改造为 Flask、FastAPI 或 Streamlit 应用。

## 注意事项

- `tourism_income.xlsx` 和 `province_info.xlsx` 中的省份名称需要能被 `main.py` 统一映射到 ECharts 中国地图名称。
- 如果新增图片，需要同时在 `province_to_asset_key()` 中补充映射关系。
- 生成 HTML 时会直接写入本地相对路径，因此不要随意移动 `assets/` 和 `assets_js/`，否则页面图片或地图脚本可能无法加载。
