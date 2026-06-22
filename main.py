from __future__ import annotations

from pathlib import Path
import json
import re

import pandas as pd
from pyecharts import options as opts
from pyecharts.charts import Bar, Map
from pyecharts.commons.utils import JsCode


BASE_DIR = Path(__file__).resolve().parent
EXCEL_FILE = BASE_DIR / "tourism_income.xlsx"
INFO_FILE = BASE_DIR / "province_info.xlsx"
ASSETS_DIR = BASE_DIR / "assets"
OUTPUT_HTML = BASE_DIR / "china_tourism_income_2023_2025.html"
LEGACY_OUTPUT_HTML = BASE_DIR / "china_tourism_income_2023.html"
ECHARTS_JS = "./assets_js/echarts.min.js"
CHINA_MAP_JS = "./assets_js/maps/china.js"
PAGE_TITLE = "中国各省旅游收入可视化（2023-2025）"
WARM_COLORS = ["#fff5eb", "#fdcc8a", "#fc8d59", "#e34a33", "#b30000"]
TRANSITION_DURATION_MS = 1400


def normalize_province_name(name: str) -> str:
    """Normalize province names to the full names used by china.js."""
    text = str(name).strip()
    province_name_map = {
        "北京": "北京市",
        "北京市": "北京市",
        "天津": "天津市",
        "天津市": "天津市",
        "上海": "上海市",
        "上海市": "上海市",
        "重庆": "重庆市",
        "重庆市": "重庆市",
        "河北": "河北省",
        "河北省": "河北省",
        "山西": "山西省",
        "山西省": "山西省",
        "辽宁": "辽宁省",
        "辽宁省": "辽宁省",
        "吉林": "吉林省",
        "吉林省": "吉林省",
        "黑龙江": "黑龙江省",
        "黑龙江省": "黑龙江省",
        "江苏": "江苏省",
        "江苏省": "江苏省",
        "浙江": "浙江省",
        "浙江省": "浙江省",
        "安徽": "安徽省",
        "安徽省": "安徽省",
        "福建": "福建省",
        "福建省": "福建省",
        "江西": "江西省",
        "江西省": "江西省",
        "山东": "山东省",
        "山东省": "山东省",
        "河南": "河南省",
        "河南省": "河南省",
        "湖北": "湖北省",
        "湖北省": "湖北省",
        "湖南": "湖南省",
        "湖南省": "湖南省",
        "广东": "广东省",
        "广东省": "广东省",
        "海南": "海南省",
        "海南省": "海南省",
        "四川": "四川省",
        "四川省": "四川省",
        "贵州": "贵州省",
        "贵州省": "贵州省",
        "云南": "云南省",
        "云南省": "云南省",
        "陕西": "陕西省",
        "陕西省": "陕西省",
        "甘肃": "甘肃省",
        "甘肃省": "甘肃省",
        "青海": "青海省",
        "青海省": "青海省",
        "台湾": "台湾省",
        "台湾省": "台湾省",
        "内蒙古": "内蒙古自治区",
        "内蒙古自治区": "内蒙古自治区",
        "广西": "广西壮族自治区",
        "广西壮族自治区": "广西壮族自治区",
        "西藏": "西藏自治区",
        "西藏自治区": "西藏自治区",
        "宁夏": "宁夏回族自治区",
        "宁夏回族自治区": "宁夏回族自治区",
        "新疆": "新疆维吾尔自治区",
        "新疆维吾尔自治区": "新疆维吾尔自治区",
        "香港": "香港特别行政区",
        "香港特别行政区": "香港特别行政区",
        "澳门": "澳门特别行政区",
        "澳门特别行政区": "澳门特别行政区",
    }
    return province_name_map.get(text, text)


def province_to_asset_key(province_name: str) -> str | None:
    """Map normalized province names to image file stems."""
    asset_map = {
        "北京市": "beijing",
        "天津市": "tianjin",
        "上海市": "shanghai",
        "重庆市": "chongqing",
        "河北省": "hebei",
        "山西省": "shanxi",
        "内蒙古自治区": "neimenggu",
        "辽宁省": "liaoning",
        "吉林省": "jilin",
        "黑龙江省": "heilongjiang",
        "江苏省": "suzhou",
        "浙江省": "zhejiang",
        "安徽省": "anhui",
        "福建省": "fujian",
        "江西省": "jiangxi",
        "山东省": "shandong",
        "河南省": "henan",
        "湖北省": "hubei",
        "湖南省": "hunan",
        "广东省": "guangdong",
        "广西壮族自治区": "guangxi",
        "海南省": "hainan",
        "四川省": "sichuan",
        "贵州省": "guizhou",
        "云南省": "yunnan",
        "西藏自治区": "xizang",
        "陕西省": "shaanxi",
        "甘肃省": "gansu",
        "青海省": "qinghai",
        "宁夏回族自治区": "ningxia",
        "新疆维吾尔自治区": "xinjiang",
    }
    return asset_map.get(province_name)


def load_tourism_dataframe(excel_file: Path) -> pd.DataFrame:
    """Load and normalize the tourism income workbook."""
    if not excel_file.exists():
        raise FileNotFoundError(f"未找到数据文件：{excel_file}")

    try:
        df = pd.read_excel(excel_file)
    except ImportError as exc:
        raise SystemExit("缺少 openpyxl，请先运行：pip install openpyxl") from exc

    df.columns = [str(column).strip() for column in df.columns]
    if "省份" not in df.columns:
        raise ValueError("tourism_income.xlsx 中缺少“省份”列。")

    df["省份"] = df["省份"].astype(str).map(normalize_province_name)
    return df


def load_province_info(info_file: Path) -> dict[str, dict[str, str | None]]:
    """Load province info workbook for tooltip card details."""
    if not info_file.exists():
        raise FileNotFoundError(f"未找到信息文件：{info_file}")

    try:
        df = pd.read_excel(info_file)
    except ImportError as exc:
        raise SystemExit("缺少 openpyxl，请先运行：pip install openpyxl") from exc

    df.columns = [str(column).strip() for column in df.columns]
    if "省份" not in df.columns:
        raise ValueError("province_info.xlsx 中缺少“省份”列。")

    def pick_column(possible_names: list[str]) -> str:
        for name in possible_names:
            if name in df.columns:
                return name
        raise ValueError(f"province_info.xlsx 缺少列：{possible_names}")

    city_column = pick_column(["著名旅游城市", "强旅游城市"])
    scenic_column = pick_column(["著名旅游景点", "热门景点"])
    food_column = pick_column(["地方美食", "特色美食"])

    info_map: dict[str, dict[str, str | None]] = {}

    for _, row in df.iterrows():
        province = normalize_province_name(row["省份"])
        image_path = None
        asset_key = province_to_asset_key(province)
        if asset_key:
            image_file = ASSETS_DIR / f"{asset_key}.jpg"
            if image_file.exists():
                image_path = f"./assets/{image_file.name}"

        info_map[province] = {
            "cities": str(row[city_column]).strip(),
            "scenic": str(row[scenic_column]).strip(),
            "food": str(row[food_column]).strip(),
            "image": image_path,
        }

    return info_map


def extract_year_columns(df: pd.DataFrame) -> list[tuple[str, str]]:
    """Return [(year, original_column_name), ...] sorted by year."""
    year_columns: list[tuple[str, str]] = []

    for column in df.columns:
        if column == "省份":
            continue
        match = re.search(r"(20\d{2})", column)
        if match:
            year_columns.append((match.group(1), column))

    year_columns.sort(key=lambda item: item[0])

    if not year_columns:
        raise ValueError("未找到年份收入列，请检查 tourism_income.xlsx。")

    return year_columns


def build_year_payloads(df: pd.DataFrame) -> list[dict]:
    """Prepare the yearly payloads used by the page script."""
    payloads: list[dict] = []

    for year, column in extract_year_columns(df):
        year_df = df[["省份", column]].copy()
        year_df[column] = pd.to_numeric(year_df[column], errors="coerce")
        year_df = year_df.dropna(subset=["省份", column])

        values = year_df[column].tolist()
        payloads.append(
            {
                "year": year,
                "min": float(min(values)),
                "max": float(max(values)),
                "data": list(zip(year_df["省份"].tolist(), values)),
            }
        )

    return payloads


def get_top10_pairs(frame: dict) -> list[tuple[str, float]]:
    """Return top 10 provinces sorted by income in descending order."""
    sorted_data = sorted(frame["data"], key=lambda item: item[1], reverse=True)
    return sorted_data[:10]


def build_base_map(initial_payload: dict) -> Map:
    """Create the base pyecharts map style for the first year."""
    return (
        Map(
            init_opts=opts.InitOpts(
                width="820px",
                height="720px",
                page_title=PAGE_TITLE,
                bg_color="rgba(255, 252, 247, 0.96)",
                js_host="./assets_js/",
                animation_opts=opts.AnimationOpts(
                    animation_duration=900,
                    animation_duration_update=TRANSITION_DURATION_MS,
                    animation_easing="cubicOut",
                    animation_easing_update="cubicInOut",
                ),
            )
        )
        .add(
            series_name=f"{initial_payload['year']}年旅游收入",
            data_pair=initial_payload["data"],
            maptype="china",
            is_map_symbol_show=False,
        )
        .set_global_opts(
            title_opts=opts.TitleOpts(is_show=False),
            legend_opts=opts.LegendOpts(is_show=False),
            visualmap_opts=opts.VisualMapOpts(
                min_=initial_payload["min"],
                max_=initial_payload["max"],
                is_calculable=True,
                range_text=["高", "低"],
                range_color=WARM_COLORS,
                pos_left="2%",
                pos_bottom="12%",
                textstyle_opts=opts.TextStyleOpts(
                    color="#7c2d12",
                    font_size=14,
                    font_weight="bold",
                ),
            ),
            tooltip_opts=opts.TooltipOpts(
                formatter=JsCode(
                    "function (params) { return window.__tourismTooltipFormatter ? window.__tourismTooltipFormatter(params) : ''; }"
                ),
                background_color="rgba(255,255,255,0)",
                border_width=0,
                padding=0,
            ),
        )
        .set_series_opts(label_opts=opts.LabelOpts(is_show=False))
    )


def build_base_bar(initial_payload: dict) -> Bar:
    """Create the base pyecharts horizontal Top10 bar chart."""
    top10 = get_top10_pairs(initial_payload)
    province_names = [name for name, _ in top10]
    income_values = [value for _, value in top10]

    return (
        Bar(
            init_opts=opts.InitOpts(
                width="340px",
                height="285px",
                bg_color="rgba(255, 252, 247, 0)",
                js_host="./assets_js/",
                animation_opts=opts.AnimationOpts(
                    animation_duration=900,
                    animation_duration_update=TRANSITION_DURATION_MS,
                    animation_easing="cubicOut",
                    animation_easing_update="cubicInOut",
                ),
            )
        )
        .add_xaxis(province_names)
        .add_yaxis(
            series_name="旅游收入",
            y_axis=income_values,
            category_gap="36%",
            label_opts=opts.LabelOpts(
                is_show=True,
                position="right",
                formatter="{c} 亿元",
                color="#7c2d12",
                font_size=11,
                font_weight="bold",
            ),
            itemstyle_opts=opts.ItemStyleOpts(
                color=JsCode(
                    """new echarts.graphic.LinearGradient(0, 0, 1, 0, [
                        {offset: 0, color: '#fdba74'},
                        {offset: 0.55, color: '#fb923c'},
                        {offset: 1, color: '#dc2626'}
                    ])"""
                )
            ),
        )
        .reversal_axis()
        .set_global_opts(
            title_opts=opts.TitleOpts(is_show=False),
            legend_opts=opts.LegendOpts(is_show=False),
            xaxis_opts=opts.AxisOpts(
                type_="value",
                axislabel_opts=opts.LabelOpts(color="#9a3412", font_size=10),
                axisline_opts=opts.AxisLineOpts(
                    linestyle_opts=opts.LineStyleOpts(color="rgba(154, 52, 18, 0.22)")
                ),
                splitline_opts=opts.SplitLineOpts(
                    is_show=True,
                    linestyle_opts=opts.LineStyleOpts(color="rgba(251, 146, 60, 0.16)")
                ),
            ),
            yaxis_opts=opts.AxisOpts(
                type_="category",
                is_inverse=True,
                axislabel_opts=opts.LabelOpts(color="#7c2d12", font_size=11, font_weight="bold"),
                axisline_opts=opts.AxisLineOpts(is_show=False),
                axistick_opts=opts.AxisTickOpts(is_show=False),
            ),
            tooltip_opts=opts.TooltipOpts(
                trigger="axis",
                axis_pointer_type="shadow",
                formatter="{b}<br/>旅游收入：{c} 亿元",
                background_color="rgba(255, 252, 247, 0.96)",
                border_color="rgba(251, 146, 60, 0.35)",
                textstyle_opts=opts.TextStyleOpts(color="#7c2d12"),
            ),
        )
    )


def build_html(
    payloads: list[dict],
    province_info: dict[str, dict[str, str | None]],
    base_map_option: str,
    base_bar_option: str,
) -> str:
    """Render the complete fourth-version page with map, info card, and Top10 bar."""
    years = [payload["year"] for payload in payloads]
    payloads_json = json.dumps(payloads, ensure_ascii=False)
    province_info_json = json.dumps(province_info, ensure_ascii=False)
    year_buttons = "".join(
        f'<button class="year-item{" active" if idx == 0 else ""}" data-index="{idx}">{year}</button>'
        for idx, year in enumerate(years)
    )

    return f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>{PAGE_TITLE}</title>
    <link rel="icon" href="data:,">
    <script src="{ECHARTS_JS}"></script>
    <script src="{CHINA_MAP_JS}"></script>
    <style>
        html, body {{
            margin: 0;
            min-height: 100%;
            background:
                radial-gradient(circle at top left, rgba(251, 191, 36, 0.22), transparent 28%),
                radial-gradient(circle at top right, rgba(239, 68, 68, 0.16), transparent 22%),
                linear-gradient(180deg, #fff8ef 0%, #fff3df 100%);
            font-family: "Microsoft YaHei", "PingFang SC", sans-serif;
        }}
        body {{
            padding: 14px 0 18px;
        }}
        .page-shell {{
            width: 1260px;
            margin: 0 auto;
            position: relative;
            border-radius: 28px;
            background: rgba(255, 252, 247, 0.96);
            box-shadow: 0 26px 70px rgba(194, 65, 12, 0.12);
            border: 1px solid rgba(251, 146, 60, 0.16);
            overflow: hidden;
        }}
        .page-header {{
            padding: 26px 34px 0;
            text-align: left;
        }}
        .page-title {{
            margin: 0;
            color: #7c2d12;
            font-size: 26px;
            line-height: 1.2;
            font-weight: 800;
            letter-spacing: 0.3px;
            white-space: nowrap;
            overflow: visible;
        }}
        .content-row {{
            display: grid;
            grid-template-columns: 840px 340px;
            column-gap: 16px;
            align-items: start;
            padding: 6px 18px 18px 0;
        }}
        .map-stage {{
            position: relative;
            width: 840px;
        }}
        .chart-container {{
            width: 820px;
            height: 720px;
        }}
        .map-hover-tooltip {{
            min-width: 132px;
            padding: 10px 12px;
            border-radius: 14px;
            background: rgba(255, 252, 247, 0.96);
            border: 1px solid rgba(251, 146, 60, 0.3);
            box-shadow: 0 14px 32px rgba(124, 45, 18, 0.18);
            color: #7c2d12;
            font-family: "Microsoft YaHei", "PingFang SC", sans-serif;
            line-height: 1.45;
        }}
        .map-hover-name {{
            font-size: 15px;
            font-weight: 800;
            margin-bottom: 4px;
        }}
        .map-hover-value {{
            color: #c2410c;
            font-size: 13px;
            font-weight: 700;
        }}
        .year-label {{
            position: absolute;
            top: 8px;
            left: 410px;
            transform: translateX(-50%);
            color: #c2410c;
            font-size: 18px;
            font-weight: 700;
            letter-spacing: 0.5px;
            z-index: 20;
            pointer-events: none;
        }}
        .side-nav {{
            position: absolute;
            top: 50%;
            transform: translateY(-50%);
            width: 58px;
            height: 58px;
            border: 1px solid rgba(154, 52, 18, 0.22);
            border-radius: 999px;
            background: rgba(255, 247, 237, 0.98);
            color: #9a3412;
            font-size: 32px;
            font-weight: bold;
            line-height: 54px;
            text-align: center;
            cursor: pointer;
            box-shadow: 0 14px 36px rgba(154, 52, 18, 0.16);
            user-select: none;
            transition: transform 0.18s ease, box-shadow 0.18s ease, background 0.18s ease;
            backdrop-filter: blur(10px);
            z-index: 20;
        }}
        .side-nav:hover {{
            transform: translateY(-50%) scale(1.03);
            background: #fff1dd;
            box-shadow: 0 18px 40px rgba(154, 52, 18, 0.2);
        }}
        .side-nav:disabled {{
            opacity: 0.45;
            cursor: not-allowed;
            transform: translateY(-50%);
            box-shadow: 0 10px 24px rgba(154, 52, 18, 0.08);
        }}
        .side-nav.prev {{
            left: 10px;
        }}
        .side-nav.next {{
            right: 10px;
        }}
        .bottom-toolbar {{
            position: absolute;
            left: 410px;
            bottom: 18px;
            transform: translateX(-50%);
            width: 600px;
            z-index: 20;
        }}
        .year-track {{
            width: 100%;
            display: flex;
            justify-content: space-between;
            align-items: center;
            position: relative;
        }}
        .year-track::before {{
            content: "";
            position: absolute;
            left: 42px;
            right: 42px;
            top: 18px;
            height: 3px;
            background: #fdba74;
            border-radius: 999px;
        }}
        .year-item {{
            position: relative;
            z-index: 1;
            border: none;
            background: transparent;
            color: #9a3412;
            font-size: 15px;
            font-weight: 700;
            cursor: pointer;
            padding: 24px 8px 0;
            min-width: 96px;
            transition: transform 0.18s ease, color 0.18s ease;
        }}
        .year-item::before {{
            content: "";
            position: absolute;
            left: 50%;
            top: 9px;
            transform: translateX(-50%);
            width: 14px;
            height: 14px;
            border-radius: 999px;
            background: #fb923c;
            border: 2px solid #ea580c;
            box-sizing: border-box;
            transition: transform 0.22s ease, background 0.22s ease, box-shadow 0.22s ease;
        }}
        .year-item:hover {{
            color: #7c2d12;
        }}
        .year-item.active {{
            color: #7c2d12;
            font-size: 16px;
        }}
        .year-item.active::before {{
            background: #dc2626;
            border-color: #b91c1c;
            box-shadow: 0 0 0 6px rgba(220, 38, 38, 0.18);
            transform: translateX(-50%) scale(1.18);
        }}
        .info-panel-wrap {{
            padding-top: 8px;
            display: grid;
            grid-template-rows: auto 1fr;
            gap: 14px;
        }}
        .info-panel {{
            width: 100%;
            border-radius: 20px;
            overflow: hidden;
            background: rgba(255, 252, 247, 0.97);
            box-shadow: 0 18px 44px rgba(124, 45, 18, 0.16);
            border: 1px solid rgba(251, 146, 60, 0.18);
            backdrop-filter: blur(8px);
            transition: transform 0.22s ease, box-shadow 0.22s ease, opacity 0.22s ease;
        }}
        .info-panel.image-missing .info-panel-image-placeholder {{
            display: flex;
        }}
        .info-panel.image-missing .info-panel-image {{
            display: none;
        }}
        .info-panel-image {{
            width: 100%;
            height: 155px;
            object-fit: cover;
            display: block;
            background: linear-gradient(135deg, #fde7c8 0%, #ffd6a5 100%);
        }}
        .info-panel-image-placeholder {{
            display: none;
            width: 100%;
            height: 155px;
            align-items: center;
            justify-content: center;
            background: linear-gradient(135deg, #fde7c8 0%, #ffd6a5 100%);
            color: #9a3412;
            font-size: 14px;
            font-weight: 700;
        }}
        .info-panel-body {{
            padding: 14px 14px 12px;
        }}
        .info-panel-head {{
            display: flex;
            align-items: baseline;
            justify-content: space-between;
            gap: 10px;
            margin-bottom: 10px;
        }}
        .info-panel-title {{
            color: #7c2d12;
            font-size: 18px;
            font-weight: 800;
            line-height: 1.2;
        }}
        .info-panel-year {{
            color: #c2410c;
            font-size: 12px;
            font-weight: 700;
            white-space: nowrap;
        }}
        .info-panel-income {{
            margin-bottom: 8px;
            padding: 8px 10px;
            border-radius: 12px;
            background: linear-gradient(135deg, rgba(254, 215, 170, 0.45), rgba(254, 242, 242, 0.82));
            color: #9a3412;
            font-size: 14px;
            font-weight: 700;
        }}
        .info-panel-grid {{
            display: grid;
            gap: 8px;
        }}
        .info-panel-row {{
            display: grid;
            gap: 3px;
            padding: 8px 10px;
            border-radius: 12px;
            background: rgba(255, 247, 237, 0.9);
            border: 1px solid rgba(251, 146, 60, 0.1);
        }}
        .info-panel-label {{
            color: #c2410c;
            font-size: 11px;
            font-weight: 800;
            letter-spacing: 0.4px;
            text-transform: uppercase;
        }}
        .info-panel-value {{
            color: #7c2d12;
            font-size: 13px;
            line-height: 1.45;
        }}
        .rank-panel {{
            width: 100%;
            border-radius: 20px;
            background: rgba(255, 252, 247, 0.97);
            box-shadow: 0 18px 44px rgba(124, 45, 18, 0.13);
            border: 1px solid rgba(251, 146, 60, 0.18);
            backdrop-filter: blur(8px);
            overflow: hidden;
        }}
        .rank-panel-head {{
            padding: 14px 16px 4px;
            display: flex;
            align-items: baseline;
            justify-content: space-between;
            gap: 10px;
        }}
        .rank-panel-title {{
            color: #7c2d12;
            font-size: 16px;
            font-weight: 800;
            line-height: 1.2;
        }}
        .rank-panel-year {{
            color: #c2410c;
            font-size: 12px;
            font-weight: 700;
            white-space: nowrap;
        }}
        .rank-chart {{
            width: 340px;
            height: 285px;
        }}
    </style>
</head>
    <body>
    <div class="page-shell">
        <div class="page-header">
            <h1 class="page-title">{PAGE_TITLE}</h1>
        </div>
        <div class="content-row">
            <div class="map-stage">
                <div class="year-label" id="year-label">2023年</div>
                <button class="side-nav prev" id="prev-btn">&#8249;</button>
                <button class="side-nav next" id="next-btn">&#8250;</button>
                <div id="chart-container" class="chart-container"></div>
                <div class="bottom-toolbar">
                    <div class="year-track" id="year-track">
                        {year_buttons}
                    </div>
                </div>
            </div>
            <div class="info-panel-wrap">
                <div class="info-panel image-missing" id="info-panel">
                    <img class="info-panel-image" id="info-panel-image" src="" alt="省份图片" />
                    <div class="info-panel-image-placeholder" id="info-panel-placeholder">悬停省份查看详细旅游信息</div>
                    <div class="info-panel-body">
                        <div class="info-panel-head">
                            <div class="info-panel-title" id="info-panel-title">旅游信息卡片</div>
                            <div class="info-panel-year" id="info-panel-year">2023年</div>
                        </div>
                        <div class="info-panel-income" id="info-panel-income">请将鼠标移动到地图省份上</div>
                        <div class="info-panel-grid">
                            <div class="info-panel-row">
                                <div class="info-panel-label">强旅游城市</div>
                                <div class="info-panel-value" id="info-panel-cities">这里会显示对应省份的热门城市</div>
                            </div>
                            <div class="info-panel-row">
                                <div class="info-panel-label">热门景点</div>
                                <div class="info-panel-value" id="info-panel-scenic">这里会显示对应省份的代表景点</div>
                            </div>
                            <div class="info-panel-row">
                                <div class="info-panel-label">特色美食</div>
                                <div class="info-panel-value" id="info-panel-food">这里会显示对应省份的特色美食</div>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="rank-panel">
                    <div class="rank-panel-head">
                        <div class="rank-panel-title">当前年份旅游收入 Top10</div>
                        <div class="rank-panel-year" id="rank-panel-year">2023年</div>
                    </div>
                    <div id="rank-chart" class="rank-chart"></div>
                </div>
            </div>
        </div>
    </div>
    <script>
        const frames = {payloads_json};
        const provinceInfo = {province_info_json};
        const mapChart = echarts.init(document.getElementById("chart-container"), "white", {{ renderer: "canvas", locale: "ZH" }});
        const barChart = echarts.init(document.getElementById("rank-chart"), "white", {{ renderer: "canvas", locale: "ZH" }});
        const baseMapOption = {base_map_option};
        const baseBarOption = {base_bar_option};
        const yearLabel = document.getElementById("year-label");
        const rankPanelYear = document.getElementById("rank-panel-year");
        const prevBtn = document.getElementById("prev-btn");
        const nextBtn = document.getElementById("next-btn");
        const yearItems = Array.from(document.querySelectorAll(".year-item"));
        const infoPanel = document.getElementById("info-panel");
        const panelTitle = document.getElementById("info-panel-title");
        const panelYear = document.getElementById("info-panel-year");
        const panelIncome = document.getElementById("info-panel-income");
        const panelCities = document.getElementById("info-panel-cities");
        const panelScenic = document.getElementById("info-panel-scenic");
        const panelFood = document.getElementById("info-panel-food");
        const panelImage = document.getElementById("info-panel-image");
        const panelPlaceholder = document.getElementById("info-panel-placeholder");
        let currentIndex = 0;
        let animationToken = 0;

        function escapeHtml(value) {{
            return String(value ?? "")
                .replace(/&/g, "&amp;")
                .replace(/</g, "&lt;")
                .replace(/>/g, "&gt;")
                .replace(/"/g, "&quot;")
                .replace(/'/g, "&#39;");
        }}

        function buildSeriesData(frame) {{
            return frame.data.map(([name, value]) => ({{ name, value }}));
        }}

        function buildTop10(frame) {{
            return frame.data
                .slice()
                .sort((left, right) => right[1] - left[1])
                .slice(0, 10);
        }}

        function syncControls() {{
            yearLabel.textContent = frames[currentIndex].year + "年";
            rankPanelYear.textContent = frames[currentIndex].year + "年";
            yearItems.forEach((item, index) => item.classList.toggle("active", index === currentIndex));
            prevBtn.disabled = currentIndex === 0;
            nextBtn.disabled = currentIndex === frames.length - 1;
        }}

        function setMapFrame(frame) {{
            mapChart.setOption({{
                visualMap: {{
                    min: frame.min,
                    max: frame.max
                }},
                series: [{{
                    name: frame.year + "年旅游收入",
                    data: buildSeriesData(frame)
                }}]
            }});
        }}

        function setBarFrame(frame) {{
            const top10 = buildTop10(frame);
            barChart.setOption({{
                yAxis: {{
                    data: top10.map(([name]) => name),
                    inverse: true
                }},
                series: [{{
                    name: frame.year + "年旅游收入 Top10",
                    data: top10.map(([, value]) => Number(value.toFixed(2)))
                }}]
            }});
        }}

        function setFrame(frame) {{
            setMapFrame(frame);
            setBarFrame(frame);
        }}

        function updateInfoPanel(provinceName, value) {{
            const frame = frames[currentIndex];
            const info = provinceInfo[provinceName] || {{}};
            const displayValue = value === undefined || value === null || value === "-" ? "暂无数据" : value + " 亿元";

            panelTitle.textContent = provinceName || "旅游信息卡片";
            panelYear.textContent = frame.year + "年";
            panelIncome.textContent = "旅游收入： " + displayValue;
            panelCities.textContent = info.cities || "暂无信息";
            panelScenic.textContent = info.scenic || "暂无信息";
            panelFood.textContent = info.food || "暂无信息";

            if (info.image) {{
                panelImage.src = info.image;
                panelImage.alt = provinceName;
                infoPanel.classList.remove("image-missing");
            }} else {{
                panelImage.removeAttribute("src");
                panelImage.alt = "暂无省份图片";
                infoPanel.classList.add("image-missing");
            }}

            infoPanel.style.transform = "translateY(-2px)";
            setTimeout(() => {{
                infoPanel.style.transform = "translateY(0)";
            }}, 120);
        }}

        function animateToIndex(nextIndex) {{
            if (nextIndex < 0 || nextIndex >= frames.length || nextIndex === currentIndex) {{
                return;
            }}

            const fromFrame = frames[currentIndex];
            const toFrame = frames[nextIndex];
            const toMap = new Map(toFrame.data);
            const duration = {TRANSITION_DURATION_MS};
            const token = ++animationToken;
            const start = performance.now();

            currentIndex = nextIndex;
            syncControls();

            const step = (now) => {{
                if (token !== animationToken) {{
                    return;
                }}

                const t = Math.min((now - start) / duration, 1);
                const eased = t < 0.5
                    ? 4 * t * t * t
                    : 1 - Math.pow(-2 * t + 2, 3) / 2;

                const interpolatedFrame = {{
                    year: toFrame.year,
                    min: fromFrame.min + (toFrame.min - fromFrame.min) * eased,
                    max: fromFrame.max + (toFrame.max - fromFrame.max) * eased,
                    data: fromFrame.data.map(([name, startValue]) => {{
                        const endValue = toMap.get(name);
                        const value = startValue + (endValue - startValue) * eased;
                        return [name, Number(value.toFixed(2))];
                    }})
                }};

                setFrame(interpolatedFrame);

                if (t < 1) {{
                    requestAnimationFrame(step);
                }} else {{
                    setFrame(toFrame);
                }}
            }};

            requestAnimationFrame(step);
        }}

        window.__tourismTooltipFormatter = function (params) {{
            updateInfoPanel(params.name, params.value);
            const value = params.value === undefined || params.value === null || params.value === "-"
                ? "暂无数据"
                : params.value + " 亿元";
            return `
                <div class="map-hover-tooltip">
                    <div class="map-hover-name">${{escapeHtml(params.name)}}</div>
                    <div class="map-hover-value">旅游收入：${{escapeHtml(value)}}</div>
                </div>
            `;
        }};

        baseMapOption.series[0].name = frames[0].year + "年旅游收入";
        baseMapOption.series[0].data = buildSeriesData(frames[0]);
        baseMapOption.series[0].label = {{
            show: false,
            color: "#7c2d12",
            fontFamily: "Microsoft YaHei, PingFang SC, sans-serif",
            fontSize: 12,
            fontWeight: "bold"
        }};
        baseMapOption.series[0].emphasis = {{
            label: {{
                show: true,
                color: "#7c2d12",
                fontFamily: "Microsoft YaHei, PingFang SC, sans-serif",
                fontSize: 15,
                fontWeight: "bold",
                textBorderColor: "rgba(255, 252, 247, 0.95)",
                textBorderWidth: 4,
                textShadowColor: "rgba(124, 45, 18, 0.18)",
                textShadowBlur: 8
            }},
            itemStyle: {{
                areaColor: "#fb923c",
                borderColor: "#9a3412",
                borderWidth: 1.2,
                shadowColor: "rgba(124, 45, 18, 0.28)",
                shadowBlur: 14
            }}
        }};
        baseMapOption.visualMap.min = frames[0].min;
        baseMapOption.visualMap.max = frames[0].max;
        mapChart.setOption(baseMapOption);

        baseBarOption.yAxis[0].inverse = true;
        baseBarOption.grid = {{
            left: "30%",
            right: "18%",
            top: "8%",
            bottom: "8%"
        }};
        barChart.setOption(baseBarOption);
        setBarFrame(frames[0]);
        syncControls();

        if (frames[0] && frames[0].data.length) {{
            const [provinceName, value] = frames[0].data[0];
            updateInfoPanel(provinceName, value);
        }}

        prevBtn.addEventListener("click", () => animateToIndex(currentIndex - 1));
        nextBtn.addEventListener("click", () => animateToIndex(currentIndex + 1));
        yearItems.forEach((item, index) => {{
            item.addEventListener("click", () => animateToIndex(index));
        }});
        window.addEventListener("resize", () => {{
            mapChart.resize();
            barChart.resize();
        }});
    </script>
</body>
</html>
"""


def main() -> None:
    tourism_df = load_tourism_dataframe(EXCEL_FILE)
    province_info = load_province_info(INFO_FILE)
    payloads = build_year_payloads(tourism_df)
    base_map = build_base_map(payloads[0])
    base_bar = build_base_bar(payloads[0])
    base_map_option = base_map.dump_options()
    base_bar_option = base_bar.dump_options()
    html = build_html(payloads, province_info, base_map_option, base_bar_option)
    OUTPUT_HTML.write_text(html, encoding="utf-8")
    LEGACY_OUTPUT_HTML.write_text(html, encoding="utf-8")
    print(f"地图已生成：{OUTPUT_HTML}")


if __name__ == "__main__":
    main()
