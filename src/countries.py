"""Country normalization shared by all sources."""
import re
import unicodedata

_COUNTRY_ALIASES = {
    "china": "China",
    "prc": "China",
    "people's republic of china": "China",
    "people s republic of china": "China",
    "usa": "USA",
    "u.s.a.": "USA",
    "united states": "USA",
    "united states of america": "USA",
    "uk": "UK",
    "u.k.": "UK",
    "united kingdom": "UK",
    "britain": "UK",
    "great britain": "UK",
    "germany": "Germany",
    "france": "France",
    "japan": "Japan",
    "south korea": "South Korea",
    "korea": "South Korea",
    "republic of korea": "South Korea",
    "india": "India",
    "italy": "Italy",
    "spain": "Spain",
    "switzerland": "Switzerland",
    "sweden": "Sweden",
    "netherlands": "Netherlands",
    "the netherlands": "Netherlands",
    "australia": "Australia",
    "canada": "Canada",
    "brazil": "Brazil",
    "russia": "Russia",
    "russian federation": "Russia",
    "singapore": "Singapore",
    "israel": "Israel",
    "iran": "Iran",
    "saudi arabia": "Saudi Arabia",
    "uae": "UAE",
    "united arab emirates": "UAE",
    "argentina": "Argentina",
    "chile": "Chile",
    "mexico": "Mexico",
    "poland": "Poland",
    "austria": "Austria",
    "belgium": "Belgium",
    "denmark": "Denmark",
    "finland": "Finland",
    "norway": "Norway",
    "ireland": "Ireland",
    "portugal": "Portugal",
    "greece": "Greece",
    "turkey": "Turkey",
    "türkiye": "Turkey",
    "czech republic": "Czech Republic",
    "czechia": "Czech Republic",
    "hungary": "Hungary",
    "south africa": "South Africa",
    "egypt": "Egypt",
    "thailand": "Thailand",
    "malaysia": "Malaysia",
    "indonesia": "Indonesia",
    "vietnam": "Vietnam",
    "taiwan": "Taiwan",
    "hong kong": "Hong Kong",
    "p.r. china": "China",
    "new zealand": "New Zealand",
}


def _country_key(value):
    value = unicodedata.normalize('NFKD', value.replace('’', "'").lower())
    return re.sub(r"[^a-z ']", '', value).strip()


_COUNTRY_ALIASES = {_country_key(key): value for key, value in _COUNTRY_ALIASES.items()}


def parse_country(affiliation: str) -> str:
    """从 Affiliation 文本中提取国家。

    规则：
    1. 取逗号分隔的最后一个非空 segment（多数单位末段就是国家）
    2. 清洗后与 _COUNTRY_ALIASES 匹配
    3. 匹配不到则返回原文末段加 "?" 前缀（标记未识别，方便人工 review）
    """
    if not affiliation:
        return ""
    # 去掉邮编、数字
    s = re.sub(r"\b\d{4,6}\b", "", affiliation).strip(" ,;")
    parts = [p.strip() for p in s.split(",") if p.strip()]
    if not parts:
        return ""
    # 从末尾向前找：连续两个 segment 都可能是国家（如 "TX, USA" → USA）
    for cand in reversed(parts[-2:]):
        key = _country_key(cand)
        if key in _COUNTRY_ALIASES:
            return _COUNTRY_ALIASES[key]
        # 也尝试整段去空格的常见变体
        if "china" in key and "taiwan" not in key and "hong" not in key and "macau" not in key:
            return "China"
    # 兜底：返回最后一段原文，加 "?" 前缀让 Excel 一眼看出未识别
    return f"?{parts[-1]}"


def is_china_country(country: str) -> bool | None:
    return None if not country or country.startswith("?") else country == "China"
