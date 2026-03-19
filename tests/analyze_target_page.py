import re
from pathlib import Path
from urllib.request import Request, urlopen

TARGET_URL = "https://std.samr.gov.cn/gb/search/gbDetailed?id=234D7936AB54E194E06397BE0A0AA0A9&review=true"
OUT_HTML = Path("target_234D_review_true.html")
OUT_REPORT = Path("analysis_234D_report.txt")

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0 Safari/537.36"
    ),
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
    "Referer": "https://openstd.samr.gov.cn/",
}


def strip_tags(text: str) -> str:
    text = re.sub(r"<script[\\s\\S]*?</script>", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"<style[\\s\\S]*?</style>", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"<[^>]+>", " ", text)
    text = re.sub(r"\\s+", " ", text)
    return text.strip()


def extract_title(html: str) -> tuple[str, str]:
    h4 = re.search(r"<h4[^>]*>([\\s\\S]*?)</h4>", html, flags=re.IGNORECASE)
    h5 = re.search(r"<h5[^>]*>([\\s\\S]*?)</h5>", html, flags=re.IGNORECASE)
    zh = strip_tags(h4.group(1)) if h4 else ""
    en = strip_tags(h5.group(1)) if h5 else ""
    return zh, en


def extract_section_titles(html: str) -> list[str]:
    titles = re.findall(r"<h2[^>]*class=\"title-text\"[^>]*>([\\s\\S]*?)</h2>", html, flags=re.IGNORECASE)
    return [strip_tags(t) for t in titles]


def extract_dt_dd_pairs(html: str) -> list[tuple[str, str]]:
    pairs = []
    for m in re.finditer(r"<dt[^>]*class=\"basicInfo-item name\"[^>]*>([\\s\\S]*?)</dt>\\s*<dd[^>]*class=\"basicInfo-item value\"[^>]*>([\\s\\S]*?)</dd>", html, flags=re.IGNORECASE):
        dt = strip_tags(m.group(1))
        dd = strip_tags(m.group(2))
        pairs.append((dt, dd))
    return pairs


def extract_caibiao_candidates(html: str) -> list[str]:
    text = strip_tags(html)
    patterns = [
        r"采标情况[:：]?[\\s\\S]{0,220}[。；]",
        r"本(?:标准|文件)[\\s\\S]{0,160}(?:等同采用|修改采用|非等效采用|采用)[\\s\\S]{0,220}[。；]",
        r"采用(?:IEC|ISO|EN|ASTM|IEEE|ITU)[\\s\\S]{0,220}[。；]",
        r"采用国际标准[\\s\\S]{0,220}[。；]",
    ]
    out = []
    for pat in patterns:
        for mm in re.finditer(pat, text, flags=re.IGNORECASE):
            s = mm.group(0).strip()
            if s not in out:
                out.append(s)
    return out


def main() -> None:
    req = Request(TARGET_URL, headers=HEADERS)
    with urlopen(req, timeout=30) as resp:
        raw = resp.read().decode("utf-8", errors="replace")
        final_url = resp.geturl()

    OUT_HTML.write_text(raw, encoding="utf-8")

    zh_title, en_title = extract_title(raw)
    sections = extract_section_titles(raw)
    pairs = extract_dt_dd_pairs(raw)
    caibiao_hits = extract_caibiao_candidates(raw)

    marks = {
        "contains_采标情况": ("采标情况" in raw),
        "contains_IEC": ("IEC" in raw),
        "contains_等同采用": ("等同采用" in raw),
        "contains_login_word": (("登录" in raw) or ("login" in raw.lower())),
        "contains_review_autoclick_js": ("indexOf(\"review=true\")" in raw),
    }

    report_lines = []
    report_lines.append(f"request_url: {TARGET_URL}")
    report_lines.append(f"final_url: {final_url}")
    report_lines.append(f"html_length: {len(raw)}")
    report_lines.append(f"title_zh: {zh_title}")
    report_lines.append(f"title_en: {en_title}")
    report_lines.append("marks:")
    for k, v in marks.items():
        report_lines.append(f"  - {k}: {v}")

    report_lines.append("section_titles:")
    for t in sections:
        report_lines.append(f"  - {t}")

    report_lines.append("basic_info_pairs(sample up to 20):")
    for dt, dd in pairs[:20]:
        report_lines.append(f"  - {dt}: {dd}")

    report_lines.append("caibiao_candidates:")
    if caibiao_hits:
        for c in caibiao_hits:
            report_lines.append(f"  - {c}")
    else:
        report_lines.append("  - (none)")

    OUT_REPORT.write_text("\n".join(report_lines), encoding="utf-8")
    print(f"saved_html={OUT_HTML}")
    print(f"saved_report={OUT_REPORT}")


if __name__ == "__main__":
    main()
