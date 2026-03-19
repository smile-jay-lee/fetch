import argparse
import hashlib
import html
import re
import sys
from datetime import datetime
from html.parser import HTMLParser
from pathlib import Path
from urllib.error import HTTPError
from urllib.parse import parse_qsl, urlencode, urljoin, urlparse, urlunparse
from urllib.request import Request, urlopen


class FkButtonParser(HTMLParser):
	def __init__(self) -> None:
		super().__init__()
		self.data_values: list[str] = []

	def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
		if tag.lower() != "button":
			return

		attrs_dict = {k: (v or "") for k, v in attrs}
		class_name = attrs_dict.get("class", "")
		if "fk_btn" not in class_name.split():
			return

		value = attrs_dict.get("data-value", "").strip()
		if value:
			self.data_values.append(value)


class TextExtractor(HTMLParser):
	def __init__(self) -> None:
		super().__init__()
		self.parts: list[str] = []
		self._skip_level = 0

	def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
		if tag.lower() in {"script", "style"}:
			self._skip_level += 1

	def handle_endtag(self, tag: str) -> None:
		if tag.lower() in {"script", "style"} and self._skip_level > 0:
			self._skip_level -= 1

	def handle_data(self, data: str) -> None:
		if self._skip_level == 0 and data:
			self.parts.append(data)


def find_excel_file(base_dir: Path, explicit_path: str | None) -> Path:
	if explicit_path:
		p = Path(explicit_path)
		if not p.is_absolute():
			p = base_dir / p
		if not p.exists():
			raise FileNotFoundError(f"Excel file not found: {p}")
		return p

	candidates = [
		p
		for p in base_dir.glob("*.xlsx")
		if p.is_file() and not p.name.startswith("~$")
	]
	if not candidates:
		raise FileNotFoundError("No .xlsx file found in current directory")

	candidates.sort()
	if len(candidates) > 1:
		names = ", ".join(p.name for p in candidates)
		raise RuntimeError(
			"Multiple .xlsx files found. Please specify one with --excel. "
			f"Candidates: {names}"
		)

	return candidates[0]


def save_workbook_with_fallback(wb: object, desired_path: Path) -> Path:
	try:
		wb.save(desired_path)
		return desired_path
	except PermissionError:
		ts = datetime.now().strftime("%Y%m%d_%H%M%S")
		fallback = desired_path.with_name(f"{desired_path.stem}.autosave_{ts}{desired_path.suffix}")
		wb.save(fallback)
		print(
			f"[WARN] Cannot write '{desired_path}' (likely open in Excel). "
			f"Saved to '{fallback}' instead."
		)
		return fallback


def normalize_url(value: object) -> str | None:
	if value is None:
		return None
	text = str(value).strip()
	if not text:
		return None

	# Ignore common header-like placeholders.
	if text.lower() in {"url", "article_url", "link"}:
		return None

	if not text.lower().startswith(("http://", "https://")):
		text = "http://" + text

	parsed = urlparse(text)
	if not parsed.netloc:
		return None

	host = parsed.netloc.split("@")[-1].split(":")[0]
	if "." not in host and host.lower() != "localhost":
		return None

	return text


def build_filename(url: str, row: int) -> str:
	parsed = urlparse(url)
	raw = f"{parsed.netloc}{parsed.path}".strip("/")
	slug = re.sub(r"[^A-Za-z0-9._-]+", "_", raw).strip("_")
	if not slug:
		slug = "page"
	slug = slug[:80]
	digest = hashlib.md5(url.encode("utf-8"), usedforsecurity=False).hexdigest()[:8]
	return f"row{row:04d}_{slug}_{digest}.html"


def build_request_headers(url: str) -> dict[str, str]:
	headers = {
		"User-Agent": (
			"Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
			"AppleWebKit/537.36 (KHTML, like Gecko) "
			"Chrome/124.0 Safari/537.36"
		),
		"Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
	}

	host = urlparse(url).netloc.lower()
	if "std.samr.gov.cn" in host:
		headers["Referer"] = "https://openstd.samr.gov.cn/"

	return headers


def remove_review_true(url: str) -> str:
	parsed = urlparse(url)
	query_pairs = parse_qsl(parsed.query, keep_blank_values=True)
	filtered = [
		(k, v)
		for k, v in query_pairs
		if not (k.lower() == "review" and v.lower() == "true")
	]
	if len(filtered) == len(query_pairs):
		return url
	return urlunparse(parsed._replace(query=urlencode(filtered)))


def fetch_url_content(url: str, timeout: int) -> str:
	req = Request(url, headers=build_request_headers(url))

	with urlopen(req, timeout=timeout) as resp:
		data = resp.read()
		content_type = resp.headers.get_content_charset()
		encoding = content_type or "utf-8"
		return data.decode(encoding, errors="replace")


def fetch_url_content_resilient(url: str, timeout: int) -> tuple[str, str]:
	try:
		return fetch_url_content(url, timeout), url
	except HTTPError as exc:
		fallback_url = remove_review_true(url)
		if exc.code >= 500 and fallback_url != url:
			return fetch_url_content(fallback_url, timeout), fallback_url
		raise


def html_to_text(raw_html: str) -> str:
	parser = TextExtractor()
	parser.feed(raw_html)
	text = " ".join(parser.parts)
	text = html.unescape(text)
	text = text.replace("\u3000", " ")
	text = re.sub(r"\s+", " ", text)
	return text.strip()


def extract_section_text_by_title(raw_html: str, section_title: str) -> str:
	# Match the section block from the title to the next para-title/footer/body end.
	pattern = re.compile(
		rf'<h2[^>]*class="title-text"[^>]*>\s*{re.escape(section_title)}\s*</h2>([\s\S]*?)(?=<div\s+class="para-title"|<div\s+class="footer-body"|</body>)',
		flags=re.IGNORECASE,
	)
	match = pattern.search(raw_html)
	if not match:
		return ""

	section_html = match.group(1)
	section_text = html_to_text(section_html)
	# Remove repeated section title if present in text extraction.
	section_text = section_text.replace(section_title, "").strip()
	return section_text


def extract_caibiao_from_structured_section(raw_html: str) -> str:
	section_text = extract_section_text_by_title(raw_html, "采标情况")
	if not section_text:
		return ""

	# Prefer the most informative sentence(s) in the section.
	hits = []
	for pat in [
		r"本(?:标准|文件)[^。；\n]{0,200}(?:等同采用|修改采用|非等效采用|采用)[^。；\n]{0,260}[。；]?",
		r"采用(?:IEC|ISO|EN|ASTM|IEEE|ITU)[^。；\n]{0,260}[。；]?",
		r"采标中文名称[:：]?[^。；\n]{1,220}[。；]?",
		r"采标程度[\s:：]?[^。；\n]{1,220}[。；]?",
	]:
		for m in re.finditer(pat, section_text, flags=re.IGNORECASE):
			v = m.group(0).strip()
			if v and v not in hits:
				hits.append(v)

	if hits:
		return " ".join(hits)

	return section_text[:300]


def extract_caibiao_info(raw_html: str) -> tuple[bool, str]:
	section_hit = extract_caibiao_from_structured_section(raw_html)
	if section_hit:
		return True, section_hit

	text = html_to_text(raw_html)
	has_module = "采标情况" in text

	patterns = [
		r"(本(?:标准|文件)[^。；\n]{0,160}(?:等同采用|修改采用|非等效采用|采用)[^。；\n]{0,240}[。；]?)",
		r"(采用(?:IEC|ISO|EN|ASTM|IEEE|ITU)[^。；\n]{0,240}[。；]?)",
		r"(采标情况[:：]?[^。；\n]{1,260}[。；]?)",
		r"(采用国际标准[^。；\n]{0,240}[。；]?)",
	]

	for pattern in patterns:
		match = re.search(pattern, text, flags=re.IGNORECASE)
		if match:
			return has_module or ("采标" in match.group(1)), match.group(1).strip()

	if has_module:
		idx = text.find("采标情况")
		snippet = text[idx: idx + 180]
		snippet = re.split(r"(起草单位|起草人|归口单位|发布日期|实施日期|ICS|中国标准分类号)", snippet)[0]
		return True, snippet.strip()

	return False, ""


def resolve_caibiao_column_index(ws: object, column_spec: str) -> int:
	if column_spec.upper() == "AUTO":
		return ws.max_column + 1

	from openpyxl.utils import column_index_from_string

	return column_index_from_string(column_spec.upper())


def n_column_hint_has_caibiao(value: object) -> bool:
	if value is None:
		return False
	text = str(value).strip()
	if not text:
		return False
	return "采" in text


def extract_feedback_ids(html: str) -> list[str]:
	parser = FkButtonParser()
	parser.feed(html)

	# Keep order while de-duplicating.
	ordered_unique: list[str] = []
	seen: set[str] = set()
	for value in parser.data_values:
		if value not in seen:
			ordered_unique.append(value)
			seen.add(value)
	return ordered_unique


def build_review_url(base_url: str, hcno: str) -> str:
	parsed = urlparse(base_url)
	origin = f"{parsed.scheme}://{parsed.netloc}"
	return urljoin(origin, f"/bzgk/gb/review?hcno={hcno}")


def build_review_filename(row: int, hcno: str) -> str:
	safe_hcno = re.sub(r"[^A-Za-z0-9._-]+", "_", hcno).strip("_")[:80] or "unknown"
	return f"row{row:04d}_review_{safe_hcno}.html"


def extract_hcnos_for_review(url: str, html_content: str) -> list[str]:
	ids = extract_feedback_ids(html_content)
	if ids:
		return ids

	parsed = urlparse(url)
	query = dict(parse_qsl(parsed.query, keep_blank_values=True))
	hcno = query.get("hcno", "").strip()
	if hcno:
		return [hcno]
	return []


def run(
	excel_path: Path,
	output_dir: Path,
	sheet_name: str | None,
	timeout: int,
	max_rows: int | None,
	fetch_review: bool,
	extract_caibiao: bool,
	caibiao_column: str,
	caibiao_header: str,
	excel_output: Path | None,
	enable_n_hint_validation: bool,
	n_hint_column: str,
	n_hint_validation_column: str,
	n_hint_validation_header: str,
) -> int:
	try:
		from openpyxl import load_workbook
	except ImportError as exc:
		raise RuntimeError(
			"Missing dependency: openpyxl. Install with: pip install openpyxl"
		) from exc

	wb = load_workbook(excel_path)
	ws = wb[sheet_name] if sheet_name else wb.active

	caibiao_col_idx: int | None = None
	n_hint_col_idx: int | None = None
	n_hint_validation_col_idx: int | None = None
	if extract_caibiao:
		caibiao_col_idx = resolve_caibiao_column_index(ws, caibiao_column)
		if caibiao_header and not ws.cell(row=1, column=caibiao_col_idx).value:
			ws.cell(row=1, column=caibiao_col_idx, value=caibiao_header)

	if enable_n_hint_validation:
		n_hint_col_idx = resolve_caibiao_column_index(ws, n_hint_column)
		if n_hint_validation_column.upper() != "NONE":
			n_hint_validation_col_idx = resolve_caibiao_column_index(ws, n_hint_validation_column)
			if n_hint_validation_header and not ws.cell(row=1, column=n_hint_validation_col_idx).value:
				ws.cell(row=1, column=n_hint_validation_col_idx, value=n_hint_validation_header)

	output_dir.mkdir(parents=True, exist_ok=True)
	review_dir = output_dir / "review"
	if fetch_review:
		review_dir.mkdir(parents=True, exist_ok=True)

	success = 0
	failed = 0
	processed = 0
	review_found = 0
	review_success = 0
	review_failed = 0
	caibiao_module_found = 0
	caibiao_text_found = 0
	caibiao_from_review_found = 0
	n_hint_positive = 0
	n_hint_positive_and_predicted = 0
	n_hint_positive_but_not_predicted = 0
	n_hint_negative_but_predicted = 0

	for row in range(1, ws.max_row + 1):
		if max_rows is not None and processed >= max_rows:
			break

		raw = ws[f"O{row}"].value
		url = normalize_url(raw)
		if not url:
			continue

		processed += 1
		filename = build_filename(url, row)
		target = output_dir / filename

		try:
			content, final_url = fetch_url_content_resilient(url, timeout=timeout)
			target.write_text(content, encoding="utf-8")
			success += 1
			print(f"[OK] row={row} saved={target.name} url={final_url}")

			predicted_has_caibiao = False
			if extract_caibiao and caibiao_col_idx is not None:
				has_module, caibiao_text = extract_caibiao_info(content)
				predicted_has_caibiao = has_module or bool(caibiao_text)
				best_caibiao_text = caibiao_text

				# Fallback to review -> gbDetailed content when newGbInfo page has no direct hit.
				if not predicted_has_caibiao:
					hcnos = extract_hcnos_for_review(url, content)
					for hcno in hcnos:
						review_url = build_review_url(url, hcno)
						try:
							review_content, _ = fetch_url_content_resilient(review_url, timeout=timeout)
							r_has_module, r_text = extract_caibiao_info(review_content)
							if r_has_module or r_text:
								predicted_has_caibiao = True
								has_module = r_has_module
								if r_text:
									best_caibiao_text = r_text
								caibiao_from_review_found += 1
								if fetch_review:
									review_target = review_dir / build_review_filename(row, hcno)
									review_target.write_text(review_content, encoding="utf-8")
								break
						except Exception:
							continue

				if has_module:
					caibiao_module_found += 1
				if best_caibiao_text:
					caibiao_text_found += 1
					ws.cell(row=row, column=caibiao_col_idx, value=best_caibiao_text)
				elif has_module:
					ws.cell(row=row, column=caibiao_col_idx, value="存在采标情况模块")
				else:
					ws.cell(row=row, column=caibiao_col_idx, value="")

			if enable_n_hint_validation and n_hint_col_idx is not None:
				n_hint_value = ws.cell(row=row, column=n_hint_col_idx).value
				expected_by_hint = n_column_hint_has_caibiao(n_hint_value)
				if expected_by_hint:
					n_hint_positive += 1
					if predicted_has_caibiao:
						n_hint_positive_and_predicted += 1
					else:
						n_hint_positive_but_not_predicted += 1
				elif predicted_has_caibiao:
					n_hint_negative_but_predicted += 1

				if n_hint_validation_col_idx is not None:
					if expected_by_hint and predicted_has_caibiao:
						status = "N列命中且已提取"
					elif expected_by_hint and not predicted_has_caibiao:
						status = "N列疑似采标但未提取"
					elif (not expected_by_hint) and predicted_has_caibiao:
						status = "N列未标采但提取为采标"
					else:
						status = "N列未标采且未提取"
					ws.cell(row=row, column=n_hint_validation_col_idx, value=status)

			if fetch_review:
				hcnos = extract_feedback_ids(content)
				review_found += len(hcnos)
				for hcno in hcnos:
					review_url = build_review_url(url, hcno)
					review_target = review_dir / build_review_filename(row, hcno)
					try:
						review_content = fetch_url_content(review_url, timeout=timeout)
						review_target.write_text(review_content, encoding="utf-8")
						review_success += 1
						print(f"[OK][REVIEW] row={row} saved={review_target.name} url={review_url}")
					except Exception as review_exc:
						review_failed += 1
						print(f"[FAIL][REVIEW] row={row} url={review_url} error={review_exc}")
		except Exception as exc:
			failed += 1
			print(f"[FAIL] row={row} url={url} error={exc}")
			if extract_caibiao and caibiao_col_idx is not None:
				ws.cell(row=row, column=caibiao_col_idx, value=f"抓取失败: {exc}")

	if extract_caibiao:
		out_excel = excel_output or excel_path
		out_excel = save_workbook_with_fallback(wb, out_excel)

	print("\nDone")
	print(f"Excel: {excel_path}")
	print(f"Sheet: {ws.title}")
	print(f"Processed URLs: {processed}")
	print(f"Success: {success}")
	print(f"Failed: {failed}")
	if fetch_review:
		print(f"Review IDs found: {review_found}")
		print(f"Review success: {review_success}")
		print(f"Review failed: {review_failed}")
		print(f"Review output folder: {review_dir}")
	if extract_caibiao:
		print(f"Caibiao module found: {caibiao_module_found}")
		print(f"Caibiao text extracted: {caibiao_text_found}")
		print(f"Caibiao hits from review fallback: {caibiao_from_review_found}")
	if enable_n_hint_validation:
		print(f"N-hint positive (N列含'采'): {n_hint_positive}")
		print(f"N-hint matched by extraction: {n_hint_positive_and_predicted}")
		print(f"N-hint possible misses: {n_hint_positive_but_not_predicted}")
		print(f"N-hint negative but extracted: {n_hint_negative_but_predicted}")
	if extract_caibiao:
		print(f"Excel updated: {out_excel}")
	print(f"Output folder: {output_dir}")
	if failed == 0 and (not fetch_review or review_failed == 0):
		return 0
	return 1


def main() -> int:
	parser = argparse.ArgumentParser(
		description="Read URLs from column O in Excel and save webpage content to content/ folder"
	)
	parser.add_argument("--excel", help="Excel file path (default: first .xlsx in current folder)")
	parser.add_argument("--sheet", help="Sheet name (default: active sheet)")
	parser.add_argument("--timeout", type=int, default=15, help="Request timeout in seconds")
	parser.add_argument("--max", type=int, default=None, help="Only process first N non-empty URLs")
	parser.add_argument(
		"--no-review",
		action="store_true",
		help="Only fetch URLs in column O, do not fetch /bzgk/gb/review pages",
	)
	parser.add_argument(
		"--no-caibiao",
		action="store_true",
		help="Skip checking/extracting \"采标情况\" information",
	)
	parser.add_argument(
		"--caibiao-column",
		default="AUTO",
		help="Target Excel column for extracted 采标情况 text (e.g. P, Q, AUTO)",
	)
	parser.add_argument(
		"--caibiao-header",
		default="采标情况提取",
		help="Header name for the output 采标情况 column",
	)
	parser.add_argument(
		"--excel-out",
		help="Output Excel path (default: overwrite source Excel)",
	)
	parser.add_argument(
		"--n-hint-validation",
		action="store_true",
		help="Enable internal validation: use column N containing '采' as a weak positive hint",
	)
	parser.add_argument(
		"--n-hint-column",
		default="N",
		help="Hint source column for weak label validation (default: N)",
	)
	parser.add_argument(
		"--n-hint-validation-column",
		default="NONE",
		help="Write per-row validation status to this column (e.g. R). Use NONE to skip writing",
	)
	parser.add_argument(
		"--n-hint-validation-header",
		default="采标校验(N列弱标签)",
		help="Header name for N-hint validation output column",
	)
	args = parser.parse_args()

	base_dir = Path(__file__).resolve().parent
	try:
		excel_path = find_excel_file(base_dir, args.excel)
		output_dir = base_dir / "content"
		excel_out = None
		if args.excel_out:
			excel_out = Path(args.excel_out)
			if not excel_out.is_absolute():
				excel_out = base_dir / excel_out
		return run(
			excel_path=excel_path,
			output_dir=output_dir,
			sheet_name=args.sheet,
			timeout=args.timeout,
			max_rows=args.max,
			fetch_review=not args.no_review,
			extract_caibiao=not args.no_caibiao,
			caibiao_column=args.caibiao_column,
			caibiao_header=args.caibiao_header,
			excel_output=excel_out,
			enable_n_hint_validation=args.n_hint_validation,
			n_hint_column=args.n_hint_column,
			n_hint_validation_column=args.n_hint_validation_column,
			n_hint_validation_header=args.n_hint_validation_header,
		)
	except Exception as exc:
		print(f"Error: {exc}")
		return 2


if __name__ == "__main__":
	sys.exit(main())
