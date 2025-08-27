import re
import fitz  # PyMuPDF
import io
from PIL import Image
import pytesseract
from slugify import slugify

# 現場で使える簡潔な対策へマッピング
_DEF_MITIGATIONS = [
    "換気の良い場所で使用",
    "有機ガス用防毒マスクの着用",
    "保護手袋・保護眼鏡を着用",
    "皮膚・目に付着した場合は速やかに洗浄",
    "静電気防止・火気厳禁",
    "使用後は手洗い・作業衣の持ち出し禁止",
]


def _extract_text_pymupdf(pdf_bytes: bytes) -> str:
    doc = fitz.open(stream=pdf_bytes, filetype='pdf')
    texts = []
    for page in doc:
        texts.append(page.get_text())
    return '\n'.join(texts).strip()


def _extract_text_ocr(pdf_bytes: bytes, dpi: int = 300) -> str:
    doc = fitz.open(stream=pdf_bytes, filetype='pdf')
    chunks = []
    for page in doc:
        pix = page.get_pixmap(dpi=dpi)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        txt = pytesseract.image_to_string(img, lang='jpn')
        chunks.append(txt)
    return '\n'.join(chunks).strip()


def _clean(s: str) -> str:
    return re.sub(r"[\u3000\t ]+", " ", s).strip()


def _between(text: str, start_pat: str, end_pat: str) -> str:
    s = re.search(start_pat, text, re.S)
    if not s:
        return ""
    start = s.end()
    e = re.search(end_pat, text[start:], re.S)
    end = start + (e.start() if e else len(text) - start)
    return text[start:end]


def _parse_components(seg: str):
    lines = [l.strip() for l in seg.splitlines() if l.strip()]
    items = []
    for l in lines:
        # 例: 二酸化チタン 10% / キシレン 1〜5% / PRTR
        m = re.search(r"([^:：%]+?)\s*[:：]?[\s\t]*([0-9\.\-〜~]+)\s*%", l)
        if m:
            items.append(_clean(f"{m.group(1)} {m.group(2)}%"))
        else:
            # % が無いが成分らしい行も一応拾う
            if re.search(r"(ベンゼン|トルエン|キシレン|アルコール|ナフサ|酸化|シリカ|タルク|樹脂|溶剤|炭化水素)", l):
                items.append(_clean(l))
    # 重複除去
    seen = set(); out = []
    for it in items:
        if it not in seen:
            seen.add(it); out.append(it)
    return out[:50]  # 安全側


def _parse_ghs(text: str):
    seg = _between(text, r"(GHS分類|GHS\s*区分|ラベル要素).*?\n", r"\n\s*(第\s*4|4\.|応急|取扱|取り扱い|取り扱い及び保管)")
    if not seg:
        seg = _between(text, r"(危険有害性の要約|危険有害性情報).*?\n", r"\n\s*(第\s*4|4\.|応急)")
    lines = [l.strip() for l in seg.splitlines() if l.strip()]
    out = []
    for l in lines:
        if '区分' in l or re.search(r"(引火性|皮膚刺激|眼刺激|発がん|変異原|生殖毒性|STOT|誤えん|水生|吸入毒性)", l):
            out.append(_clean(l))
    return out[:50]


def _parse_hazards(text: str):
    seg = _between(text, r"(危険有害性の要約|危険有害性情報).*?\n", r"\n\s*(第\s*3|3\.|組成|成分|応急|取扱)")
    lines = [l.strip(' ・*\t') for l in seg.splitlines() if l.strip()]
    out = []
    for l in lines:
        if re.search(r"(危険|有害|毒性|刺激|眠気|めまい|中枢|肝|腎|呼吸|環境|発がん|生殖)", l):
            out.append(_clean(l))
    # 最低限カバー
    if not out:
        out = [
            "引火性がある",
            "皮膚・眼に刺激",
            "蒸気/ミストの吸入で健康障害のおそれ",
            "水生環境へ有害"
        ]
    return out[:30]


def _risk_level(ghs_list):
    text = ' '.join(ghs_list)
    if re.search(r"区分\s*1|発がん|生殖毒性|STOT.*区分\s*1", text):
        return "高"
    if re.search(r"引火性|区分\s*2|区分\s*3", text):
        return "中"
    return "中"


def _product_name(text: str, original_filename: str) -> str:
    for pat in [r"(製品名(?:称)?|品名)\s*[:：]\s*(.+)", r"(製品名(?:称)?|品名)\s*(.+)"]:
        m = re.search(pat, text)
        if m:
            name = m.group(2).strip()
            # 行末だけ
            return name.splitlines()[0].strip()
    # Fallback: ファイル名から
    base = original_filename.rsplit('/', 1)[-1].rsplit('\\', 1)[-1]
    return base.replace('.pdf', '').replace('.PDF', '') or "不明製品"


def parse_sds(pdf_bytes: bytes, original_filename: str = "") -> dict:
    text = _extract_text_pymupdf(pdf_bytes)
    if len(re.sub(r"\s+", "", text)) < 80:
        # 文字化け/アウトラインPDF → OCR
        text = _extract_text_ocr(pdf_bytes)

    name = _product_name(text, original_filename)

    # 第3項 組成・成分情報
    comp_seg = _between(text, r"(第\s*3\s*項.*?組成|3\.|組成\s*[・/ ]\s*成分)", r"\n\s*(第\s*4|4\.|応急|救急)")
    components = _parse_components(comp_seg)

    ghs = _parse_ghs(text)
    hazards = _parse_hazards(text)

    # 作業員向けに簡潔化（方針に準拠）
    mitigations = list(_DEF_MITIGATIONS)
    if not any('引火' in g for g in ghs):
        # 引火性明記なしなら火気厳禁は残しつつ控えめに
        pass

    ra = {
        '作業場所': '',
        '作業日': '',
        '使用製品名': name,
        '主な成分': components or ["（SDSから成分が読取不能のため、後日追記）"],
        'SDSの有無': '有',
        'GHSの分類（区分付き）': ghs or ["（SDSからGHS区分が読取不能のため、後日追記）"],
        'リスクレベルの判定': _risk_level(ghs),
        '主なリスクの内容': hazards,
        'リスク低減措置の検討': mitigations,
        '作成者': '若宮良裕',
        '所属会社': '株式会社若宮塗装工業所',
    }

    # ファイル名（日本語OK + ASCIIフォールバック）
    slug = slugify(name) or 'product'
    ra['filename_xlsx'] = f"{name}_リスクアセスメント.xlsx"
    ra['filename_pdf']  = f"{name}_リスクアセスメント.pdf"
    ra['fallback_xlsx'] = f"{slug}_risk_assessment.xlsx"
    ra['fallback_pdf']  = f"{slug}_risk_assessment.pdf"

    return ra
