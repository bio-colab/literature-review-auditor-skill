#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
recency_profile.py — الملفّ الزمنيّ الحتميّ لمِحَكّ الأدبيات والتأصيل.

يعطي دليلاً (لا حكماً) على البنية الزمنية لمراجع البحث:
  • الوسيط، وأحدث سنة، وأقدمها (ميلادياً).
  • الفجوة بين أحدث مرجعٍ مستشهَد وتاريخ التقديم  → أثرُ «الصمت المفاجئ».
  • حصّة مراجع آخر N سنة، وعدد ما قبل عتبةٍ يحدّدها المستخدم.
  • عزلُ السنوات الهجرية (مصادر تراثية أوّلية غالباً) عن التوزيع الميلاديّ،
    صَوناً لتمييز «الأوّليّ/الثانويّ» — فالتقادم يُطبَّق على الثانويّ لا الأوّليّ.

الحكمُ بالتقادم ليس من شأن السكربت: العائلةُ المعرفية وتمييزُ الأوّليّ/الثانويّ
يحدّدانه (انظر temporal-conceptual.md). السكربت يرسم التوزيع فقط.

المُدخَلات (يكتشفها تلقائياً بالامتداد/المحتوى):
  • ملف .docx           — يُستخرَج نصّه وتُلتقَط السنوات منه.
  • ملف .json من        detect_mechanism.py --refs  (سجلّاتٌ فيها حقل year).
  • ملف .txt/قائمة نصّية — تُلتقَط السنوات من سطورها.

الاستعمال:
    python recency_profile.py refs.json --submission-year 2024
    python recency_profile.py thesis.docx --recent-window 5 --old-threshold 2010
    python recency_profile.py list.txt --json out.json

لا يتصل بالإنترنت. مكتبة قياسية فقط.
"""
import sys, os, re, json, argparse, statistics, zipfile
import xml.etree.ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# نطاقٌ معقولٌ للسنة الميلادية في مرجعٍ أكاديميّ.
MIN_YEAR = 1500
MAX_YEAR_SLACK = 1  # نسمح بسنةٍ فوق الحالية (منشوراتٌ مبكرة)

# مؤشّرات التاريخ الهجريّ: رقمٌ يتبعه هـ / ه / AH، أو داخل قوسٍ بعده هـ.
HIJRI_RE = re.compile(r"(?<!\d)(\d{3,4})\s*(?:هـ?\b|ه\.?\s*ق|AH\b)")
GREG_RE = re.compile(r"(?<!\d)(1[5-9]\d{2}|20\d{2}|2100)(?!\d)")


def _current_year():
    import datetime
    return datetime.date.today().year


def extract_text_from_docx(path):
    """يستخرج نصّ الوثيقة من document.xml بلا مكتبات خارجية."""
    with zipfile.ZipFile(path) as z:
        try:
            xml = z.read("word/document.xml")
        except KeyError:
            return ""
    # أزل الوسوم واحتفظ بالنصّ داخل <w:t>.
    root = ET.fromstring(xml)
    texts = [node.text or "" for node in root.iter(f"{{{W}}}t")]
    return " ".join(texts)


def collect_years_from_text(text, cap):
    """يعيد (سنوات_ميلادية, عدد_هجرية). الهجرية تُعزَل ولا تُحسب في التوزيع."""
    # التقط الهجرية أولاً كي لا تُلتقَط أرقامُها ميلادياً.
    hijri = HIJRI_RE.findall(text)
    hijri_spans = [m.span() for m in HIJRI_RE.finditer(text)]

    greg = []
    for m in GREG_RE.finditer(text):
        # تجاهل ما وقع ضمن مطابقةٍ هجرية.
        if any(s <= m.start() < e for s, e in hijri_spans):
            continue
        y = int(m.group(1))
        if MIN_YEAR <= y <= cap:
            greg.append(y)
    return greg, len(hijri)


def collect_years_from_refs_json(data, cap):
    """من مخرَج detect_mechanism.py --refs: قائمةُ سجلّاتٍ فيها حقل year (قد يكون هجرياً كنصّ)."""
    if isinstance(data, dict):
        data = data.get("references", data.get("refs", []))
    greg, hijri = [], 0
    for rec in data or []:
        raw = str(rec.get("year", "") if isinstance(rec, dict) else rec).strip()
        if not raw:
            continue
        if HIJRI_RE.search(raw + "هـ") and not GREG_RE.search(raw):
            # سنةٌ صغيرة بلا صيغةٍ ميلادية واضحة قد تكون هجرية؛ عالِجها بحذر:
            pass
        gm = GREG_RE.search(raw)
        hm = HIJRI_RE.search(raw)
        if hm and not gm:
            hijri += 1
        elif gm:
            y = int(gm.group(1))
            if MIN_YEAR <= y <= cap:
                greg.append(y)
    return greg, hijri


def build_profile(years, hijri_count, submission_year, recent_window, old_threshold, cap):
    years = sorted(years)
    prof = {
        "gregorian_count": len(years),
        "hijri_count": hijri_count,
        "note_hijri": "السنوات الهجرية معزولةٌ — غالباً مصادرُ تراثية أوّلية لا يُقاس تقادمها.",
    }
    if not years:
        prof["warning"] = "لم تُلتقَط سنواتٌ ميلادية. تحقّق من المُدخَل أو مرّره من detect_mechanism --refs."
        return prof

    newest, oldest = years[-1], years[0]
    prof.update({
        "oldest": oldest,
        "newest": newest,
        "median": int(statistics.median(years)),
        "mean": round(statistics.mean(years), 1),
        "span_years": newest - oldest,
    })

    # حصّة آخر نافذة قياساً بأحدث سنةٍ حاضرة (أو التقديم إن أُعطي).
    anchor = submission_year or newest
    recent_cut = anchor - recent_window
    recent = [y for y in years if y >= recent_cut]
    prof["recent_window"] = recent_window
    prof["recent_anchor"] = anchor
    prof["recent_share"] = round(100 * len(recent) / len(years), 1)

    if old_threshold:
        pre = [y for y in years if y < old_threshold]
        prof["old_threshold"] = old_threshold
        prof["pre_threshold_count"] = len(pre)
        prof["pre_threshold_share"] = round(100 * len(pre) / len(years), 1)

    # مؤشّر «الصمت المفاجئ»: الفجوة بين أحدث مرجعٍ وتاريخ التقديم.
    if submission_year:
        prof["submission_year"] = submission_year
        prof["gap_newest_to_submission"] = submission_year - newest

    # توزيعٌ عشريّ مختصر.
    buckets = {}
    for y in years:
        b = (y // 5) * 5
        buckets[b] = buckets.get(b, 0) + 1
    prof["distribution_5yr"] = dict(sorted(buckets.items()))
    return prof


def human_summary(prof):
    L = []
    L.append("── الملفّ الزمنيّ (دليلٌ لا حكم) ──")
    L.append(f"مراجع ميلادية مُلتقَطة: {prof.get('gregorian_count', 0)}"
             + (f"  ·  هجرية معزولة: {prof['hijri_count']}" if prof.get("hijri_count") else ""))
    if prof.get("hijri_count"):
        L.append(f"  ⓘ {prof['note_hijri']}")
    if "warning" in prof:
        L.append("⚠ " + prof["warning"]); return "\n".join(L)
    L.append(f"المدى: {prof['oldest']}–{prof['newest']}  ·  الوسيط: {prof['median']}  ·  المتوسط: {prof['mean']}")
    L.append(f"حصّة آخر {prof['recent_window']} سنة (قياساً بـ{prof['recent_anchor']}): {prof['recent_share']}%")
    if "pre_threshold_share" in prof:
        L.append(f"ما قبل {prof['old_threshold']}: {prof['pre_threshold_count']} مرجعاً ({prof['pre_threshold_share']}%)")
    if "gap_newest_to_submission" in prof:
        g = prof["gap_newest_to_submission"]
        flag = "  ← افحص «الصمت المفاجئ»" if g >= 3 else ""
        L.append(f"الفجوة بين أحدث مرجع ({prof['newest']}) والتقديم ({prof['submission_year']}): {g} سنة{flag}")
    dist = "  ".join(f"{k}:{v}" for k, v in prof["distribution_5yr"].items())
    L.append("التوزيع (خماسيّ): " + dist)
    L.append("")
    L.append("تذكير: التقادم لا يُحكَم به هنا. طبّق عدسة العائلة وميّز الأوّليّ من الثانويّ")
    L.append("قبل أيّ تنبيه (temporal-conceptual.md).")
    return "\n".join(L)


def main():
    ap = argparse.ArgumentParser(description="الملفّ الزمنيّ لمراجع البحث (حتميّ، دون إنترنت).")
    ap.add_argument("input", help="ملف .docx أو .json (من detect_mechanism --refs) أو .txt")
    ap.add_argument("--submission-year", type=int, default=None, help="سنة تقديم الرسالة (لقياس الفجوة).")
    ap.add_argument("--recent-window", type=int, default=5, help="نافذة «الحديث» بالسنوات (افتراضي 5).")
    ap.add_argument("--old-threshold", type=int, default=None, help="عتبةُ «القديم» لعدّ ما قبلها.")
    ap.add_argument("--json", dest="json_out", default=None, help="اكتب الملفّ الزمنيّ إلى JSON.")
    args = ap.parse_args()

    cap = _current_year() + MAX_YEAR_SLACK
    path = args.input
    if not os.path.exists(path):
        sys.exit(f"لا يوجد ملف: {path}")

    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx":
        text = extract_text_from_docx(path)
        years, hijri = collect_years_from_text(text, cap)
    elif ext == ".json":
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        years, hijri = collect_years_from_refs_json(data, cap)
    else:  # نصّ عاديّ
        with open(path, encoding="utf-8", errors="ignore") as f:
            text = f.read()
        years, hijri = collect_years_from_text(text, cap)

    prof = build_profile(years, hijri, args.submission_year,
                         args.recent_window, args.old_threshold, cap)

    print(human_summary(prof))
    if args.json_out:
        with open(args.json_out, "w", encoding="utf-8") as f:
            json.dump(prof, f, ensure_ascii=False, indent=2)
        print(f"\n→ JSON: {args.json_out}")


if __name__ == "__main__":
    main()
