"""
Karadel calc worker
====================
Небольшой Python-сервис, который снимает с Lovable тяжёлый расчёт.

Поток:
  Lovable -> POST /process {order_id}   (отвечаем 202 за доли секунды)
  фоновая задача: скачать чертёж и шаблон из Supabase -> Claude по чертежу
                  -> заполнить xlsx правкой XML в zip -> залить результат
                  -> статус заказа = "done"

Ничего не загружает целиком через ExcelJS/openpyxl: правятся ТОЛЬКО
два листа (sheet1.xml = Расчет, sheet6.xml = Сводная), остальные ~774
части копируются байт-в-байт. Выпадающие списки и условное
форматирование (extLst/x14) при этом сохраняются.

Запуск локально:  uvicorn main:app --reload
Деплой:           Railway (см. Procfile и README.md)
"""

import os
import io
import re
import json
import time
import base64
import logging
import zipfile
import traceback
from datetime import datetime, timezone

# Метка версии — видно по /health. Если после деплоя /health не показывает
# этот номер, значит на сервере СТАРЫЙ файл (загрузка не доехала).
WORKER_VERSION = "v7-comment-row1-2026-06-10"
from xml.sax.saxutils import escape

import fitz  # PyMuPDF
from PIL import Image
import anthropic
from supabase import create_client
from fastapi import FastAPI, BackgroundTasks, Header, HTTPException
from fastapi.responses import JSONResponse
from pydantic import BaseModel

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(message)s")
log = logging.getLogger("worker")

# ---------------------------------------------------------------------------
# КОНФИГ — подгоните под свою схему Supabase, если имена отличаются
# ---------------------------------------------------------------------------
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
SUPABASE_URL = os.environ.get("SUPABASE_URL", "")
SUPABASE_SERVICE_ROLE_KEY = os.environ.get("SUPABASE_SERVICE_ROLE_KEY", "")
WORKER_SECRET = os.environ.get("WORKER_SECRET", "")
CLAUDE_MODEL = os.environ.get("CLAUDE_MODEL", "claude-haiku-4-5-20251001")

# таблица заказов и её колонки
ORDERS_TABLE = "orders"
ID_COL = "id"
STATUS_COL = "status"
LOG_COL = "processing_log"          # JSON-массив
DRAWING_PATH_COL = "drawing_url"    # путь файла чертежа внутри бакета DRAWINGS_BUCKET
RESULT_PATH_COL = "result_url"      # сюда запишем путь готового файла
STATUS_DONE = "completed"           # статус при успехе (у вас "completed", не "done")
ERROR_MSG_COL = "error_message"     # сюда пишем текст ошибки

# бакеты хранилища
DRAWINGS_BUCKET = "drawings"
TEMPLATES_BUCKET = "templates"
RESULTS_BUCKET = "results"

# путь активного шаблона в бакете TEMPLATES_BUCKET (задаётся env-переменной)
TEMPLATE_PATH = os.environ.get("TEMPLATE_PATH", "")

MAX_PAGES = int(os.environ.get("MAX_PAGES", "15"))
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# ВАЖНО: вставьте сюда ваш рабочий промпт распознавания чертежа из Lovable.
# Главное требование — Claude должен вернуть ТОЛЬКО JSON указанного вида.
EXTRACTION_PROMPT = """
Ты — инженер-расчётчик металлоконструкций. На изображениях — листы чертежа КМ.
Главный источник данных — таблица «Техническая спецификация металла»
(обычно лист 1–2). Извлеки из неё перечень стального проката.

КАК УСТРОЕНА ТАБЛИЦА:
- Крайний ЛЕВЫЙ столбец — ТИП профиля (напр. «Швеллеры по ГОСТ 8240-97»,
  «Прокат листовой горячекатаный», «Прокат угловой равнополочный»,
  «Квадратные трубы», «Двутавр по ГОСТ 26020-83»). Он РАСТЯНУТ на несколько строк
  и относится КО ВСЕМ строкам-размерам под ним, до следующего типа.
- Столбец «Наименование или марка металла» — марка стали (С255, С345 и т.п.).
- Столбец «Номер или размер профиля в мм» — конкретный размер: [8, [10, [12;
  t5, t8, t10, t14…; L45x5, L90x6, L100x10; []60x5, []100x5; 30К1, 23К1, 30Б1.
- Средние столбцы «Масса металла по элементам конструкций» (Колонна, Ферма,
  Прогоны, связи, Раскос, Стеновые прогоны) — НЕ бери их по отдельности.
- Крайний ПРАВЫЙ столбец «Общая масса» — ИТОГОВАЯ масса этого размера.
  БЕРИ МАССУ ТОЛЬКО ОТСЮДА.

ЕДИНИЦЫ МАССЫ — ОЧЕНЬ ВАЖНО:
- В шапке столбца массы всегда подписана единица: «Общая масса, кг» ИЛИ
  «Общая масса, т» (тонны). Сначала ОПРЕДЕЛИ единицу по шапке.
- В ответе mass_kg и control_total_kg ВСЕГДА указывай В КИЛОГРАММАХ.
- Если шапка в тоннах («т», «тн», «тонн») — УМНОЖЬ каждое значение НА 1000
  (45,04 т → 45040). Если в кг — оставь как есть.
- Десятичная запятая = точка (45,04 → 45.04), затем перевод в кг.
- Подстраховка: итоговая масса металлоконструкции обычно ТЫСЯЧИ кг.
  Если после перевода сумма вышла в десятках (напр. 45) — ты забыл умножить
  тонны на 1000, перечитай шапку единиц.

ЧТО ДЕЛАТЬ (по каждой строке-размеру):
1) Если в столбце «Общая масса» стоит число — это позиция. Возьми её (в кг).
2) Если «Общая масса» пустая (нет массы) — ПРОПУСТИ строку (не вноси).
3) Тип бери из левого растянутого столбца, размер — из столбца размера,
   массу — из «Общая масса» (переведённую в кг), марку — из столбца марки.

ЧТО ПРОПУСКАТЬ ВСЕГДА:
- Строки «Всего профиля» (это подытоги — НЕ нужны).
- Строку «Итого масса металла» (общий итог — НЕ нужна).
- Блок «в том числе по маркам» (С235 / С255 / С345 — НЕ нужен).
- Строки без массы в столбце «Общая масса».

ПЕРЕВОД ТИПА В НАЗВАНИЕ (как в нашей базе «Сводная»); это база синонимов,
сопоставляй гибко:
- «Швеллеры …», «Швеллер», «[N»            → "Швеллер <номер>"  (напр. [10 → "Швеллер 10")
- «Прокат листовой …», «Лист», «t<N>»       → "Лист t<толщина>"  (t10 → "Лист t10")
- «Прокат угловой равнополочный …», «Уголок», «L…» → "Уголок <размер>" (L100x10 → "Уголок 100x10")
- ТРУБЫ. Круглая или профильная — определяй ТОЛЬКО по ЛЕВОМУ столбцу (вид
  профиля/ГОСТ), а НЕ по количеству чисел в размере. Это критично: и круглая,
  и квадратная труба могут писаться двумя числами (273x4 и 250x6 выглядят
  одинаково, но это РАЗНЫЕ трубы).
  • КВАДРАТНАЯ / ПРЯМОУГОЛЬНАЯ (профильная): левый столбец «Трубы квадратные…»,
    «Трубы прямоугольные…», ГОСТ 30245 / 8639 / 8645, слова «профильная»/«[]…».
    → "Труба проф. <размер>". Квадратная может быть записана как «250x6»
    (это 250×250×6) или «250×250×6»; прямоугольная — тремя числами «160×120×4».
  • КРУГЛАЯ: левый столбец «электросварные»/«бесшовные»/«водогазопроводные/ВГП»,
    ГОСТ 10704 / 10705 / 8732 / 8734 / 3262, символы «Ф»/«Ø», размер диаметр×толщина
    (273x4, Ф159х4). → "Труба <диаметр>x<толщина>" (273x4 → "Труба 273x4").
    НЕ пиши такие как «Труба проф.».
- «Двутавр …», «Балка», «<N>К<n>», «<N>Б<n>»→ "Балка <обозначение>" (30К1 → "Балка 30К1", 30Б1 → "Балка 30Б1")
ВАЖНО: обозначения вида «30К1», «23К1», «30Б1» — это ДВУТАВРЫ (балки),
а НЕ листы. Никогда не пиши их как «Лист».

ПОЛЯ ПОЗИЦИИ:
- sortament: переведённое название (как в «Сводной»).
- grade: марка стали.
- mass_kg: число из «Общая масса», ПЕРЕВЕДЁННОЕ В КИЛОГРАММЫ (обязательно).
- qty: если есть, иначе пропусти.
- uncertain: true, ЕСЛИ ты не уверен в этой позиции (чертёж размыт, цифры/название
  плохо читаются, неоднозначно). НЕ ФАНТАЗИРУЙ — лучше поставь uncertain:true.
  Если всё чётко — uncertain:false или не указывай.

ЗАПРЕЩЕНО: выдумывать профили/размеры/массы, использовать римские цифры,
повторять один и тот же размер дважды, включать позиции без массы.
Если число не разобрать — поставь uncertain:true (НЕ угадывай значение).

Самопроверка: сумма всех mass_kg (в кг) должна примерно совпасть с «Итого масса
металла» из таблицы (тоже переведённой в кг). Если расходится в ~1000 раз —
ты перепутал тонны и килограммы, перечитай шапку единиц столбца массы.
Дополнительно верни это контрольное число в поле "control_total_kg"
(значение строки «Итого масса металла» В КИЛОГРАММАХ; если в шапке тонны —
умножь на 1000; если строки нет — null).

Верни СТРОГО JSON без пояснений и markdown:
{
  "drawing": "обозначение из штампа, если видно",
  "control_total_kg": 70103.84,
  "positions": [
    {"sortament": "Швеллер 10", "grade": "С255", "mass_kg": 5803.68, "qty": 1},
    {"sortament": "Лист t10", "grade": "С345", "mass_kg": 8021.71, "qty": 1},
    {"sortament": "Балка 30К1", "grade": "С345", "mass_kg": 15152, "qty": 1}
  ]
}
Только JSON, ничего больше.
""".strip()

sb = create_client(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY)
app = FastAPI(title="Karadel calc worker")


# ---------------------------------------------------------------------------
# Вспомогательное
# ---------------------------------------------------------------------------
def now_iso():
    return datetime.now(timezone.utc).isoformat()


def col_to_idx(letters):
    n = 0
    for c in letters:
        n = n * 26 + (ord(c) - 64)
    return n


def idx_to_col(idx):
    s = ""
    while idx > 0:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


# стальные сортаменты живут в колонках N..BA (40 штук)
STEEL_COLUMNS = [idx_to_col(i) for i in range(col_to_idx("N"), col_to_idx("BA") + 1)]


def fetch_order(order_id):
    res = sb.table(ORDERS_TABLE).select("*").eq(ID_COL, order_id).single().execute()
    return res.data


def db_update(order_id, fields):
    sb.table(ORDERS_TABLE).update(fields).eq(ID_COL, order_id).execute()


def log_run(*, drawing_name, check_status,
            positions_count=None, sum_mass=None, expected_total=None,
            review=None, cost_cents=None, request_id=None, log_text=None):
    """Пишет одну строку-журнал в worker_runs по итогам обработки чертежа.
    Никогда не роняет основную обработку: ошибку журнала просто логируем."""
    review = review or {}
    payload = {
        "drawing_name":   drawing_name,
        "request_id":     request_id,
        "worker_version": WORKER_VERSION,
        "model":          CLAUDE_MODEL,
        "check_status":   check_status,        # "OK" / "MISMATCH" / "NO_CONTROL"
        "positions_count": positions_count,
        "sum_mass":       sum_mass,
        "expected_total": expected_total,
        "review_total":   review.get("total"),
        "review_yellow":  review.get("yellow"),
        "review_orange":  review.get("orange"),
        "review_red":     review.get("red"),
        "cost_cents":     cost_cents,
        "log_text":       log_text,
    }
    try:
        sb.table("worker_runs").insert(payload).execute()
    except Exception as e:
        log.error("[log_run] не удалось записать журнал прогона: %s", e)


def get_template_bytes():
    if not TEMPLATE_PATH:
        raise RuntimeError("TEMPLATE_PATH не задан (путь активного шаблона в бакете templates)")
    return sb.storage.from_(TEMPLATES_BUCKET).download(TEMPLATE_PATH)


def download_drawing(order):
    path = order.get(DRAWING_PATH_COL)
    if not path:
        raise RuntimeError(f"в заказе нет {DRAWING_PATH_COL}")
    return sb.storage.from_(DRAWINGS_BUCKET).download(path)


def upload_result(path, data):
    # upsert=true чтобы перезапись при повторном расчёте не падала
    sb.storage.from_(RESULTS_BUCKET).upload(
        path, data, {"content-type": XLSX_MIME, "upsert": "true"}
    )


# ---------------------------------------------------------------------------
# Чертёж -> картинки для Claude
# ---------------------------------------------------------------------------
def drawing_to_images(data, filename):
    raw_images = []
    is_pdf = filename.lower().endswith(".pdf") or data[:4] == b"%PDF"
    if is_pdf:
        doc = fitz.open(stream=data, filetype="pdf")
        for i, page in enumerate(doc):
            if i >= MAX_PAGES:
                break
            long_side = max(page.rect.width, page.rect.height) or 1000
            zoom = min(3.0, 1500.0 / long_side)
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
            raw_images.append(Image.frombytes("RGB", [pix.width, pix.height], pix.samples))
    else:
        raw_images.append(Image.open(io.BytesIO(data)).convert("RGB"))

    out = []
    for img in raw_images:
        img.thumbnail((1500, 1500))
        buf = io.BytesIO()
        img.save(buf, "JPEG", quality=75)
        out.append(base64.b64encode(buf.getvalue()).decode())
    return out


# ---------------------------------------------------------------------------
# Claude -> позиции
# ---------------------------------------------------------------------------
def parse_json(text):
    t = text.strip()
    t = re.sub(r"^```(?:json)?", "", t).strip()
    t = re.sub(r"```$", "", t).strip()
    start, end = t.find("{"), t.rfind("}")
    if start != -1 and end != -1:
        t = t[start:end + 1]
    return json.loads(t)


def extract_positions(images_b64):
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    content = [
        {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": b}}
        for b in images_b64
    ]
    content.append({"type": "text", "text": EXTRACTION_PROMPT})
    resp = client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=4096,
        messages=[{"role": "user", "content": content}],
    )
    text = "".join(b.text for b in resp.content if getattr(b, "type", "") == "text")
    return parse_json(text)


# ---------------------------------------------------------------------------
# Правка xlsx через zip (без ExcelJS / openpyxl-save)
# ---------------------------------------------------------------------------
def _build_cell(ref, attrs, value, kind, style_id=None):
    if style_id is not None:
        s = f' s="{style_id}"'
    else:
        m = re.search(r'\bs="(\d+)"', attrs or "")
        s = f' s="{m.group(1)}"' if m else ""
    if kind == "n":
        return f'<c r="{ref}"{s}><v>{value}</v></c>'
    text = escape(str(value))
    return f'<c r="{ref}"{s} t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>'


def set_cell(xml, ref, value, kind, style_id=None):
    col = re.match(r"[A-Z]+", ref).group()
    row = re.search(r"\d+", ref).group()
    col_idx = col_to_idx(col)

    # 1) ячейка уже есть -> заменить
    pat = re.compile(r'<c r="' + ref + r'"([^>]*?)(?:/>|>.*?</c>)', re.S)
    if pat.search(xml):
        return pat.sub(lambda m: _build_cell(ref, m.group(1), value, kind, style_id), xml, count=1)

    # 2) ячейки нет -> вставить в нужную строку в порядке колонок
    rowpat = re.compile(r'(<row r="' + row + r'"[^>]*>)(.*?)(</row>)', re.S)
    m = rowpat.search(xml)
    new_cell = _build_cell(ref, "", value, kind, style_id)
    if not m:
        return xml.replace("</sheetData>", f'<row r="{row}">{new_cell}</row></sheetData>', 1)
    head, body, tail = m.group(1), m.group(2), m.group(3)
    insert_pos = len(body)
    for cm in re.finditer(r'<c r="([A-Z]+)\d+"', body):
        if col_to_idx(cm.group(1)) > col_idx:
            insert_pos = cm.start()
            break
    body = body[:insert_pos] + new_cell + body[insert_pos:]
    return xml[:m.start()] + head + body + tail + xml[m.end():]


def _split_xfs(body):
    """Аккуратно разбивает cellXfs на отдельные <xf> (учёт самозакрытия и <alignment>)."""
    xfs = []
    for m in re.finditer(r'<xf\b', body):
        start = m.start()
        rest = body[start:]
        gt = rest.find(">")
        if gt > 0 and rest[gt - 1] == "/":
            end = start + gt + 1
        else:
            close = rest.find("</xf>")
            end = start + close + len("</xf>")
        xfs.append(body[start:end])
    return xfs


def _ensure_red_fill(styles_xml):
    """Гарантирует наличие красной заливки в палитре. Возвращает (xml, fill_id)."""
    m = re.search(r'(<fills count=")(\d+)(">)(.*?)(</fills>)', styles_xml, re.S)
    body = m.group(4)
    # уже есть красная (FFFF0000)?
    fills = re.findall(r"<fill>.*?</fill>", body, re.S)
    for i, f in enumerate(fills):
        if 'rgb="FFFF0000"' in f:
            return styles_xml, i
    red = '<fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/><bgColor indexed="64"/></patternFill></fill>'
    new_id = len(fills)
    styles_xml = (styles_xml[:m.start()] + m.group(1) + str(int(m.group(2)) + 1) + m.group(3)
                  + body + red + m.group(5) + styles_xml[m.end():])
    return styles_xml, new_id


def add_marker_style(styles_xml, fill_id, base_index=374):
    """Добавляет в cellXfs стиль = базовый(374) + заливка fill_id. Возвращает (xml, индекс)."""
    m = re.search(r'(<cellXfs count=")(\d+)(">)(.*?)(</cellXfs>)', styles_xml, re.S)
    if not m:
        return styles_xml, None
    body = m.group(4)
    xfs = _split_xfs(body)
    base = xfs[base_index] if base_index < len(xfs) else '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    new = base
    if 'fillId="' in new:
        new = re.sub(r'fillId="\d+"', f'fillId="{fill_id}"', new, count=1)
    else:
        new = new.replace("<xf", f'<xf fillId="{fill_id}"', 1)
    if "applyFill=" in new:
        new = re.sub(r'applyFill="\d"', 'applyFill="1"', new, count=1)
    else:
        new = new.replace("<xf", '<xf applyFill="1"', 1)
    new_index = len(xfs)
    styles_xml = (styles_xml[:m.start()] + m.group(1) + str(int(m.group(2)) + 1) + m.group(3)
                  + body + new + m.group(5) + styles_xml[m.end():])
    return styles_xml, new_index


def ensure_full_calc(wb):
    if "fullCalcOnLoad" in wb:
        return wb
    m = re.search(r"<calcPr\b[^>]*?/>", wb)
    if m:
        return wb.replace(m.group(0), m.group(0)[:-2] + ' fullCalcOnLoad="1"/>', 1)
    return re.sub(r"(</sheets>)", r'\1<calcPr fullCalcOnLoad="1"/>', wb, count=1)


def read_svodnaya_names(xlsx_bytes):
    """Читает названия материалов из листа «Сводная» (столбец C, кэш-значения формул)."""
    entries, _ = read_svodnaya_full(xlsx_bytes)
    return [nm for _, nm, _ in entries]


def read_svodnaya_full(xlsx_bytes):
    """Возвращает ([(row, name, price)], [free_rows]) из листа «Сводная» (C=имя, G=цена)."""
    try:
        z = zipfile.ZipFile(io.BytesIO(xlsx_bytes))
        s6 = z.read("xl/worksheets/sheet6.xml").decode("utf-8")
    except Exception:
        return [], []

    def cv(ref):
        m = re.search(r'<c r="' + ref + r'"[^>]*?>(?:<f>[^<]*</f>)?<v>([^<]*)</v>', s6)
        return m.group(1) if m else None

    entries, free = [], []
    for r in range(8, 621):
        c = cv(f"C{r}")
        if c and any(ch.isalpha() for ch in c):
            g = cv(f"G{r}")
            try:
                price = float(g)
            except (TypeError, ValueError):
                price = None
            entries.append((r, c, price))
        elif c is None:
            free.append(r)
    return entries, free


def make_real_name(real_size, analog_name, typ):
    """Подставляет реальный размер в формат имени аналога (для дозаписи в «Сводную»)."""
    real = int(real_size) if float(real_size).is_integer() else real_size
    if typ == "БАЛКА":
        return re.sub(r"(,\s*)\d+", lambda m: m.group(1) + str(real), analog_name, count=1)
    if typ == "ТРУБА_ПРОФ":
        return re.sub(r"\d+х\d+", f"{real}х{real}", analog_name, count=1)
    return re.sub(r"\d+", str(real), analog_name, count=1)


def _parse_profile(name):
    """(тип, [размеры], марка, подтип) из названия — нашего или из «Сводной»."""
    s = name.lower().replace("ё", "е")
    grade = "09Г2С" if "09г2с" in s else ("ст3" if ("ст 3" in s or "ст3" in s) else None)

    def nums(txt):
        txt = re.sub(r"(\d),(\d)", r"\1.\2", txt)  # 4,0 -> 4.0
        return [float(x) for x in re.findall(r"\d+(?:\.\d+)?", txt)]

    if "швеллер" in s:
        return ("ШВЕЛЛЕР", nums(s.split(",")[0].replace("швеллер", ""))[:1], grade, None)
    if "уголок" in s:
        a = nums(s.split("уголок")[1].split(",")[0])
        sub = None
        if len(a) >= 3:
            if a[0] != a[1]:
                sub = "неравн"            # неравнополочный
            dims = [max(a[:-1]), a[-1]]   # бОльшая полка + толщина
        elif len(a) == 2:
            dims = [a[0], a[1]]
        else:
            dims = a[:1]
        return ("УГОЛОК", dims, grade, sub)
    if "лист" in s:
        body = s.split("лист")[1].split(",")[0].replace("ст 3", "").replace("t", "")
        return ("ЛИСТ", nums(body)[:1], grade, None)
    if "балка" in s or "двутавр" in s:
        after = s.split("балка")[-1] if "балка" in s else s.split("двутавр")[-1]
        m_our = re.search(r"(\d+)\s*([кбш])", after)  # наш формат "30к1"
        if m_our:
            return ("БАЛКА", [float(m_our.group(1))], grade, m_our.group(2).upper())
        m_sub = re.match(r"\s*([кбшм])\b", after)      # Сводная "к-1, 30"
        sub = m_sub.group(1).upper() if m_sub else None
        size = None
        for part in after.split(",")[1:]:
            d = re.findall(r"\d+", part)
            if d:
                size = float(d[0]); break
        return ("БАЛКА", [size] if size else [], grade, sub)
    if "труб" in s:
        is_prof = "проф" in s
        body = re.sub(r"(\d),(\d)", r"\1.\2", s).split(",")[0]  # отрезать ", 12м"
        n = [float(x) for x in re.findall(r"\d+(?:\.\d+)?", body)]
        if is_prof:
            if len(n) >= 3:
                sides = sorted(n[:-1], reverse=True)[:2]
                dims = [sides[0], sides[1], n[-1]]   # [большая сторона, меньшая, толщина]
            elif len(n) == 2:
                dims = [n[0], n[0], n[1]]            # квадрат «250x6» = 250×250×6
            else:
                dims = n
            return ("ТРУБА_ПРОФ", dims, grade, None)
        return ("ТРУБА_КРУГ", n, grade, None)
    return ("?", [], grade, None)


def build_svodnaya_index(names):
    index = {}
    for nm in names:
        p = _parse_profile(nm)
        index.setdefault(p[0], []).append((p, nm))
    return index


def _grade_class(g):
    """Класс марки для выбора варианта в «Сводной»: '09Г2С' (низколегир.) | 'ст3' (углерод.) | None."""
    g = (g or "").upper().replace(" ", "")
    if any(k in g for k in ("09Г2С", "10ХСНД", "С345", "С375", "С390", "С440")):
        return "09Г2С"
    if any(k in g for k in ("С235", "С245", "С255", "С285", "СТ3", "ВСТ3")):
        return "ст3"
    return None


def resolve_to_svodnaya(name, grade, index):
    """Возвращает (имя_из_Сводной, способ). способ: 'точно' | 'аналог' | 'нет'."""
    p = _parse_profile(name)
    typ, dims, _, sub = p
    if typ == "УГОЛОК" and "гнут" in name.lower():
        return name, "гнутый"
    cands = index.get(typ, [])
    if not cands:
        return name, "нет"
    gcls = _grade_class(grade)
    if typ == "ЛИСТ":
        want = gcls                       # марка листа берётся с чертежа: С245→ст3, С345→09Г2С
    elif typ == "УГОЛОК" and gcls == "09Г2С":
        want = "09Г2С"
    else:
        want = None

    def gpref(lst):
        if want:
            g = [x for x in lst if x[0][2] == want]
            if g:
                return g
        ng = [x for x in lst if x[0][2] in (None, "ст3")]
        return ng or lst

    def same_sub(pp):
        return not (typ == "БАЛКА" and sub and pp[3] and pp[3] != sub)

    exact = [(pp, nm) for pp, nm in cands if pp[1] == dims and same_sub(pp)]
    if exact:
        return gpref(exact)[0][1], "точно"

    # ближайший больший
    if typ == "ТРУБА_ПРОФ" and len(dims) >= 3:
        big, small, wall = dims[0], dims[1], dims[2]
        # 1) то же сечение (обе стороны), толщина >= нужной — ближайшая большая толщина
        same_sec = [(pp[1][2], nm) for pp, nm in cands
                    if len(pp[1]) >= 3 and pp[1][0] == big and pp[1][1] == small and pp[1][2] >= wall]
        if same_sec:
            return min(same_sec)[1], "аналог"
        # 2) обе стороны >= нужных и толщина >= нужной — ближайший по площади сечения
        bigger = [(pp[1][0] * pp[1][1], nm) for pp, nm in cands
                  if len(pp[1]) >= 3 and pp[1][0] >= big and pp[1][1] >= small and pp[1][2] >= wall]
        if bigger:
            return min(bigger)[1], "аналог"
        # 3) обе стороны >= нужных, толщина любая
        bigger2 = [(pp[1][0] * pp[1][1], nm) for pp, nm in cands
                   if len(pp[1]) >= 3 and pp[1][0] >= big and pp[1][1] >= small]
        if bigger2:
            return min(bigger2)[1], "аналог"
    elif typ == "УГОЛОК" and len(dims) >= 2:
        leg, th = dims[0], dims[1]
        legs = sorted({pp[1][0] for pp, nm in cands if pp[1]})
        bigger = [L for L in legs if L >= leg]
        if bigger:
            std_leg = bigger[0]
            same = gpref([(pp, nm) for pp, nm in cands
                          if pp[1] and pp[1][0] == std_leg and len(pp[1]) >= 2 and pp[1][1] == th])
            if same:
                return same[0][1], "аналог"
        # нет стандартного с такой толщиной на ближайшей полке -> гнутый
        return name, "гнутый"
    else:
        big = sorted([(pp[1][0], pp, nm) for pp, nm in cands if pp[1] and dims and pp[1][0] >= dims[0] and same_sub(pp)], key=lambda x: x[0])
        big = gpref([(pp, nm) for _, pp, nm in big])
        if big:
            return big[0][1], "аналог"
    return name, "нет"


def match_sortament(name):
    # подбор по «Сводной» выполняется в build_raschet_cells (нужен индекс).
    return (name or "").strip()


def _norm_key(name):
    # ключ для дедупа: убрать "ГОСТ ...", лишние пробелы, привести к нижнему регистру
    s = re.sub(r"\s*ГОСТ[\s\S]*$", "", name, flags=re.I)
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s


def build_raschet_cells(positions, drawing_name, svod_index=None, price_map=None):
    """
    Возвращает (cells, notes, dopisat, summary).
    cells[ref] = (значение, тип, mark)  где mark = None | "yellow" | "orange" | "red"
    comment-ячейки (строка 7) тоже в cells (без цвета).
    dopisat = [(имя, цена)] — что дописать в «Сводную» (оранжевые).
    summary = {"yellow":n,"orange":n,"red":n}
    """
    price_map = price_map or {}
    cells = {
        "C6": ("Металлоконструкции", "s", None),
        "D6": (drawing_name or "", "s", None),
        "I6": ("шт", "s", None),
        "J6": (1, "n", None),
    }
    agg, info, order, notes, dopisat = {}, {}, [], [], []
    summary = {"yellow": 0, "orange": 0, "red": 0}

    for p in positions:
        raw = match_sortament(p.get("sortament") or p.get("name") or "")
        if not raw:
            continue
        try:
            mass = float(p.get("mass_kg") or p.get("mass") or 0)
        except (TypeError, ValueError):
            mass = 0.0
        if mass <= 0:
            continue
        grade = str(p.get("grade") or "")
        uncertain = bool(p.get("uncertain"))
        typ = _parse_profile(raw)[0]
        dim0 = (_parse_profile(raw)[1] or [None])[0]

        name, mark, comment = raw, None, ""
        if uncertain:
            name, mark = raw, "red"
            comment = "ВНИМАНИЕ: неуверенно прочитано с чертежа. Проверить вручную."
        elif svod_index:
            resolved, how = resolve_to_svodnaya(raw, grade, svod_index)
            if how == "точно":
                name, mark = resolved, None
                if typ == "УГОЛОК" and _parse_profile(raw)[3] == "неравн":
                    mark = "yellow"
                    comment = f"Неравнополочный уголок сведён к равнополочному {resolved}. Проверить."
            elif how == "гнутый":
                # гнутый уголок: стандартного с такой толщиной нет — оставляем
                # реальную геометрию, помечаем красным, цену ставит менеджер
                a = [x.replace(",", ".") for x in re.findall(r"\d+(?:[.,]\d+)?", raw)]
                def _f(v):
                    return str(int(float(v))) if float(v).is_integer() else v
                body = "*".join(_f(x) for x in a) if a else ""
                real = f"Уголок {body}".strip()
                if "гнут" not in real.lower():
                    real += " гнутый"
                name, mark = real, "red"
                comment = "Гнутый уголок (нет стандартного с такой толщиной полки). Запросить цену у менеджера."
            elif how == "аналог":
                if typ == "ЛИСТ":
                    name, mark = resolved, "yellow"
                    comment = f"Лист заменён на ближайший: {resolved}."
                elif typ in ("ТРУБА_ПРОФ", "ТРУБА_КРУГ"):
                    name, mark = raw, "orange"
                    price = price_map.get(resolved)
                    dopisat.append((raw, price))
                    comment = f"Нет в базе. Цена от аналога {resolved}. Проверить."
                elif typ == "УГОЛОК":
                    # полка округлена до ближайшей стандартной; resolved РЕАЛЬНО есть
                    # в «Сводной» со своей ценой -> жёлтый, дозапись не нужна
                    name, mark = resolved, "yellow"
                    comment = f"Уголок заменён на ближайший стандартный: {resolved}. Цена от него."
                else:
                    real = make_real_name(dim0, resolved, typ) if dim0 else raw
                    name, mark = real, "orange"
                    price = price_map.get(resolved)
                    dopisat.append((real, price))
                    comment = f"Нет в базе. Цена от аналога {resolved}. Проверить."
            else:  # не нашли совсем
                name, mark = raw, "red"
                comment = "Не найдено в базе. Проверить и добавить цену."
        else:
            name, mark = raw, None

        key = _norm_key(name)
        if key not in agg:
            agg[key] = 0.0
            info[key] = (name, mark, comment)
            order.append(key)
        agg[key] = max(agg[key], mass)
        if mark in ("аналог",):
            pass
        if mark:
            notes.append({"исходное": raw, "имя": name, "пометка": mark, "коммент": comment})

    total = 0.0
    for i, key in enumerate(order):
        if i >= len(STEEL_COLUMNS):
            break
        c = STEEL_COLUMNS[i]
        name, mark, comment = info[key]
        cells[f"{c}5"] = (name, "s", mark)             # имя (с цветом)
        cells[f"{c}6"] = (round(agg[key], 2), "n", None)  # масса
        if comment:
            cells[f"{c}1"] = (comment, "s", None)      # комментарий сверху, над названием
        if mark in summary:
            summary[mark] += 1
        total += agg[key]
    cells["K6"] = (round(total, 2), "n", None)
    return cells, notes, dopisat, summary


def fill_template(xlsx_bytes, positions, drawing_name):
    zin = zipfile.ZipFile(io.BytesIO(xlsx_bytes))
    parts = {n: zin.read(n) for n in zin.namelist()}

    entries, free_rows = read_svodnaya_full(xlsx_bytes)
    price_map = {nm: pr for _, nm, pr in entries}
    svod_index = build_svodnaya_index([nm for _, nm, _ in entries])
    cells, notes, dopisat, summary = build_raschet_cells(positions, drawing_name, svod_index, price_map)

    # стили цветов: для каждой помечаемой ячейки клонируем ЕЁ СОБСТВЕННЫЙ стиль
    # (шрифт, перенос, рамка) и только меняем заливку — формат ячейки сохраняется
    styles = parts["xl/styles.xml"].decode("utf-8")
    s1 = parts["xl/worksheets/sheet1.xml"].decode("utf-8")

    used = {m for _, _, m in cells.values() if m}
    fill_for = {}
    if "red" in used:
        styles, red_fill = _ensure_red_fill(styles)
    for mark in used:
        fill_for[mark] = red_fill if mark == "red" else {"yellow": 2, "orange": 6}[mark]

    def cur_style(ref):
        mm = re.search(r'<c r="' + ref + r'"([^>]*?)(?:/>|>)', s1)
        if mm:
            sm = re.search(r'\bs="(\d+)"', mm.group(1))
            if sm:
                return int(sm.group(1))
        return 0

    # лист «Расчет»
    marker_cache = {}   # (исходный стиль ячейки, заливка) -> индекс нового стиля
    for ref, (val, kind, mark) in cells.items():
        sid = None
        if mark:
            ckey = (cur_style(ref), fill_for[mark])
            if ckey not in marker_cache:
                styles, idx = add_marker_style(styles, ckey[1], base_index=ckey[0])
                marker_cache[ckey] = idx
            sid = marker_cache[ckey]
        s1 = set_cell(s1, ref, val, kind, sid)

    parts["xl/styles.xml"] = styles.encode("utf-8")
    parts["xl/worksheets/sheet1.xml"] = s1.encode("utf-8")

    # дозапись оранжевых в «Сводную» (реальное имя + цена аналога) в свободные строки
    if dopisat and free_rows:
        s6 = parts["xl/worksheets/sheet6.xml"].decode("utf-8")
        for (real_name, price), row in zip(dopisat, list(free_rows)):
            s6 = set_cell(s6, f"C{row}", real_name, "s")
            if price is not None:
                s6 = set_cell(s6, f"G{row}", price, "n")
        parts["xl/worksheets/sheet6.xml"] = s6.encode("utf-8")

    wb = parts["xl/workbook.xml"].decode("utf-8")
    parts["xl/workbook.xml"] = ensure_full_calc(wb).encode("utf-8")

    # Мы перезаписали формульные ячейки (K6, C6, N5/N6 …) значениями, а
    # xl/calcChain.xml всё ещё считает их формулами — из-за этого Excel при
    # открытии показывает «Ошибка в части содержимого… восстановить?».
    # Удаляем calcChain (и две ссылки на него) — Excel строит цепочку заново,
    # тем более что fullCalcOnLoad уже включён.
    parts.pop("xl/calcChain.xml", None)
    ct = parts["[Content_Types].xml"].decode("utf-8")
    ct = re.sub(r'<Override[^>]*calcChain\.xml[^>]*/>', "", ct)
    parts["[Content_Types].xml"] = ct.encode("utf-8")
    rels = parts["xl/_rels/workbook.xml.rels"].decode("utf-8")
    rels = re.sub(r'<Relationship[^>]*calcChain\.xml[^>]*/>', "", rels)
    parts["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")

    out = io.BytesIO()
    zout = zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED)
    for zi in zin.infolist():
        if zi.filename not in parts:      # пропускаем удалённый calcChain
            continue
        zout.writestr(zi, parts[zi.filename])
    zout.close()
    return out.getvalue(), notes, summary


# ---------------------------------------------------------------------------
# Фоновая обработка
# ---------------------------------------------------------------------------
def process_order(order_id):
    order = fetch_order(order_id)
    entries = order.get(LOG_COL) or []
    t0 = time.time()

    def phase(name, **extra):
        entries.append({"ts": now_iso(), "step": name, **extra})
        log.info("[worker] %s %s", name, extra or "")
        db_update(order_id, {LOG_COL: entries})

    try:
        phase("LOAD_TEMPLATE")
        tpl = get_template_bytes()

        phase("LOAD_DRAWING")
        draw = download_drawing(order)

        phase("RASTERIZE")
        images = drawing_to_images(draw, order.get(DRAWING_PATH_COL, ""))

        phase("CLAUDE_CALL", pages=len(images), model=CLAUDE_MODEL)
        t = time.time()
        data = extract_positions(images)
        positions = data.get("positions") or []
        phase("CLAUDE_DONE", ms=int((time.time() - t) * 1000), positions=len(positions))

        # --- Самопроверка: сумма масс vs «Итого масса металла» из таблицы ---
        sum_mass = 0.0
        for p in positions:
            try:
                m = float(p.get("mass_kg") or p.get("mass") or 0)
            except (TypeError, ValueError):
                m = 0.0
            if m > 0:
                sum_mass += m
        sum_mass = round(sum_mass, 2)
        try:
            control = data.get("control_total_kg")
            control = float(control) if control is not None else None
        except (TypeError, ValueError):
            control = None

        check = {"sum_mass_kg": sum_mass, "control_total_kg": control}
        if control and control > 0:
            diff_pct = round(abs(sum_mass - control) / control * 100, 1)
            check["diff_pct"] = diff_pct
            # допуск 1% — масса в спецификации дана с учётом отходов/швов
            check["ok"] = diff_pct <= 1.0
            check_status = "CHECK_OK" if check["ok"] else "CHECK_MISMATCH"
        else:
            check["ok"] = None
            check["note"] = "контрольный итог не найден — проверьте вручную"
            check_status = "CHECK_NO_CONTROL"
        phase(check_status, **check)
        needs_review = check.get("ok") is not True
        # --------------------------------------------------------------------

        phase("FILL_EXCEL")
        drawing_name = data.get("drawing") or order.get(DRAWING_PATH_COL, "")
        filled, notes, summary = fill_template(tpl, positions, drawing_name)
        if notes:
            phase("SVODNAYA_ANALOGS", count=len(notes), items=notes)
        review_total = summary["yellow"] + summary["orange"] + summary["red"]
        phase("REVIEW_SUMMARY", **summary, total=review_total)

        phase("SAVE", bytes=len(filled))
        path = f"{order_id}.xlsx"
        upload_result(path, filled)
        result_fields = {RESULT_PATH_COL: path, STATUS_COL: STATUS_DONE}
        # сводка на проверку для дашборда (поля могут отсутствовать в схеме — пишем мягко)
        try:
            db_update(order_id, {
                **result_fields,
                "review_total": review_total,
                "review_yellow": summary["yellow"],
                "review_orange": summary["orange"],
                "review_red": summary["red"],
            })
        except Exception:
            db_update(order_id, result_fields)  # если полей нет в схеме — хотя бы статус

        phase("DONE", ms=int((time.time() - t0) * 1000),
              needs_review=needs_review, mass_check=check)

        log_run(
            drawing_name    = drawing_name,
            request_id      = order_id,
            check_status    = "OK" if check.get("ok") is True
                              else ("MISMATCH" if check.get("ok") is False else "NO_CONTROL"),
            positions_count = len(positions),
            sum_mass        = sum_mass,
            expected_total  = control,
            review          = {"total":  review_total,
                               "yellow": summary["yellow"],
                               "orange": summary["orange"],
                               "red":    summary["red"]},
        )
    except Exception as e:
        tb = traceback.format_exc()
        log.error("[worker] FATAL %s", tb)
        entries.append({"ts": now_iso(), "step": "ERROR", "message": str(e), "traceback": tb[-1500:]})
        db_update(order_id, {LOG_COL: entries, STATUS_COL: "error", ERROR_MSG_COL: str(e)})


# ---------------------------------------------------------------------------
# HTTP
# ---------------------------------------------------------------------------
class ProcessReq(BaseModel):
    order_id: str


@app.get("/health")
def health():
    return {"ok": True, "version": WORKER_VERSION}


@app.post("/process")
def process(req: ProcessReq, background: BackgroundTasks,
            x_worker_secret: str = Header(default="", alias="X-Worker-Secret")):
    if not WORKER_SECRET or x_worker_secret != WORKER_SECRET:
        raise HTTPException(status_code=401, detail="bad secret")
    try:
        order = fetch_order(req.order_id)
    except Exception:
        order = None
    if not order:
        raise HTTPException(status_code=404, detail="order not found")

    entries = order.get(LOG_COL) or []
    entries.append({"ts": now_iso(), "step": "WORKER_START", "orderId": req.order_id})
    db_update(req.order_id, {STATUS_COL: "processing", LOG_COL: entries})

    background.add_task(process_order, req.order_id)
    return JSONResponse(status_code=202, content={"accepted": True})


@app.get("/result/{order_id}")
def get_result(order_id: str):
    """Прямое скачивание готового файла: отдаёт xlsx из бакета results.
    Открой в браузере: https://<worker>/result/<order_id>
    """
    from fastapi.responses import Response
    try:
        order = fetch_order(order_id)
    except Exception:
        order = None
    if not order:
        raise HTTPException(status_code=404, detail="order not found")
    path = order.get(RESULT_PATH_COL)
    if not path:
        raise HTTPException(status_code=404, detail="result not ready")
    path = path.lstrip("/").replace("results/", "", 1) if path.startswith("results/") else path.lstrip("/")
    try:
        data = sb.storage.from_(RESULTS_BUCKET).download(path)
    except Exception as e:
        raise HTTPException(status_code=404, detail=f"file not found: {e}")
    fname = f"{order.get('order_number') or order_id}.xlsx"
    return Response(
        content=data,
        media_type=XLSX_MIME,
        headers={"Content-Disposition": f'attachment; filename="{fname}"'},
    )
