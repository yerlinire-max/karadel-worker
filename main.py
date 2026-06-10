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
WORKER_VERSION = "v4-svodnaya-match-2026-06-09"
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
- Крайний ПРАВЫЙ столбец «Общая масса кг» — ИТОГОВАЯ масса этого размера.
  БЕРИ МАССУ ТОЛЬКО ОТСЮДА.

ЧТО ДЕЛАТЬ (по каждой строке-размеру):
1) Если в столбце «Общая масса кг» стоит число — это позиция. Возьми её.
2) Если «Общая масса» пустая (нет массы) — ПРОПУСТИ строку (не вноси).
3) Тип бери из левого растянутого столбца, размер — из столбца размера,
   массу — из «Общая масса кг», марку — из столбца марки.

ЧТО ПРОПУСКАТЬ ВСЕГДА:
- Строки «Всего профиля» (это подытоги — НЕ нужны).
- Строку «Итого масса металла» (общий итог — НЕ нужна).
- Блок «в том числе по маркам» (С235 / С255 / С345 — НЕ нужен).
- Строки без массы в «Общая масса кг».

ПЕРЕВОД ТИПА В НАЗВАНИЕ (как в нашей базе «Сводная»); это база синонимов,
сопоставляй гибко:
- «Швеллеры …», «Швеллер», «[N»            → "Швеллер <номер>"  (напр. [10 → "Швеллер 10")
- «Прокат листовой …», «Лист», «t<N>»       → "Лист t<толщина>"  (t10 → "Лист t10")
- «Прокат угловой равнополочный …», «Уголок», «L…» → "Уголок <размер>" (L100x10 → "Уголок 100x10")
- «Квадратные трубы …», «Труба», «[]…»      → "Труба проф. <размер>" ([]100x5 → "Труба проф. 100x5")
- «Двутавр …», «Балка», «<N>К<n>», «<N>Б<n>»→ "Балка <обозначение>" (30К1 → "Балка 30К1", 30Б1 → "Балка 30Б1")
ВАЖНО: обозначения вида «30К1», «23К1», «30Б1» — это ДВУТАВРЫ (балки),
а НЕ листы. Никогда не пиши их как «Лист».

ПОЛЯ ПОЗИЦИИ:
- sortament: переведённое название (как в «Сводной»).
- grade: марка стали.
- mass_kg: число из «Общая масса кг» (обязательно).
- qty: если есть, иначе пропусти.

ЗАПРЕЩЕНО: выдумывать профили/размеры/массы, использовать римские цифры,
повторять один и тот же размер дважды, включать позиции без массы.
Если число не разобрать — пропусти позицию.

Самопроверка: сумма всех mass_kg должна примерно совпасть с «Итого масса металла»
из таблицы. Если сильно расходится — перечитай столбец «Общая масса кг».
Дополнительно верни это контрольное число в поле "control_total_kg"
(значение строки «Итого масса металла, кг» из таблицы; если его нет — null).

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
def _build_cell(ref, attrs, value, kind):
    m = re.search(r'\bs="(\d+)"', attrs or "")
    s = f' s="{m.group(1)}"' if m else ""
    if kind == "n":
        return f'<c r="{ref}"{s}><v>{value}</v></c>'
    text = escape(str(value))
    return f'<c r="{ref}"{s} t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>'


def set_cell(xml, ref, value, kind):
    col = re.match(r"[A-Z]+", ref).group()
    row = re.search(r"\d+", ref).group()
    col_idx = col_to_idx(col)

    # 1) ячейка уже есть -> заменить, сохранив её стиль (s=)
    pat = re.compile(r'<c r="' + ref + r'"([^>]*?)(?:/>|>.*?</c>)', re.S)
    if pat.search(xml):
        return pat.sub(lambda m: _build_cell(ref, m.group(1), value, kind), xml, count=1)

    # 2) ячейки нет -> вставить в нужную строку в порядке колонок
    rowpat = re.compile(r'(<row r="' + row + r'"[^>]*>)(.*?)(</row>)', re.S)
    m = rowpat.search(xml)
    new_cell = _build_cell(ref, "", value, kind)
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


def ensure_full_calc(wb):
    if "fullCalcOnLoad" in wb:
        return wb
    m = re.search(r"<calcPr\b[^>]*?/>", wb)
    if m:
        return wb.replace(m.group(0), m.group(0)[:-2] + ' fullCalcOnLoad="1"/>', 1)
    return re.sub(r"(</sheets>)", r'\1<calcPr fullCalcOnLoad="1"/>', wb, count=1)


def read_svodnaya_names(xlsx_bytes):
    """Читает названия материалов из листа «Сводная» (столбец C, кэш-значения формул)."""
    try:
        z = zipfile.ZipFile(io.BytesIO(xlsx_bytes))
        data = z.read("xl/worksheets/sheet6.xml").decode("utf-8")
    except Exception:
        return []
    out = []
    for m in re.finditer(r'<c r="C(\d+)"[^>]*>(?:<f>[^<]*</f>)?<v>([^<]*)</v></c>', data):
        val = m.group(2)
        if val and any(ch.isalpha() for ch in val):
            out.append(val)
    return out


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
        return ("УГОЛОК", nums(s.split("уголок")[1].split(",")[0])[:2], grade, None)
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
            return ("ТРУБА_ПРОФ", ([n[0], n[-1]] if len(n) >= 2 else n), grade, None)
        return ("ТРУБА_КРУГ", n, grade, None)
    return ("?", [], grade, None)


def build_svodnaya_index(names):
    index = {}
    for nm in names:
        p = _parse_profile(nm)
        index.setdefault(p[0], []).append((p, nm))
    return index


def resolve_to_svodnaya(name, grade, index):
    """Возвращает (имя_из_Сводной, способ). способ: 'точно' | 'аналог' | 'нет'."""
    p = _parse_profile(name)
    typ, dims, _, sub = p
    cands = index.get(typ, [])
    if not cands:
        return name, "нет"
    want = "09Г2С" if typ == "ЛИСТ" else ("09Г2С" if grade == "С345" and typ == "УГОЛОК" else None)

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
    if typ == "ТРУБА_ПРОФ" and len(dims) >= 2:
        side, wall = dims
        sw = [(pp[1][0], nm) for pp, nm in cands if len(pp[1]) >= 2 and pp[1][1] == wall and pp[1][0] >= side]
        if sw:
            return min(sw)[1], "аналог"
        any_ = [(pp[1][0], nm) for pp, nm in cands if pp[1] and pp[1][0] >= side]
        if any_:
            return min(any_)[1], "аналог"
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


def build_raschet_cells(positions, drawing_name, svod_index=None):
    cells = {
        "C6": ("Металлоконструкции", "s"),
        "D6": (drawing_name or "", "s"),
        "I6": ("шт", "s"),
        "J6": (1, "n"),
    }
    agg, display, order, notes = {}, {}, [], []
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
        # подбор имени из «Сводной» (точное или ближайший аналог)
        if svod_index:
            resolved, how = resolve_to_svodnaya(raw, grade, svod_index)
        else:
            resolved, how = raw, "нет-индекса"
        if how in ("аналог", "нет"):
            notes.append({"исходное": raw, "подставлено": resolved, "тип": how})
        key = _norm_key(resolved)
        if key not in agg:
            agg[key] = 0.0
            display[key] = resolved
            order.append(key)
        agg[key] = max(agg[key], mass)
    total = 0.0
    for i, key in enumerate(order):
        if i >= len(STEEL_COLUMNS):
            break
        c = STEEL_COLUMNS[i]
        cells[f"{c}5"] = (display[key], "s")
        cells[f"{c}6"] = (round(agg[key], 2), "n")
        total += agg[key]
    cells["K6"] = (round(total, 2), "n")
    return cells, notes


def fill_template(xlsx_bytes, positions, drawing_name):
    zin = zipfile.ZipFile(io.BytesIO(xlsx_bytes))
    parts = {n: zin.read(n) for n in zin.namelist()}

    svod_index = build_svodnaya_index(read_svodnaya_names(xlsx_bytes))

    s1 = parts["xl/worksheets/sheet1.xml"].decode("utf-8")  # Расчет
    cells, notes = build_raschet_cells(positions, drawing_name, svod_index)
    for ref, (val, kind) in cells.items():
        s1 = set_cell(s1, ref, val, kind)
    parts["xl/worksheets/sheet1.xml"] = s1.encode("utf-8")

    wb = parts["xl/workbook.xml"].decode("utf-8")
    parts["xl/workbook.xml"] = ensure_full_calc(wb).encode("utf-8")

    out = io.BytesIO()
    zout = zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED)
    for zi in zin.infolist():
        zout.writestr(zi, parts[zi.filename])
    zout.close()
    return out.getvalue(), notes


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
        filled, notes = fill_template(tpl, positions, drawing_name)
        if notes:
            phase("SVODNAYA_ANALOGS", count=len(notes), items=notes)

        phase("SAVE", bytes=len(filled))
        path = f"{order_id}.xlsx"
        upload_result(path, filled)
        db_update(order_id, {RESULT_PATH_COL: path, STATUS_COL: STATUS_DONE})

        phase("DONE", ms=int((time.time() - t0) * 1000),
              needs_review=needs_review, mass_check=check)
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
