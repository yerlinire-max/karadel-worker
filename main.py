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
Твоя задача — извлечь перечень стального проката из ТАБЛИЦ-СПЕЦИФИКАЦИЙ,
а НЕ из видов, схем и эскизов.

ГДЕ ИСКАТЬ (в порядке приоритета):
1) Таблица «Техническая спецификация металла» (обычно лист 1–2) — сводная
   ведомость: профиль, марка стали, размер, общая масса.
2) Таблицы «Спецификация элементов» (на разных листах) — столбцы:
   Поз., Обозначение, Наименование, Кол., Масса ед., Масса (всего).
3) «Ведомость элементов» — марки и сечения.

ЧТО ИЗВЛЕЧЬ для каждой строки спецификации:
- sortament: тип профиля + размер, КОРОТКО, без слов «ГОСТ», «ТУ» и номеров стандартов.
  Примеры: "Швеллер 10П", "Уголок L100x10", "Уголок L75x5", "Двутавр 30Б1",
  "Труба квадратная 100х100х5".
- grade: марка стали (С235, С255, С345, Ст3 и т.п.).
- mass_kg: ОБЩАЯ масса этой позиции в кг (число; бери из столбца общей массы).
- qty: количество, если указано (иначе пропусти).

ЛИСТОВОЙ ПРОКАТ — ОСОБОЕ ПРАВИЛО:
- Листы считаются ПО ТОЛЩИНЕ, а не по марке. На КАЖДУЮ толщину — ОТДЕЛЬНАЯ позиция
  со своей массой: "Лист t8", "Лист t10", "Лист t14", "Лист t20" и т.д.
- НЕ объединяй все листы в одну строку по марке (НЕ "Лист С255").

СТРОГИЕ ПРАВИЛА:
- Бери ТОЛЬКО то, что реально видно в таблицах-спецификациях. НИЧЕГО НЕ ВЫДУМЫВАЙ.
- У КАЖДОЙ позиции ОБЯЗАТЕЛЬНА масса (mass_kg). Нет массы — НЕ включай позицию.
- НЕ повторяй одну и ту же позицию. Каждый профиль/толщина — РОВНО ОДИН раз.
  Если профиль встречается в нескольких таблицах — это та же позиция, верни её один раз
  с итоговой массой из сводной «Техническая спецификация металла».
- Если массу или название не разобрать — пропусти позицию, НЕ угадывай.
- «Формат A4», «Формат A3» — это формат листа, НЕ номер чертежа и НЕ профиль. Игнорируй.
- Римских цифр в марках профиля не бывает — не используй их.

Верни СТРОГО JSON без пояснений и markdown:
{
  "drawing": "обозначение/марка из штампа, если видно",
  "positions": [
    {"sortament": "Швеллер 10П", "grade": "С255", "mass_kg": 5803.68, "qty": 1}
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


def match_sortament(name):
    # TODO: здесь будет подбор по «Сводной» + аналоги + раскраска (отдельные правила).
    # Пока что имя проходит как есть.
    return (name or "").strip()


def _norm_key(name):
    # ключ для дедупа: убрать "ГОСТ ...", лишние пробелы, привести к нижнему регистру
    s = re.sub(r"\s*ГОСТ[\s\S]*$", "", name, flags=re.I)
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s


def build_raschet_cells(positions, drawing_name):
    cells = {
        "C6": ("Металлоконструкции", "s"),
        "D6": (drawing_name or "", "s"),
        "I6": ("шт", "s"),
        "J6": (1, "n"),
    }
    agg, display, order = {}, {}, []
    for p in positions:
        raw = match_sortament(p.get("sortament") or p.get("name") or "")
        if not raw:
            continue
        # масса обязательна: без массы — это мусор/галлюцинация, пропускаем
        try:
            mass = float(p.get("mass_kg") or p.get("mass") or 0)
        except (TypeError, ValueError):
            mass = 0.0
        if mass <= 0:
            continue
        key = _norm_key(raw)
        if key not in agg:
            agg[key] = 0.0
            display[key] = raw          # показываем первое встреченное имя
            order.append(key)
        # один профиль из разных таблиц не суммируем (избегаем двойного счёта):
        # берём максимум — это и есть итог из сводной спецификации
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
    return cells


def fill_template(xlsx_bytes, positions, drawing_name):
    zin = zipfile.ZipFile(io.BytesIO(xlsx_bytes))
    parts = {n: zin.read(n) for n in zin.namelist()}

    s1 = parts["xl/worksheets/sheet1.xml"].decode("utf-8")  # Расчет
    for ref, (val, kind) in build_raschet_cells(positions, drawing_name).items():
        s1 = set_cell(s1, ref, val, kind)
    parts["xl/worksheets/sheet1.xml"] = s1.encode("utf-8")

    # TODO: правка sheet6.xml (Сводная) — добавление недостающих сортаментов.
    # parts["xl/worksheets/sheet6.xml"] правится тем же set_cell.

    wb = parts["xl/workbook.xml"].decode("utf-8")
    parts["xl/workbook.xml"] = ensure_full_calc(wb).encode("utf-8")

    out = io.BytesIO()
    zout = zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED)
    for zi in zin.infolist():
        zout.writestr(zi, parts[zi.filename])
    zout.close()
    return out.getvalue()


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

        phase("FILL_EXCEL")
        drawing_name = data.get("drawing") or order.get(DRAWING_PATH_COL, "")
        filled = fill_template(tpl, positions, drawing_name)

        phase("SAVE", bytes=len(filled))
        path = f"{order_id}.xlsx"
        upload_result(path, filled)
        db_update(order_id, {RESULT_PATH_COL: path, STATUS_COL: STATUS_DONE})

        phase("DONE", ms=int((time.time() - t0) * 1000))
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
    return {"ok": True}


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
