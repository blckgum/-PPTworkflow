"""严格解析制表符内容，保留重复字段与正文顺序。"""
from dataclasses import dataclass
from pathlib import Path
import math
import re

from .strings import tr
from .theme import THEMES

COURSE_KEYS = {"title", "subtitle", "org", "docid", "deckname", "theme",
               "time_target_min", "time_target_max"}
REPEAT_KEYS = {"bullet", "body", "head", "kv", "row", "card", "step", "num",
               "image", "image2", "caption", "caption2", "column", "reason", "source"}
COMMON_KEYS = {"layout", "title", "subtitle", "key", "footer", "time", "note",
               "image", "image2", "caption", "caption2"}
LAYOUT_KEYS = {
    "cover": {"subtitle2", "question"}, "statement": set(),
    "points": {"bullet", "body"}, "blocks": {"head", "body"},
    "columns": {"column"}, "table": {"row"}, "kv": {"kv"},
    "chain": {"step"}, "flow": {"step"}, "diagram": {"body"},
    "grid": {"card"}, "dark": {"statement", "reason"}, "summary": {"num"},
    "closing": {"title2"}, "sources": {"source"}, "qa": {"body"},
    "section": set(),
}
ALL_KEYS = COMMON_KEYS | set().union(*LAYOUT_KEYS.values())


class WorkflowError(Exception):
    def __init__(self, code, message, page=None, field=None):
        self.code, self.page, self.field = code, page, field
        location = tr("page", page=page) + " " if page else ""
        if field:
            location += tr("field", field=field) + " "
        super().__init__(f"{code} {location}{message}")


@dataclass
class Page:
    number: int
    entries: list[tuple[str, str]]
    notes: str

    def all(self, key):
        return [value for name, value in self.entries if name == key]

    def get(self, key, default=""):
        values = self.all(key)
        return values[0] if values else default

    @property
    def layout(self):
        return self.get("layout")

    @property
    def seconds(self):
        return int(self.get("time", "90"))

    @property
    def images(self):
        result = []
        for key in ("image", "image2"):
            captions = self.all("caption" if key == "image" else "caption2")
            result.extend((value, captions[i] if i < len(captions) else "")
                          for i, value in enumerate(self.all(key)))
        return result


@dataclass
class Course:
    root: Path
    meta: dict[str, str]
    pages: list[Page]

    @property
    def assets(self):
        return self.root / tr("assets_dir")

    @property
    def output(self):
        return self.root / tr("output_dir")

    @property
    def basename(self):
        name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", self.meta["title"]).rstrip(" .")
        # 留出课程目录、讲稿后缀与平台路径长度的空间。
        return name[:70] or "course"

    @property
    def pptx_path(self):
        return self.output / (self.basename + ".pptx")

    @property
    def docx_path(self):
        return self.output / (self.basename + tr("notes_suffix") + ".docx")

    def image_path(self, name, page=None):
        path = self.assets / name
        if Path(name).is_absolute() or not path.resolve().is_relative_to(self.assets.resolve()):
            raise WorkflowError("V8", tr("outside_assets", path=path), page, "image")
        return path


def read_text(path):
    try:
        return path.read_text(encoding="utf-8-sig")
    except (OSError, UnicodeError) as exc:
        raise WorkflowError("INPUT", tr("read_error", path=path, error=exc)) from exc


def parse_line(line, path, lineno, allowed, page=None):
    if "\t" not in line:
        raise WorkflowError("INPUT", tr("tab_required", path=path, line=lineno), page)
    key, value = line.split("\t", 1)
    key, value = key.strip(), value.strip()
    if key not in allowed:
        raise WorkflowError("INPUT", tr("unknown_key", path=path, line=lineno, key=key), page, key)
    if not value:
        raise WorkflowError("INPUT", tr("empty_value", path=path, line=lineno), page, key)
    if any(ord(char) < 32 and char != "\t" for char in value):
        raise WorkflowError("INPUT", tr("control_char"), page, key)
    return key, value


def load_course(root):
    root = Path(root).resolve()
    meta = {}
    for lineno, line in enumerate(read_text(root / "course.txt").splitlines(), 1):
        if not line.strip() or line.lstrip().startswith("#"):
            continue
        key, value = parse_line(line, root / "course.txt", lineno, COURSE_KEYS)
        if key in meta:
            raise WorkflowError("INPUT", tr("duplicate_key", key=key), field=key)
        meta[key] = value
    for key in ("title", "org", "time_target_min", "time_target_max"):
        if key not in meta:
            raise WorkflowError("INPUT", tr("required"), field=key)
    try:
        minimum, maximum = (float(meta[k]) for k in ("time_target_min", "time_target_max"))
        if not (math.isfinite(minimum) and math.isfinite(maximum) and 0 < minimum <= maximum):
            raise ValueError
    except ValueError as exc:
        raise WorkflowError("INPUT", tr("time_range"), field="time_target_min/max") from exc
    meta.setdefault("theme", "military_light")
    meta.setdefault("subtitle", "")
    meta.setdefault("org", "")
    meta.setdefault("docid", root.name)
    meta.setdefault("deckname", root.name)
    if meta["theme"] not in THEMES:
        raise WorkflowError("INPUT", tr("unknown_theme", name=meta["theme"]), field="theme")
    pages = []
    entries = None
    for lineno, line in enumerate(read_text(root / "deck.txt").splitlines(), 1):
        if not line.strip() or line.lstrip().startswith("#"):
            continue
        match = re.fullmatch(r"\[(\d{2})\]", line.strip())
        if match:
            number = int(match[1])
            if number != len(pages) + 1:
                raise WorkflowError("INPUT", tr("page_sequence", line=lineno), number)
            entries = []
            note_path = root / "notes" / f"{number:02}.txt"
            if not note_path.resolve().is_relative_to(root / "notes"):
                raise WorkflowError("V3", tr("read_error", path=note_path, error="path"), number)
            notes = read_text(note_path) if note_path.is_file() else ""
            pages.append(Page(number, entries, notes))
        else:
            if entries is None:
                raise WorkflowError("INPUT", tr("page_required", line=lineno))
            key, value = parse_line(line, root / "deck.txt", lineno, ALL_KEYS, len(pages))
            if key not in REPEAT_KEYS and any(name == key for name, _ in entries):
                raise WorkflowError("INPUT", tr("duplicate_key", key=key), len(pages), key)
            entries.append((key, value))
    if not pages:
        raise WorkflowError("INPUT", tr("empty_deck"))
    course = Course(root, meta, pages)
    for page in pages:
        validate_page(page)
    if pages[0].layout != "cover":
        raise WorkflowError("V5", tr("cover_first"), 1, "layout")
    maximum_pages = THEMES[meta["theme"]]["geometry"]["max_pages"]
    if len(pages) > maximum_pages:
        raise WorkflowError("INPUT", tr("page_limit", count=maximum_pages))
    return course


def validate_page(page):
    if page.layout not in LAYOUT_KEYS:
        raise WorkflowError("INPUT", tr("unknown_layout", name=page.layout), page.number, "layout")
    for key in ("title", "layout"):
        if not page.get(key):
            raise WorkflowError("INPUT", tr("required"), page.number, key)
    for key, value in page.entries:
        if key not in COMMON_KEYS | LAYOUT_KEYS[page.layout]:
            raise WorkflowError("INPUT", tr("layout_field", layout=page.layout), page.number, key)
    try:
        if page.seconds <= 0:
            raise ValueError
    except ValueError as exc:
        raise WorkflowError("INPUT", tr("positive_seconds"), page.number, "time") from exc
    for field, delimiter in (("column", "|"), ("card", "|"), ("kv", "=")):
        for value in page.all(field):
            parts = value.split(delimiter, 1)
            if len(parts) != 2 or not all(part.strip() for part in parts):
                raise WorkflowError("INPUT", tr("pair_format", delimiter=delimiter), page.number, field)
    if page.layout == "columns" and not 2 <= len(page.all("column")) <= 3:
        raise WorkflowError("INPUT", tr("column_count"), page.number, "column")
    if page.layout == "blocks":
        sequence = [key for key, _ in page.entries if key in {"head", "body"}]
        if not sequence or sequence != ["head", "body"] * (len(sequence) // 2):
            raise WorkflowError("INPUT", tr("head_body_pairs"), page.number, "head/body")
    if page.layout == "table":
        rows = page.all("row")
        if len(rows) < 2 or len(rows[0].split("|")) < 2 or len({len(row.split("|")) for row in rows}) != 1:
            raise WorkflowError("INPUT", tr("table_shape"), page.number, "row")
    required = {"cover": "image", "points": "bullet", "table": "row", "kv": "kv",
                "chain": "step", "flow": "step", "diagram": "image", "grid": "card",
                "dark": "statement", "summary": "num", "sources": "source",
                "closing": "title2"}
    field = required.get(page.layout)
    if field and not page.get(field):
        raise WorkflowError("INPUT", tr("required"), page.number, field)
    if page.layout == "cover" and len(page.images) != 1:
        raise WorkflowError("INPUT", tr("cover_one_image"), page.number, "image")
    for image_key, caption_key in (("image", "caption"), ("image2", "caption2")):
        if len(page.all(caption_key)) > len(page.all(image_key)):
            raise WorkflowError("INPUT", tr("caption_count"), page.number, caption_key)
