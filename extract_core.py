# extract_core.py
import re
from pathlib import Path
from typing import Dict, Optional, List, Tuple
import pandas as pd
import pdfplumber

HOURS_INLINE_RX = r"(\d{1,3})\s*h\b"
HOURS_WORD_RX = r"(\d{1,3})\s*horas?\b"

MODALITY_HINTS = {
    "presencial": re.compile(r"\bpresencial\b", re.I),
    "à distância": re.compile(r"\b(a|à)\s*dist[aâ]ncia\b|ead\b|on\s*line|online", re.I),
    "misto": re.compile(r"\bmisto\b|h[ií]brido|blended", re.I),
}

EXCLUDE_PREFIXES = (
    "Saldo de",
    "com carga-horária de",
    "concluiu o curso online",
    "mergulhe em programação",
    "18/01/2025 a 17/02/2025",
)


def normalize_text(txt: str) -> str:
    lines = [re.sub(r"\s+", " ", line).strip() for line in txt.splitlines()]
    return "\n".join([l for l in lines if l])


def extract_text_pages(pdf_path: Path) -> List[str]:
    pages: List[str] = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            pages.append(normalize_text(page.extract_text() or ""))
    return pages


def first_match(patterns: List[re.Pattern], text: str) -> Optional[str]:
    for rx in patterns:
        m = rx.search(text)
        if m:
            return re.sub(r"\s+", " ", m.group(1)).strip()
    return None


# --- helpers para Lotação ---

def _sanitize_lotacao(val: Optional[str]) -> Optional[str]:
    """
    Limpa 'Ramal: 123' e números soltos no final (ex.: 'SECOF 107').
    Normaliza espaços e traços.
    """
    if not val:
        return val
    # remove "Ramal: 123" em qualquer posição
    val = re.sub(r"\bRamal\b\s*:\s*\d+\b", "", val, flags=re.I)
    # remove número solto no final: "SECOF 107" / "SECOF - 107"
    val = re.sub(r"[\s\-–—]*\d+\s*$", "", val)
    # normaliza espaços e traços finais
    val = re.sub(r"\s+", " ", val).strip(" -–—").strip()
    return val or None


def extract_lotacao_from_lines(all_text: str) -> Optional[str]:
    """
    Varre por linhas para achar 'Lotação:' e capturar o valor mesmo se estiver
    na mesma linha ou nas próximas 1–2 linhas. Remove qualquer 'Ramal: ...'.
    """
    lines = [re.sub(r"\s+", " ", l).strip() for l in all_text.splitlines()]
    for i, line in enumerate(lines):
        if re.search(r"\bLota[çc][aã]o\s*:", line, re.I):
            # trecho após 'Lotação:' na MESMA linha
            after = re.split(r"\bLota[çc][aã]o\s*:\s*", line, maxsplit=1, flags=re.I)[1].strip()
            # corta 'Ramal: ...' se vier colado
            after = re.split(r"\bRamal\b\s*:\s*", after, maxsplit=1, flags=re.I)[0].strip(" -–—").strip()
            if after and not re.match(r"^Ramal\b", after, re.I):
                return _sanitize_lotacao(after)

            # senão, tenta nas próximas duas linhas
            for j in (1, 2):
                if i + j < len(lines):
                    cand = re.sub(r"^\s*(Ramal\s*:\s*)?", "", lines[i + j], flags=re.I).strip()
                    if not cand or re.match(r"^Ramal\b", cand, re.I):
                        continue
                    # exige pelo menos uma letra para evitar capturar só números
                    if re.search(r"[A-Za-zÁÉÍÓÚÃÕÇ]", cand):
                        return _sanitize_lotacao(cand)
    return None



def extract_header(all_text: str) -> Dict[str, Optional[str]]:
    patterns = {
        "requerente": [
            re.compile(r"\brequerente\s*:\s*(.+)", re.I),
            re.compile(r"\bnome\s*:\s*(.+)", re.I),
        ],
        "matricula": [re.compile(r"\bmatr[ií]cula\s*:\s*([^\n]+)", re.I)],
        "cargo":     [re.compile(r"\bcargo\s*:\s*([^\n]+)", re.I)],
        "lotacao":   [re.compile(r"\blota[çc][aã]o\s*:\s*([^\n]+)", re.I)],  # base (pode vir suja)
    }
    out = {k: first_match(v, all_text) for k, v in patterns.items()}

    # Fallback nome + matrícula (MAIÚSCULAS + números)
    if not out["requerente"] or not out["matricula"]:
        m = re.search(r"\n([A-ZÁÉÍÓÚÃÕÇ ]{3,})\s+(\d{3,})\b", "\n" + all_text)
        if m:
            out["requerente"] = out["requerente"] or m.group(1).strip()
            out["matricula"]  = out["matricula"]  or m.group(2).strip()

    # Limpezas básicas
    if out.get("matricula"):
        out["matricula"] = re.sub(r"^\s*Matr[ií]cula:\s*", "", out["matricula"], flags=re.I).strip()
    if out.get("cargo"):
        out["cargo"] = re.sub(r"^\s*Cargo:\s*", "", out["cargo"], flags=re.I).strip()
    if out.get("requerente"):
        out["requerente"] = re.sub(r"^\s*(Requerente|Matr[ií]cula)\s*:\s*", "", out["requerente"], flags=re.I).strip()

    # -------- Lotação robusta --------
    lot_by_lines = extract_lotacao_from_lines(all_text)  # leitura por linhas (mais confiável)

    # saneia a captura "crua" (uma linha) se existir
    lot_raw = out.get("lotacao") or ""
    lot_raw = re.split(r"\bRamal\b\s*:\s*", lot_raw, maxsplit=1, flags=re.I)[0].strip(" -–—").strip()
    if re.match(r"^Ramal\b", lot_raw, re.I):
        lot_raw = ""

    out["lotacao"] = lot_by_lines or _sanitize_lotacao(lot_raw)

    return out



def hours_in_line(line: str) -> Optional[int]:
    m = re.search(HOURS_INLINE_RX, line)
    if m:
        return int(m.group(1))
    m = re.search(HOURS_WORD_RX, line, re.I)
    if m:
        return int(m.group(1))
    return None


def collect_lines_with_y(page) -> List[Tuple[str, float]]:
    words = page.extract_words(use_text_flow=True, keep_blank_chars=False)
    if not words:
        return []
    rows: Dict[int, List[dict]] = {}
    for w in words:
        key = int(round(w["top"]))
        rows.setdefault(key, []).append(w)
    lines: List[Tuple[str, float]] = []
    for key in sorted(rows.keys()):
        ws = sorted(rows[key], key=lambda x: x["x0"])
        text = " ".join(w["text"] for w in ws).strip()
        if not text:
            continue
        y_mid = sum(w["top"] + (w["bottom"] - w["top"]) / 2 for w in ws) / len(ws)
        lines.append((text, y_mid))
    return lines


def find_course_rows_with_y(
    page, y_range: Tuple[float, float]
) -> List[Tuple[str, int, float]]:
    y_min, y_max = y_range
    lines = [(t, y) for (t, y) in collect_lines_with_y(page) if y_min <= y <= y_max]
    raw = []
    for text, y in lines:
        m = re.search(HOURS_INLINE_RX, text) or re.search(HOURS_WORD_RX, text, re.I)
        if not m:
            continue
        if text.startswith(EXCLUDE_PREFIXES):
            continue
        title = text[: m.start()].strip(" -–—;:,")
        if len(title) < 6:
            continue
        hrs = int(m.group(1))
        raw.append((title, hrs, y))
    # dedup
    seen, out = set(), []
    for title, hrs, y in raw:
        key = (int(round(y)), re.sub(r"\s+", " ", title.lower()), hrs)
        if key in seen:
            continue
        seen.add(key)
        out.append((title, hrs, y))
    return out


def detect_checkbox_modality_by_coords(
    page, y_mid: float, x_cols: Dict[str, Tuple[float, float]], y_tol: int
) -> Optional[str]:
    y0, y1 = y_mid - y_tol, y_mid + y_tol
    marked = {label: False for label in x_cols.keys()}

    for ch in page.chars:
        cy0, cy1 = ch.get("top", 0), ch.get("bottom", 0)
        cx0, cx1 = ch.get("x0", 0), ch.get("x1", 0)
        if cy1 < y0 or cy0 > y1:
            continue
        for label, (x0c, x1c) in x_cols.items():
            if cx0 >= x0c and cx1 <= x1c:
                marked[label] = True

    for ln in getattr(page, "lines", []):
        cx = (ln.get("x0", 0) + ln.get("x1", 0)) / 2
        cy = (ln.get("y0", 0) + ln.get("y1", 0)) / 2
        if cy < y0 or cy > y1:
            continue
        for label, (x0c, x1c) in x_cols.items():
            if x0c <= cx <= x1c:
                marked[label] = True

    for r in page.rects:
        rx0, ry0, rx1, ry1 = (
            r.get("x0", 0),
            r.get("top", 0),
            r.get("x1", 0),
            r.get("bottom", 0),
        )
        if ry1 < y0 or ry0 > y1:
            continue
        if (rx1 - rx0) < 20 and (ry1 - ry0) < 20:
            for label, (x0c, x1c) in x_cols.items():
                if not (rx1 < x0c or rx0 > x1c):
                    marked[label] = True

    for label in ["presencial", "misto", "à distância"]:
        if marked[label]:
            return label
    return None


def process_pdf(
    pdf_path: Path,
    course_pages: List[int],
    course_y_range: Tuple[float, float],
    checkbox_columns: Dict[str, Tuple[float, float]],
    y_tolerance: int,
    export_annotations: bool = False,
    annotations_dir: Optional[Path] = None,
) -> List[Dict[str, Optional[str]]]:

    rows: List[Dict[str, Optional[str]]] = []
    pages_text = extract_text_pages(pdf_path)
    header = extract_header("\n".join(pages_text))

    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages, start=1):
            if course_pages and (page_idx not in course_pages):
                continue

            course_rows = find_course_rows_with_y(page, course_y_range)

            if export_annotations and annotations_dir:
                try:
                    annotations_dir.mkdir(exist_ok=True)
                    img = page.to_image(resolution=160)
                    img.draw_rect(
                        (0, course_y_range[0], page.width, course_y_range[1]),
                        stroke="blue",
                        fill=None,
                    )
                    for x0, x1 in checkbox_columns.values():
                        img.draw_rect((x0, 0, x1, page.height), stroke="red", fill=None)
                    for _, _, y in course_rows:
                        img.draw_line((0, y, page.width, y), stroke="green")
                    img.save(str(annotations_dir / f"{pdf_path.stem}_p{page_idx}.png"))
                except Exception:
                    pass

            for title, hours, y in course_rows:
                modality = detect_checkbox_modality_by_coords(
                    page, y, checkbox_columns, y_tolerance
                )
                if modality is None:
                    if MODALITY_HINTS["presencial"].search(title):
                        modality = "presencial"
                    elif MODALITY_HINTS["misto"].search(title):
                        modality = "misto"
                    elif MODALITY_HINTS["à distância"].search(title):
                        modality = "à distância"

                # REQUERENTE (nome + matrícula)
                nome = (header.get("requerente") or "").strip()
                matr = (header.get("matricula") or "").strip()
                if nome and matr:
                    requerente_display = f"{nome} {matr}"
                elif nome:
                    requerente_display = nome
                else:
                    requerente_display = matr  # fallback

                rows.append(
                    {
                        "arquivo": pdf_path.name,
                        "pagina": page_idx,
                        "requerente": requerente_display,
                        "matricula": header.get("matricula"),
                        "cargo": header.get("cargo"),
                        "lotacao": header.get("lotacao"),
                        "curso_titulo": title,
                        "curso_horas": hours,
                        "modalidade": modality,
                    }
                )
    return rows


def run_batch(
    input_dir: Path,
    output_xlsx: Path,
    course_pages: List[int],
    course_y_range: Tuple[float, float],
    checkbox_columns: Dict[str, Tuple[float, float]],
    y_tolerance: int,
    export_annotations: bool = False,
    annotations_dir: Optional[Path] = None,
) -> pd.DataFrame:

    all_rows: List[Dict[str, Optional[str]]] = []
    pdfs = sorted(input_dir.glob("*.pdf"))
    for pdf in pdfs:
        try:
            all_rows.extend(
                process_pdf(
                    pdf,
                    course_pages,
                    course_y_range,
                    checkbox_columns,
                    y_tolerance,
                    export_annotations,
                    annotations_dir,
                )
            )
        except Exception as e:
            all_rows.append(
                {
                    "arquivo": pdf.name,
                    "pagina": None,
                    "requerente": None,
                    "matricula": None,
                    "cargo": None,
                    "lotacao": None,
                    "curso_titulo": None,
                    "curso_horas": None,
                    "modalidade": None,
                    "_erro": str(e),
                }
            )

    df = pd.DataFrame(
        all_rows,
        columns=[
            "arquivo",
            "pagina",
            "requerente",
            "matricula",
            "cargo",
            "lotacao",
            "curso_titulo",
            "curso_horas",
            "modalidade",
            "_erro",
        ],
    )
    df.to_excel(output_xlsx, index=False)
    return df
