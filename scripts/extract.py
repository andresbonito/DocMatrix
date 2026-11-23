"""Extração determinística de documentos comuns, sem LLM nem execução de conteúdo."""
from __future__ import annotations
import csv
import hashlib
import html
from html.parser import HTMLParser
import json
import re
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET
from .normalize import normalize_value

SUPPORTED = {'.pdf', '.docx', '.txt', '.md', '.csv', '.tsv', '.json', '.xml', '.html', '.htm', '.xlsx'}
NS_W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
NS_S = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
NS_REL = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
NS_PKG = '{http://schemas.openxmlformats.org/package/2006/relationships}'
KEY_VALUE = re.compile(r'^\s*([\wÀ-ÿ][^:=]{1,85}?)\s*[:=]\s*(.{1,2000})\s*$')
DATES = re.compile(r'(?<!\d)(\d{2}/\d{2}/\d{4}|\d{4}-\d{2}-\d{2})(?!\d)')


def _sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open('rb') as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b''):
            digest.update(chunk)
    return digest.hexdigest()


def _decode(path: Path) -> str:
    data = path.read_bytes()
    for enc in ('utf-8-sig', 'utf-16', 'cp1252'):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode('utf-8', errors='replace')


def _add_block(doc: dict, text: str, *, page: int | None = None, location: str = '', kind: str = 'texto') -> None:
    text = str(text).strip()
    if not text:
        return
    doc['blocos'].append({'id': f"{doc['id']}:b{len(doc['blocos'])+1:06d}", 'pagina': page,
                          'localizacao': location, 'tipo': kind, 'texto': text})


def _add_table(doc: dict, rows: list[list[Any]], *, page: int | None = None, location: str = '') -> None:
    if not rows:
        return
    width = max((len(row) for row in rows), default=0)
    if width == 0:
        return
    normalized = [(list(row) + [None] * (width - len(row))) for row in rows]
    raw_head = [str(x).strip() if x is not None else '' for x in normalized[0]]
    # Não inventar cabeçalho se a primeira linha for apenas dados.
    has_header = len(normalized) > 1 and all(h and h not in raw_head[:i] for i,h in enumerate(raw_head)) and any(isinstance(x,str) for x in normalized[0])
    headers = raw_head if has_header else [f'coluna_{i+1}' for i in range(width)]
    data = normalized[1:] if has_header else normalized
    doc['tabelas'].append({'id': f"{doc['id']}:t{len(doc['tabelas'])+1:04d}", 'pagina': page,
                           'localizacao': location, 'cabecalho_identificado': has_header,
                           'colunas': headers, 'linhas': data})


def _text_lines(doc: dict, text: str, location: str = 'corpo') -> None:
    _add_block(doc, text, location=location)
    lines = text.splitlines()
    md_table: list[list[str]] = []
    for line in lines:
        if '|' in line and line.count('|') >= 2:
            cells = [s.strip() for s in line.strip().strip('|').split('|')]
            if all(re.fullmatch(r':?-{3,}:?', c or '') for c in cells):
                continue
            md_table.append(cells)
        elif md_table:
            if len(md_table) >= 2:
                _add_table(doc, md_table, location='tabela markdown')
            md_table = []
    if len(md_table) >= 2:
        _add_table(doc, md_table, location='tabela markdown')


def _parse_csv(doc: dict, path: Path) -> None:
    sample = _decode(path)
    dialect = csv.excel_tab if path.suffix.lower() == '.tsv' else csv.excel
    try:
        dialect = csv.Sniffer().sniff(sample[:8192], delimiters=',;\t|')
    except csv.Error:
        pass
    rows = list(csv.reader(sample.splitlines(), dialect=dialect))
    _add_table(doc, rows, location='linhas CSV/TSV')
    doc['paginas'] = None


def _check_office_zip(z: zipfile.ZipFile) -> None:
    members = z.infolist()
    if len(members) > 10000:
        raise ValueError('Arquivo Office com número excessivo de membros ZIP')
    total = sum(item.file_size for item in members)
    if total > 512*1024*1024:
        raise ValueError('Arquivo Office descompactado acima do limite de segurança de 512 MiB')
    if any(item.file_size > 256*1024*1024 for item in members):
        raise ValueError('Componente Office individual grande demais')


def _parse_docx(doc: dict, path: Path) -> None:
    with zipfile.ZipFile(path) as z:
        _check_office_zip(z)
        root = ET.fromstring(z.read('word/document.xml'))
    body = root.find(f'{NS_W}body')
    if body is None:
        raise ValueError('DOCX sem corpo legível')
    for idx, item in enumerate(body, start=1):
        tag = item.tag.rsplit('}',1)[-1]
        if tag == 'p':
            value = ''.join((n.text or '') for n in item.iter(f'{NS_W}t')).strip()
            _add_block(doc, value, location=f'paragrafo {idx}')
        elif tag == 'tbl':
            rows = []
            for tr in item.findall(f'{NS_W}tr'):
                row = []
                for tc in tr.findall(f'{NS_W}tc'):
                    ps = [ ''.join((n.text or '') for n in p.iter(f'{NS_W}t')) for p in tc.findall(f'.//{NS_W}p')]
                    row.append('\n'.join(x for x in ps if x))
                rows.append(row)
            _add_table(doc, rows, location=f'tabela DOCX {idx}')


def _read_xlsx_cell(cell: ET.Element, shared: list[str]) -> Any:
    ctype = cell.get('t', '')
    if ctype == 'inlineStr':
        return ''.join(t.text or '' for t in cell.iter(f'{NS_S}t'))
    val = cell.find(f'{NS_S}v')
    f = cell.find(f'{NS_S}f')
    if f is not None:
        return '=' + (f.text or '')  # Preserve fórmula, mas jamais executar.
    if val is None or val.text is None:
        return None
    if ctype == 's':
        try: return shared[int(val.text)]
        except (ValueError, IndexError): return val.text
    if ctype == 'b':
        return val.text == '1'
    if ctype in ('str', 'e'):
        return val.text
    try:
        fl = float(val.text)
        return int(fl) if fl.is_integer() else fl
    except ValueError:
        return val.text


def _col_index(ref: str) -> int:
    letters = re.match(r'[A-Z]+', ref or '')
    out = 0
    for ch in (letters.group(0) if letters else 'A'):
        out = out * 26 + ord(ch) - 64
    return out - 1


def _parse_xlsx(doc: dict, path: Path) -> None:
    with zipfile.ZipFile(path) as z:
        _check_office_zip(z)
        names = set(z.namelist())
        shared: list[str] = []
        if 'xl/sharedStrings.xml' in names:
            root = ET.fromstring(z.read('xl/sharedStrings.xml'))
            for si in root.findall(f'{NS_S}si'):
                shared.append(''.join(n.text or '' for n in si.iter(f'{NS_S}t')))
        wb = ET.fromstring(z.read('xl/workbook.xml'))
        rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        relmap = {r.get('Id'):r.get('Target') for r in rels.findall(f'{NS_PKG}Relationship')}
        for sheet in wb.findall(f'.//{NS_S}sheet'):
            title = sheet.get('name', 'Planilha')
            target = relmap.get(sheet.get(f'{NS_REL}id'))
            if not target:
                continue
            if target.startswith('/'):
                filename = target.lstrip('/')
            elif target.startswith('xl/'):
                filename = target
            else:
                filename = 'xl/' + target
            if filename not in names:
                continue
            root = ET.fromstring(z.read(filename))
            rows: list[list[Any]] = []
            # Respeita coordenadas de coluna e células ausentes.
            for row in root.findall(f'.//{NS_S}sheetData/{NS_S}row'):
                cells: list[Any] = []
                for c in row.findall(f'{NS_S}c'):
                    idx = _col_index(c.get('r', 'A1'))
                    if idx > 16383:
                        continue
                    if idx >= len(cells):
                        cells.extend([None]*(idx + 1 - len(cells)))
                    cells[idx] = _read_xlsx_cell(c, shared)
                rows.append(cells)
            _add_table(doc, rows, location=f'aba {title}')


def _parse_pdf(doc: dict, path: Path) -> None:
    try:
        import fitz  # PyMuPDF, opcional.
        with fitz.open(path) as pdf:
            doc['paginas'] = len(pdf)
            for i, page in enumerate(pdf, 1):
                _add_block(doc, page.get_text('text') or '', page=i, location=f'pagina {i}')
                if hasattr(page, 'find_tables'):
                    try:
                        for table in page.find_tables().tables:
                            _add_table(doc, table.extract(), page=i, location=f'tabela pagina {i}')
                    except Exception as exc:
                        doc['avisos'].append(f'Página {i}: extração tabular limitada ({type(exc).__name__}).')
    except ImportError:
        try:
            from pypdf import PdfReader
        except ImportError as exc:
            raise RuntimeError('PDF exige PyMuPDF (recomendado) ou pypdf; veja README.md.') from exc
        reader = PdfReader(str(path))
        doc['paginas'] = len(reader.pages)
        for i, page in enumerate(reader.pages, 1):
            _add_block(doc, page.extract_text() or '', page=i, location=f'pagina {i}')
    if not doc['blocos']:
        doc['avisos'].append('Nenhum texto extraível: possível PDF digitalizado. OCR requer processo autorizado em separado.')


class _HTMLText(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.tokens: list[str] = []
        self.skip = 0
    def handle_starttag(self, tag: str, attrs: list[tuple[str,str|None]]) -> None:
        if tag in ('script', 'style'): self.skip += 1
        if tag in ('p','br','li','tr','h1','h2','h3'): self.tokens.append('\n')
    def handle_endtag(self, tag: str) -> None:
        if tag in ('script','style') and self.skip: self.skip -= 1
    def handle_data(self, data: str) -> None:
        if not self.skip: self.tokens.append(data)


def _parse_json(doc: dict, path: Path) -> None:
    data = json.loads(_decode(path))
    def walk(obj: Any, prefix: str = '$', depth: int = 0) -> None:
        if depth > 30:
            doc['avisos'].append('JSON muito aninhado; profundidade máxima de 30 atingida.')
            return
        if isinstance(obj, dict):
            for k,v in obj.items(): walk(v, f'{prefix}.{k}', depth+1)
        elif isinstance(obj, list):
            if obj and all(isinstance(x, dict) for x in obj):
                keys = list(dict.fromkeys(k for x in obj for k in x.keys()))
                _add_table(doc, [keys] + [[x.get(k) if not isinstance(x.get(k),(dict,list)) else json.dumps(x[k], ensure_ascii=False) for k in keys] for x in obj], location=prefix)
            else:
                for i,v in enumerate(obj): walk(v, f'{prefix}[{i}]', depth+1)
        else:
            _add_block(doc, f'{prefix}: {obj}', location=prefix, kind='campo')
    walk(data)


def _parse_xml(doc: dict, path: Path) -> None:
    raw = path.read_bytes()
    if b'<!DOCTYPE' in raw.upper() or b'<!ENTITY' in raw.upper():
        raise ValueError('XML com DTD/ENTITY bloqueado por segurança')
    root = ET.fromstring(raw)
    def walk(el: ET.Element, prefix: str, depth: int) -> None:
        if depth > 40:
            doc['avisos'].append('XML muito aninhado; profundidade máxima de 40 atingida.')
            return
        tag = el.tag.rsplit('}', 1)[-1]
        loc = f'{prefix}/{tag}'
        text = (el.text or '').strip()
        if text: _add_block(doc, f'{loc}: {text}', location=loc, kind='campo')
        for attr, val in el.attrib.items():
            _add_block(doc, f'{loc}/@{attr}: {val}', location=loc, kind='atributo')
        for child in el: walk(child, loc, depth+1)
    walk(root, '', 0)


def _derive_records(doc: dict) -> None:
    def add_record(b: dict, field: str, raw: str, line_no: int) -> None:
        val, tp = normalize_value(raw)
        doc['registros'].append({'id': f"{doc['id']}:r{len(doc['registros'])+1:06d}",
          'documento_id': doc['id'], 'campo': field.strip(), 'valor_original': raw.strip(),
          'valor_normalizado': val, 'tipo': tp, 'pagina': b['pagina'],
          'localizacao': f"{b['localizacao']}; linha {line_no}", 'bloco_id': b['id'],
          'origem': 'extracao_deterministica'})
    for b in doc['blocos']:
        # JSON/XML têm caminhos de origem como '$.foo' ou '/root/foo',
        # não necessariamente chaves iniciadas por letras.
        is_structured = b['tipo'] in ('campo', 'atributo') and bool(b['localizacao'])
        if is_structured and ': ' in b['texto']:
            _prefix, _, raw = b['texto'].partition(': ')
            add_record(b,b['localizacao'],raw,1)
        # Nas páginas longas, preservar todas as linhas no bloco; só campos claros viram registros.
        for line_no, line in enumerate(b['texto'].splitlines(), 1):
            m = KEY_VALUE.match(line)
            if m and not is_structured:
                add_record(b,m.group(1),m.group(2),line_no)
            # Apenas menções a datas, não inferir evento confirmado.
            for dm in DATES.finditer(line):
                val, kind = normalize_value(dm.group(1))
                if kind == 'data':
                    doc['eventos'].append({'id': f"{doc['id']}:e{len(doc['eventos'])+1:06d}",
                      'documento_id': doc['id'], 'data': val, 'descricao': line.strip()[:300],
                      'tipo': 'mencao_de_data', 'pagina': b['pagina'], 'bloco_id': b['id'],
                      'confirmado': False})


def extract_document(path: str | Path, doc_id: str, *, max_bytes: int = 100*1024*1024) -> dict[str, Any]:
    path = Path(path).expanduser().resolve(strict=True)
    if not path.is_file(): raise ValueError(f'Não é arquivo: {path}')
    ext = path.suffix.lower()
    if ext not in SUPPORTED: raise ValueError(f'Formato não suportado: {ext}')
    if path.stat().st_size > max_bytes:
        raise ValueError(f'Arquivo excede o limite de segurança ({max_bytes} bytes). Aumente explicitamente --max-mb.')
    doc = {'id': doc_id, 'nome': path.name, 'formato': ext.lstrip('.'), 'sha256': _sha(path),
           'tamanho_bytes': path.stat().st_size, 'paginas': None, 'status': 'processado',
           'avisos': [], 'blocos': [], 'tabelas': [], 'registros': [], 'eventos': []}
    try:
        if ext in ('.txt', '.md'):
            _text_lines(doc, _decode(path))
        elif ext in ('.csv', '.tsv'):
            _parse_csv(doc, path)
        elif ext == '.docx':
            _parse_docx(doc, path)
        elif ext == '.xlsx':
            _parse_xlsx(doc, path)
        elif ext == '.pdf':
            _parse_pdf(doc, path)
        elif ext == '.json':
            _parse_json(doc, path)
        elif ext == '.xml':
            _parse_xml(doc, path)
        elif ext in ('.html','.htm'):
            parser = _HTMLText(); parser.feed(_decode(path))
            _text_lines(doc, html.unescape(''.join(parser.tokens)), 'texto HTML')
        _derive_records(doc)
    except Exception as exc:
        doc['status'] = 'erro'
        doc['avisos'].append(f'{type(exc).__name__}: {exc}')
    if doc['status'] == 'processado' and not (doc['blocos'] or doc['tabelas']):
        doc['status'] = 'sem_dados'
        doc['avisos'].append('Não foram obtidos dados estruturáveis. Conferência manual pode ser necessária.')
    return doc
