"""Exportação de Excel multiplataforma: artifact_tool ou XlsxWriter.

No backend artifact_tool escrevemos matrizes por intervalos; no XlsxWriter,
usamos write_row sequencial e constant_memory para estabilidade de memória.
Nunca executamos fórmulas vindas de documentos.
"""
from __future__ import annotations

from collections import Counter
import json
from pathlib import Path
from typing import Any

from .normalize import excel_safe, normalize_value

MAX_ROWS = 1_048_576
MAX_COLS = 16_384
CHUNK_ROWS = 1500
HEADER_COLOR = '#16324F'
HEADER_STYLE = {
    'fill': HEADER_COLOR,
    'font': {'color': '#FFFFFF', 'bold': True},
    'row_height': 24,
    'vertical_alignment': 'center',
}
SUPPORTED_BACKENDS = ('auto', 'artifact_tool', 'xlsxwriter')


def _simple_value(value: Any) -> Any:
    if value is None or isinstance(value, (float, int, bool)):
        return value
    if isinstance(value, str):
        return excel_safe(value[:32767])
    return excel_safe(json.dumps(value, ensure_ascii=False, default=str)[:32767])


def _col_letter(n: int) -> str:
    result = ''
    while n:
        n, r = divmod(n - 1, 26)
        result = chr(65 + r) + result
    return result


def _sheet_name(name: str) -> str:
    for bad in '[]:*?/\\':
        name = name.replace(bad, '_')
    return name[:31] or 'Dados'


def _unique_name(candidate: str, existing: set[str]) -> str:
    name = _sheet_name(candidate)
    if name.casefold() not in existing:
        existing.add(name.casefold())
        return name
    for i in range(2, 10000):
        name = _sheet_name(f'{candidate[:26]}_{i}')
        if name.casefold() not in existing:
            existing.add(name.casefold())
            return name
    raise ValueError('Muitas abas com o mesmo nome')


def _select_backend(requested: str) -> str:
    if requested not in SUPPORTED_BACKENDS:
        raise ValueError(f'Backend inválido: {requested}. Escolha: {SUPPORTED_BACKENDS}')
    if requested in ('auto', 'artifact_tool'):
        try:
            import artifact_tool  # noqa: F401
            return 'artifact_tool'
        except ImportError as exc:
            if requested == 'artifact_tool':
                raise RuntimeError('Backend artifact_tool solicitado, mas a biblioteca não está disponível.') from exc
    try:
        import xlsxwriter  # noqa: F401
        return 'xlsxwriter'
    except ImportError as exc:
        raise RuntimeError(
            'Nenhum backend Excel disponível. Instale XlsxWriter com: '
            'python -m pip install -r requirements.txt'
        ) from exc


class ExcelWriter:
    """Encapsula seleção de biblioteca; dados e abas são comuns aos backends."""

    def __init__(self, output: str | Path, backend: str = 'auto') -> None:
        self.output = Path(output)
        self.output.parent.mkdir(parents=True, exist_ok=True)
        self.backend = _select_backend(backend)
        self.counts: dict[str, int] = {}
        self.names: set[str] = set()
        if self.backend == 'artifact_tool':
            from artifact_tool import Workbook
            self.workbook = Workbook.create()
        else:
            import xlsxwriter
            self.workbook = xlsxwriter.Workbook(str(self.output), {
                'constant_memory': True,
                'strings_to_formulas': False,
                'strings_to_urls': False,
                'nan_inf_to_errors': True,
            })
            self.header_format = self.workbook.add_format({
                'bold': True, 'font_color': '#FFFFFF', 'bg_color': HEADER_COLOR,
                'valign': 'vcenter', 'text_wrap': True,
            })
            self.default_format = self.workbook.add_format({'font_name': 'Aptos', 'font_size': 10})

    def write_tab(self, name: str, headers: list[str], rows: list[list[Any]]) -> None:
        headers = headers or ['informacao']
        if len(headers) > MAX_COLS:
            raise ValueError(f'Excesso de colunas ({len(headers)}) na aba {name}')
        # Uma linha do Excel precisa ser reservada ao cabeçalho.
        part_size = MAX_ROWS - 1
        parts = max(1, (len(rows) + part_size - 1) // part_size)
        for part in range(parts):
            subset = rows[part * part_size : (part + 1) * part_size]
            title = _unique_name(name if parts == 1 else f'{name}_{part+1}', self.names)
            if self.backend == 'artifact_tool':
                self._artifact_sheet(title, headers, subset)
            else:
                self._xlsxwriter_sheet(title, headers, subset)
            self.counts[title] = len(subset)

    def _artifact_sheet(self, title: str, headers: list[str], rows: list[list[Any]]) -> None:
        sheet = self.workbook.worksheets.add(title)
        ncols = len(headers)
        last_col = _col_letter(ncols)
        header = sheet.get_range(f'A1:{last_col}1')
        header.values = [[_simple_value(h) for h in headers]]
        header.format = HEADER_STYLE
        sheet.freeze_panes.freeze_rows(1)
        sheet.get_range(f'A:{last_col}').format.column_width = 22
        for start in range(0, len(rows), CHUNK_ROWS):
            batch = rows[start:start + CHUNK_ROWS]
            matrix = [[_simple_value(row[i]) if i < len(row) else None for i in range(ncols)]
                      for row in batch]
            first = start + 2
            sheet.get_range(f'A{first}:{last_col}{first + len(matrix) - 1}').values = matrix
        if rows and len(rows) <= 5000 and ncols <= 40:
            try:
                sheet.tables.add(f'A1:{last_col}{len(rows)+1}', True, f'DocM_{len(self.names)}')
            except Exception:
                # Tabelas nativas são opcionais; os valores já foram escritos.
                pass

    def _xlsxwriter_sheet(self, title: str, headers: list[str], rows: list[list[Any]]) -> None:
        sheet = self.workbook.add_worksheet(title)
        ncols = len(headers)
        sheet.set_column(0, ncols - 1, 22)
        sheet.set_row(0, 25)
        sheet.freeze_panes(1, 0)
        # Autofiltrar é suportado em constant_memory; add_table() não é.
        if rows:
            sheet.autofilter(0, 0, len(rows), ncols - 1)
        sheet.write_row(0, 0, [_simple_value(h) for h in headers], self.header_format)
        for rownum, row in enumerate(rows, 1):
            values = [_simple_value(row[i]) if i < len(row) else None for i in range(ncols)]
            sheet.write_row(rownum, 0, values, self.default_format)

    def close(self) -> None:
        if self.backend == 'artifact_tool':
            from artifact_tool import SpreadsheetFile
            SpreadsheetFile.export_xlsx(self.workbook).save(str(self.output))
        else:
            self.workbook.close()


def export_excel(payload: dict[str, Any], output: str | Path,
                 backend: str = 'auto') -> tuple[dict[str, int], str]:
    """Gera workbook e retorna (quantidade de linhas por aba, backend utilizado)."""
    builder = ExcelWriter(output, backend)
    docs = payload.get('documentos', [])
    stats = Counter()
    for d in docs:
        for key in ('blocos', 'tabelas', 'registros', 'eventos'):
            stats[key] += len(d.get(key, []))
    overview = [
        ['Versão do schema', payload.get('schema_version')],
        ['Skill', 'DocMatrix'],
        ['Modo', payload.get('execucao', {}).get('modo')],
        ['Foco', payload.get('execucao', {}).get('foco')],
        ['Documentos', len(docs)],
        ['Blocos de texto', stats['blocos']],
        ['Tabelas identificadas', stats['tabelas']],
        ['Campos extraídos', stats['registros']],
        ['Menções de datas', stats['eventos']],
        ['Validações', len(payload.get('validacoes', []))],
        ['Inconsistências', len(payload.get('inconsistencias', []))],
        ['Backend Excel', builder.backend],
        ['Observação', 'Dados inferidos e dados originais são identificados separadamente.'],
    ]
    builder.write_tab('Visão Geral', ['Indicador', 'Valor'], overview)
    records = []
    evidences = []
    for d in docs:
        for r in d.get('registros', []):
            records.append([r.get('id'), d.get('nome'), r.get('campo'), r.get('valor_original'),
                            r.get('valor_normalizado'), r.get('tipo'), r.get('origem')])
            evidences.append([r.get('id'), d.get('nome'), r.get('pagina'), r.get('localizacao'),
                              r.get('bloco_id'), 'campo_extraido'])
    for r in payload.get('anotacoes', {}).get('registros', []):
        records.append([r.get('id'), '[anotacao IA]', r.get('campo'), r.get('valor_original'),
                        r.get('valor_normalizado'), r.get('tipo'), 'interpretacao_ia'])
        evidences.append([r.get('id'), r.get('documento_id'), r.get('pagina'), r.get('localizacao'),
                          r.get('bloco_id'), 'interpretacao_ia'])
    # Quando não houver campos reconhecíveis, os blocos substituem linhas artificiais.
    if not records:
        for d in docs:
            for b in d.get('blocos', []):
                records.append([b['id'], d['nome'], 'trecho_documental', b['texto'][:300],
                                None, 'texto', 'trecho'])
                evidences.append([b['id'], d['nome'], b.get('pagina'), b.get('localizacao'),
                                  b['id'], 'trecho'])
    builder.write_tab('Dados Extraídos',
                      ['ID', 'Documento', 'Campo', 'Valor original', 'Valor normalizado', 'Tipo', 'Origem'],
                      records)
    builder.write_tab('Evidências',
                      ['Registro', 'Documento', 'Página', 'Localização', 'Bloco', 'Tipo'], evidences)
    validations = []
    for v in payload.get('validacoes', []):
        validations.append([v.get('id'), v.get('documento_id'), v.get('tipo'), v.get('status'),
                            v.get('descricao'), v.get('origem')])
    for v in payload.get('inconsistencias', []):
        validations.append([v.get('id'), v.get('documento_id'), v.get('tipo'), 'inconsistencia',
                            v.get('descricao'), v.get('origem')])
    builder.write_tab('Validações',
                      ['ID', 'Documento', 'Tipo', 'Status', 'Descrição', 'Origem'], validations)
    comparison = payload.get('comparativo', [])
    if comparison:
        comp_rows = [[x.get('campo'), x.get('documento_base'), x.get('valor_base'),
                      x.get('documento_comparado'), x.get('valor_comparado'), x.get('status'),
                      x.get('evidencia_base'), x.get('evidencia_comparada')] for x in comparison]
        builder.write_tab('Comparativo',
                          ['Campo', 'Documento A', 'Valor A', 'Documento B', 'Valor B',
                           'Status', 'Evidência A', 'Evidência B'], comp_rows)
    events = []
    for d in docs:
        for e in d.get('eventos', []):
            events.append([e.get('data'), d.get('nome'), e.get('descricao'), e.get('tipo'),
                           e.get('pagina'), e.get('bloco_id'), e.get('confirmado')])
    for e in payload.get('anotacoes', {}).get('eventos', []):
        events.append([e.get('data'), e.get('documento_id'), e.get('descricao'), e.get('tipo'),
                       e.get('pagina'), e.get('bloco_id'), e.get('confirmado')])
    if events:
        events.sort(key=lambda x: str(x[0] or ''))
        builder.write_tab('Cronologia',
                          ['Data', 'Documento', 'Descrição', 'Tipo', 'Página', 'Bloco', 'Confirmado'], events)
    analysis = payload.get('analise', {})
    if analysis:
        builder.write_tab('Síntese', ['Aspecto', 'Conteúdo'], [[k, v] for k, v in analysis.items()])
    table_number = 0
    table_index = []
    for d in docs:
        for table in d.get('tabelas', []):
            table_number += 1
            headers = table.get('colunas', [])
            raw_rows = table.get('linhas', [])
            rows = [[normalize_value(cell)[0] if isinstance(cell, str) else cell
                     for cell in row] for row in raw_rows]
            if not rows and not headers:
                continue
            title = f'Tabela_{table_number:03d}'
            builder.write_tab(title, headers, rows)
            table_index.append([title, table.get('id'), d.get('nome'), table.get('pagina'),
                                table.get('localizacao'), len(rows), len(headers)])
    if table_index:
        builder.write_tab('Índice Tabelas',
                          ['Aba', 'ID tabela', 'Documento', 'Página', 'Localização', 'Linhas', 'Colunas'],
                          table_index)
    builder.close()
    return builder.counts, builder.backend
