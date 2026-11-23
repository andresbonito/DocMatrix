"""Verificação leve dos artefatos sem recalcular planilhas nem ler tudo no contexto."""
from __future__ import annotations
import json
from pathlib import Path
import zipfile
from xml.etree import ElementTree as ET


def verify_outputs(json_path: str | Path, xlsx_path: str | Path) -> dict:
    jp,xp=Path(json_path),Path(xlsx_path)
    errors=[]
    if not jp.is_file() or jp.stat().st_size==0: errors.append('JSON ausente ou vazio')
    if not xp.is_file() or xp.stat().st_size==0: errors.append('Excel ausente ou vazio')
    if errors: return {'status':'erro','erros':errors}
    try:
        data=json.loads(jp.read_text(encoding='utf-8'))
        if data.get('schema_version')!='1.0.0': errors.append('Versão do JSON inesperada')
        if not isinstance(data.get('documentos'),list): errors.append('documentos não é lista')
    except Exception as exc:
        errors.append(f'JSON inválido: {exc}')
    sheet_names=[]
    try:
        with zipfile.ZipFile(xp) as z:
            if z.testzip() is not None: errors.append('Erro de integridade ZIP XLSX')
            names=set(z.namelist())
            for required in ('xl/workbook.xml','[Content_Types].xml'):
                if required not in names: errors.append(f'XLSX sem {required}')
            root=ET.fromstring(z.read('xl/workbook.xml'))
            namespace='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
            sheet_names=[s.get('name') for s in root.findall(f'.//{namespace}sheet')]
    except Exception as exc:
        errors.append(f'XLSX inválido: {exc}')
    for required in ('Visão Geral','Dados Extraídos','Evidências','Validações'):
        if required not in sheet_names: errors.append(f'Aba ausente: {required}')
    return {'status':'aprovado' if not errors else 'erro','erros':errors,'abas':sheet_names,
            'documentos':len(data.get('documentos',[])) if 'data' in locals() else None}
