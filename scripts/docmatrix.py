#!/usr/bin/env python3
"""CLI DocMatrix. Use: python scripts/docmatrix.py --help"""
from __future__ import annotations
import argparse
from datetime import datetime, timezone
import json
from pathlib import Path
import sys
from typing import Any

ROOT=Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0,str(ROOT))
from scripts.extract import extract_document
from scripts.context import make_context
from scripts.validate import build_comparison, merge_annotations, validate_payload
from scripts.export_json import export_json
from scripts.export_excel import export_excel
from scripts.verify import verify_outputs

MODES={'auto','extracao','comparacao','validacao','cronologia','sintese'}


def prepare(inputs: list[str], output_dir: str, mode: str='auto', focus: str='',
            context_chars: int=12000, max_mb: int=100) -> tuple[Path,Path]:
    if mode not in MODES: raise ValueError(f'Modo inválido: {mode}')
    if not inputs: raise ValueError('Informe ao menos um arquivo de entrada')
    dest=Path(output_dir).expanduser().resolve(); dest.mkdir(parents=True,exist_ok=True)
    paths=[Path(s).expanduser().resolve(strict=True) for s in inputs]
    docs=[extract_document(p,f'doc_{i:03d}',max_bytes=max_mb*1024*1024) for i,p in enumerate(paths,1)]
    modes_applied=['extracao','validacao']
    if mode in ('comparacao','auto') and len(docs)>1: modes_applied.append('comparacao')
    if mode in ('cronologia','auto') and any(d.get('eventos') for d in docs): modes_applied.append('cronologia')
    if mode=='sintese': modes_applied.append('sintese')
    payload: dict[str,Any]={
        'schema_version':'1.0.0', 'skill':'docmatrix',
        'execucao':{'modo':mode,'foco':focus,'estrategia':'hibrida',
                    'gerado_em_utc':datetime.now(timezone.utc).isoformat(),
                    'modos_aplicados':list(dict.fromkeys(modes_applied)),
                    'interpretacao_ia':'pendente'},
        'documentos':docs,'validacoes':[],'inconsistencias':[],
        'comparativo':[], 'anotacoes':{'registros':[],'eventos':[]},'analise':{}
    }
    payload=validate_payload(payload)
    if 'comparacao' in modes_applied:
        payload['comparativo']=build_comparison(payload)
    base=export_json(payload,dest/'extracao_base.json')
    context=dest/'contexto_ia.md'
    context.write_text(make_context(payload,max_chars=context_chars),encoding='utf-8')
    return base,context


def finalize(base_file: str, output_dir: str, annotations: str | None=None,
             basename: str='docmatrix_resultado', excel_backend: str='auto') -> dict[str,Any]:
    dest=Path(output_dir).expanduser().resolve(); dest.mkdir(parents=True,exist_ok=True)
    payload=json.loads(Path(base_file).read_text(encoding='utf-8'))
    if annotations:
        notes=json.loads(Path(annotations).read_text(encoding='utf-8'))
        payload=merge_annotations(payload,notes)
        payload['execucao']['interpretacao_ia']='fornecida_pelo_agente'
    else:
        payload['execucao']['interpretacao_ia']='nao_executada'
        payload['analise']={'observacao':'Extração determinística. Interpretação semântica não foi executada.'}
    # Não confiar em nome arbitrário que altere a pasta de saída.
    filename=Path(basename).name
    if filename in ('','.','..'): raise ValueError('Nome inválido de resultado')
    json_path=export_json(payload,dest/f'{filename}.json')
    xlsx_path=dest/f'{filename}.xlsx'
    sheet_counts,used_backend=export_excel(payload,xlsx_path,backend=excel_backend)
    verification=verify_outputs(json_path,xlsx_path)
    report={'excel':str(xlsx_path),'json':str(json_path),'abas':sheet_counts,'backend_excel':used_backend,'verificacao':verification}
    (dest/'verificacao.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    if verification['status']!='aprovado':
        raise RuntimeError('Falha na verificação final: '+ '; '.join(verification['erros']))
    return report


def main() -> int:
    parser=argparse.ArgumentParser(description='DocMatrix: extração, estruturação e exportação Excel + JSON')
    sub=parser.add_subparsers(dest='command',required=True)
    p=sub.add_parser('prepare',help='Lê documentos e gera JSON bruto + contexto reduzido')
    r=sub.add_parser('run',help='Pipeline determinístico completo: prepare + finalize')
    for cmd in (p,r):
        cmd.add_argument('--input',nargs='+',required=True,help='Um ou mais arquivos, com caminhos entre aspas')
        cmd.add_argument('--output-dir',required=True)
        cmd.add_argument('--mode',choices=sorted(MODES),default='auto')
        cmd.add_argument('--focus',default='',help='Objetivo da análise, para priorizar contexto')
        cmd.add_argument('--context-chars',type=int,default=12000)
        cmd.add_argument('--max-mb',type=int,default=100,help='Limite por arquivo; amplie conscientemente')
        if cmd is r:
            cmd.add_argument('--excel-backend',choices=['auto','artifact_tool','xlsxwriter'],default='auto')
    f=sub.add_parser('finalize',help='Mescla análise opcional e cria Excel + JSON')
    f.add_argument('--base',required=True,help='Caminho do extracao_base.json')
    f.add_argument('--output-dir',required=True)
    f.add_argument('--annotations',help='JSON interpretativo criado pelo agente')
    f.add_argument('--basename',default='docmatrix_resultado')
    f.add_argument('--excel-backend',choices=['auto','artifact_tool','xlsxwriter'],default='auto')
    v=sub.add_parser('verify',help='Confere integridade de Excel e JSON')
    v.add_argument('--json',required=True); v.add_argument('--excel',required=True)
    args=parser.parse_args()
    try:
        if args.command in ('prepare','run'):
            if args.context_chars<1000 or args.max_mb<1:
                raise ValueError('--context-chars >= 1000 e --max-mb >= 1')
            base,ctx=prepare(args.input,args.output_dir,args.mode,args.focus,args.context_chars,args.max_mb)
            if args.command=='prepare':
                print(json.dumps({'base':str(base),'contexto_ia':str(ctx)},ensure_ascii=False))
            else:
                print(json.dumps(finalize(str(base),args.output_dir,excel_backend=args.excel_backend),ensure_ascii=False))
        elif args.command=='finalize':
            print(json.dumps(finalize(args.base,args.output_dir,args.annotations,args.basename,args.excel_backend),ensure_ascii=False))
        elif args.command=='verify':
            check=verify_outputs(args.json,args.excel)
            print(json.dumps(check,ensure_ascii=False))
            if check['status']!='aprovado': return 2
    except Exception as exc:
        print(json.dumps({'status':'erro','erro':f'{type(exc).__name__}: {exc}'},ensure_ascii=False),file=sys.stderr)
        return 1
    return 0

if __name__=='__main__':
    raise SystemExit(main())
