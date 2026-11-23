"""Validações estruturais e comparação conservadora entre documentos."""
from __future__ import annotations
from collections import defaultdict
from typing import Any


def validate_payload(payload: dict[str, Any]) -> dict[str, Any]:
    docs = payload.get('documentos') or []
    checks: list[dict[str, Any]] = []
    issues: list[dict[str, Any]] = []
    def check(kind: str, status: str, description: str, doc_id: str | None = None) -> None:
        checks.append({'id': f'v{len(checks)+1:05d}', 'tipo': kind, 'status': status,
                       'descricao': description, 'documento_id': doc_id,
                       'origem': 'python'})
    for d in docs:
        st = d.get('status')
        check('leitura', 'aprovado' if st == 'processado' else 'atencao',
              f"Leitura de {d.get('nome')}: {st}", d.get('id'))
        if st != 'processado':
            issues.append({'tipo': 'extracao_incompleta', 'documento_id': d.get('id'),
                           'descricao': '; '.join(d.get('avisos',[])) or 'Leitura incompleta',
                           'severidade': 'atencao'})
        if d.get('formato') == 'pdf' and d.get('paginas') and not d.get('blocos'):
            check('cobertura_pdf','atencao','PDF sem texto digital extraído; avaliar OCR.',d.get('id'))
        for t in d.get('tabelas', []):
            expected = len(t.get('colunas',[]))
            bad = sum(1 for r in t.get('linhas', []) if len(r) != expected)
            check('integridade_tabela','aprovado' if not bad else 'atencao',
                  f"Tabela {t.get('id')}: {len(t.get('linhas',[]))} linhas, {expected} colunas, {bad} linhas desalinhadas.",d.get('id'))
            if bad:
                issues.append({'tipo':'tabela_irregular', 'documento_id': d.get('id'),
                               'descricao':f"Tabela {t.get('id')} contém {bad} linhas com largura diferente.",
                               'severidade':'atencao'})
    payload['validacoes'] = checks + payload.get('validacoes',[])
    payload['inconsistencias'] = issues + payload.get('inconsistencias',[])
    return payload


def build_comparison(payload: dict[str, Any]) -> list[dict[str, Any]]:
    """Compara campos apenas quando são únicos em cada documento (sem falsos conflitos)."""
    docs = payload.get('documentos',[])
    if len(docs) < 2:
        return []
    by_doc = {}
    for doc in docs:
        grouped: dict[str,list[dict[str,Any]]] = defaultdict(list)
        for record in doc.get('registros',[]):
            key = str(record.get('campo','')).strip().casefold()
            if key: grouped[key].append(record)
        by_doc[doc['id']] = {k:v[0] for k,v in grouped.items() if len(v)==1}
    # Critério: comparação de cada documento com o primeiro.
    base = docs[0]['id']; result=[]
    for doc in docs[1:]:
        current = doc['id']
        for key in sorted(by_doc[base].keys() & by_doc[current].keys()):
            left, right = by_doc[base][key], by_doc[current][key]
            a,b = left.get('valor_normalizado'),right.get('valor_normalizado')
            if type(a) is type(b) and a == b:
                st = 'igual'
            else:
                st = 'diferente'
            result.append({'campo':left.get('campo'),'documento_base':base,'documento_comparado':current,
                           'valor_base':a, 'valor_comparado':b,'status':st,
                           'evidencia_base':left.get('id'),'evidencia_comparada':right.get('id'),
                           'criterio':'campo unico em ambos documentos'})
    return result


def merge_annotations(payload: dict[str, Any], annotations: dict[str, Any]) -> dict[str, Any]:
    """Enriquecimento do agente: resultados são rotulados e NUNCA alteram fontes brutas."""
    doc_ids={d.get('id') for d in payload.get('documentos',[])}
    block_ids={b.get('id') for d in payload.get('documentos',[]) for b in d.get('blocos',[])}
    for key in ('registros','eventos','validacoes','inconsistencias','comparativo'):
        values = annotations.get(key,[])
        if not isinstance(values,list): raise ValueError(f'Anotações: {key} deve ser uma lista')
        normalized=[]
        for item in values:
            if not isinstance(item,dict): raise ValueError(f'Anotações: elementos de {key} devem ser objetos')
            if key in ('registros','eventos'):
                if item.get('documento_id') not in doc_ids:
                    raise ValueError(f'Anotações {key}: documento_id inexistente')
                evidence_ids = item.get('evidencias') or []
                if not isinstance(evidence_ids,list):
                    raise ValueError('evidencias deve ser lista')
                candidate_refs = evidence_ids + ([item['bloco_id']] if item.get('bloco_id') else [])
                if not candidate_refs:
                    raise ValueError(f'Anotações {key}: informe pelo menos uma evidência documental')
                if any(ref not in block_ids for ref in candidate_refs):
                    raise ValueError(f'Anotações {key}: referência a bloco inexistente')
            normalized.append({**item,'origem':'interpretacao_ia'})
        if key in ('registros','eventos'):
            payload.setdefault('anotacoes',{}).setdefault(key,[]).extend(normalized)
        else:
            payload.setdefault(key,[]).extend(normalized)
    summary = annotations.get('analise',{})
    if summary:
        if not isinstance(summary,dict): raise ValueError('analise deve ser objeto')
        payload['analise'] = {**summary, 'origem':'interpretacao_ia'}
    return payload
