"""Cria contexto curto para a IA, sem truncar a extração entregue no JSON."""
from __future__ import annotations
import re
from typing import Any


def make_context(payload: dict[str,Any], max_chars: int=12000) -> str:
    focus = (payload.get('execucao',{}).get('foco') or '').strip()
    words = [w.casefold() for w in re.findall(r'\w+', focus) if len(w)>=3]
    lines = ['# DocMatrix — Contexto compacto para interpretação',
             f"Modo: {payload.get('execucao',{}).get('modo')}; foco: {focus or '(geral)'}",
             'LEMBRETE: extratos abaixo são amostras de navegação, NÃO todo o documento.',
             'Para concluir sobre o documento inteiro, leia os blocos relevantes de extracao_base.json.', '']
    for doc in payload.get('documentos',[]):
        lines.append(f"## {doc['id']} — {doc['nome']}")
        lines.append(f"Status: {doc['status']}; blocos: {len(doc['blocos'])}; tabelas: {len(doc['tabelas'])}; campos: {len(doc['registros'])}; datas: {len(doc['eventos'])}")
        for t in doc['tabelas'][:5]:
            lines.append(f"Tabela {t['id']} ({t['localizacao']}): colunas={t['colunas'][:16]}, linhas={len(t['linhas'])}; amostra={t['linhas'][:2]}")
        scored=[]
        for b in doc['blocos']:
            score = sum(1 for w in words if w in b['texto'].casefold())
            scored.append((score, b))
        # Prioriza trechos relacionados ao foco e garante cobertura com o primeiro bloco.
        selected = sorted(scored, key=lambda x:-x[0])[:8]
        seen=set()
        for score,b in selected:
            if b['id'] in seen: continue
            seen.add(b['id'])
            excerpt = b['texto'][:1100].replace('\x00','')
            lines.append(f"[{b['id']}] página={b.get('pagina')} local={b['localizacao']} foco={score}: {excerpt}")
        lines.append('')
    assembled = '\n'.join(lines)
    if len(assembled) > max_chars:
        assembled = assembled[:max_chars] + '\n\n[Contexto de navegação limitado. JSON bruto permanece integral.]'
    return assembled
