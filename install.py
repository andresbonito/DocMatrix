#!/usr/bin/env python3
"""Copia a Skill DocMatrix para o diretório de Skills do Claude Code ou Codex.

Não instala dependências e não altera configurações dos agentes.
"""
from __future__ import annotations
import argparse
from pathlib import Path
import shutil

HOST_FOLDER = {'claude': Path('.claude/skills'), 'codex': Path('.agents/skills')}
IGNORE = shutil.ignore_patterns('.venv', '__pycache__', '*.pyc', '.DS_Store', '.pytest_cache', 'saida')


def install(host: str, scope: str = 'user', project: str | None = None,
            overwrite: bool = False) -> Path:
    source = Path(__file__).resolve().parent
    if host not in HOST_FOLDER or scope not in ('user', 'project'):
        raise ValueError('host deve ser claude/codex e scope deve ser user/project')
    if scope == 'user':
        base = Path.home()
    else:
        if not project:
            raise ValueError('Informe --project quando --scope project')
        base = Path(project).expanduser().resolve()
        if not base.is_dir():
            raise ValueError(f'Projeto não encontrado: {base}')
    target = base / HOST_FOLDER[host] / 'docmatrix'
    if target.resolve() == source:
        raise ValueError('O pacote já está no diretório de instalação')
    if target.exists():
        if not overwrite:
            raise FileExistsError(f'Instalação existente: {target}. Use --force para atualizar.')
        if not (target.is_dir() and (target / 'SKILL.md').is_file()):
            raise ValueError('Destino não parece ser uma Skill DocMatrix; substituição bloqueada')
        # Atualiza substituindo a pasta inteira, sem misturar arquivos de versões.
        shutil.rmtree(target)
    target.parent.mkdir(parents=True, exist_ok=True)
    shutil.copytree(source, target, ignore=IGNORE)
    if not (target / 'SKILL.md').is_file():
        raise RuntimeError('Falha: SKILL.md não encontrado na instalação')
    return target


def main() -> None:
    parser = argparse.ArgumentParser(description='Instalador local DocMatrix (Claude Code / Codex)')
    parser.add_argument('--host', choices=sorted(HOST_FOLDER), required=True)
    parser.add_argument('--scope', choices=['user', 'project'], default='user')
    parser.add_argument('--project', help='Pasta do projeto quando --scope project')
    parser.add_argument('--force', action='store_true', help='Substituir instalação DocMatrix existente')
    args = parser.parse_args()
    try:
        target = install(args.host, args.scope, args.project, args.force)
    except (ValueError, FileExistsError, OSError) as exc:
        parser.error(str(exc))
    print(f'DocMatrix instalado em: {target}')
    print('Instale dependências conforme README.md e reinicie a sessão do agente, se necessário.')


if __name__ == '__main__':
    main()
