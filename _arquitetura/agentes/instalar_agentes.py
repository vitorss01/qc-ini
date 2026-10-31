# -*- coding: utf-8 -*-
"""Instala os agentes do QC_INI na pasta em que o Claude Code é aberto.

Fonte única (versionada): qc-ini/.claude/{agents, skills/auditoria-qc, workflows/auditoria-qc.js}.
O Claude Code procura .claude/ na pasta de abertura da sessão; aqui ela é QC_INI_PRONTO, um nível acima do
repositório. Este script copia a fonte para lá. Edite sempre em qc-ini/.claude e rode de novo.

Uso: python instalar_agentes.py [pasta_de_abertura]   (padrão: a pasta acima do repositório)
     python instalar_agentes.py --conferir             (só diz se as cópias estão iguais à fonte)
"""
import filecmp
import os
import shutil
import sys

REPO = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..'))
FONTE = os.path.join(REPO, '.claude')
ITENS = [os.path.join('agents', n) for n in ('usuario-simulado.md', 'especialista-cq.md', 'ux-ui.md', 'tester-qa.md')]
ITENS += [os.path.join('skills', 'auditoria-qc', 'SKILL.md'), os.path.join('workflows', 'auditoria-qc.js')]


def main(argv):
    conferir = '--conferir' in argv
    resto = [a for a in argv if not a.startswith('--')]
    destino = os.path.join(os.path.abspath(resto[0]) if resto else os.path.dirname(REPO), '.claude')
    diferentes = 0
    for rel in ITENS:
        src, dst = os.path.join(FONTE, rel), os.path.join(destino, rel)
        igual = os.path.exists(dst) and filecmp.cmp(src, dst, shallow=False)
        if igual:
            print('IGUAL     ', rel)
            continue
        diferentes += 1
        if conferir:
            print('DIFERENTE ', rel)
            continue
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        shutil.copy2(src, dst)
        print('INSTALADO ', rel)
    print('destino:', destino)
    return 1 if (conferir and diferentes) else 0


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
