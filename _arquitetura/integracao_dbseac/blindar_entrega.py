# -*- coding: utf-8 -*-
"""blindar_entrega.py -- QUALITY_GATE 3.3/3.5 na arquitetura do ADR-057: o arquivo SALVO sai travado.

Uso:  python blindar_entrega.py <arquivo.xlsm> [--so-verificar]

O blindar_artefato.ps1 (scripts_fase3) fazia isto no build antigo; os instaladores do
ADR-057 em diante nao passavam por ele, e o arquivo voltou a ser salvo com todas as abas
visiveis e a estrutura aberta (Bioquimica: 35 de 35 visiveis). Com macros habilitadas nada
mudava -- Workbook_Open chama LockApp --, mas quem abre com macros DESABILITADAS (o auditor
cauteloso, o curioso) via a tabela de usuarios, a Audit_Log e a configuracao.

Faz:
  - toda aba protegida; a que ja estava protegida mantem as permissoes que tinha;
    Audit_* com filtro e ordenacao (o mesmo que mAuditoria.ProtegerAudit);
  - toda aba xlSheetVeryHidden, menos Login (visivel e ativa);
  - estrutura da pasta protegida (impede reexibir aba, inclusive pela janela de
    propriedades do VBE).
Nao mexe no comportamento em uso: com macros habilitadas, LockApp/UnlockApp continuam
decidindo o que cada papel ve. NAO desliga a AutoRecuperacao (seria gravado no arquivo e
valeria para o laboratorio).

Confere LENDO O XML do .xlsm salvo, nao perguntando ao Excel: a pergunta e "o que se ve
com macros desligadas", e o Excel com macros ligadas nao responde isso.
"""
import os
import re
import sys
import zipfile

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
sys.path.insert(0, os.path.join(AQUI, '..', 'etapa1_multilote'))
import xlh  # noqa: E402
import instalar_integracao as ii  # noqa: E402

LOGIN = 'Login'


def log(*a):
    print(*a, flush=True)


def proteger_audit(ws):
    # = mAuditoria.ProtegerAudit, sem UserInterfaceOnly (nao persiste mesmo)
    ws.Protect(ii.SENHA, False, True, True, False, False, True, False, False, False, False, False, False,
               True, True, False)


def blindar(caminho):
    ex = xlh.Excel()
    log(f'EXCEL_PID {ex.pid}')
    try:
        wb = ex.abrir(caminho)
        if wb.ReadOnly:
            raise SystemExit(f'arquivo aberto somente leitura (outra instancia o segura): {caminho}')
        try:
            wb.Unprotect(ii.SENHA)
        except Exception:
            pass
        login = [ws for ws in wb.Worksheets if ws.Name == LOGIN]
        if not login:
            raise SystemExit('aba Login nao encontrada -- nao da para travar o arquivo com seguranca')
        login[0].Visible = -1
        login[0].Activate()
        n_prot = n_ocul = 0
        for ws in wb.Worksheets:
            if ws.Name.startswith('Audit_'):
                try:
                    ws.Unprotect(ii.SENHA)
                except Exception:
                    pass
                proteger_audit(ws)
                n_prot += 1
            elif not ws.ProtectContents:
                ws.Protect(ii.SENHA, False, True, True)
                n_prot += 1
            if ws.Name != LOGIN:
                ws.Visible = 2
                n_ocul += 1
        wb.Protect(ii.SENHA, True, False)
        wb.Save()
        log(f'   protecao aplicada/ajustada em {n_prot} aba(s); {n_ocul} ocultas (veryHidden); estrutura protegida; salvo')
    finally:
        ex.fechar()


def verificar(caminho):
    """Le o XML do .xlsm: o estado que vale com macros desligadas."""
    with zipfile.ZipFile(caminho) as z:
        wbx = z.read('xl/workbook.xml').decode('utf-8')
        rels = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
        alvo = dict(re.findall(r'<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
        alvo.update({k: v for v, k in re.findall(r'<Relationship[^>]*Target="([^"]+)"[^>]*Id="([^"]+)"', rels)})
        abas = re.findall(r'<sheet\b([^>]*)/>', wbx)
        visiveis, sem_prot = [], []
        for a in abas:
            nome = re.search(r'name="([^"]+)"', a).group(1)
            estado = (re.search(r'state="([^"]+)"', a) or [None, 'visible'])[1]
            rid = re.search(r'r:id="([^"]+)"', a).group(1)
            xml = z.read('xl/' + alvo[rid].lstrip('/').replace('xl/', '')).decode('utf-8')
            if estado == 'visible':
                visiveis.append(nome)
            if '<sheetProtection' not in xml:
                sem_prot.append(nome)
        estrutura = bool(re.search(r'<workbookProtection[^>]*lockStructure="(1|true)"', wbx))
    ok = visiveis == [LOGIN] and not sem_prot and estrutura
    log(f'   XML: {len(abas)} abas | visiveis {visiveis} | sem protecao {sem_prot} | estrutura travada {estrutura}'
        f' -> {"TRAVADO" if ok else "NAO TRAVADO"}')
    return ok


if __name__ == '__main__':
    arq = os.path.abspath(sys.argv[1])
    if '--so-verificar' not in sys.argv:
        blindar(arq)
    sys.exit(0 if verificar(arq) else 1)
