# Parecer UX/UI - ciclo 2026-10-31_ensaio (ENSAIO DE INTEGRAÇÃO)

Nada foi executado. Não abri Excel, não copiei arquivos e não rodei suítes. Testes: NAO_EXECUTADO (motivo: ensaio).

No ciclo real, eu faria o seguinte:
1. Leria o ADR-087, `visual_telas.py`, `testes/auditoria_abas.py`, `testes/auditoria_janelas.py` e o `GUIA_DE_USO.md` para fixar o padrão visual.
2. Leria a evidência de cada ocorrência nos relatórios das frentes (F1, F2), sem abrir Excel.
3. Julgaria cada ocorrência do ponto de vista da rotina do laboratório (pressa, luvas, interrupções) e daria o veredito PROCEDE, PROCEDE_PARCIAL ou NAO_PROCEDE.
4. Proporia uma correção viável em Excel/VBA (célula, rótulo, texto exato da mensagem, largura), respeitando a aba protegida e sem `Names.Add` com TRUE.
5. Marcaria como dependente de autorização qualquer correção que mexa em regra de negócio.
6. Definiria o critério de aceite visual ou de fluxo para o reteste (Usuário Simulado + auditoria visual).

## Ocorrências deste ensaio
- "[ENSAIO] Ocorrência fictícia de entrada de CIQ" (F1, Bioquímica): NAO_PROCEDE. É fictícia, não tem evidência e é do tipo funcional. Descartar.
- "[ENSAIO] Ocorrencia ficticia de usabilidade na janela do CEQ" (F2, Bioquímica): NAO_PROCEDE. É fictícia e não tem evidência. Descartar.
