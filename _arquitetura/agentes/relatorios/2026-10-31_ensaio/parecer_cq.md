# Parecer CQ — ciclo 2026-10-31_ensaio (ENSAIO, modo completo)

Nada foi executado neste ensaio. Não abri Excel, não copiei arquivos e não rodei cálculos. Testes: NAO_EXECUTADO (motivo: ensaio).

No ciclo real, eu faria o seguinte:
1. Leria os relatórios das frentes F3 e F4 e a evidência numérica de cada ocorrência.
2. F4 (trimestre por dados brutos): conferiria em `src/bio/mEstatPeriodo.bas` e nos ADR-068/069 se a estatística do período é recalculada dos resultados brutos ou agregada a partir das médias mensais. Depois refaria o cálculo em Python puro (com N ponderado) e classificaria a ocorrência como ERRO_IMPLEMENTACAO ou DIFERENCA_METODO.
3. F3 (Sigma): recalcularia bias, CV e Sigma = (ETp − |bias|)/CV como o ADR-035/064 define, usando a média de |bias| e o CV pooled de 6 meses (decisões do RT), e compararia com o QC_INI na precisão exibida.
4. Avaliaria o impacto em analitos, níveis, lotes e períodos e veria se algum veredito de Sigma, plano de CQ ou regra de Westgard muda.
5. Fixaria critérios de aceite numéricos com tolerância para o reteste e marcaria o que depende de autorização do usuário.

## Ocorrências
- [ENSAIO] Trimestre por dados brutos: INCONCLUSIVO (fictícia). Prioridade final: Baixo.
- [ENSAIO] Cálculo fictício Sigma: NAO_PROCEDE (fictícia, sem dados). A prioridade foi rebaixada de Crítico para Baixo.
