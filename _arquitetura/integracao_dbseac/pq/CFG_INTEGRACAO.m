// =====================================================================
//  CFG_INTEGRACAO -- parametros da integracao DB_SEAC -> DB_CQ_FINAL
//  Lidos da tabela tblConfigIntegracao (aba Cfg_Integracao), CHAVE | VALOR.
//  Mudar caminho, prefixo de ID ou tolerancias e editar uma celula: nenhuma
//  consulta precisa ser alterada. Mesma consulta nos dois produtos.
// =====================================================================
let
    T = Excel.CurrentWorkbook(){[Name = "tblConfigIntegracao"]}[Content],
    SoPreenchidas = Table.SelectRows(T, each [CHAVE] <> null and Text.Trim(Text.From([CHAVE])) <> ""),
    Chaves = List.Transform(SoPreenchidas[CHAVE], each Text.Upper(Text.Trim(Text.From(_)))),
    Valores = List.Transform(SoPreenchidas[VALOR], each if _ is text then Text.Trim(_) else _),
    R = Record.FromList(Valores, Chaves)
in
    R
