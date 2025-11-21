let
    Fonte = Excel.CurrentWorkbook(){[Name="Tabela_Consolidado"]}[Content],
    textoAparado = Table.TransformColumns(
        Table.TransformColumnTypes(Fonte, {{"telefone", type text}}, "pt-BR"),
        {{"telefone", Text.Trim, type text}}
    ),
    textoLimpo = Table.TransformColumns(textoAparado, {{"telefone", Text.Clean, type text}}),
    removidosEspacos = Table.ReplaceValue(textoLimpo, " ", "", Replacer.ReplaceText, {"telefone"}),
    removidosTracos = Table.ReplaceValue(removidosEspacos, "-", "", Replacer.ReplaceText, {"telefone"}),
    addColTemporaria = Table.DuplicateColumn(removidosTracos, "telefone", "t_tmp", Int64.Type),
    addColTamanho = Table.AddColumn(addColTemporaria, "t_length", each Text.Length([telefone]), Int64.Type),
    alteradoTipoColTemp = Table.TransformColumnTypes(addColTamanho, {{"t_tmp", Int64.Type}, {"t_length", Int64.Type}}),
    filtroSobreTemp = Table.RemoveRowsWithErrors(alteradoTipoColTemp, {"t_tmp"}),

    // ---------------------------------------------
    // Carrega tabela de DDD válidos (tb_ddd[CN])
    // ---------------------------------------------
    tbDDD = Excel.CurrentWorkbook(){[Name="tb_ddd"]}[Content],
    tbDDD_CN_Texto = Table.TransformColumns(
        Table.TransformColumnTypes(tbDDD, {{"CN", type text}}, "pt-BR"),
        {{"CN", Text.Trim, type text}}
    ),
    validDDDList = tbDDD_CN_Texto[CN],

    // ---------------------------------------------------
    // 1) Cria record com status, ddi, ddd e número base
    // ---------------------------------------------------
    addDadosTelefone = Table.AddColumn(
        filtroSobreTemp,
        "dadosTelefone",
        each
            let
                tel    = [telefone],
                len    = [t_length],
                first1 = if len >= 1 then Text.Start(tel, 1) else "",
                first2 = if len >= 2 then Text.Start(tel, 2) else "",

                // DDI
                ddi =
                    if (len = 10 or len = 11) then
                        "55"
                    else if (List.Contains({12, 13}, len) and first2 = "55") then
                        "55"
                    else
                        null,

                // DDD
                ddd =
                    if (len = 10 or len = 11) then
                        Text.Start(tel, 2)
                    else if (List.Contains({12, 13}, len) and first2 = "55") then
                        Text.Range(tel, 2, 2)
                    else
                        null,

                // Número local
                numero =
                    if (len = 10 or len = 11) then
                        Text.Range(tel, 2, len - 2)
                    else if (List.Contains({12, 13}, len) and first2 = "55") then
                        Text.Range(tel, 4, len - 4)
                    else if (len = 8 and (first1 = "8" or first1 = "9")) then
                        "9" & tel
                    else if (len = 9) then
                        tel
                    else
                        null,

                // Status base (sem checar DDD ainda)
                statusBase =
                    if (len = 10 or len = 11) then
                        "OK"
                    else if (len = 8 and (first1 = "8" or first1 = "9")) then
                        "Falta DDD"
                    else if (len = 9) then
                        "Falta DDD"
                    else if (List.Contains({12, 13}, len) and first2 <> "55") then
                        "Erro DDI"
                    else if (List.Contains({12, 13}, len) and first2 = "55") then
                        "OK"
                    else if (len > 13 or len <= 7) then
                        "Erro"
                    else
                        "Erro",

                // Validação do DDD com base na tb_ddd
                isDDDValid = if ddd <> null then List.Contains(validDDDList, ddd) else true,

                    // Status final mais amigável
                statusFinalRaw =
                    if (ddd <> null and not isDDDValid and statusBase = "OK") then
                        "DDD errado"
                    else
                        statusBase,

                // Conversão para nomes mais amigáveis
                statusFinal =
                    if statusFinalRaw = "OK" then "OK"
                    else if statusFinalRaw = "Falta DDD" then "Falta DDD"
                    else if statusFinalRaw = "Erro" then "Formato inválido"
                    else if statusFinalRaw = "Erro DDI" then "DDI inválido"
                    else if statusFinalRaw = "DDD errado" then "DDD inválido"
                    else "Formato inválido"
            in
                [status = statusFinal, ddi = ddi, ddd = ddd, numero = numero]
    ),

    // ---------------------------------------------------
    // 2) Expande o record em colunas
    // ---------------------------------------------------
    expandeDados = Table.ExpandRecordColumn(
        addDadosTelefone,
        "dadosTelefone",
        {"status", "ddi", "ddd", "numero"},
        {"t_status", "t_ddi", "t_ddd", "t_numero"}
    ),

    // ---------------------------------------------------
    // 3) Coluna final normalizada: DDI (DDD) NUMERO
    //     – só para status OK
    // ---------------------------------------------------
    addTelefoneNormalizado = Table.AddColumn(
        expandeDados,
        "telefone_normalizado",
        each
            if [t_status] = "OK" and [t_ddi] <> null and [t_ddd] <> null and [t_numero] <> null then
                [t_ddi] & " (" & [t_ddd] & ") " & [t_numero]
            else
                null,
        type text
    ),    
    // ---------------------------------------------------
    // 4) Formato CRM: +55 99 999999999
    // ---------------------------------------------------
    addFormatoCRM = Table.AddColumn(
        addTelefoneNormalizado,
        "telefone_CRM",
        each if [t_status] = "OK" then
                "+" & [t_ddi] & " " & [t_ddd] & " " & [t_numero]
             else
                null,
        type text
    ),

    // ---------------------------------------------------
    // 5) Formato Bot: 5599999999999
    // ---------------------------------------------------
    addFormatoBot = Table.AddColumn(
        addFormatoCRM,
        "telefone_Bot",
        each if [t_status] = "OK" then
                [t_ddi] & [t_ddd] & [t_numero]
             else
                null,
        type text
    ),
    #"Tipo Alterado" = Table.TransformColumnTypes(addFormatoBot,{{"nome", type text}, {"t_status", type text}, {"t_ddi", type text}, {"t_ddd", type text}, {"t_numero", type text}, {"origem", type text}}),
    #"Outras Colunas Removidas" = Table.SelectColumns(#"Tipo Alterado",{"nome", "origem", "telefone", "t_status", "telefone_normalizado", "telefone_CRM", "telefone_Bot"}),
    #"Colocar Cada Palavra Em Maiúscula" = Table.TransformColumns(#"Outras Colunas Removidas",{{"nome", Text.Proper, type text}}),
    #"Duplicatas Removidas" = Table.Distinct(#"Colocar Cada Palavra Em Maiúscula", {"telefone"})
in
    #"Duplicatas Removidas"