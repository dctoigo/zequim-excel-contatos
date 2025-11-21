Attribute VB_Name = "auxiliar"
Sub CriarAbaInstrucoesComBotoes()
    Dim wsInstrucao As Worksheet
    Dim btnFormatar As Shape
    Dim btnAtualizar As Shape
    
    ' Cria ou limpa a aba Instruções
    On Error Resume Next
    Set wsInstrucao = ThisWorkbook.Worksheets("Instruções")
    If wsInstrucao Is Nothing Then
        Set wsInstrucao = ThisWorkbook.Worksheets.Add
        wsInstrucao.Name = "Instruções"
    Else
        wsInstrucao.Cells.Clear
    End If
    On Error GoTo 0
    
    ' Texto explicativo
    wsInstrucao.Range("A1").Value = "INSTRUÇÕES PARA USO DA PLANILHA"
    wsInstrucao.Range("A1").Font.Bold = True
    wsInstrucao.Range("A3").Value = "1. Habilitar Macros: Clique em 'Habilitar Conteúdo' ao abrir o arquivo."
    wsInstrucao.Range("A4").Value = "2. Botão Formatar Planilhas: Cria as tabelas e ajusta a estrutura. Execute se adicionar nova planilha."
    wsInstrucao.Range("A5").Value = "3. Botão Atualizar Extração: Atualiza a Tabela Consolidado com nomes e telefones corrigidos ou novos."
    wsInstrucao.Range("A6").Value = "Importante: Não altere manualmente a Tabela Consolidado, pois será sobrescrita na atualização."
    
    ' Ajusta largura
    wsInstrucao.Columns("A").ColumnWidth = 100
    
    ' Cria Botão Formatar Planilhas
    Set btnFormatar = wsInstrucao.Shapes.AddShape(msoShapeRoundedRectangle, 20, 100, 200, 40)
    With btnFormatar
        .TextFrame.Characters.Text = "Formatar Planilhas"
        .OnAction = "CriarTabelasEmTodasAsPlanilhas"
        .Fill.ForeColor.RGB = RGB(0, 176, 80)
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .TextFrame.HorizontalAlignment = xlHAlignCenter
    End With
    
    ' Cria Botão Atualizar Extração
    Set btnAtualizar = wsInstrucao.Shapes.AddShape(msoShapeRoundedRectangle, 20, 160, 200, 40)
    With btnAtualizar
        .TextFrame.Characters.Text = "Atualizar Extração"
        .OnAction = "ConsolidarDadosDasTabelas"
        .Fill.ForeColor.RGB = RGB(0, 112, 192)
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .TextFrame.HorizontalAlignment = xlHAlignCenter
    End With
    
    MsgBox "Aba 'Instruções' criada com botões!", vbInformation
End Sub


Sub CriarPlanilhaDDD_UF()
    Dim wsDDD As Worksheet
    Dim tbl As ListObject
    Dim rngTabela As Range
    Dim dados As Variant
    Dim i As Long
    
    ' Cria ou limpa a planilha tb_dddUF
    On Error Resume Next
    Set wsDDD = ThisWorkbook.Worksheets("tb_dddUF")
    If wsDDD Is Nothing Then
        Set wsDDD = ThisWorkbook.Worksheets.Add
        wsDDD.Name = "tb_dddUF"
    Else
        wsDDD.Cells.Clear
    End If
    On Error GoTo 0
    
    ' Cabeçalho
    wsDDD.Range("A1").Value = "UF"
    wsDDD.Range("B1").Value = "CN"
    
    ' Dados fornecidos
    dados = Array( _
        Array("AC", 68), Array("AL", 82), Array("AM", 97), Array("AM", 92), Array("AP", 96), _
        Array("BA", 77), Array("BA", 75), Array("BA", 73), Array("BA", 74), Array("BA", 71), _
        Array("CE", 88), Array("CE", 85), Array("DF", 61), Array("ES", 27), Array("ES", 28), _
        Array("GO", 62), Array("GO", 64), Array("MA", 99), Array("MA", 98), _
        Array("MG", 34), Array("MG", 37), Array("MG", 31), Array("MG", 33), Array("MG", 35), _
        Array("MG", 32), Array("MG", 38), Array("MS", 67), Array("MT", 65), Array("MT", 66), _
        Array("PA", 91), Array("PA", 94), Array("PA", 93), Array("PB", 83), Array("PE", 81), _
        Array("PE", 87), Array("PI", 89), Array("PI", 86), _
        Array("PR", 43), Array("PR", 41), Array("PR", 44), Array("PR", 46), Array("PR", 45), _
        Array("PR", 42), Array("PR", 49), Array("PR", 47), _
        Array("RJ", 24), Array("RJ", 22), Array("RJ", 21), _
        Array("RN", 84), Array("RO", 69), Array("RR", 95), _
        Array("RS", 53), Array("RS", 54), Array("RS", 55), Array("RS", 51), _
        Array("SC", 48), Array("SE", 79), _
        Array("SP", 18), Array("SP", 17), Array("SP", 19), Array("SP", 14), Array("SP", 15), _
        Array("SP", 16), Array("SP", 11), Array("SP", 12), Array("SP", 13), _
        Array("TO", 63))
    
    ' Preenche os dados na planilha
    For i = LBound(dados) To UBound(dados)
        wsDDD.Cells(i + 2, 1).Value = dados(i)(0)
        wsDDD.Cells(i + 2, 2).Value = dados(i)(1)
    Next i
    
    ' Define intervalo e cria tabela
    Dim ultimaLinha As Long
    ultimaLinha = wsDDD.Cells(wsDDD.Rows.Count, "A").End(xlUp).Row
    Set rngTabela = wsDDD.Range("A1:B" & ultimaLinha)
    
    On Error Resume Next
    Set tbl = wsDDD.ListObjects.Add(xlSrcRange, rngTabela, , xlYes)
    tbl.Name = "tb_dddUF"
    On Error GoTo 0
    
    MsgBox "Planilha 'tb_dddUF' criada com sucesso!", vbInformation
End Sub


