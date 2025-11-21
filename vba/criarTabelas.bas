Attribute VB_Name = "criarTabelas"

Sub CriarTabelasEmTodasAsPlanilhas()
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim primeiraLinhaDados As Long
    Dim tbl As ListObject
    Dim rngTabela As Range
    
    For Each ws In ThisWorkbook.Worksheets
        ' Ignora aba Instruções e Consolidado
        If ws.Name <> "Instruções" And ws.Name <> "Consolidado" And ws.Name <> "tb_ddd" And ws.Name <> "_extracao" Then
            
            ' Verifica se já existe tabela; se sim, ignora
            If ws.ListObjects.Count = 0 Then
                
                ' Identifica primeira linha com dados
                If Application.WorksheetFunction.CountA(ws.Columns("A")) > 0 Then
                    primeiraLinhaDados = ws.Columns("A").Find(What:="*", LookIn:=xlValues, _
                                                              LookAt:=xlWhole, SearchOrder:=xlByRows, _
                                                              SearchDirection:=xlNext).Row
                Else
                    primeiraLinhaDados = 1
                End If
                
                ' Se primeira linha for 1, insere nova linha no topo
                If primeiraLinhaDados = 1 Then
                    ws.Rows(1).Insert Shift:=xlDown
                    primeiraLinhaDados = 1
                Else
                    ws.Rows(primeiraLinhaDados).Insert Shift:=xlDown
                End If
                
                ' Verifica se cabeçalho já existe
                If LCase(ws.Cells(primeiraLinhaDados, 1).Value) <> "mes" Or _
                   LCase(ws.Cells(primeiraLinhaDados, 2).Value) <> "nome" Or _
                   LCase(ws.Cells(primeiraLinhaDados, 3).Value) <> "telefone" Then
                   
                    ws.Cells(primeiraLinhaDados, 1).Value = "mes"
                    ws.Cells(primeiraLinhaDados, 2).Value = "nome"
                    ws.Cells(primeiraLinhaDados, 3).Value = "telefone"
                End If
                
                ' Última linha com dados
                ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
                
                ' Define intervalo e cria tabela
                Set rngTabela = ws.Range(ws.Cells(primeiraLinhaDados, 1), ws.Cells(ultimaLinha, 3))
                On Error Resume Next
                Set tbl = ws.ListObjects.Add(xlSrcRange, rngTabela, , xlYes)
                tbl.Name = "Tabela_" & ws.Name
                On Error GoTo 0
            End If
        End If
    Next ws
    
    MsgBox "Formatação concluída!", vbInformation
End Sub
