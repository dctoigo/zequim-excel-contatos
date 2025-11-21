Attribute VB_Name = "extrairDados"

Sub ConsolidarDadosDasTabelas()
    Dim ws As Worksheet, tbl As ListObject, novaWs As Worksheet
    Dim ultimaLinhaNova As Long, linhaTabela As ListRow
    
    ' Cria ou limpa aba Consolidado
    On Error Resume Next
    Set novaWs = ThisWorkbook.Worksheets("Consolidado")
    If novaWs Is Nothing Then
        Set novaWs = ThisWorkbook.Worksheets.Add
        novaWs.Name = "Consolidado"
    Else
        novaWs.Cells.Clear
    End If
    On Error GoTo 0
    
    ' Cabeçalho
    novaWs.Range("A1:D1").Value = Array("mes", "nome", "telefone", "origem")
    ultimaLinhaNova = 2
    
    ' Percorre planilhas
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Instruções" And ws.Name <> novaWs.Name Then
            If ws.ListObjects.Count > 0 Then
                For Each tbl In ws.ListObjects
                    For Each linhaTabela In tbl.ListRows
                        Dim mesVal As Variant, nomeVal As Variant, telVal As Variant
                        mesVal = linhaTabela.Range.Cells(1, 1).Value
                        nomeVal = linhaTabela.Range.Cells(1, 2).Value
                        telVal = linhaTabela.Range.Cells(1, 3).Value
                        
                        If Trim(nomeVal) <> "" And Trim(telVal) <> "" Then
                            novaWs.Cells(ultimaLinhaNova, 1).Value = mesVal
                            novaWs.Cells(ultimaLinhaNova, 2).Value = nomeVal
                            novaWs.Cells(ultimaLinhaNova, 3).Value = telVal
                            novaWs.Cells(ultimaLinhaNova, 4).Value = ws.Name
                            ultimaLinhaNova = ultimaLinhaNova + 1
                        End If
                    Next linhaTabela
                Next tbl
            End If
        End If
    Next ws
    
    ' Cria tabela consolidada
    Dim rngConsolidado As Range
    Set rngConsolidado = novaWs.Range("A1:D" & ultimaLinhaNova - 1)
    Dim tblConsolidado As ListObject
    Set tblConsolidado = novaWs.ListObjects.Add(xlSrcRange, rngConsolidado, , xlYes)
    tblConsolidado.Name = "Tabela_Consolidado"
    
    MsgBox "Extração atualizada com sucesso!", vbInformation
End Sub
