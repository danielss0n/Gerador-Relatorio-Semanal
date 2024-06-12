Attribute VB_Name = "GerarRelatorio"
Public objWord As Object
Public objDoc As Object
Public DataDir As String
Public wb As Workbook
Public ws As Worksheet
Public NumeroSemanaPassada As String

Sub Init()
    DataDir = "diretorio/para/os/dados"
    NumeroSemanaPassada = CStr(Application.WorksheetFunction.WeekNum(DateAdd("d", -7, Date)))

    AbrirWord
    AtualizarDataSource
End Sub

Sub AtualizarDataSource()

    Call EscreverParagrafo("T�tulo 1", "Pir�mide de Heinrich")
    Set wb = Workbooks.Open(DataDir + "SEGURAN�A.xlsx")

    Set ws = wb.Worksheets(1)
    Call CopiarTabela("I1:M8")
    
    
    Call EscreverParagrafo("T�tulo 1", "Incidentes da semana")
    Call CopiarIncidentes
    
    
    Call EscreverParagrafo("T�tulo 1", "Cart�es de seguran�a")
    Set ws = wb.Worksheets(4)
    Dim lrow As Integer
    lrow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Call CopiarTabela("A1:E" & lrow)
    
    
    Call EscreverParagrafo("T�tulo 1", "Fatores de trabalho")
    Set ws = wb.Worksheets(5)
    lrow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Call CopiarTabela("A1:E" & lrow)
    
    wb.Save
    wb.Close
    
    Call EscreverParagrafo("T�tulo 1", "Qualidade")
    Call EscreverParagrafo("T�tulo 2", "QMs")

    Set wb = Workbooks.Open(DataDir + "QM.xlsx")
    Set ws = wb.Worksheets(1)
    Call CopiarTabela("L1:P3")
    
    Call EscreverParagrafo("T�tulo 2", "N�o conformidades causadas")
    Set ws = wb.Worksheets(2)
    lrow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Call CopiarTabela("A1:F" & lrow)
    
    wb.Save
    wb.Close
    
    Call EscreverParagrafo("T�tulo 2", "BIQs")
    Set wb = Workbooks.Open(DataDir + "BIQ.xlsx")
    Set ws = wb.Worksheets(1)
    Call CopiarTabela("G1:K3")
    
    wb.Save
    wb.Close
    
    Call EscreverParagrafo("T�tulo 1", "Absenteismo")
    Set wb = Workbooks.Open(DataDir + "ABSENTEISMO.xlsx")
    Set ws = wb.Worksheets(1)
    Call CopiarTabela("U1:Y27")
    
    wb.Save
    wb.Close
End Sub

Sub AbrirWord()
    Set objWord = CreateObject("Word.Application")
    objWord.Visible = True
    Set objDoc = objWord.Documents.Open("Relat�rio em branco.docx")
End Sub

Sub CopiarTabela(rngString)
    Dim Rng As range
    Set Rng = ws.range(rngString)
    Rng.Copy
    
    objDoc.Paragraphs(objDoc.Paragraphs.Count).range.PasteExcelTable LinkedToExcel:=True, WordFormatting:=False, RTF:=False
    objDoc.Content.InsertAfter ""
    
    Set tbl = objDoc.Tables(objDoc.Tables.Count)
    tbl.AutoFitBehavior (1)
End Sub

Sub CopiarIncidentes()
    Set wb = Workbooks.Open(DataDir + "SEGURAN�A.xlsx")
    Set ws = wb.Worksheets(1)
       
    For i = 1 To 100
        If ws.Cells(i, 5).Value = NumeroSemanaPassada Then
            ' Pegar textos das linhas
            dataIncidente = ws.Cells(i, 1)
            secaoIncidente = ws.Cells(i, 2)
            tipoIncidente = ws.Cells(i, 3)
            descricaoIncidente = ws.Cells(i, 4)
            
            ' Colar no word
            With objDoc.Content
                .Collapse Direction:=0
                .Style = objDoc.Styles("T�tulo 2")
                .InsertAfter tipoIncidente & " - " & dataIncidente & vbCrLf
                .Collapse Direction:=0
                .Style = objDoc.Styles("Normal")
                .InsertAfter "Caldeiraria - " & secaoIncidente & " - " & descricaoIncidente & vbCrLf
            End With
            objDoc.Content.InsertAfter ""
            objDoc.Content.InsertAfter ""
        End If
    Next
End Sub

Sub EscreverParagrafo(tipoParagrafo, textoParagrafo)
    With objDoc.Content
        .Collapse Direction:=0
        .Style = objDoc.Styles(tipoParagrafo)
        .InsertAfter textoParagrafo & vbCrLf
        .Collapse Direction:=0
        .Style = objDoc.Styles("Normal")
    End With
End Sub

Sub FecharSalvarWord()
    objDoc.SaveAs "diretorio/para/salvar/documento"
    objWord.Quit
    
    Set objDoc = Nothing
    Set objWord = Nothing
End Sub
