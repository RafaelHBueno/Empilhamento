Sub TransformData()
    
    Application.ScreenUpdating = False ' Desativa a atualização da interface gráfica
    
       
    Dim wsOriginal As Worksheet
    Dim wsResultado As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim newRow As Long
    Dim i As Long, j As Long, k As Long
    Dim novaAba As Worksheet
    
    'Cria uma nova aba
    Set novaAba = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    novaAba.Name = "Apontamento_Horas_copia"
    
    'Cola Valores da planilha Apontamento de Horas Original
    Sheets("Apontamento_Horas").Select
    Cells.Select
    Selection.Copy
    Sheets("Apontamento_Horas_copia").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
            
    'Cria Nova Planilha Com resultados
    Set novaAbaEmpilhadas = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    novaAbaEmpilhadas.Name = "Empilhadas"
        
     
        
    ' Defina as planilhas de origem e destino
    Set wsOriginal = ThisWorkbook.Sheets("Apontamento_Horas_copia")
    Set wsResultado = ThisWorkbook.Sheets("Empilhadas")
    
    ' Encontre a última linha e coluna da planilha original
    lastRow = wsOriginal.Cells(wsOriginal.Rows.Count, "A").End(xlUp).Row - 3
    lastCol = wsOriginal.Cells(1, wsOriginal.Columns.Count).End(xlToLeft).Column
    
    'Mover Colunas
    
    wsOriginal.Select
    Columns("F:F").Select
    Selection.Insert Shift:=x1ToRight
    Selection.Insert Shift:=x1ToRight
    Cells.Find(What:="R$/Mês", After:=ActiveCell, LookIn:=xlFormulas2, LookAt:= _
        xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    Columns(ActiveCell.Column).Select
    Selection.Cut Destination:=Columns("F:F")
        
    Cells.Find(What:="Horas", After:=ActiveCell, LookIn:=xlFormulas2, LookAt:= _
        xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    Columns(ActiveCell.Column).Select
    Selection.Cut Destination:=Columns("G:G")
    
    'deleta colunas vazias
    Range("AF:AF,AG:AG,BF:BF").Select
    Selection.Delete Shift:=xlToLeft
    
    'Cria Primeiro bloco de cabeçalho
    wsOriginal.Select
    Range(Cells(4, 1), Cells(lastRow + 3, 7)).Select
    Selection.Copy
    wsResultado.Select
    Range("A1").Select
    ActiveSheet.Paste
           
    'Selecionar e colar colunas DE: Projeto Para: Linhas
    newRow = 4
    i = 1
    j = 7
        
    For i = i To lastRow + 2
    For j = j + 1 To lastCol + 2
         
    
    wsResultado.Cells((lastRow * (i - 1)) + 1, 8).Value = wsOriginal.Cells(1, j).Value
    wsResultado.Cells((lastRow * (i - 1)) + 1, 9).Value = wsOriginal.Cells(2, j).Value
    wsResultado.Cells((lastRow * (i - 1)) + 1, 10).Value = wsOriginal.Cells(3, j).Value
    Range(Cells((lastRow * (i - 1)) + 1, 8), Cells((lastRow * (i - 1)) + 1, 10)).Select
    wsResultado.Select
    Selection.AutoFill Destination:=Range(Cells((lastRow * (i - 1)) + 1, 8), Cells(lastRow * i, 10)), Type:=xlFillCopy
    
                    
    'Preencha os dados na planilha de resultados
    wsResultado.Range(wsResultado.Cells((lastRow * (i - 1)) + 1, 11), wsResultado.Cells((lastRow * i), 11)).Value = wsOriginal.Range(wsOriginal.Cells(4, j), wsOriginal.Cells(lastRow + 3, j)).Value

    wsOriginal.Select
    Range(Cells(4, 1), Cells(lastRow + 3, 7)).Select
    Selection.Copy
    wsResultado.Select
    Cells(((lastRow * i) + 1), 1).Select
    ActiveSheet.Paste
                   
               
    newRow = newRow + 1  ' Avance para a próxima linha da planilha de resultados
    i = i + 1  ' +1i
            
    Next j
    Next i
    
        
        
    'Deleta ultimo bloco de cabeçalho
    Rows("28441:28441").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    
    
    'Cria cabeçalho linha 1
    Sheets("Empilhadas").Select
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown
    ActiveCell.Select
    Sheets("Apontamento_Horas_copia").Select
    Range("A3:G3").Select
    Selection.Copy
    Sheets("Empilhadas").Select
    ActiveSheet.Paste
    ActiveCell.Offset(0, 7).Range("A1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Tipo"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Setor"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "CC/Para"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Valor"
    ActiveCell.Offset(1, 0).Range("A1").Select
        
    Application.CutCopyMode = False
    Sheets("Empilhadas").Select
    wsResultado.Cells.Select
    Selection.ClearFormats
   
    Application.ScreenUpdating = True ' Reativa a atualização da interface gráfica
    
End Sub
