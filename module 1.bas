Attribute VB_Name = "module 1"
Sub addBorders(r As Range)
    r.Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub

'print statement
Sub 客户对账单()

    Application.ScreenUpdating = False

    'delete original data
    Sheets("客服对账单").Select
    Dim statementsRow As Integer
    Dim tableStart As String
    Dim r As Range
    Set r = Range("A:A").Find("出货日期", Range("A1"), xlValues, xlWhole, xlByColumns, xlNext, True)
    tableStart = "A" & r.row
    statementsRow = r.row + 1
    While Range("A" & statementsRow) <> ""
        Rows(statementsRow).Select
        Selection.Delete Shift:=xlUp
    Wend
    Rows(statementsRow + 1).Delete Shift:=xlUp
    
    'generate statement id
    Range("I3").Value = Date
    Dim today As String
    today = Format(Date, "YYYYMMDD")
    Dim todayStatementCount As Integer
    todayStatementCount = 1
    While Not Sheets("对账单汇总").Range("A:A").Find(today & "-" & todayStatementCount, Range("A1"), xlValues, xlWhole, xlByColumns, xlNext, True) Is Nothing
        todayStatementCount = todayStatementCount + 1
    Wend
    Dim statementsId As String
    statementsId = today & "-" & todayStatementCount
    Range("I4").Value = statementsId

    'fill in statement according to export
    Dim id As Integer
    id = 4
    Dim ignored As Object
    Set ignored = CreateObject("System.Collections.ArrayList")
    While Range("C" & id).Value <> ""
        Dim numbering As String
        numbering = Range("C" & id).Value
        
        Sheets("出货明细").Select
        Dim row As Integer
        row = 2
        While Range("C" & row).Value <> ""
            If Range("C" & row).Value = numbering Then
            
                If Range("K" & row).Value <> "" Then
                    ignored.Add Range("C" & row).Value
                    Sheets("客户对账单").Select
                    Range("C" & id).Value = ""
                    GoTo nextIteration
                End If
            
                Range("K" & row).Value = statementsId
                Range("B" & row & ":K" & row).Select
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .color = 5296274
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
        
                
                Dim itemId As String
                itemId = Range("D" & row).Value
                
                Dim outDate, itemName, color, size, note As String
                Dim count As Integer
                Dim unitPrice, totalPrice As Double
                outDate = Range("B" & row).Value
                itemName = Application.WorksheetFunction.XLookup(itemId, Sheets("零件汇总表").Range("D:D"), Sheets("零件汇总表").Range("E:E"))
                note = Range("I" & row).Value
                color = Application.WorksheetFunction.XLookup(itemId, Sheets("零件汇总表").Range("D:D"), Sheets("零件汇总表").Range("J:J"))
                size = Application.WorksheetFunction.XLookup(itemId, Sheets("零件汇总表").Range("D:D"), Sheets("零件汇总表").Range("K:K"))
                count = Range("H" & row).Value
                unitPrice = Application.WorksheetFunction.XLookup(itemId, Sheets("零件汇总表").Range("D:D"), Sheets("零件汇总表").Range("N:N"))
                totalPrice = count * unitPrice
            
                Sheets("客户对账单").Select
                Rows(statementsRow).Insert
                Range("A" & statementsRow).Value = outDate
                Range("B" & statementsRow).Value = numbering
                Range("C" & statementsRow).Value = itemId
                Range("D" & statementsRow).Value = itemName
                Range("E" & statementsRow).Value = color
                Range("F" & statementsRow).Value = size
                Range("G" & statementsRow).Value = count
                Range("H" & statementsRow).Value = unitPrice
                Range("I" & statementsRow).Value = totalPrice
                Range("J" & statementsRow).Value = note
                statementsRow = statementsRow + 1
            End If
            row = row + 1
            Sheets("出货明细").Select
        Wend
nextIteration:
        Sheets("客户对账单").Select
        id = id + 1
    Wend
    
    'calculate totals
    Sheets("客户对账单").Select
    Dim total As Double
    Range("A" & (statementsRow + 1)).Value = "合计"
    Range("G" & (statementsRow + 1)).Value = Application.WorksheetFunction.Sum(Range("G" & (r.row + 1) & ":G" & (statementsRow - 1)))
    total = Application.WorksheetFunction.Sum(Range("I" & (r.row + 1) & ":I" & (statementsRow - 1)))
    Range("I" & (statementsRow + 1)).Value = total
    Dim tableEnd As String
    tableEnd = "J" & (statementsRow + 1)
    
    'fill in summaries
    Sheets("对账单汇总").Select
    Dim writeRow As Integer
    writeRow = Sheets("对账单汇总").UsedRange.Rows.count + 1
    Range("A" & writeRow).Value = statementsId
    Range("C" & writeRow).Value = total
    Call addBorders(Range("A" & writeRow & ":E" & writeRow))
    Rows(writeRow).RowHeight = 24
    
    
    'format
    Sheets("客户对账单").Select
    Call addBorders(Range(tableStart & ":" & tableEnd))
    Range("A" & statementsRow & ":" & tableEnd).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Rows(statementsRow + 1).RowHeight = 24
    
    
    If ignored.count <> 0 Then
        MsgBox ("以下出货编号已对账， 已跳过:" & vbNewLine & Join(ignored.ToArray, vbNewLine))
    End If
    
    'print
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        "E:\" & Sheets("客户对账单").Range("C3") & Sheets("客户对账单").Range("I4") & "对账单.pdf"
    
    Range("C4:C" & (id - 1)).ClearContents
    
    Dim idEnd As Integer
    idEnd = r.row - 2
    
    If idEnd > 6 Then
        Range(7 & ":" & idEnd).Delete xlUp
    End If
    
    While idEnd < 6
        Rows(5).Insert
        idEnd = idEnd + 1
    Wend
    

End Sub

'print supplier order
Sub 供应商订单打印()

    Application.ScreenUpdating = False
    
    Dim currentSheet As String
    currentSheet = ActiveSheet.name
    
    'generate id
    Range("F3").Value = Date
    Dim today As String
    today = Format(Date, "YYYYMMDD")
    Dim todayOrderCount As Integer
    todayOrderCount = 1
    While Not Sheets("��Ӧ�̶���").Range("C:C").Find(today & "-" & todayOrderCount, Range("C1"), xlValues, xlWhole, xlByColumns, xlNext, True) Is Nothing
        todayOrderCount = todayOrderCount + 1
    Wend
    Dim orderId As String
    orderId = today & "-" & todayOrderCount
    Range("B3").Value = orderId
    
    'look up items
    Dim current As Integer
    current = 12
    While Range("A" & current).Value <> ""
        Dim itemId As String
        itemId = Range("A" & current).Value
        
        Dim name, color, size As String
        name = Range("B" & current).Value
        color = Range("C" & current).Value
        size = Range("D" & current).Value
        
        Dim outDate, note As String
        Dim count As Integer
        outDate = Range("E" & current).Value
        count = Range("F" & current).Value
        note = Range("G" & current).Value
        
        Sheets("供应商订单").Select
        Dim writeRow As Integer
        writeRow = Sheets("供应商订单").UsedRange.Rows.count + 1
        Range("A" & writeRow).Value = Date
        Range("B" & writeRow).Value = todayOrderCount
        Range("C" & writeRow).Value = orderId
        Range("D" & writeRow).Value = itemId
        Range("E" & writeRow).Value = name
        Range("F" & writeRow).Value = color
        Range("G" & writeRow).Value = size
        Range("H" & writeRow).Value = outDate
        Range("I" & writeRow).Value = count
        Range("J" & writeRow).Value = note
        Call addBorders(Range("A" & writeRow & ":J" & writeRow))
        Rows(writeRow).RowHeight = 24
        
        current = current + 1
        Sheets(currentSheet).Select
    Wend
    
    Sheets(currentSheet).Select
    Dim totalRow As Integer
    totalRow = Range("A:A").Find("订单要求：", Range("A1"), xlValues, xlWhole, xlByColumns, xlNext, True).row - 1
    Range("F" & totalRow).Value = Application.WorksheetFunction.Sum(Range("F12:F" & (totalRow - 2)))
    
    'print
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        "E:\" & Sheets("供应商订单").Range("C" & Sheets("供应商订单").UsedRange.Rows.count) & ".pdf"
    
    Range("A12:A" & current).ClearContents
    Range("E12:F" & current).ClearContents
    Range("F" & totalRow).ClearContents
    
    If current > 15 Then
        Range("15:" & (current - 1)).Delete xlUp
    End If
    While current < 15
        Rows(12).Insert
        current = current + 1
    Wend
End Sub

'print export
Sub 出货单()

    Application.ScreenUpdating = False
    
    Dim currentSheet As String
    currentSheet = ActiveSheet.name
    
    'generate id
    Range("F3").Value = Date
    Dim today As String
    today = Format(Date, "YYYYMMDD")
    Dim todayOrderCount As Integer
    todayOrderCount = 1
    While Not Sheets("出货明细").Range("C:C").Find(today & "-" & todayOrderCount, Range("C1"), xlValues, xlWhole, xlByColumns, xlNext, True) Is Nothing
        todayOrderCount = todayOrderCount + 1
    Wend
    Dim orderId As String
    orderId = today & "-" & todayOrderCount
    Range("B3").Value = orderId
    
    'look up items
    Dim current As Integer
    current = 12
    While Range("A" & current).Value <> ""
        Dim itemId As String
        itemId = Range("A" & current).Value
        
        Dim name, color, size As String
        name = Range("B" & current).Value
        color = Range("C" & current).Value
        size = Range("D" & current).Value
        
        Dim note As String
        Dim count As Integer
        count = Range("E" & current).Value
        note = Range("F" & current).Value
        
        Sheets("出货明细").Select
        Dim writeRow As Integer
        writeRow = Sheets("出货明细").UsedRange.Rows.count + 1
        Range("A" & writeRow).Value = todayOrderCount
        Range("B" & writeRow).Value = Date
        Range("C" & writeRow).Value = orderId
        Range("D" & writeRow).Value = itemId
        Range("E" & writeRow).Value = name
        Range("F" & writeRow).Value = color
        Range("G" & writeRow).Value = size
        Range("H" & writeRow).Value = count
        Range("I" & writeRow).Value = note
        Call addBorders(Range("A" & writeRow & ":K" & writeRow))
        
        current = current + 1
        Sheets(currentSheet).Select
    Wend
    
    Sheets(currentSheet).Select
    Dim totalRow As Integer
    totalRow = Range("A:A").Find("合计", Range("A1"), xlValues, xlWhole, xlByColumns, xlNext, True).row
    Range("E" & totalRow).Value = Application.WorksheetFunction.Sum(Range("E12:E" & (totalRow - 2)))
        
    'print
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        "E:\" & Sheets("出货单").Range("B3") & ".pdf"
    
    Range("A12:A" & current).ClearContents
    Range("E12:F" & current).ClearContents
    Range("E" & totalRow).ClearContents
    
    If current > 15 Then
        Range("15:" & (current - 1)).Delete xlUp
    End If
    While current < 15
        Rows(12).Insert
        current = current + 1
    Wend
End Sub

'print sample
Sub 样品寄样单()

    Application.ScreenUpdating = False
    
    Dim currentSheet As String
    currentSheet = ActiveSheet.name
    
    'delete original data
    Dim writeRow As Integer
    writeRow = 12
    Dim tableStart As String
    tableStart = "A12"
    While Range("A" & writeRow) <> ""
        Rows(writeRow).Select
        Selection.Delete Shift:=xlUp
    Wend
    Rows(writeRow + 1).Delete Shift:=xlUp
    
    'generate id
    Dim orderId As String
    orderId = Range("B3").Value
    
    'look up items
    Dim current As Integer
    current = 2
    Sheets("样品零件汇总表").Select
    While Range("D" & current).Value <> ""
    
        If Range("M" & current).Value = orderId Then
        
            Dim itemId, name, blueprint, factoryVersion, note, color, size, sendDate, sendId, version As String
            
            itemId = Range("D" & current).Value
            name = Range("E" & current).Value
            blueprint = Range("G" & current).Value
            factoryVersion = Range("H" & current).Value
            note = Range("I" & current).Value
            color = Range("J" & current).Value
            size = Range("K" & current).Value
            sendDate = Range("L" & current).Value
            sendId = Range("M" & current).Value
            version = Range("N" & current).Value
            
            Sheets(currentSheet).Select
            Rows(writeRow).Insert
            Range("A" & writeRow).Value = itemId
            Range("B" & writeRow).Value = name
            Range("D" & writeRow).Value = blueprint
            Range("E" & writeRow).Value = factoryVersion
            Range("F" & writeRow).Value = note
            Range("G" & writeRow).Value = color
            Range("H" & writeRow).Value = size
            Range("I" & writeRow).Value = sendDate
            Range("J" & writeRow).Value = sendId
            Range("K" & writeRow).Value = version
            
            writeRow = writeRow + 1
            
            Sheets("样品零件汇总表").Select
            
        End If
        current = current + 1
    Wend
    
    Sheets(currentSheet).Select
    Dim tableEnd As String
    tableEnd = "K" & (writeRow + 1)
    Call addBorders(Range(tableStart & ":" & tableEnd))
    Range("A" & writeRow & ":" & tableEnd).Select
        With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorDark2
                .TintAndShade = 0
                .PatternTintAndShade = 0
        End With
    Rows(writeRow + 1).RowHeight = 24
        
    'print
    'ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
       "E:\" & Sheets("样品寄样单").Range("B3") & ".pdf"
    
    Range("B3").ClearContents
End Sub

'print sample price
Sub 样品报价单()
    
    Application.ScreenUpdating = False
    
    Dim currentSheet As String
    currentSheet = ActiveSheet.name
    
    'delete original data
    Dim writeRow As Integer
    writeRow = 12
    Dim tableStart As String
    tableStart = "A12"
    While Range("A" & writeRow) <> ""
        Rows(writeRow).Select
        Selection.Delete Shift:=xlUp
    Wend
    Rows(writeRow + 1).Delete Shift:=xlUp
    
    'generate id
    Dim orderId As String
    orderId = Range("B3").Value
    
    'look up items
    Dim current As Integer
    current = 2
    Sheets("样品零件汇总表").Select
    While Range("D" & current).Value <> ""
    
        If Range("T" & current).Value = orderId Then
            Dim itemId, name, blueprint, note, color, size As String
            Dim price As Double
            
            itemId = Range("D" & current).Value
            
            name = Range("E" & current).Value
            blueprint = Range("G" & current).Value
            note = Range("I" & current).Value
            color = Range("J" & current).Value
            size = Range("K" & current).Value
            price = Range("R" & current).Value
                    
            Sheets(currentSheet).Select
            Rows(writeRow).Insert
            Range("A" & writeRow).Value = itemId
            Range("B" & writeRow).Value = name
            Range("D" & writeRow).Value = blueprint
            Range("F" & writeRow).Value = note
            Range("G" & writeRow).Value = color
            Range("H" & writeRow).Value = size
            Range("O" & writeRow).Value = price
            
            writeRow = writeRow + 1
                
            Sheets("样品零件汇总表").Select
            
        End If
        current = current + 1
    Wend
    
    Sheets(currentSheet).Select
    Dim tableEnd As String
    tableEnd = "O" & (writeRow + 1)
    Call addBorders(Range(tableStart & ":" & tableEnd))
    Range("A" & writeRow & ":" & tableEnd).Select
        With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorDark2
                .TintAndShade = 0
                .PatternTintAndShade = 0
        End With
    Rows(writeRow + 1).RowHeight = 24
    
    'print
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        "E:\" & Sheets("报价单").Range("B3") & ".pdf"
    
    Range("B3").ClearContents
End Sub



