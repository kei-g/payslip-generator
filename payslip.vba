'
' Copyright (c) 2019-, kei-g
' All rights reserved.
'
' Redistribution and use in source and binary forms, with or without
' modification, are permitted provided that the following conditions are met:
'
' 1. Redistributions of source code must retain the above copyright notice, this
'    list of conditions and the following disclaimer.
'
' 2. Redistributions in binary form must reproduce the above copyright notice,
'    this list of conditions and the following disclaimer in the documentation
'    and/or other materials provided with the distribution.
'
' 3. Neither the name of the copyright holder nor the names of its
'    contributors may be used to endorse or promote products derived from
'    this software without specific prior written permission.
'
' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
' AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
' IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
' DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
' FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
' DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
' SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
' CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
' OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
' OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'

Function Interpolate(ByVal format As String, ParamArray args() As Variant) As String
    Dim a() As String: a = Split(format, "{")
    Dim result As String: result = a(0)
    For i = 1 To UBound(a)
        Dim b() As String: b = Split(a(i), "}")
        Dim idx As Integer: idx = CInt(b(0))
        result = result & args(idx) & b(1)
    Next i
    Interpolate = result
End Function

Function IsBlankAll(ByVal Target As Range) As Boolean
    For Each c In Target
        If c.Value <> "" Then
            IsBlankAll = False
            Exit Function
        End If
    Next c
    IsBlankAll = True
End Function

Function IsBlankAny(ByVal Target As Range) As Boolean
    For Each c In Target
        If c.Value = "" Then
            IsBlankAny = True
            Exit Function
        End If
    Next c
    IsBlankAny = False
End Function

Function Round15A(ByVal vTime As Double) As Double
    Dim iH As Integer, iM As Integer
    iH = Hour(vTime)
    iM = Minute(vTime)
    If iM <= 5 Then
        iM = 0
    ElseIf iM <= 20 Then
        iM = 15
    ElseIf iM <= 35 Then
        iM = 30
    ElseIf iM <= 50 Then
        iM = 45
    Else
        iH = iH + 1
        iM = 0
    End If
    Round15A = TimeSerial(iH, iM, 0)
End Function

Function Round15B(ByVal vTime As Double) As Double
    Dim iH As Integer, iM As Integer
    iH = Hour(vTime)
    iM = Minute(vTime)
    If iM < 10 Then
        iM = 0
    ElseIf iM < 25 Then
        iM = 15
    ElseIf iM < 40 Then
        iM = 30
    ElseIf iM < 55 Then
        iM = 45
    Else
        iH = iH + 1
        iM = 0
    End If
    Round15B = TimeSerial(iH, iM, 0)
End Function

Function SumIfBlank(ByVal rTarget As Range, ByVal rSum As Range) As Double
    Dim vSum As Double: vSum = 0
    For i = 1 To rTarget.Rows.Count
        If rTarget(i, 1) = "" Then
            For j = 1 To rSum.Columns.Count
                Dim v As Variant: v = rSum(i, j)
                If IsNumeric(v) Then
                    vSum = vSum + v
                End If
            Next j
        End If
    Next i
    SumIfBlank = vSum
End Function

Function SumUnlessBlank(ByVal rTarget As Range, ByVal rSum As Range) As Double
    Dim vSum As Double: vSum = 0
    For i = 1 To rTarget.Rows.Count
        If rTarget(i, 1) <> "" Then
            For j = 1 To rSum.Columns.Count
                Dim v As Variant: v = rSum(i, j)
                If IsNumeric(v) Then
                    vSum = vSum + v
                End If
            Next j
        End If
    Next i
    SumUnlessBlank = vSum
End Function

Function UnlessBlankAll(ByVal Target As Range, ByVal Expr As Variant) As Variant
    If IsBlankAll(Target) Then
        UnlessBlankAll = ""
    Else
        UnlessBlankAll = Expr
    End If
End Function

Function UnlessBlankAny(ByVal Target As Range, ByVal Expr As Variant) As Variant
    If IsBlankAny(Target) Then
        UnlessBlankAny = ""
    Else
        UnlessBlankAny = Expr
    End If
End Function

Function IsNationalHoliday(ByVal dtDate As Date) As Boolean
    Dim cNH As New CNationalHoliday
    IsNationalHoliday = cNH.IsNationalHoliday(dtDate)
End Function

Function GetNationalHolidayName(ByVal dtDate As Date) As String
    Dim cNH As New CNationalHoliday
    Dim sName As String
    If cNH.isNationalHoliday2(dtDate, sName) Then
        GetNationalHolidayName = sName
    Else
        GetNationalHolidayName = ""
    End If
End Function

Function GetMaxRow(wSheet As Worksheet, targetCol As Long) As Long
    GetMaxRow = wSheet.Cells(wSheet.Rows.Count, targetCol).End(xlUp).Row
End Function

Private Function DoesWorksheetExist(ByVal sName As String) As Boolean
    Dim wSheet As Worksheet
    For Each wSheet In Sheets
        If wSheet.Name = sName Then
            DoesWorksheetExist = True
            Exit Function
        End If
    Next wSheet
    DoesWorksheetExist = False
End Function

Private Sub SetHeader(ByVal wSheet As Worksheet, ByVal iRow As Long, ByVal iColumnStart As Long, ByVal sCSV As String)
    Dim aCSV As Variant
    aCSV = Split(sCSV, ",")
    Dim i As Long
    i = 0
    For Each v In aCSV
        wSheet.Cells(iRow, iColumnStart + i) = v
        i = i + 1
    Next v
End Sub

Sub GenerateTimesheet(ByRef cNH As CNationalHoliday, ByVal wSheet As Worksheet, ByVal sName As String, ByVal iFound As Long)
    Dim wTimesheet As Worksheet: Set wTimesheet = Worksheets.Add(after:=Worksheets(Worksheets.Count))
    wTimesheet.Name = sName
    wSheet.Activate
    wTimesheet.Range("A2:C2").Merge
    wTimesheet.Cells(2, 1) = "日付"
    wTimesheet.Cells(2, 1).HorizontalAlignment = xlCenter
    Call SetHeader(wTimesheet, 2, 4, "出勤,退勤,出勤,退勤,勤務時間,18時まで,18時以降,出勤,退勤,出勤,退勤,勤務時間,18時まで,18時以降,昼勤,夜勤,★日給★")
    Dim dStart As Date, dCurrent As Date, dEnd As Date
    dStart = DateSerial(Year(Now), Month(Now) - 1, 1)
    dCurrent = dStart
    dEnd = DateSerial(Year(Now), Month(Now), 0)
    For i = 1 To Day(dEnd)
        Dim j As Integer: j = i + 2
        wTimesheet.Cells(j, 1) = dCurrent
        wTimesheet.Cells(j, 1).NumberFormatLocal = "m""月""d""日"""
        wTimesheet.Cells(j, 2) = WeekdayName(Weekday(dCurrent), True)
        Select Case Weekday(dCurrent)
        Case vbWednesday
            wTimesheet.Cells(j, 3) = "定休日"
            wTimesheet.Range(Interpolate("A{0}:C{0}", j)).Font.Color = RGB(0, 128, 255)
        Case vbSaturday
            wTimesheet.Range(Interpolate("A{0}:C{0}", j)).Font.Color = RGB(0, 0, 255)
        Case vbSunday
            wTimesheet.Cells(j, 3) = "休日"
            wTimesheet.Range(Interpolate("A{0}:C{0}", j)).Font.Color = RGB(255, 0, 0)
        End Select
        Dim sHolidayName As String
        If cNH.isNationalHoliday2(dCurrent, sHolidayName) Then
            wTimesheet.Cells(j, 3) = sHolidayName
            wTimesheet.Range(Interpolate("A{0}:C{0}", j)).Font.Color = RGB(255, 0, 0)
        End If
        dCurrent = dCurrent + 1
    Next i
    j = Day(dEnd) + 2
    wTimesheet.Range("D3:E" & j & ",K3:L" & j).Locked = False
    wTimesheet.Range("F3:F" & j & ",M3:M" & j).Formula = "=IF(ISBLANK(D3),"""",Round15A(D3))"
    wTimesheet.Range("G3:G" & j & ",N3:N" & j).Formula = "=IF(ISBLANK(E3),"""",Round15B(E3))"
    wTimesheet.Range("H3:H" & j & ",O3:O" & j).Formula = "=UnlessBlankAny(F3:G3,G3-F3)"
    wTimesheet.Range("I3:I" & j & ",P3:P" & j).Formula = "=UnlessBlankAny(F3:G3,MIN(G3,18/24)-MIN(F3,18/24))"
    wTimesheet.Range("J3:J" & j & ",Q3:Q" & j).Formula = "=UnlessBlankAny(F3:G3,MIN(MAX(G3,18/24)-18/24,H3))"
    wTimesheet.Range("R3:S" & j).Formula = "=UnlessBlankAll((I3,P3),SUM(I3,P3))"
    wTimesheet.Range("T3:T" & j).Formula = Interpolate("=UnlessBlankAll(R3:S3,IF(ISBLANK(C3),(R3*{0}!C${1}+S3*({0}!C${1}+100))*24,(R3+S3)*({0}!C${1}+100)*24))", wSheet.Name, iFound)
    j = j + 1
    wTimesheet.Range("F3:J" & j & ",M3:T" & j).Interior.Color = RGB(192, 192, 192)
    wTimesheet.Range(Interpolate("H{0}:J{0},O{0}:S{0}", j)).Formula = Interpolate("=SUM(H3:H{0})", j - 1)
    wTimesheet.Range("F3:S" & j).NumberFormatLocal = "[h]:mm"
    wTimesheet.Cells(j, 20) = "=ROUNDDOWN(SUM(T3:T" & (j - 1) & "),0)"
    wTimesheet.Cells(j + 1, 19) = "交通費"
    wTimesheet.Cells(j + 1, 20) = Interpolate("=COUNTA(D3:D{0})*{1}!D{2}", j - 1, wSheet.Name, iFound)
    wTimesheet.Range(Interpolate("S{0}:T{0}", j + 1)).Interior.Color = RGB(192, 192, 192)
    wTimesheet.Columns("A:T").AutoFit
    wTimesheet.Protect
End Sub

Private Sub SetLineStyle(ByVal bs As Border, ByVal iLineStyle As Integer, ByVal iLineWeight As Integer)
    bs.LineStyle = iLineStyle
    bs.Weight = iLineWeight
End Sub

Private Sub SetBordersEdge(ByVal rTarget As Range, ByVal iLineStyle As Integer, ByVal iLineWeight As Integer)
    Call SetLineStyle(rTarget.Borders(xlEdgeBottom), iLineStyle, iLineWeight)
    Call SetLineStyle(rTarget.Borders(xlEdgeLeft), iLineStyle, iLineWeight)
    Call SetLineStyle(rTarget.Borders(xlEdgeTop), iLineStyle, iLineWeight)
    Call SetLineStyle(rTarget.Borders(xlEdgeRight), iLineStyle, iLineWeight)
End Sub

Private Sub SetBordersInside(ByVal rTarget As Range, ByVal iLineStyle As Integer, ByVal iLineWeight As Integer)
    Call SetLineStyle(rTarget.Borders(xlInsideHorizontal), iLineStyle, iLineWeight)
    Call SetLineStyle(rTarget.Borders(xlInsideVertical), iLineStyle, iLineWeight)
End Sub

Sub GeneratePayslip(ByVal wSheet As Worksheet, ByVal sName As String, ByVal iFound As Long)
    Dim dtStart As Date: dtStart = DateSerial(Year(Now()), Month(Now()) - 1, 1)
    Dim dtEnd As Date: dtEnd = DateSerial(Year(Now()), Month(Now()), 0)
    Dim iDays As Integer: iDays = Day(dtEnd)

    Dim wPayslip As Worksheet: Set wPayslip = Worksheets.Add(after:=Worksheets(Worksheets.Count))
    wPayslip.Name = wSheet.Cells(iFound, 2) & "_給与明細"
    wSheet.Activate

    wPayslip.Cells.Font.Name = "ＭＳ Ｐゴシック"
    wPayslip.Cells.Font.Size = 12

    wPayslip.Range("A1:A1").ColumnWidth = 3
    wPayslip.Range("B1:B1").ColumnWidth = 5.25
    wPayslip.Range("C1:C1").ColumnWidth = 16.38
    wPayslip.Range("D1:F1").ColumnWidth = 4.38
    wPayslip.Range("G1:G1").ColumnWidth = 6
    wPayslip.Range("H1:J1").ColumnWidth = 4.5
    wPayslip.Range("K1:L1").ColumnWidth = 5.13
    wPayslip.Range("A1:A1").RowHeight = 24
    wPayslip.Range("A2,A4,A6,A9:A10").RowHeight = 15
    wPayslip.Range("A3:A3").RowHeight = 21
    wPayslip.Range("A5,A12:A27").RowHeight = 23.25
    wPayslip.Range("A7:A8").RowHeight = 18
    wPayslip.Range("A11").RowHeight = 13.5

    wPayslip.Range("C1:K1").Merge
    wPayslip.Range("C1") = "給 料 支 払 明 細 書"
    wPayslip.Range("C1").Font.Bold = True
    wPayslip.Range("C1").Font.Size = 16
    wPayslip.Range("C1").HorizontalAlignment = xlCenter

    wPayslip.Range("D2:H2").Borders(xlEdgeTop).LineStyle = xlContinuous

    wPayslip.Range("C3,D3") = dtStart
    wPayslip.Range("C3").NumberFormatLocal = "ggg"
    wPayslip.Range("D3").NumberFormatLocal = "e"
    wPayslip.Range("E3") = "年"
    wPayslip.Range("E3").HorizontalAlignment = xlCenter
    wPayslip.Range("F3") = dtStart
    wPayslip.Range("F3").NumberFormatLocal = "M"
    wPayslip.Range("G3") = "月分"
    wPayslip.Range("G3").HorizontalAlignment = xlLeft
    wPayslip.Range("I3,K3") = DateSerial(Year(Now()), Month(Now()), 5)
    wPayslip.Range("I3").NumberFormatLocal = "M"
    wPayslip.Range("J3") = "月"
    wPayslip.Range("K3").NumberFormatLocal = "d"
    wPayslip.Range("L3") = "日"

    wPayslip.Range("C5") = "=" & wSheet.Name & "!B" & iFound
    wPayslip.Range("C5").Borders(xlEdgeBottom).LineStyle = xlContinuous
    wPayslip.Range("C5,D5,H5").Font.Bold = True
    wPayslip.Range("C5,H5").Font.Size = 14
    wPayslip.Range("D5") = "様"
    wPayslip.Range("H5") = "西脇大橋ラーメン"

    wPayslip.Range("C7:C8").Merge
    wPayslip.Range("C7") = "労働日数"
    wPayslip.Range("C7").HorizontalAlignment = xlCenter
    wPayslip.Range("D7") = "自"
    wPayslip.Range("D8") = "至"
    wPayslip.Range("E7,G7") = dtStart
    wPayslip.Range("E8,G8") = dtEnd
    wPayslip.Range("E7:E8").NumberFormatLocal = "M"
    wPayslip.Range("G7:G8").NumberFormatLocal = "d"
    wPayslip.Range("F7:F8") = "月"
    wPayslip.Range("H7:H8") = "日"
    wPayslip.Range("I7:I8").Merge
    wPayslip.Range("I7") = "=COUNTA(" & sName & "!D3:D33)"
    wPayslip.Range("J7:J8").Merge
    wPayslip.Range("J7") = "日"
    wPayslip.Range("C9:C10").Merge
    wPayslip.Range("C9") = "労働時間"
    wPayslip.Range("C9").HorizontalAlignment = xlCenter
    wPayslip.Range("D9:F10").Merge
    wPayslip.Range("D9") = "=SUM(" & sName & "!R" & (iDays + 3) & ":S" & (iDays + 3) & ")*24"
    wPayslip.Range("G9:H10").Merge
    wPayslip.Range("G9") = "時間"
    wPayslip.Range("G9").HorizontalAlignment = xlCenter
    wPayslip.Range("I9:I10").Merge
    wPayslip.Range("J9:J10").Merge
    Call SetBordersEdge(wPayslip.Range("C7:J10"), xlContinuous, xlMedium)
    wPayslip.Range("C7:C10").Borders(xlInsideHorizontal).LineStyle = xlContinuous
    wPayslip.Range("C7:C10").Borders(xlEdgeRight).LineStyle = xlContinuous
    wPayslip.Range("D8:J10").Borders(xlInsideHorizontal).LineStyle = xlContinuous

    wPayslip.Range("B12:B23").Merge
    wPayslip.Range("B12") = "支" & vbCrLf & "給" & vbCrLf & "額"
    wPayslip.Range("B12").HorizontalAlignment = xlCenter
    For i = 12 To 22
        wPayslip.Range("C" & i & ":D" & i).Merge
        wPayslip.Range("E" & i & ":F" & i).Merge
        wPayslip.Range("G" & i & ":H" & i).Merge
        wPayslip.Range("I" & i & ":K" & i).Merge
    Next i
    wPayslip.Range("C12") = "摘要"
    wPayslip.Range("E12") = "時間"
    wPayslip.Range("G12") = "単価"
    wPayslip.Range("I12") = "金額"
    wPayslip.Range("C12,E12,G12,I12,C13:C27").HorizontalAlignment = xlCenter
    wPayslip.Range("C13") = "基本給"
    wPayslip.Range("C14") = "所定時間外"
    wPayslip.Range("C15") = "家族手当"
    wPayslip.Range("C16") = "日・祝祭日手当"
    wPayslip.Range("E16") = "=SumUnlessBlank(" & sName & "!C3:C" & (iDays + 2) & "," & sName & "!R3:S" & (iDays + 2) & ")*24"
    wPayslip.Range("G16") = "=" & wSheet.Name & "!C" & iFound & "+100"
    wPayslip.Range("C17") = "昼出勤手当"
    wPayslip.Range("E17") = "=SumIfBlank(" & sName & "!C3:C" & (iDays + 2) & "," & sName & "!R3:R" & (iDays + 2) & ")*24"
    wPayslip.Range("G17") = "=" & wSheet.Name & "!C" & iFound
    wPayslip.Range("C18") = "夜出勤手当"
    wPayslip.Range("E18") = "=SumIfBlank(" & sName & "!C3:C" & (iDays + 2) & "," & sName & "!S3:S" & (iDays + 2) & ")*24"
    wPayslip.Range("G18") = "=" & wSheet.Name & "!C" & iFound & "+100"
    wPayslip.Range("C19") = "特別手当"
    wPayslip.Range("C20") = "臨時手当"
    wPayslip.Range("C21") = "通勤手当"
    wPayslip.Range("E21") = "=COUNTA(" & sName & "!D3:D" & (iDays + 2) & ")"
    wPayslip.Range("G21") = "=" & wSheet.Name & "!D" & iFound
    wPayslip.Range("I13:I22").Formula = "=UnlessBlankAny((E13,G13),ROUNDDOWN(E13*G13,0))"
    wPayslip.Range("I23") = "=SUM(I13:I22)"
    wPayslip.Range("C23:H23").Merge
    wPayslip.Range("C23") = "支給額合計"
    wPayslip.Range("I23:K23").Merge

    wPayslip.Range("B24:B26").Merge
    wPayslip.Range("B24") = "控" & vbCrLf & "除" & vbCrLf & "額"
    wPayslip.Range("B24").HorizontalAlignment = xlCenter
    For i = 24 To 25
        wPayslip.Range("C" & i & ":D" & i).Merge
        wPayslip.Range("E" & i & ":F" & i).Merge
        wPayslip.Range("G" & i & ":H" & i).Merge
        wPayslip.Range("I" & i & ":K" & i).Merge
    Next i
    wPayslip.Range("C24") = "年末調整"
    wPayslip.Range("I24") = 0
    wPayslip.Range("C25") = "特別徴収（市民税・県民税）"
    wPayslip.Range("I25") = "=" & wSheet.Name & "!E" & iFound
    wPayslip.Range("C25").Font.Size = 10
    wPayslip.Range("C26:H26").Merge
    wPayslip.Range("C26") = "控除額合計"
    wPayslip.Range("I26:K26").Merge
    wPayslip.Range("I26") = "=SUM(I24:I25)"

    wPayslip.Range("C27:H27").Merge
    wPayslip.Range("C27") = "差引支給額"
    wPayslip.Range("I27:K27").Merge
    wPayslip.Range("I27") = "=I23-I26"

    Call SetBordersEdge(wPayslip.Range("B12:K26"), xlContinuous, xlMedium)
    Call SetBordersInside(wPayslip.Range("B12:K26"), xlContinuous, xlThin)
    Call SetBordersEdge(wPayslip.Range("C23:K23"), xlContinuous, xlMedium)
    Call SetBordersEdge(wPayslip.Range("C26:K26"), xlContinuous, xlMedium)
    Call SetBordersEdge(wPayslip.Range("C27:K27"), xlContinuous, xlMedium)
    Call SetBordersInside(wPayslip.Range("C27:K27"), xlContinuous, xlThin)

    wPayslip.PageSetup.Orientation = xlPortrait
    wPayslip.PageSetup.PaperSize = xlPaperB5
    wPayslip.PageSetup.LeftMargin = Application.CentimetersToPoints(1)
    wPayslip.PageSetup.RightMargin = Application.CentimetersToPoints(0.5)
End Sub

Sub 一括生成()
    Dim cNH As New CNationalHoliday
    Dim wSheet As Worksheet: Set wSheet = Worksheets("Sheet1")
    Dim iEnd As Long
    iEnd = GetMaxRow(wSheet, 1)
    Dim sErrorMsg As String: sErrorMsg = ""
    For i = 5 To iEnd
        Dim sNameTimesheet As String: sNameTimesheet = Cells(i, 2) & "_タイムシート"
        Dim sNamePayslip As String: sNamePayslip = Cells(i, 2) & "_給与明細"
        If DoesWorksheetExist(sNameTimesheet) Then
            sErrorMsg = sNameTimesheet & "は既に存在します" & vbCrLf
        ElseIf DoesWorksheetExist(sNamePayslip) Then
            sErrorMsg = sNamePayslip & "は既に存在します" & vbCrLf
        Else
            Call GenerateTimesheet(cNH, wSheet, sNameTimesheet, i)
            Call GeneratePayslip(wSheet, sNameTimesheet, i)
        End If
    Next i
    If sErrorMsg <> "" Then
        MsgBox sErrorMsg
    End If
End Sub

