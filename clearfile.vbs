Sub CreateFloorList()
    Dim floorRangeSelect As Range
    Dim floorTextSelect As Range
    Dim floorRange As Range
    Dim floorText As Range
    Set floorText = ActiveSheet.Range("E2:E850")
    Set floorRange = ActiveSheet.Range("H2:H850")
    
    For Each floorTextSelect In floorText.Cells
        If InStr(1, floorTextSelect.Value, "CSI") Then
            floorTextSelect.Offset(0, 3).Value = "CSI"
        ElseIf InStr(1, floorTextSelect.Value, "Eulalia") Then
            floorTextSelect.Offset(0, 3).Value = "Eulalia"
        ElseIf InStr(1, floorTextSelect.Value, "/1") Or InStr(1, floorTextSelect.Value, "/1 -") Or InStr(1, floorTextSelect.Value, "ITS") Then
            floorTextSelect.Offset(0, 3).Value = 1
        ElseIf InStr(1, floorTextSelect.Value, "/2 -") Or InStr(1, floorTextSelect.Value, "/2") Then
            floorTextSelect.Offset(0, 3).Value = 2
        ElseIf InStr(1, floorTextSelect.Value, "/3 -") Or InStr(1, floorTextSelect.Value, "/3") Then
            floorTextSelect.Offset(0, 3).Value = 3
        ElseIf InStr(1, floorTextSelect.Value, "/4 -") Then
            floorTextSelect.Offset(0, 3).Value = 4
        ElseIf InStr(1, floorTextSelect.Value, "/5 -") Then
            floorTextSelect.Offset(0, 3).Value = 5
        ElseIf InStr(1, floorTextSelect.Value, "/6 -") Then
            floorTextSelect.Offset(0, 3).Value = 6
        ElseIf InStr(1, floorTextSelect.Value, "/7 -") Then
            floorTextSelect.Offset(0, 3).Value = 7
        ElseIf InStr(1, floorTextSelect.Value, "/8 -") Or InStr(1, floorTextSelect.Value, "TEL") Or InStr(1, floorTextSelect.Value, "/8") Then
            floorTextSelect.Offset(0, 3).Value = 8
        ElseIf InStr(1, floorTextSelect.Value, "/B") Then
            floorTextSelect.Offset(0, 3).Value = "Basement"
        ElseIf InStr(1, floorTextSelect.Value, "/R") Then
            floorTextSelect.Offset(0, 3).Value = "Roof"
        Else
            floorTextSelect.Offset(0, 3).Value = ""
        End If
    Next
End Sub

Sub STEP_ONE_RemoveNonWindows7()
    'step one will clear out non windows 7
    
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    ActiveSheet.Range("H1").Value = "Floor"
    Columns("A:I").Select
    Selection.AutoFilter
    
    ActiveSheet.Range("$A$1:$H$1222").AutoFilter Field:=7, Criteria1:=Array( _
        "aix_version_5.2", "cisco_ios_version_12_2_55_se", "Linux", _
        "Microsoft Windows 10 Enterprise", "SuSE(Linux)", "Ubuntu(Linux)", "vmnix-x86", _
        "Windows 2000", "Windows 2003", "Windows 2003 R2", "Windows 2008", _
        "Windows 2008 R2", "Windows 2012 R2", "Windows 2012 Standard", "="), Operator:= _
        xlFilterValues
    Range("A9:G1222").Select
    Selection.EntireRow.Delete
    ActiveSheet.Range("$A$1:$H$835").AutoFilter Field:=7
    
    CreateFloorList
End Sub

Sub STEP_TWO_RemoveNonAllCTs()
    Dim pcNames As Range
    Dim pcName As Range
    Set pcNames = ActiveSheet.Range("B2:B900")
    
    For Each pcName In pcNames.Cells
        If InStr(1, pcName.Value, "gmhe") Then
            pcName.Offset(0, 8).Value = "Delete"
        ElseIf InStr(1, pcName.Value, "gmhs") Then
            pcName.Offset(0, 8).Value = "Delete"
        ElseIf InStr(1, pcName.Value, "gmhc") Then
            pcName.Offset(0, 8).Value = "Delete"
        ElseIf InStr(1, pcName.Value, "gmhy") Then
            pcName.Offset(0, 8).Value = "Delete"
        ElseIf InStr(1, pcName.Value, "gmhz") Then
            pcName.Offset(0, 8).Value = "Delete"
        Else
            pcName.Offset(0, 8).Value = ""
        End If
    Next
    DeleteNonCTS
End Sub

Sub DeleteNonCTS()
    Dim LR As Long, i As Long
    LR = Range("J" & Rows.count).End(xlUp).Row
    For i = LR To 2 Step -1
        If Range("J" & i).Value Like "Delete" Then Rows(i).Delete
    Next i
End Sub

Sub DeleteNonDeployed()
    Dim LR As Long, i As Long
    LR = Range("I" & Rows.count).End(xlUp).Row
    For i = LR To 2 Step -1
        If Not Range("I" & i).Value Like "Deployed" Then Rows(i).Delete
    Next i

End Sub


Sub Stats()
    
    Dim x As Range
    Dim x1 As Range
    Dim count As Integer
    
    ActiveSheet.Range("J7").Value = "Floor:"
    
    Dim lblArr() As Variant
    lblArr = Array("1", "2", "3", "4", "5", "6", "7", "8", "Basement", "Roof", "CSI", "Eulalia", "", "Desktop", "Laptop")
    
    Dim formulaArr() As Variant
    formulaArr = Array("=COUNTIF(R[-6]C[-3]:R[892]C[-3], ""1"")", _
    "=COUNTIF(R[-7]C[-3]:R[891]C[-3],""2"")", _
    "=COUNTIF(R[-8]C[-3]:R[890]C[-3],""3"")", _
    "=COUNTIF(R[-9]C[-3]:R[889]C[-3],""4"")", _
    "=COUNTIF(R[-10]C[-3]:R[888]C[-3],""5"")", _
    "=COUNTIF(R[-11]C[-3]:R[887]C[-3],""6"")", _
    "=COUNTIF(R[-12]C[-3]:R[886]C[-3],""7"")", _
    "=COUNTIF(R[-13]C[-3]:R[885]C[-3],""8"")", _
    "=COUNTIF(R[-14]C[-3]:R[884]C[-3],""Basement"")", _
    "=COUNTIF(R[-15]C[-3]:R[883]C[-3],""Roof"")", _
    "=COUNTIF(R[-16]C[-3]:R[882]C[-3],""CSI"")", _
    "=COUNTIF(R[-17]C[-3]:R[881]C[-3], ""Eulalia"")")
    count = 0
    Set x1 = ActiveSheet.Range("J8:J19")
    For Each x In x1
        x.Value = lblArr(count)
        x.Offset(0, 1).Value = formulaArr(count)
        count = count + 1
    Next
    
    
End Sub

Sub CreateFloorList()
    Dim floorRangeSelect As Range
    Dim floorTextSelect As Range
    Dim floorRange As Range
    Dim floorText As Range
    Set floorText = ActiveSheet.Range("E2:E850")
    Set floorRange = ActiveSheet.Range("H2:H850")
    
    For Each floorTextSelect In floorText.Cells
        If InStr(1, floorTextSelect.Value, "CSI") Then
            floorTextSelect.Offset(0, 3).Value = "CSI"
        ElseIf InStr(1, floorTextSelect.Value, "Eulalia") Then
            floorTextSelect.Offset(0, 3).Value = "Eulalia"
        ElseIf InStr(1, floorTextSelect.Value, "/1") Or InStr(1, floorTextSelect.Value, "/1 -") Or InStr(1, floorTextSelect.Value, "ITS") Then
            floorTextSelect.Offset(0, 3).Value = 1
        ElseIf InStr(1, floorTextSelect.Value, "/2 -") Or InStr(1, floorTextSelect.Value, "/2") Then
            floorTextSelect.Offset(0, 3).Value = 2
        ElseIf InStr(1, floorTextSelect.Value, "/3 -") Or InStr(1, floorTextSelect.Value, "/3") Then
            floorTextSelect.Offset(0, 3).Value = 3
        ElseIf InStr(1, floorTextSelect.Value, "/4 -") Then
            floorTextSelect.Offset(0, 3).Value = 4
        ElseIf InStr(1, floorTextSelect.Value, "/5 -") Then
            floorTextSelect.Offset(0, 3).Value = 5
        ElseIf InStr(1, floorTextSelect.Value, "/6 -") Then
            floorTextSelect.Offset(0, 3).Value = 6
        ElseIf InStr(1, floorTextSelect.Value, "/7 -") Then
            floorTextSelect.Offset(0, 3).Value = 7
        ElseIf InStr(1, floorTextSelect.Value, "/8 -") Or InStr(1, floorTextSelect.Value, "TEL") Or InStr(1, floorTextSelect.Value, "/8") Then
            floorTextSelect.Offset(0, 3).Value = 8
        ElseIf InStr(1, floorTextSelect.Value, "/B") Then
            floorTextSelect.Offset(0, 3).Value = "Basement"
        ElseIf InStr(1, floorTextSelect.Value, "/R") Then
            floorTextSelect.Offset(0, 3).Value = "Roof"
        Else
            floorTextSelect.Offset(0, 3).Value = ""
        End If
    Next
End Sub

Sub STEP_ONE_RemoveNonWindows7()
    'step one will clear out non windows 7
    
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    ActiveSheet.Range("H1").Value = "Floor"
    Columns("A:I").Select
    Selection.AutoFilter
    
    ActiveSheet.Range("$A$1:$H$1222").AutoFilter Field:=7, Criteria1:=Array( _
        "aix_version_5.2", "cisco_ios_version_12_2_55_se", "Linux", _
        "Microsoft Windows 10 Enterprise", "SuSE(Linux)", "Ubuntu(Linux)", "vmnix-x86", _
        "Windows 2000", "Windows 2003", "Windows 2003 R2", "Windows 2008", _
        "Windows 2008 R2", "Windows 2012 R2", "Windows 2012 Standard", "="), Operator:= _
        xlFilterValues
    Range("A9:G1222").Select
    Selection.EntireRow.Delete
    ActiveSheet.Range("$A$1:$H$835").AutoFilter Field:=7
    
    CreateFloorList
End Sub

Sub STEP_TWO_RemoveNonAllCTs()
    Dim pcNames As Range
    Dim pcName As Range
    Set pcNames = ActiveSheet.Range("B2:B900")
    
    For Each pcName In pcNames.Cells
        If InStr(1, pcName.Value, "gmhe") Then
            pcName.Offset(0, 8).Value = "Delete"
        ElseIf InStr(1, pcName.Value, "gmhs") Then
            pcName.Offset(0, 8).Value = "Delete"
        ElseIf InStr(1, pcName.Value, "gmhc") Then
            pcName.Offset(0, 8).Value = "Delete"
        ElseIf InStr(1, pcName.Value, "gmhy") Then
            pcName.Offset(0, 8).Value = "Delete"
        ElseIf InStr(1, pcName.Value, "gmhz") Then
            pcName.Offset(0, 8).Value = "Delete"
        Else
            pcName.Offset(0, 8).Value = ""
        End If
    Next
    DeleteNonCTS
End Sub

Sub DeleteNonCTS()
    Dim LR As Long, i As Long
    LR = Range("J" & Rows.count).End(xlUp).Row
    For i = LR To 2 Step -1
        If Range("J" & i).Value Like "Delete" Then Rows(i).Delete
    Next i
End Sub

Sub DeleteNonDeployed()
    Dim LR As Long, i As Long
    LR = Range("I" & Rows.count).End(xlUp).Row
    For i = LR To 2 Step -1
        If Not Range("I" & i).Value Like "Deployed" Then Rows(i).Delete
    Next i

End Sub


Sub Stats()
    
    Dim x As Range
    Dim x1 As Range
    Dim count As Integer
    
    ActiveSheet.Range("J7").Value = "Floor:"
    
    Dim lblArr() As Variant
    lblArr = Array("1", "2", "3", "4", "5", "6", "7", "8", "Basement", "Roof", "CSI", "Eulalia", "", "Desktop", "Laptop")
    
    Dim formulaArr() As Variant
    formulaArr = Array("=COUNTIF(R[-6]C[-3]:R[892]C[-3], ""1"")", _
    "=COUNTIF(R[-7]C[-3]:R[891]C[-3],""2"")", _
    "=COUNTIF(R[-8]C[-3]:R[890]C[-3],""3"")", _
    "=COUNTIF(R[-9]C[-3]:R[889]C[-3],""4"")", _
    "=COUNTIF(R[-10]C[-3]:R[888]C[-3],""5"")", _
    "=COUNTIF(R[-11]C[-3]:R[887]C[-3],""6"")", _
    "=COUNTIF(R[-12]C[-3]:R[886]C[-3],""7"")", _
    "=COUNTIF(R[-13]C[-3]:R[885]C[-3],""8"")", _
    "=COUNTIF(R[-14]C[-3]:R[884]C[-3],""Basement"")", _
    "=COUNTIF(R[-15]C[-3]:R[883]C[-3],""Roof"")", _
    "=COUNTIF(R[-16]C[-3]:R[882]C[-3],""CSI"")", _
    "=COUNTIF(R[-17]C[-3]:R[881]C[-3], ""Eulalia"")")
    count = 0
    Set x1 = ActiveSheet.Range("J8:J19")
    For Each x In x1
        x.Value = lblArr(count)
        x.Offset(0, 1).Value = formulaArr(count)
        count = count + 1
    Next
    
    
End Sub

