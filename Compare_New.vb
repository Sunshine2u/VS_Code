' ############################################################################################################
' ส่วนที่ 1: Function สำหรับประมวลผลและกรองข้อมูลออกมาเป็น Array (Logic)
' ############################################################################################################
Public Function GetFilteredLocationArray(ByVal Mode As String, ByVal Prov As String, Optional ByVal Amp As String = "") As Variant
    Dim wsCommon As Worksheet
    Dim rawData As Variant
    Dim resultData() As String
    Dim lastRow As Long, i As Long, count As Long
    
    Set wsCommon = ThisWorkbook.Worksheets("CF_Common")
    
    ' 1. หาบรรทัดสุดท้ายและดึงข้อมูลเข้า Array (C=จังหวัด, D=อำเภอ, E=ตำบล)
    lastRow = wsCommon.Cells(wsCommon.Rows.Count, "C").End(xlUp).Row
    If lastRow < 2 Then 
        GetFilteredLocationArray = Empty
        Exit Function
    End If
    
    rawData = wsCommon.Range("C2:E" & lastRow).Value
    ReDim resultData(1 To UBound(rawData, 1), 1 To 1)
    count = 0
    
    ' 2. วนลูปกรองข้อมูลตามเงื่อนไข
    For i = 1 To UBound(rawData, 1)
        If Mode = "Amphoe" Then
            ' กรองอำเภอ (คอลัมน์ที่ 2 ของ Array)
            If CStr(rawData(i, 1)) = Prov And rawData(i, 2) <> "" Then
                If Not IsInArray(CStr(rawData(i, 2)), resultData, count) Then
                    count = count + 1
                    resultData(count, 1) = rawData(i, 2)
                End If
            End If
        ElseIf Mode = "Tambon" Then
            ' กรองตำบล (คอลัมน์ที่ 3 ของ Array)
            If CStr(rawData(i, 1)) = Prov And CStr(rawData(i, 2)) = Amp And rawData(i, 3) <> "" Then
                If Not IsInArray(CStr(rawData(i, 3)), resultData, count) Then
                    count = count + 1
                    resultData(count, 1) = rawData(i, 3)
                End If
            End If
        End If
    Next i
    
    ' 3. ส่งค่ากลับเป็น Array หากพบข้อมูล
    If count > 0 Then
        GetFilteredLocationArray = resultData
    Else
        GetFilteredLocationArray = Empty
    End If
End Function

' ############################################################################################################
' ส่วนที่ 2: Sub สำหรับจัดการชีทและเขียนข้อมูลลงไป (Action)
' ############################################################################################################
Public Sub UpdateLocationList(ByVal Mode As String, ByVal Prov As String, Optional ByVal Amp As String = "")
    Dim wsTarget As Worksheet
    Dim arrResults As Variant
    Dim targetCol As String, colAmphoe As String, colTambon As String
    Dim targetLastRow As Long
    
    ' --- CONFIG: กำหนดเป้าหมาย ---
    Set wsTarget = ThisWorkbook.Worksheets("CF_อยู่ดีมีสุข")
    colAmphoe = "U" ' คอลัมน์สำหรับ LIST_อำเภอ
    colTambon = "V" ' คอลัมน์สำหรับ LIST_ตำบล
    
    targetCol = IIf(Mode = "Amphoe", colAmphoe, colTambon)
    
    Application.ScreenUpdating = False
    Call SetSheetProtection(wsTarget, False)
    
    ' 1. ล้างข้อมูลเก่า (Cleanup)
    targetLastRow = wsTarget.Cells(wsTarget.Rows.Count, targetCol).End(xlUp).Row
    If targetLastRow >= 2 Then
        wsTarget.Range(wsTarget.Cells(2, targetCol), wsTarget.Cells(targetLastRow, targetCol)).ClearContents
    End If
    
    ' พิเศษ: หากเลือกจังหวัดใหม่ ให้ล้างรายการตำบลเดิมในชีทปลายทางทิ้งด้วย
    If Mode = "Amphoe" Then
        Dim lastRowV As Long
        lastRowV = wsTarget.Cells(wsTarget.Rows.Count, colTambon).End(xlUp).Row
        If lastRowV >= 2 Then wsTarget.Range(wsTarget.Cells(2, colTambon), wsTarget.Cells(lastRowV, colTambon)).ClearContents
    End If
    
    ' 2. เรียกใช้ Function เพื่อขอข้อมูล Array
    arrResults = GetFilteredLocationArray(Mode, Prov, Amp)
    
    ' 3. เขียนข้อมูลลงชีท
    If Not IsEmpty(arrResults) Then
        ' กรองเอาเฉพาะแถวที่มีข้อมูลจริงมาวาง (Resize ตามจำนวน count ที่ได้จาก Function)
        ' ในที่นี้ Function คืนค่า Array ขนาดใหญ่ที่มีช่องว่าง ดังนั้นเราต้องหาจำนวนจริง
        Dim actualCount As Long, i As Long
        For i = 1 To UBound(arrResults, 1)
            If arrResults(i, 1) = "" Then Exit For
            actualCount = actualCount + 1
        Next i
        
        If actualCount > 0 Then
            wsTarget.Cells(2, targetCol).Resize(actualCount, 1).Value = arrResults
        End If
    End If
    
CleanUp:
    Call SetSheetProtection(wsTarget, FileLockSetting)
    Application.ScreenUpdating = True
End Sub