' ############################################################################################################
' Fucntion และ Sub Routine กลางสำหรับการประมวลผลข้อมูลและจัดการการล็อกแผ่นงาน 
' โดยไม่มีการอ้างอิงถึงเซลล์หรือแผ่นงานใดโดยตรง (เพื่อความยืดหยุ่นในการเรียกใช้จากหลายๆ Sub)
' ############################################################################################################


' =======================================================================================
' Function สำหรับประมวลผลและกรองข้อมูลออกมาเป็น Array (Logic)
' เหมาะสำหรับการดึงข้อมูลจากฐานข้อมูลและนำไปใช้ต่อใน Sub ต่างๆ เช่น การอัปเดตรายชื่อจังหวัด/อำเภอ/ตำบล
' =======================================================================================   
Public Function GetFilteredLocationArray(ByVal Mode As String, ByVal Prov As String, Optional ByVal Amp As String = "") As Variant
    Dim wsCommon As Worksheet
    Dim rawData As Variant
    Dim resultData() As String
    Dim lastRow As Long, i As Long, count As Long
    
    Set wsCommon = ThisWorkbook.Worksheets("CF_Common") '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    ' 1. หาบรรทัดสุดท้ายและดึงข้อมูลเข้า Array (C=จังหวัด, D=อำเภอ, E=ตำบล)
    lastRow = wsCommon.Cells(wsCommon.Rows.count, "C").End(xlUp).Row
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


' ======================================================================================
' GetPackageValidation: ฟังก์ชันคำนวณหน้างาน (เรียกใช้ภายใน)
' รับค่า: ทุนประกัน (Double) และ ช่วงเซลล์ตารางอ้างอิง (Range)
' คืนค่า: Array 5 ช่อง [สถานะ, ค่าต่ำกว่า, ค่าสูงกว่า, ต่ำสุด, สูงสุด]
' ======================================================================================
Public Function GetPackageValidation(ByVal totalVal As Double, ByVal tableRef As Range) As Variant
    Dim result(1 To 5) As Variant
    Dim matchIdx As Variant
    
    ' STEP 1: ตรวจสอบความปลอดภัย (หากส่ง Range มาว่างเปล่าให้เด้งออก)
    If tableRef Is Nothing Then
        result(1) = "Error: No Table Reference"
        GetPackageValidation = result
        Exit Function
    End If
    
    ' STEP 2: คำนวณขอบเขตล่าง-บนของตารางจริง (Dynamic Min-Max)
    result(4) = Application.WorksheetFunction.Min(tableRef)
    result(5) = Application.WorksheetFunction.Max(tableRef)

    ' STEP 3: ตรวจสอบว่าทุน "หลุดช่วง" ที่รับประกันหรือไม่
    If totalVal < result(4) Or totalVal > result(5) Then
        result(1) = "OutOfRange"
        GetPackageValidation = result
        Exit Function
    End If

    ' STEP 4: ตรวจสอบแบบตรงตัว (Exact Match)
    matchIdx = Application.Match(totalVal, tableRef, 0)
    
    If Not IsError(matchIdx) Then
        ' พบแผนที่ตรงเป๊ะ
        result(1) = "Valid"
    Else
        ' ไม่ตรงแผนเป๊ะ -> เริ่มกระบวนการหาแผนใกล้เคียง (Suggestion)
        result(1) = "Invalid"
        On Error Resume Next
        ' หาแผนที่ทุนน้อยกว่าที่ใกล้ที่สุด
        result(2) = Application.WorksheetFunction.Lookup(totalVal, tableRef)
        ' หาแผนที่ทุนมากกว่าที่ใกล้ที่สุด
        matchIdx = Application.Match(totalVal, tableRef, 1)
        If Not IsError(matchIdx) And matchIdx < tableRef.Rows.count Then
            result(3) = tableRef.Cells(matchIdx + 1, 1).Value
        Else
            result(3) = result(2)
        End If
        On Error GoTo 0
    End If
    
    GetPackageValidation = result
End Function

' ======================================================================================
' IsPremiumValid: ฟังก์ชันคืนค่า True/False
' วัตถุประสงค์: ใช้ขวางการทำงานถ้าข้อมูลไม่ถูกต้อง เช่น "ถ้าไม่ผ่าน ห้าม Save PDF"
' ======================================================================================
Public Function IsPremiumValid(ByVal InputVal As Double, ByVal tableRef As Range) As Boolean
    Dim valResult As Variant
    ' เรียกใช้ Logic กลางเพื่อขอข้อมูลผลการตรวจสอบ
    valResult = GetPackageValidation(InputVal, tableRef) ' ระบุตารางอ้างอิงด้วย
    
    ' แยกแยะผลลัพธ์เพื่อแจ้งเตือนผ่านกล่องข้อความ (MsgBox)
    Select Case valResult(1)
        Case "Valid"
            IsPremiumValid = True ' ผ่านฉลุย
            
        Case "OutOfRange"
            ' แจ้งเตือนกรณีทุนหลุดขอบเขต (ใช้ตัวเลข Min/Max จาก Array มาแสดง)
            MsgBox "ไม่สามารถดำเนินการได้!" & vbCrLf & _
                   "ทุนประกันต้องอยู่ระหว่าง " & Format(valResult(4), "#,##0") & _
                   " ถึง " & Format(valResult(5), "#,##0") & " บาท", vbCritical, "ทุนประกันไม่อยู่ในเงื่อนไข"
            IsPremiumValid = False
            
        Case "Invalid"
            ' แจ้งเตือนกรณีทุนไม่ตรงแผน (เอาค่าแนะนำ Floor/Ceiling มาโชว์)
            MsgBox "ทุนประกันไม่ตรงตาม Package ที่จำหน่าย" & vbCrLf & _
                   "แนะนำให้ปรับเป็น: " & Format(valResult(2), "#,##0") & _
                   " หรือ " & Format(valResult(3), "#,##0"), vbCritical, "ไม่พบ Package"
            IsPremiumValid = False
    End Select
End Function

' ======================================================================================
' Function ค้นหาตำแหน่งคอลัมน์ (Index) จากข้อความหัวตาราง
' รับค่า: ชื่อแผ่นงาน, แถวหัวตาราง, ข้อความหัวตารางที่ต้องการค้นหา
' คืนค่า: เลขคอลัมน์ที่พบ (Long) หรือ 0 ถ้าไม่พบ
' ======================================================================================
Public Function FindHeaderColumn(sheetName As String, headerRow As Long, headerText As String) As Long
    Dim ws As Worksheet, c As Range, lastCol As Long
    Set ws = ThisWorkbook.Worksheets(sheetName)
    
    ' หาคอลัมน์สุดท้ายที่มีข้อมูลในแถวนั้น เพื่อจำกัดขอบเขตการวนลูป
    lastCol = ws.Cells(headerRow, ws.Columns.count).End(xlToLeft).Column
    
    ' วนลูปตรวจเช็คทีละเซลล์ในแถวหัวตาราง
    For Each c In ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, lastCol))
        ' เปรียบเทียบข้อความ (ตัดช่องว่างออกและไม่สนใจพิมพ์เล็ก-ใหญ่)
        If StrComp(Trim$(CStr(c.Value)), headerText, vbTextCompare) = 0 Then
            ' หากเจอ ให้ส่งเลขคอลัมน์กลับทันที
            FindHeaderColumn = c.Column
            Exit Function
        End If
    Next c
    ' หากวนจนจบแล้วไม่เจอ ให้ส่งค่า 0 กลับไป
    FindHeaderColumn = 0
End Function

' ======================================================================================
' Function ดึงข้อมูลจากคอลัมน์มาเป็น Range (เพื่อนำไปใช้ต่อใน Match)
' รับค่า: ชื่อแผ่นงาน, แถวหัวตาราง, ข้อความหัวตารางที่ต้องการค้นหา
' ======================================================================================
Public Function GetListRange(ByVal sheetName As String, ByVal headerRow As Long, ByVal headerText As String) As Range
    Dim ws As Worksheet, colIndex As Long, lastRow As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Function
    
    ' หาเลขคอลัมน์
    colIndex = FindHeaderColumn(sheetName, headerRow, headerText)
    
    If colIndex > 0 Then
        lastRow = ws.Cells(ws.Rows.count, colIndex).End(xlUp).Row
        ' ส่งค่ากลับเป็นวัตถุ Range (เพื่อให้ Match ทำงานได้แม่นยำกว่า Array 2 มิติ)
        If lastRow > headerRow Then
            Set GetListRange = ws.Range(ws.Cells(headerRow + 1, colIndex), ws.Cells(lastRow, colIndex))
        End If
    End If
End Function

' ======================================================================================
' Function ตรวจสอบพื้นที่เสี่ยง (ตัวที่ Sub ภายนอกจะเรียกใช้)
' รับค่า: ชื่อจังหวัด (String)
' ======================================================================================
Public Function IsFloodRisk(ByVal provinceName As String) As Boolean
    Dim riskRange As Range
    Dim result As Variant
    Dim cleanProv As String: cleanProv = Trim$(CStr(provinceName))
    
    ' หากไม่ระบุชื่อจังหวัด ให้ถือว่าไม่เสี่ยง
    If cleanProv = "" Then IsFloodRisk = False: Exit Function
    
    ' ขั้นตอนที่ 1: ดึง Range รายชื่อจังหวัดเสี่ยงภัยมาจากฐานข้อมูล
    ' ปรับ "CF_อยู่ดีมีสุข", 1, "จังหวัดยกเว้นน้ำท่วม" ให้ตรงตามหน้างานจริง
    Set riskRange = GetListRange("CF_Common", 1, "จังหวัดยกเว้นน้ำท่วม1")
    
    ' ขั้นตอนที่ 2: ตรวจสอบความปลอดภัย (หากไม่พบตารางรายชื่อ ให้ผ่านไปก่อน)
    If riskRange Is Nothing Then
        IsFloodRisk = False
        Exit Function
    End If
    
    ' ขั้นตอนที่ 3: ใช้ Match ค้นหาชื่อจังหวัดใน Range ที่ดึงมา
    result = Application.Match(cleanProv, riskRange, 0)
    
    ' ขั้นตอนที่ 4: สรุปผล (ถ้าเจอเลขลำดับ = เสี่ยงภัย)
    IsFloodRisk = Not IsError(result)
End Function

' ======================================================================================
' ฟังก์ชัน IsInArray (ตรวจสอบว่าค่าที่กำหนดอยู่ใน Array หรือไม่)
' ใช้สำหรับตรวจสอบข้อมูลที่ดึงมาจากตารางฐานข้อมูล (เช่น รายชื่อจังหวัดเสี่ยงภัย)
' ======================================================================================
Public Function IsInArray(ByVal val As String, ByRef arr() As String, ByVal currentCount As Long) As Boolean
    Dim j As Long
    If currentCount = 0 Then IsInArray = False: Exit Function
    For j = 1 To currentCount
        If arr(j, 1) = val Then
            IsInArray = True
            Exit Function
        End If
    Next j
    IsInArray = False
End Function

' ======================================================================================
' SetSheetProtection: ฟังก์ชันกลางสำหรับจัดการการล็อก/ปลดล็อก
' targetSheet: ส่ง Object ของ Worksheet ที่ต้องการจัดการเข้ามา
' IsLock: True = ล็อกข้อมูล, False = ปลดล็อกข้อมูล
' ======================================================================================
Public Sub SetSheetProtection(ByVal targetSheet As Worksheet, ByVal IsLock As Boolean)
    
    ' ตรวจสอบเบื้องต้นว่ามีการส่ง Sheet เข้ามาจริงหรือไม่ (ป้องกัน Error)
    If targetSheet Is Nothing Then Exit Sub

    If IsLock Then
        ' ---- กรณีสั่งให้ "ล็อก" (Protect) ----
        ' คุณสามารถตั้งค่า Property การล็อกเพิ่มเติมได้ที่นี่
        targetSheet.Protect Password:=myPassword, _
                            DrawingObjects:=True, _
                            Contents:=True, _
                            Scenarios:=True, _
                            DrawingObjects:=True, _
                            UserInterfaceOnly:=True
        ThisWorkbook.Protect Password:=myPassword, Structure:=True, Windows:=False
    Else
        ' ---- กรณีสั่งให้ "ปลดล็อก" (Unprotect) ----
        targetSheet.Unprotect Password:=myPassword
        ThisWorkbook.Unprotect Password:=myPassword
    End If
    
End Sub

Public Sub SetWorkbookProtection(ByVal IsLock As Boolean)
    
     If IsLock Then
        ' ---- กรณีสั่งให้ "ล็อก" (Protect) ----
        ThisWorkbook.Protect Password:=myPassword, Structure:=True, Windows:=False
    Else
        ' ---- กรณีสั่งให้ "ปลดล็อก" (Unprotect) ----
        ThisWorkbook.Unprotect Password:=myPassword
    End If
    
End Sub


' ======================================================================================
' UnlockAllSheets: ฟังก์ชันสำหรับปลดล็อกทุกแผ่นงานใน Workbook (ใช้ในกรณีฉุกเฉิน)
' เหมาะสำหรับสถานการณ์ที่คุณลืมรหัสผ่านหรือมีปัญหากับการล็อกแผ่นงาน
' =======================================================================================
Public Sub UnlockAllSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect Password:=myPassword
    Next ws
    ThisWorkbook.Unprotect Password:=myPassword
End Sub

' ======================================================================================
' ---------- Sub Routine สำหรับรีเซ็ตระบบด้วยตนเอง ----------
' เหมาะสำหรับสถานการณ์ที่ระบบ Event หรือ Calculation มีปัญหา เช่น ค้าง, ไม่ตอบสนอง, หรือเกิดข้อผิดพลาดที่ทำให้ Excel อยู่ในสถานะไม่ปกติ
' =======================================================================================
Public Sub ResetExcelEvents()
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
End Sub



