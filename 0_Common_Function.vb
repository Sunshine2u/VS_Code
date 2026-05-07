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
        If Not IsError(matchIdx) And matchIdx < tableRef.Rows.Count Then
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

' Function ค้นหาตำแหน่งคอลัมน์ (Index) จากข้อความหัวตาราง
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

' Function ดึงข้อมูลทั้งหมดในคอลัมน์นั้นๆ มาเก็บไว้ใน Array (แบบ 2 มิติ)
Public Function GetList(sheetName As String, headerRow As Long, headerText As String) As Variant
    Dim ws As Worksheet, colIndex As Long, lastRow As Long
    Set ws = ThisWorkbook.Worksheets(sheetName)
    
    ' หาตำแหน่งคอลัมน์ก่อน
    colIndex = FindHeaderColumn(sheetName, headerRow, headerText)
    
    ' หากไม่พบคอลัมน์ที่ต้องการ ให้คืนค่าเป็น Array ว่าง
    If colIndex = 0 Then GetList = Array(): Exit Function
    
    ' หาแถวสุดท้ายที่มีข้อมูลในคอลัมน์นั้น
    lastRow = ws.Cells(ws.Rows.count, colIndex).End(xlUp).Row
    
    ' ตรวจสอบว่ามีข้อมูลอยู่ใต้หัวตารางหรือไม่
    If lastRow <= headerRow Then GetList = Array(): Exit Function

    ' ดึงข้อมูลทั้งช่วง (Range) เข้าสู่ตัวแปร Variant ทีเดียว (เร็วกว่าวนลูปเก็บทีละเซลล์)
    GetList = ws.Range(ws.Cells(headerRow + 1, colIndex), ws.Cells(lastRow, colIndex)).Value
End Function

' Function ตรวจสอบว่าจังหวัดที่ระบุ อยู่ในพื้นที่เสี่ยงภัย (ยกเว้นความคุ้มครองน้ำท่วม) หรือไม่
Public Function IsFloodRisk(ByVal provinceName As String) As Boolean
    Dim list As Variant, result As Variant
    Dim cleanProv As String: cleanProv = Trim$(CStr(provinceName))
    
    ' หากไม่ระบุชื่อจังหวัด ให้ถือว่าไม่เสี่ยงภัย
    If cleanProv = "" Then IsFloodRisk = False: Exit Function
    
    ' ดึงรายชื่อจังหวัดยกเว้นน้ำท่วมจากแผ่นงานฐานข้อมูล (CF_อยู่ดีมีสุข)
    list = GetList("CF_อยู่ดีมีสุข", 1, "จังหวัดยกเว้นน้ำท่วม")
    
    ' ตรวจสอบว่าผลลัพธ์เป็น Array หรือไม่
    If IsArray(list) Then
        ' ใช้ฟังก์ชัน Match ในการค้นหาจังหวัดในรายการ
        result = Application.Match(cleanProv, list, 0)
        ' หากหาเจอ (ไม่เป็น Error) แสดงว่าเป็นจังหวัดพื้นที่เสี่ยง
        If Not IsError(result) Then IsFloodRisk = True: Exit Function
    End If
    
    ' หากไม่พบในรายการพื้นที่เสี่ยง
    IsFloodRisk = False
End Function

' ฟังก์ชัน IsInArray (ใช้ตัวเดิม)
Private Function IsInArray(ByVal val As String, ByRef arr() As String, ByVal currentCount As Long) As Boolean
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
                            Scenarios:=True
    Else
        ' ---- กรณีสั่งให้ "ปลดล็อก" (Unprotect) ----
        targetSheet.Unprotect Password:=myPassword
    End If
    
End Sub
