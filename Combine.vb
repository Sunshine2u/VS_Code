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
        ThisWorkbook.Protect Password:=myPassword, Structure:=True, Windows:=False
    Else
        ' ---- กรณีสั่งให้ "ปลดล็อก" (Unprotect) ----
        targetSheet.Unprotect Password:=myPassword
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





'############################################################################################################

'==============================================
' (A) แจ้งเตือนจังหวัดกลุ่มเสี่ยงภัยน้ำท่วมที่ H28
' (B) ตรวจสอบทุนรวมที่ G43 ให้ตรงตามตาราง Premium Table แบบเป๊ะๆ
'     - หากไม่ตรง จะแสดงทุนแนะนำที่ใกล้เคียงที่สุด (Floor/Ceiling) ที่ L43
'     - ตรวจสอบช่วงทุนให้อยู่ระหว่าง 500,000 - 10,000,000
'     - ทำงานทั้งเมื่อมีการแก้ไข (Change) และเมื่อสูตรคำนวณใหม่ (Calculate)
'===================================================================

Private Sub Worksheet_Change(ByVal Target As Range)
   
    ' 1. ตรวจสอบเบื้องต้น
    If Intersect(Target, Me.Range("H28,G41:H41,G42:H42")) Is Nothing Then Exit Sub

    On Error GoTo ErrorHandler
    
    ' ปิดการแจ้งเตือนชั่วคราวเพื่อป้องกันการรันวนซ้ำ
    Application.EnableEvents = False
    
    ' ปลดล็อกชีทก่อนเริ่มทำงาน
    Call SetSheetProtection(Me, False)
    
    ' ---------- (C) ส่วนตรวจสอบการกรอก G42 (เฟอร์นิเจอร์) ----------
    If Not Intersect(Target, Me.Range("G41,G42")) Is Nothing Then

    End If
    
    ' ---------- ส่วนคำนวณ G43 อัตโนมัติ (G41 + G42) ----------
    If Not Intersect(Target, Me.Range("G41,G42")) Is Nothing Then
        
        If Len(Trim(Me.Range("G42").Text)) = 0 Then
            Me.Range("J42").Value = "ถ้าไม่มีให้กรอกเลข 0"
            Me.Range("J42").Font.Color = RGB(255, 0, 0)
        Else
            Me.Range("J42").ClearContents
        End If
        
        If IsNumeric(Me.Range("G41").Value) And IsNumeric(Me.Range("G42").Value) Then
            ' รวมค่าอาคารและเฟอร์นิเจอร์
            Me.Range("G43").Value = Me.Range("G41").Value + Me.Range("G42").Value
            Call CheckAndSuggestPremium(Me.Range("G43").Value)
        Else
            Me.Range("G43").ClearContents
        End If
        
    End If
    
    ' ---------- (A) ส่วนแจ้งเตือนจังหวัดที่ H28 ----------
    If Not Intersect(Target, Me.Range("H28")) Is Nothing Then
        Dim list As Variant
        Dim result As Variant
        Dim prov As String
        
        prov = Trim$(CStr(Me.Range("H28").Value))
        list = GetListRange("CF_Common", 1, "จังหวัดยกเว้นน้ำท่วม1")
        
        ' ใช้ Match ค้นหาชื่อจังหวัดใน Array ได้เลยไม่ต้องวน Loop
        result = Application.Match(prov, list, 0)
        
        If Not IsError(result) Then
            MsgBox "พบว่าจังหวัด " & prov & " เป็นพื้นที่เสี่ยงภัยน้ำท่วม" & vbCrLf & _
                           "โปรดติดต่อเจ้าหน้าที่ MTI ผู้ดูแลตัวแทน ในการออกใบเสนอราคา", _
                           vbExclamation, "แจ้งเตือนความเสี่ยง"
        End If
        
    End If
    
    ' ล็อกชีทคืน
    Call SetSheetProtection(Me, FileLockSetting) ' ใช้ค่าจาก Const ที่ตั้งไว้ใน 2_Product1_Sub.vb
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    MsgBox "เกิดข้อผิดพลาด: " & Err.Description, vbCritical, "Error"
    Call SetSheetProtection(Me, FileLockSetting)
    Application.EnableEvents = True
End Sub





' ##############################################################################################################################

Option Explicit

' ======================================================================================
' ส่วนที่ 1: การประกาศตัวแปรระดับ Global (Global Constants)
' ======================================================================================
' รหัสผ่าน (Password) ส่วนกลางที่ใช้สำหรับปลดล็อกแผ่นงาน (Sheet) และโครงสร้างไฟล์ (Workbook)
' ประกาศเป็น Public Const เพื่อให้ทุก Sub และ Function ในโปรเจกต์สามารถเรียกใช้ได้ทันที
Public Const myPassword As String = "QTMTI"
Public Const FileLockSetting As Boolean = True 'ตั้งค่าให้ล็อกไฟล์โดย Default (True = ล็อก, False = ไม่ล็อก)

' ======================================================================================
' ส่วนที่ 2: ระบบจัดการการกรอกข้อมูล (Data Clearing & Navigation)
' ======================================================================================

' Sub สำหรับล้างข้อมูล (Reset) ในแบบฟอร์มใบเสนอราคาหน้าหลัก
Public Sub อยู่ดีมีสุข_Clear_Input()
    ' ดักจับข้อผิดพลาด: หากเกิด Error ให้กระโดดไปที่ ErrorHandler ด้านล่าง
    On Error GoTo ErrorHandler

    ' ปิดระบบ Event ชั่วคราว: เพื่อป้องกันไม่ให้โค้ด Event อื่นๆ (เช่น Worksheet_Change) ทำงานแทรก
    ' และช่วยให้การล้างข้อมูลหลายๆ เซลล์พร้อมกันทำได้รวดเร็วขึ้น
    Application.EnableEvents = False
    
    ' อ้างอิงการทำงานกับแผ่นงานหน้าคีย์ข้อมูลใบเสนอราคา
    With Worksheets("QT_อยู่ดีมีสุข") '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        ' ปลดล็อกแผ่นงานก่อนเพื่อให้สามารถลบข้อมูลในเซลล์ที่ถูกล็อกไว้ได้
        .Unprotect Password:=myPassword
        
        ' ล้างเฉพาะเนื้อหา (ClearContents) ในช่วงเซลล์ที่กำหนด (ไม่ลบ Format หรือสูตรในเซลล์อื่น)
        ' มีการตัดบรรทัดด้วย " _" เพื่อให้โค้ดอ่านง่ายขึ้น
        .Range("G24:M24,G26:M26,H28,J28,L28,G31:H31,G33:H33,H35,L35,H36,L36,L38," & _
               "G41:H41,G42:H42,G45:J45,G49:I49,L49:M49,G51:I51,L51:M51,G53:I53,J43:L43,G43:H43").ClearContents '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        
        ' ล็อกแผ่นงานคืนหลังจากล้างข้อมูลเสร็จ เพื่อป้องกันผู้ใช้แก้ไขสูตรคำนวณโดยไม่ตั้งใจ
        .Protect Password:=myPassword
    End With

    ' เปิดระบบ Event กลับคืนสู่ปกติ
    Application.EnableEvents = True
    
    ' แสดงข้อความแจ้งผู้ใช้ว่าดำเนินการสำเร็จ
    MsgBox "ล้างข้อมูลเรียบร้อยแล้ว", vbInformation, "ระบบแจ้งเตือน"
    
    ' เรียกฟังก์ชันรีเซ็ตค่าระบบเสริม เพื่อให้มั่นใจว่า Excel กลับมาอยู่ในสถานะพร้อมทำงานปกติ
    Call ResetExcelEvents
    
    ' จบการทำงานของ Sub หลัก (ป้องกันไม่ให้โค้ดไหลไปทำงานใน ErrorHandler)
    Exit Sub

ErrorHandler:
    ' หากเกิดความผิดพลาดระหว่างทาง ให้เปิดระบบ Event คืนเสมอ เพื่อไม่ให้ Excel ค้าง
    Application.EnableEvents = True
    MsgBox "เกิดข้อผิดพลาดในการล้างข้อมูล: " & Err.Description, vbCritical, "ข้อผิดพลาดระบบ"
End Sub

' =======================================================================================
' Sub สำหรับเปิดหน้า Leaflet (แผ่นพับรายละเอียดความคุ้มครอง)
' ' ในกรณีที่แผ่นงานถูกซ่อนอยู่ จะทำการปลดล็อกโครงสร้างไฟล์เพื่อแสดงแผ่นงานนั้น และย้ายหน้าจอไปยังแผ่นงาน Leaflet
' =======================================================================================
Sub อยู่ดีมีสุข_Go_To_Leaflet()
    ' ปลดล็อกโครงสร้างไฟล์เพื่อให้สามารถเปลี่ยนสถานะการซ่อนของแผ่นงานได้
    ActiveWorkbook.Unprotect Password:=myPassword
    
    ' แสดงแผ่นงาน Leaflet (ที่อาจถูกซ่อนอยู่)
    Sheets("LL_อยู่ดีมีสุข").Visible = True '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    ' ย้ายหน้าจอไปยังแผ่นงานนั้น
    Worksheets("LL_อยู่ดีมีสุข").Activate '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    ' ล็อกโครงสร้างไฟล์คืน เพื่อป้องกันการลบหรือสลับลำดับแผ่นงาน
    ActiveWorkbook.Protect Password:=myPassword, Structure:=True, Windows:=False
End Sub

' ======================================================================================
' Sub สำหรับปิดหน้า Leaflet และกลับมายังหน้าใบเสนอราคาหลัก
' เมื่อผู้ใช้ปิดแผ่นพับรายละเอียดความคุ้มครอง จะทำการปลดล็อกโครงสร้างไฟล์เพื่อซ่อนแผ่นงาน Leaflet อีกครั้ง และย้ายหน้าจอกลับไปยังหน้าใบเสนอราคาหลัก
' ======================================================================================
Sub อยู่ดีมีสุข_Close_Leafltet()
    ' ปลดล็อกโครงสร้างไฟล์
    ActiveWorkbook.Unprotect Password:=myPassword
    
    ' ซ่อนแผ่นงาน Leaflet
    Sheets("LL_อยู่ดีมีสุข").Visible = False '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    ' กลับมายังหน้าคีย์ข้อมูลใบเสนอราคา
    Worksheets("QT_อยู่ดีมีสุข").Activate '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    ' ล็อกโครงสร้างไฟล์คืน
    ActiveWorkbook.Protect Password:=myPassword, Structure:=True, Windows:=False
End Sub

' ======================================================================================
' ส่วนที่ 3: ระบบออกเอกสารใบเสนอราคา (Quotation Issuance)
' ======================================================================================

' =======================================================================================
' Sub สำหรับแสดงตัวอย่างใบเสนอราคา (Print Preview) ก่อนพิมพ์จริง
' จะตรวจสอบเงื่อนไขพื้นที่เสี่ยงภัยน้ำท่วมและความถูกต้องของทุนประกันก่อน หากผ่านถึงจะปลดล็อกโครงสร้างไฟล์เพื่อแสดงหน้าสรุป (Report) ชั่วคราวสำหรับทำ Print Preview
' =======================================================================================
Sub อยู่ดีมีสุข_Preview_Quotation()
    Dim wsKey As Worksheet
    Dim wsQTR As Worksheet
    Dim wsCF As Worksheet 
    Set wsKey = Worksheets("QT_อยู่ดีมีสุข")
    Set wsQTR = Worksheets("QTR_อยู่ดีมีสุข")
    Set wsCF = Worksheets("CF_อยู่ดีมีสุข") '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

    ' 1. ตรวจสอบเงื่อนไขจังหวัดน้ำท่วมก่อน
    If IsFloodRisk(wsKey.Range("H28").Value) Then
        MsgBox "ไม่สามารถออกใบเสนอราคาได้!" & vbCrLf & _
               "จังหวัดที่ระบุอยู่ในพื้นที่เสี่ยงภัยน้ำท่วม" & vbCrLf & _
               "กรุณาติดต่อเจ้าหน้าที่ดูแลตัวแทน MTI เพื่อออกใบเสนอราคาให้ท่าน", vbExclamation, "ตรวจสอบพื้นที่เสี่ยงภัย"
        Exit Sub
    End If

    ' 2. ตรวจสอบเงื่อนไขความถูกต้องของทุนประกัน (แยกออกมาเป็นอีก If หนึ่ง ไม่ซ้อนกัน)
    ' ส่งค่าจาก G43 เข้าไปตรวจสอบในฟังก์ชัน IsPremiumValid ที่เราปรับปรุงไว้
    If Not IsPremiumValid(wsKey.Range("G43").Value, wsCF.Range("A2:A46")) Then
        Exit Sub ' ถ้าฟังก์ชันคืนค่า False (ทุนไม่ตรงแผน) จะหยุดทำงานทันทีพร้อมแจ้งเตือนจากในฟังก์ชันเอง
    End If
    
    ' --- หากผ่านทั้ง 2 ด่านด้านบน ถึงจะเริ่มกระบวนการ Preview ---

    ' ปลดล็อกโครงสร้างไฟล์
    ActiveWorkbook.Unprotect Password:=myPassword

    ' แสดงหน้าสรุป (Report) ชั่วคราวเพื่อใช้ในการทำ Print Preview
    wsQTR.Visible = True
    
    ' เปิดหน้าต่าง Preview
    wsQTR.PrintPreview
    
    ' ซ่อนหน้าสรุปกลับคืน
    wsQTR.Visible = False
    
    ' กลับมายังหน้าคีย์ข้อมูลหลัก
    wsKey.Activate
    
    ' ล็อกโครงสร้างไฟล์คืน
    ActiveWorkbook.Protect Password:=myPassword, Structure:=True, Windows:=False
End Sub

' =======================================================================================
'กดเพื่อสร้างใบเสนอราคา PDF
' จะตรวจสอบเงื่อนไขพื้นที่เสี่ยงภัยน้ำท่วมและความถูกต้องของทุนประกันก่อน หากผ่านถึงจะปลดล็อกโครงสร้างไฟล์เพื่อแสดงหน้าสรุป (Report) ชั่วคราวสำหรับทำการ Export เป็น PDF
' หลังจาก Export เสร็จจะล็อกโครงสร้างไฟล์คืนและซ่อนหน้าสรุปอีกครั้งเพื่อป้องกันการแก้ไขข้อมูลในหน้าสรุปโดยไม่ตั้งใจ
' =======================================================================================
Public Sub อยู่ดีมีสุข_Get_Quotation()
    Dim ws As Worksheet
    Dim filePath As String
    Dim fileName As String
    Dim wsKey As Worksheet
    Dim wsQTR As Worksheet
    Dim wsCF As Worksheet '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Set wsKey = Worksheets("QT_อยู่ดีมีสุข")
    Set wsQTR = Worksheets("QTR_อยู่ดีมีสุข")
    Set wsCF = Worksheets("CF_อยู่ดีมีสุข") '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

    ' 1. ตรวจสอบเงื่อนไขจังหวัดน้ำท่วมก่อน
    If IsFloodRisk(wsKey.Range("H28").Value) Then
        MsgBox "ไม่สามารถออกใบเสนอราคาได้!" & vbCrLf & _
               "จังหวัดที่ระบุอยู่ในพื้นที่เสี่ยงภัยน้ำท่วม" & vbCrLf & _
               "กรุณาติดต่อเจ้าหน้าที่ดูแลตัวแทน MTI เพื่อออกใบเสนอราคาให้ท่าน", vbExclamation, "ตรวจสอบพื้นที่เสี่ยงภัย"
        Exit Sub
    End If

    ' 2. ตรวจสอบเงื่อนไขความถูกต้องของทุนประกัน (แยกออกมาเป็นอีก If หนึ่ง ไม่ซ้อนกัน)
    ' ส่งค่าจาก G43 เข้าไปตรวจสอบในฟังก์ชัน IsPremiumValid ที่เราปรับปรุงไว้
    If Not IsPremiumValid(wsKey.Range("G43").Value, wsCF.Range("A2:A46")) Then
        Exit Sub ' ถ้าฟังก์ชันคืนค่า False (ทุนไม่ตรงแผน) จะหยุดทำงานทันทีพร้อมแจ้งเตือนจากในฟังก์ชันเอง
    End If
    
    ' ปลดล็อก Workbook เพื่อให้สามารถทำงานต่อได้
    ActiveWorkbook.Unprotect Password:=myPassword
    
    wsQTR.Visible = True
    
    On Error GoTo ErrorHandler
 
    
    ' 1. กำหนดชื่อและที่เก็บไฟล์
    fileName = "ใบเสนอราคา_อยู่ดีมีสุข_" & Format(Now, "yyyy-mm-dd_hhmm") & wsKey.Range("G51").Value & ".pdf"
    filePath = ThisWorkbook.Path & "\" & fileName
    
    ' 2. คำสั่ง Export เป็น PDF
    wsQTR.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=filePath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
        
    ' แจ้งเตือนเมื่อสำเร็จ
    MsgBox "บันทึกไฟล์ PDF เรียบร้อยแล้วที่: " & vbCrLf & filePath, vbInformation, "สำเร็จ"

' ส่วนนี้จะทำงานเสมอไม่ว่าจะสำเร็จหรือ Error เพื่อล็อกไฟล์คืน
Finalize:
    wsQTR.Visible = False
    ActiveWorkbook.Protect Password:=myPassword
    Exit Sub

ErrorHandler:
    MsgBox "เกิดข้อผิดพลาด: " & Err.Description, vbCritical, "ข้อผิดพลาด"
    Resume Finalize ' สั่งให้กลับไปล็อกไฟล์ที่ Finalize ก่อนจบโปรแกรม
End Sub

' ======================================================================================
' ส่วนที่ 4: ฟังก์ชันเสริมและระบบตรวจสอบ (Utilities & Validations)
' ======================================================================================
' Function สำหรับรีเซ็ตสถานะของ Excel (Event และการคำนวณอัตโนมัติ)
' มักเรียกใช้หลังจากจบ Sub ใหญ่ๆ หรือเมื่อระบบเกิดข้อผิดพลาด
Public Function ResetExcelEvents() As Boolean
    On Error GoTo ErrorHandler
    ' เปิดการทำงานของ Event (เช่น การตรวจจับการแก้ไขเซลล์)
    Application.EnableEvents = True
    ' ตั้งค่าการคำนวณสูตรให้เป็นแบบอัตโนมัติ (Automatic Calculation)
    Application.Calculation = xlCalculationAutomatic
    ' คืนค่าผลลัพธ์เป็น True เพื่อแจ้งว่ารีเซ็ตสำเร็จ
    ResetExcelEvents = True
    Exit Function
ErrorHandler:
    ResetExcelEvents = False
End Function

' ======================================================================================
' GetPackageValidation: ฟังก์ชันตัวกลางสำหรับดึงข้อมูลและคำนวณหน้างาน
' ทำหน้าที่: ตรวจสอบว่าทุนที่กรอกมา "ผ่านเกณฑ์" หรือ "ควรแนะนำค่าไหน"
' ======================================================================================
Public Function GetPackageValidation(ByVal totalVal As Double) As Variant
    Dim tableRange As Range
    Dim result(1 To 5) As Variant ' เตรียมที่เก็บข้อมูล 5 ช่อง (สถานะ, ค่าต่ำกว่า, ค่าสูงกว่า, ต่ำสุด, สูงสุด)
    Dim matchIdx As Variant
    
    ' --- ขั้นตอนที่ 1: กำหนดตารางอ้างอิง ---
    Set tableRange = ThisWorkbook.Worksheets("CF_อยู่ดีมีสุข").Range("A2:A46")
    
    ' --- ขั้นตอนที่ 2: หาขอบเขต ต่ำสุด-สูงสุด จากตารางจริง ---
    result(4) = Application.WorksheetFunction.Min(tableRange) ' เก็บค่าต่ำสุดไว้ช่อง 4
    result(5) = Application.WorksheetFunction.Max(tableRange) ' เก็บค่าสูงสุดไว้ช่อง 5

    ' --- ขั้นตอนที่ 3: เช็คว่าทุน "น้อยไป" หรือ "มากเกินไป" หรือไม่ ---
    If totalVal < result(4) Or totalVal > result(5) Then
        result(1) = "OutOfRange" ' ระบุสถานะว่า: นอกขอบเขต
        GetPackageValidation = result
        Exit Function ' จบการทำงานทันที
    End If

    ' --- ขั้นตอนที่ 4: ตรวจสอบว่า "ตรงกับแผนเป๊ะๆ" หรือไม่ ---
    matchIdx = Application.Match(totalVal, tableRange, 0)
    
    If Not IsError(matchIdx) Then
        ' ถ้าหาเจอ (ไม่ Error)
        result(1) = "Valid" ' ระบุสถานะว่า: ถูกต้อง
    Else
        ' ถ้าหาไม่เจอ (ไม่ตรงแผนเป๊ะ แต่ยังอยู่ในช่วงที่รับได้)
        result(1) = "Invalid" ' ระบุสถานะว่า: ไม่ตรงแผน (แต่มีค่าแนะนำ)
        
        On Error Resume Next
        ' หาค่าที่ "น้อยกว่าและใกล้ที่สุด" (Floor)
        result(2) = Application.WorksheetFunction.Lookup(totalVal, tableRange)
        ' หาตำแหน่งลำดับของค่าที่ใกล้เคียง
        matchIdx = Application.Match(totalVal, tableRange, 1)
        
        ' หาค่าที่ "มากกว่าและใกล้ที่สุด" (Ceiling)
        If Not IsError(matchIdx) And matchIdx < tableRange.Rows.count Then
            result(3) = tableRange.Cells(matchIdx + 1, 1).Value
        Else
            result(3) = result(2) ' กรณีสูงสุดแล้วให้ใช้ค่าเดิม
        End If
        On Error GoTo 0
    End If
    
    ' ส่งผลลัพธ์ทั้งหมดกลับไปให้คนเรียกใช้
    GetPackageValidation = result
End Function

' ======================================================================================
' CheckAndSuggestPremium: ซับรูทีนสำหรับเขียนคำแนะนำลงหน้าชีท
' วัตถุประสงค์: ให้คำแนะนำทันทีที่ผู้ใช้พิมพ์ทุนเสร็จ (พิมพ์ปุ๊บ ข้อความขึ้นปั๊บ)
' ======================================================================================
Public Sub CheckAndSuggestPremium(ByVal totalVal As Double)
    Dim valResult As Variant
    Dim QTSheet As Worksheet: Set QTSheet = ThisWorkbook.Worksheets("QT_อยู่ดีมีสุข")
    
    ' เรียกใช้ Logic กลาง
    valResult = GetPackageValidation(totalVal)
    
    ' ปลดล็อกชีทก่อนแก้ไข (ถ้ามีการป้องกันไว้)
    QTSheet.Unprotect Password:=myPassword
    
    ' เลือกจัดการข้อความในเซลล์ J43 ตามผลลัพธ์
    Select Case valResult(1)
        Case "Valid"
            ' --- กรณีถูกต้อง ---
            ' ล้างข้อความเตือนเดิมทิ้งให้สะอาด
            QTSheet.Range("J43:L43").ClearContents
            
        Case "OutOfRange"
            ' --- กรณีหลุดขอบเขต ---
            ' เขียนบอกช่วงที่ถูกต้อง และเปลี่ยนตัวอักษรเป็นสีแดงเพื่อเตือน
            QTSheet.Range("J43").Value = "**ต้องอยู่ระหว่าง " & Format(valResult(4), "#,##0") & _
                                         " ถึง " & Format(valResult(5), "#,##0")
            QTSheet.Range("J43").Font.Color = vbRed
            
        Case "Invalid"
            ' --- กรณีไม่ตรงแผนเป๊ะ ---
            ' เขียนบอกแผนที่แนะนำ และเปลี่ยนเป็นสีน้ำเงินให้ดูเป็นคำแนะนำ (ไม่ใช่ข้อผิดพลาดร้ายแรง)
            QTSheet.Range("J43").Value = "** แนะนำ: " & Format(valResult(2), "#,##0") & _
                                         " หรือ " & Format(valResult(3), "#,##0")
            QTSheet.Range("J43").Font.Color = vbBlue
    End Select
    
    ' ล็อกชีทกลับคืนหลังทำงานเสร็จ
    QTSheet.Protect Password:=myPassword
End Sub

' ======================================================================================
' UpdateLocationList: เวอร์ชันปรับปรุง (แก้ไขปัญหาลบ List ข้ามคอลัมน์)
' รับค่า: Mode (Amphoe/Tambon), Prov (จังหวัด), Amp (อำเภอ - ใช้เฉพาะกรณี Mode = Tambon)
' ======================================================================================
Public Sub UpdateLocationList(ByVal Mode As String, ByVal Prov As String, Optional ByVal Amp As String = "")
    Dim ws As Worksheet
    Dim rawData As Variant
    Dim resultData() As String
    Dim lastRow As Long, i As Long, count As Long
    Dim targetCol As String
    
    Set ws = ThisWorkbook.Worksheets("CF_อยู่ดีมีสุข") '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    targetCol = IIf(Mode = "Amphoe", "Z", "AA")
    
    Application.ScreenUpdating = False
    ws.Unprotect Password:=myPassword

    ' ----------------------------------------------------------------------
    ' แก้ไขจุดนี้: ล้างเฉพาะคอลัมน์ที่กำลังจะอัปเดตเท่านั้น
    ' ----------------------------------------------------------------------
    Dim targetLastRow As Long
    targetLastRow = ws.Cells(ws.Rows.count, targetCol).End(xlUp).Row
    
    If targetLastRow >= 2 Then
        ws.Range(ws.Cells(2, targetCol), ws.Cells(targetLastRow, targetCol)).ClearContents
    End If
    
    ' ถ้าอัปเดตอำเภอ (เลือกจังหวัดใหม่) ให้ล้างตำบลทิ้งด้วยเสมอเพราะอำเภอเปลี่ยนแล้ว
    If Mode = "Amphoe" Then
        Dim lastRowAA As Long
        lastRowAA = ws.Cells(ws.Rows.count, "AA").End(xlUp).Row
        If lastRowAA >= 2 Then ws.Range("AA2:AA" & lastRowAA).ClearContents
    End If
    ' ----------------------------------------------------------------------

    ' [ส่วนดึงข้อมูลเข้า Array เหมือนเดิม]
    lastRow = ws.Cells(ws.Rows.count, "T").End(xlUp).Row
    If lastRow < 2 Then GoTo CleanUp
    rawData = ws.Range("T2:V" & lastRow).Value

    ReDim resultData(1 To UBound(rawData, 1), 1 To 1)
    count = 0

    For i = 1 To UBound(rawData, 1)
        If Mode = "Amphoe" Then
            If rawData(i, 1) = Prov And rawData(i, 2) <> "" Then
                If Not IsInArray(CStr(rawData(i, 2)), resultData, count) Then
                    count = count + 1
                    resultData(count, 1) = rawData(i, 2)
                End If
            End If
        ElseIf Mode = "Tambon" Then
            ' กรองแบบ Strict: จังหวัด + อำเภอ
            If rawData(i, 1) = Prov And rawData(i, 2) = Amp And rawData(i, 3) <> "" Then
                If Not IsInArray(CStr(rawData(i, 3)), resultData, count) Then
                    count = count + 1
                    resultData(count, 1) = rawData(i, 3)
                End If
            End If
        End If
    Next i

    If count > 0 Then
        ws.Cells(2, targetCol).Resize(count, 1).Value = resultData
    End If

CleanUp:
    ws.Protect Password:=myPassword
    Application.ScreenUpdating = True
End Sub
ดกหฟดฟหกดกหฟฃ





