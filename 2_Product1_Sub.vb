Option Explicit

' ======================================================================================
' ส่วนที่ 1: การประกาศตัวแปรระดับ Global (Global Constants)
' ======================================================================================
' รหัสผ่าน (Password) ส่วนกลางที่ใช้สำหรับปลดล็อกแผ่นงาน (Sheet) และโครงสร้างไฟล์ (Workbook)
' ประกาศเป็น Public Const เพื่อให้ทุก Sub และ Function ในโปรเจกต์สามารถเรียกใช้ได้ทันที
Public Const myPassword As String = "QTMTI"
Public Const SheetLockSetting As Boolean = True 'ตั้งค่าให้ล็อกไฟล์โดย Default (True = ล็อก, False = ไม่ล็อก)
Public Const WorkbookLockSetting As Boolean = True 'ตั้งค่าให้ล็อกโครงสร้างไฟล์โดย Default (True = ล็อก, False = ไม่ล็อก)

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
    Call SetWorkbookProtection(False)
    
    ' แสดงแผ่นงาน Leaflet (ที่อาจถูกซ่อนอยู่)
    Sheets("LL_อยู่ดีมีสุข").Visible = True '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    ' ย้ายหน้าจอไปยังแผ่นงานนั้น
    Worksheets("LL_อยู่ดีมีสุข").Activate '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    ' ล็อกโครงสร้างไฟล์คืน เพื่อป้องกันการลบหรือสลับลำดับแผ่นงาน
    Call SetWorkbookProtection(WorkbookLockSetting)
End Sub

' ======================================================================================
' Sub สำหรับปิดหน้า Leaflet และกลับมายังหน้าใบเสนอราคาหลัก
' เมื่อผู้ใช้ปิดแผ่นพับรายละเอียดความคุ้มครอง จะทำการปลดล็อกโครงสร้างไฟล์เพื่อซ่อนแผ่นงาน Leaflet อีกครั้ง และย้ายหน้าจอกลับไปยังหน้าใบเสนอราคาหลัก
' ======================================================================================
Sub อยู่ดีมีสุข_Close_Leafltet()
    ' ปลดล็อกโครงสร้างไฟล์
    call setWorkbookProtection(False) ' ปลดล็อกโครงสร้างไฟล์ชั่วคราว
    
    ' ซ่อนแผ่นงาน Leaflet
    Sheets("LL_อยู่ดีมีสุข").Visible = False '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    ' กลับมายังหน้าคีย์ข้อมูลใบเสนอราคา
    Worksheets("QT_อยู่ดีมีสุข").Activate '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    ' ล็อกโครงสร้างไฟล์คืน
    Call SetWorkbookProtection(WorkbookLockSetting) 'ล็อกโครงสร้างไฟล์คืน
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
    Call SetWorkbookProtection(False)

    ' แสดงหน้าสรุป (Report) ชั่วคราวเพื่อใช้ในการทำ Print Preview
    wsQTR.Visible = True
    
    ' เปิดหน้าต่าง Preview
    wsQTR.PrintPreview
    
    ' ซ่อนหน้าสรุปกลับคืน
    wsQTR.Visible = False
    
    ' กลับมายังหน้าคีย์ข้อมูลหลัก
    wsKey.Activate
    
    ' ล็อกโครงสร้างไฟล์คืน
    Call SetWorkbookProtection(WorkbookLockSetting)
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
    Call SetWorkbookProtection(False)
    
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
    Call SetWorkbookProtection(WorkbookLockSetting)
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
' CheckAndSuggestPremium: ซับรูทีนสำหรับเขียนคำแนะนำลงหน้าชีท
' วัตถุประสงค์: ให้คำแนะนำทันทีที่ผู้ใช้พิมพ์ทุนเสร็จ (พิมพ์ปุ๊บ ข้อความขึ้นปั๊บ)
' ======================================================================================
Public Sub CheckAndSuggestPremium(ByVal totalVal As Double)
    Dim valResult As Variant
    Dim QTSheet As Worksheet: Set QTSheet = ThisWorkbook.Worksheets("QT_อยู่ดีมีสุข")
    Dim TableRange As Range: Set TableRange = ThisWorkbook.Worksheets("CF_อยู่ดีมีสุข").Range("A2:A46")
    
    ' เรียกใช้ Logic กลาง
    valResult = GetPackageValidation(totalVal, TableRange)
    
    ' ปลดล็อกชีทก่อนแก้ไข (ถ้ามีการป้องกันไว้)
    Call SetSheetProtection(QTSheet, False)
    
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
    Call SetSheetProtection(QTSheet, sheetLockSetting)
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
    
' --- STEP 1: ตั้งค่าเริ่มต้น และกำหนดเป้าหมาย ---
Set ws = ThisWorkbook.Worksheets("CF_อยู่ดีมีสุข")'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

' กำหนดคอลัมน์ปลายทาง: ถ้าหาอำเภอเขียนลง Z, ถ้าหาตำบลเขียนลง AA
targetCol = IIf(Mode = "Amphoe", "Z", "AA")'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

' ปิดการประมวลผลหน้าจอ (ทำให้โค้ดวิ่งเร็วขึ้น) และปลดล็อคชีท
Application.ScreenUpdating = False
call SetSheetProtection(ws, False) ' ปลดล็อกชีทชั่วคราว

' --- STEP 2: ล้างข้อมูลเก่า (Cleanup) ---
' หาบรรทัดสุดท้ายของคอลัมน์ที่จะเขียน เพื่อลบข้อมูลเดิมออกก่อนป้องกันข้อมูลค้าง
Dim targetLastRow As Long
targetLastRow = ws.Cells(ws.Rows.count, targetCol).End(xlUp).Row

If targetLastRow >= 2 Then
    ws.Range(ws.Cells(2, targetCol), ws.Cells(targetLastRow, targetCol)).ClearContents
End If

' พิเศษ: หากเป็นการเลือกจังหวัดใหม่ (Mode = Amphoe)
' ต้องล้างข้อมูลในคอลัมน์ตำบล (AA) ทิ้งด้วย เพราะอำเภอเดิมจะใช้ไม่ได้แล้ว
If Mode = "Amphoe" Then
    Dim lastRowAA As Long
    lastRowAA = ws.Cells(ws.Rows.count, "AA").End(xlUp).Row
    If lastRowAA >= 2 Then ws.Range("AA2:AA" & lastRowAA).ClearContents
End If

' --- STEP 3: ดึงข้อมูลจากฐานข้อมูลเข้า Array (เพื่อความรวดเร็ว) ---
' หาบรรทัดสุดท้ายของฐานข้อมูล (คอลัมน์ T)
lastRow = ws.Cells(ws.Rows.count, "T").End(xlUp).Row
If lastRow < 2 Then GoTo CleanUp ' ถ้าไม่มีข้อมูลเลย ให้ข้ามไปขั้นตอนสุดท้าย

' ดึงข้อมูลจังหวัด/อำเภอ/ตำบล (T-V) มาเก็บไว้ในตัวแปร Array
rawData = ws.Range("T2:V" & lastRow).Value

' เตรียมพื้นที่เก็บผลลัพธ์ (Array) ขนาดสูงสุดเท่ากับจำนวนข้อมูลที่มี
ReDim resultData(1 To UBound(rawData, 1), 1 To 1)
count = 0

' --- STEP 4: วนลูปคัดกรองข้อมูลตามเงื่อนไข ---
For i = 1 To UBound(rawData, 1)
    
    ' กรณีที่ 1: หา "อำเภอ" ของจังหวัดที่เลือก
    If Mode = "Amphoe" Then
        If rawData(i, 1) = Prov And rawData(i, 2) <> "" Then
            ' ตรวจสอบว่าชื่ออำเภอนี้ถูกเพิ่มไปหรือยัง (ป้องกันชื่อซ้ำ)
            If Not IsInArray(CStr(rawData(i, 2)), resultData, count) Then
                count = count + 1
                resultData(count, 1) = rawData(i, 2)
            End If
        End If
        
    ' กรณีที่ 2: หา "ตำบล" ของจังหวัดและอำเภอที่เลือก
    ElseIf Mode = "Tambon" Then
        If rawData(i, 1) = Prov And rawData(i, 2) = Amp And rawData(i, 3) <> "" Then
            ' ตรวจสอบชื่อซ้ำก่อนเพิ่มลงรายการ
            If Not IsInArray(CStr(rawData(i, 3)), resultData, count) Then
                count = count + 1
                resultData(count, 1) = rawData(i, 3)
            End If
        End If
    End If
Next i

' --- STEP 5: เขียนผลลัพธ์ที่กรองได้ลงใน Excel ---
If count > 0 Then
    ws.Cells(2, targetCol).Resize(count, 1).Value = resultData
End If

CleanUp:
    call SetSheetProtection(ws, sheetLockSetting) ' ล็อกชีทคืนหลังทำงานเสร็จ
    Application.ScreenUpdating = True
End Sub


