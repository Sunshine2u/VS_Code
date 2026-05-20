' ############################################################################################################
' Sub และ Function สำหรับจัดการการกรอกข้อมูลและการออกเอกสารใบเสนอราคา (Quotation) ของแผนประกันภัย "อยู่ดีมีสุข"\
' ** ต้องเปลี่ยน Sheet Name ในโค้ดให้ตรงกับชื่อแผ่นงานจริงในไฟล์ Excel ของคุณ
' มีการเขียนข้อมูลในShett "QT_อยู่ดีมีสุข" และการจัดการข้อมูลใน Sheet "CF_อยู่ดีมีสุข" เป็นหลัก
' โดยมีการจัดการทั้งในส่วนของการล้างข้อมูล, การนำทางไปยังแผ่นพับรายละเอียดความคุ้มครอง (Leaflet), การตรวจสอบเงื่อนไขก่อนออกเอกสาร, และการสร้างไฟล์ PDF ของใบเสนอราคา
' #############################################################################################################

Option Explicit

' ======================================================================================
' การประกาศตัวแปรระดับ Global (Global Constants)
' ======================================================================================
' รหัสผ่าน (Password) ส่วนกลางที่ใช้สำหรับปลดล็อกแผ่นงาน (Sheet) และโครงสร้างไฟล์ (Workbook)
' ประกาศเป็น Public Const เพื่อให้ทุก Sub และ Function ในโปรเจกต์สามารถเรียกใช้ได้ทันที
Public Const myPassword As String = "QTMTI"
Public Const SheetLockSetting As Boolean = True 'ตั้งค่าให้ล็อกไฟล์โดย Default (True = ล็อก, False = ไม่ล็อก)
Public Const WorkbookLockSetting As Boolean = False 'ตั้งค่าให้ล็อกโครงสร้างไฟล์โดย Default (True = ล็อก, False = ไม่ล็อก)

' ======================================================================================
' ระบบจัดการการกรอกข้อมูล (Data Clearing & Navigation)
' ======================================================================================

' Sub สำหรับล้างข้อมูล (Reset) ในแบบฟอร์มใบเสนอราคาหน้าหลัก
Public Sub อยู่ดีมีสุข_Clear_Input()
    Dim QTSheet As Worksheet
    
    On Error GoTo ClearErrorHandler
    Set QTSheet = ThisWorkbook.Worksheets("QT_อยู่ดีมีสุข")'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    ' ปิดระบบ Event ชั่วคราวป้องกันการขัดจังหวะขณะล้างข้อมูลหลายๆ เซลล์พร้อมกัน
    Application.EnableEvents = False

    Call setSheetProtection(QTSheet, False) ' ปลดล็อก Sheet ชั่วคราวเพื่อให้สามารถแก้ไขเซลล์ได้
    Call อยู่ดีมีสุข_Hide_Address_Rows() ' เปิดแถวที่อยู่ก่อนกรอกข้อมูล เพื่อความสวยงามของหน้าจอขณะทำงาน

   With QTSheet
        
        ' 1. ประกาศตัวแปร Array เพื่อเก็บเฉพาะ "ชื่อเซลล์แรกสุด (Top-Left Cell)" ของแต่ละช่อง
        Dim cleanRanges As Variant
        Dim cellName As Variant
        
        ' 2. นำชื่อเซลล์ทั้งหมดมารวมกันไว้ใน Array เดียว (เรียงลำดับตามต้องการได้เลย)
        ' ใส่แค่เซลล์แรก ในกรณีเป็น Merged Cells
        cleanRanges = Array( _
            "G24", "G26","H28","J28","L28", _
            "G31", "G33", "H35","H36", "L35" ,"L36" , "L38","G41", "G42", "G43","G45", _
            "G49","G57", "H51", "H53", "H59", "H61", "J51", "J53","J59","L51","L53","L59","F55", _
            "G64","G66","G68","L64" _
        ) '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        
        ' 3. ใช้ For Each เพื่อ Loop ดึงชื่อเซลล์ออกมา Set ค่าเป็นว่างทีละช่อง
        For Each cellName In cleanRanges
            ' ใช้ .Value = "" ตรงไปที่เซลล์นั้นๆ ปลอดภัยจาก Error 1004 แน่นอน
            .Range(cellName).Value = ""
        Next cellName


    End With

    Call SetSheetProtection(QTSheet, SheetLockSetting) ' ล็อก Sheet คืนตามค่าใน Const
    
    MsgBox "ล้างข้อมูลหน้าแบบฟอร์มเรียบร้อยแล้ว", vbInformation, "ล้างข้อมูล"

ClearSafeExit:
    Application.EnableEvents = True
    Exit Sub

ClearErrorHandler:
    MsgBox "เกิดข้อผิดพลาดขณะล้างข้อมูล: " & Err.Description, vbCritical, "Error"
    ' เปิดระบบคืนทุกครั้งแม้โค้ดจะทำงานผิดพลาด
    If Not QTSheet Is Nothing Then QTSheet.Protect Password:=myPassword
    Application.EnableEvents = True
End Sub

' =======================================================================================
' Sub สำหรับเปิดหน้า Leaflet (แผ่นพับรายละเอียดความคุ้มครอง)
' ' ในกรณีที่แผ่นงานถูกซ่อนอยู่ จะทำการปลดล็อกโครงสร้างไฟล์เพื่อแสดงแผ่นงานนั้น และย้ายหน้าจอไปยังแผ่นงาน Leaflet
' =======================================================================================
Sub อยู่ดีมีสุข_Go_To_Leaflet()
    ' ปลดล็อกโครงสร้างไฟล์เพื่อให้สามารถเปลี่ยนสถานะการซ่อนของแผ่นงานได้
    Call SetWorkbookProtection(False)

    Dim leafletSheet As Worksheet
    Set leafletSheet = ThisWorkbook.Worksheets("LL_อยู่ดีมีสุข") '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    'แสดงแผ่นงาน Leaflet (ที่อาจถูกซ่อนอยู่)
    leafletSheet.Visible = True 
    ' ย้ายหน้าจอไปยังแผ่นงานนั้น
    leafletSheet.Activate 
    
    ' ล็อกโครงสร้างไฟล์คืน เพื่อป้องกันการลบหรือสลับลำดับแผ่นงาน
    Call SetWorkbookProtection(WorkbookLockSetting)
End Sub

' ======================================================================================
' Sub สำหรับปิดหน้า Leaflet และกลับมายังหน้าใบเสนอราคาหลัก
' เมื่อผู้ใช้ปิดแผ่นพับรายละเอียดความคุ้มครอง จะทำการปลดล็อกโครงสร้างไฟล์เพื่อซ่อนแผ่นงาน Leaflet อีกครั้ง และย้ายหน้าจอกลับไปยังหน้าใบเสนอราคาหลัก
' ======================================================================================
Sub อยู่ดีมีสุข_Close_Leafltet()
    ' ปลดล็อกโครงสร้างไฟล์
    Call SetWorkbookProtection(False) ' ปลดล็อกโครงสร้างไฟล์ชั่วคราว
    Dim leafletSheet As Worksheet
    Dim QTSheet As Worksheet

    Set leafletSheet = ThisWorkbook.Worksheets("LL_อยู่ดีมีสุข") '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Set QTSheet = ThisWorkbook.Worksheets("QT_อยู่ดีมีสุข") '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

    ' ซ่อนแผ่นงาน Leaflet
    leafletSheet.Visible = False
    ' กลับมายังหน้าคีย์ข้อมูลใบเสนอราคา
    QTSheet.Activate
    
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
    Dim QTSheet As Worksheet
    Dim QTRSheet As Worksheet
    Dim ProvName As String
    Dim PremVal As Double
    Dim PremRange As Range

    Set QTSheet = ThisWorkbook.Worksheets("QT_อยู่ดีมีสุข") '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Set QTRSheet = ThisWorkbook.Worksheets("QTR_อยู่ดีมีสุข") '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Set PremRange = GetListRange(Sheet3, 1, "ทุนประกันภัย") '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    ProvName = QTSheet.Range("H28").Value '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    PremVal = QTSheet.Range("G43").Value '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    ' 1. ตรวจสอบเงื่อนไขจังหวัดน้ำท่วมก่อน
    If IsFloodRisk(ProvName) Then
        MsgBox "ไม่สามารถออกใบเสนอราคาได้!" & vbCrLf & _
               "จังหวัดที่ระบุอยู่ในพื้นที่เสี่ยงภัยน้ำท่วม" & vbCrLf & _
               "กรุณาติดต่อเจ้าหน้าที่ดูแลตัวแทน MTI เพื่อออกใบเสนอราคาให้ท่าน", vbExclamation, "ตรวจสอบพื้นที่เสี่ยงภัย"
        Exit Sub
    End If

    ' 2. ตรวจสอบเงื่อนไขความถูกต้องของทุนประกัน (แยกออกมาเป็นอีก If หนึ่ง ไม่ซ้อนกัน)
    ' ส่งค่าจาก G43 เข้าไปตรวจสอบในฟังก์ชัน IsPremiumValid ที่เราปรับปรุงไว้
    If Not IsPremiumValid(PremVal, PremRange) Then
        Exit Sub ' ถ้าฟังก์ชันคืนค่า False (ทุนไม่ตรงแผน) จะหยุดทำงานทันทีพร้อมแจ้งเตือนจากในฟังก์ชันเอง
    End If
    
    ' --- หากผ่านทั้ง 2 ด่านด้านบน ถึงจะเริ่มกระบวนการ Preview ---

    ' ปลดล็อกโครงสร้างไฟล์
    Call SetWorkbookProtection(False)

    ' แสดงหน้าสรุป (Report) ชั่วคราวเพื่อใช้ในการทำ Print Preview
    QTRSheet.Visible = True 
    
    ' เปิดหน้าต่าง Preview
    QTRSheet.PrintPreview
    
    ' ซ่อนหน้าสรุปกลับคืน
    QTRSheet.Visible = False
    
    ' กลับมายังหน้าคีย์ข้อมูลหลัก
    QTSheet.Activate
    
    ' ล็อกโครงสร้างไฟล์คืน
    Call SetWorkbookProtection(WorkbookLockSetting)
End Sub

'========================================================================================
' Sub สำหรับซ่อนแถวกรณีที่มีการเลือกที่อยู่รับเอกสารเป็นที่เดียวกับที่อยู่เอาประกัน (F55)
'=========================================================================================

Sub อยู่ดีมีสุข_Hide_Address_Rows(Optional flag As Boolean = False)
    Dim QTSheet As Worksheet
    Dim Checkbox As Range
    Dim HiddingRows As Range

    Set QTSheet = ThisWorkbook.Worksheets("QT_อยู่ดีมีสุข") '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Set Checkbox = QTSheet.Range("F55") '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Set HiddingRows = QTSheet.Rows("56:61") '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

    Call SetSheetProtection(QTSheet, False) ' ปลดล็อก Sheet ชั่วคราว
    If Checkbox.Value = True  Then
        HiddingRows.Hidden = True
    Else
        HiddingRows.Hidden = False
    End If
    Call SetSheetProtection(QTSheet, SheetLockSetting) ' ล็อก Sheet คืนตามค่าใน Const
End Sub


' =======================================================================================
'กดเพื่อสร้างใบเสนอราคา PDF
' จะตรวจสอบเงื่อนไขพื้นที่เสี่ยงภัยน้ำท่วมและความถูกต้องของทุนประกันก่อน หากผ่านถึงจะปลดล็อกโครงสร้างไฟล์เพื่อแสดงหน้าสรุป (Report) ชั่วคราวสำหรับทำการ Export เป็น PDF
' หลังจาก Export เสร็จจะล็อกโครงสร้างไฟล์คืนและซ่อนหน้าสรุปอีกครั้งเพื่อป้องกันการแก้ไขข้อมูลในหน้าสรุปโดยไม่ตั้งใจ
' =======================================================================================
Public Sub อยู่ดีมีสุข_Get_Quotation()
    Dim QTSheet As Worksheet
    Dim QTRSheet As Worksheet
    Dim filePath As String, fileName As String
    Dim ProvName As String
    Dim PremVal As Double
    Dim PremRange As Range
    Dim QTNumber As String

    Set QTSheet = ThisWorkbook.Worksheets("QT_อยู่ดีมีสุข") '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Set QTRSheet = ThisWorkbook.Worksheets("QTR_อยู่ดีมีสุข") '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Set PremRange = GetListRange(Sheet3, 1, "ทุนประกันภัย") '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    ProvName = QTSheet.Range("H28").Value '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    PremVal = QTSheet.Range("G43").Value '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    QTNumber = QTSheet.Range("G58").Value '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>


    ' 1. ตรวจสอบเงื่อนไขจังหวัดน้ำท่วมก่อน
    If IsFloodRisk(ProvName) Then
        MsgBox "ไม่สามารถออกใบเสนอราคาได้!" & vbCrLf & _
               "จังหวัดที่ระบุอยู่ในพื้นที่เสี่ยงภัยน้ำท่วม" & vbCrLf & _
               "กรุณาติดต่อเจ้าหน้าที่ดูแลตัวแทน MTI เพื่อออกใบเสนอราคาให้ท่าน", vbExclamation, "ตรวจสอบพื้นที่เสี่ยงภัย"
        Exit Sub
    End If

    ' 2. ตรวจสอบเงื่อนไขความถูกต้องของทุนประกัน (แยกออกมาเป็นอีก If หนึ่ง ไม่ซ้อนกัน)
    ' ส่งค่าจาก G43 เข้าไปตรวจสอบในฟังก์ชัน IsPremiumValid ที่เราปรับปรุงไว้
    If Not IsPremiumValid(PremVal, PremRange) Then 
        Exit Sub ' ถ้าฟังก์ชันคืนค่า False (ทุนไม่ตรงแผน) จะหยุดทำงานทันทีพร้อมแจ้งเตือนจากในฟังก์ชันเอง
    End If
    
    ' ปลดล็อก Workbook เพื่อให้สามารถทำงานต่อได้
    Call SetWorkbookProtection(False)
    
    QTRSheet.Visible = True
    
    On Error GoTo ErrorHandler
 
    
    ' 1. กำหนดชื่อและที่เก็บไฟล์
    fileName = "ใบเสนอราคา_อยู่ดีมีสุข_" & Format(Now, "yyyy-mm-dd_hhmm") & QTNumber & ".pdf" 
    filePath = ThisWorkbook.Path & "\" & fileName
    
    ' 2. คำสั่ง Export เป็น PDF
    QTRSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=filePath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
        
        QTSheet.Activate ' กลับมายังหน้าคีย์ข้อมูลหลัก
        ' แจ้งเตือนเมื่อสำเร็จ
    MsgBox "บันทึกไฟล์ PDF เรียบร้อยแล้วที่: " & vbCrLf & filePath, vbInformation, "สำเร็จ"


' ส่วนนี้จะทำงานเสมอไม่ว่าจะสำเร็จหรือ Error เพื่อล็อกไฟล์คืน
Finalize:
    QTRSheet.Visible = False
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
    Dim QTSheet As Worksheet
    Dim CFSheet As Worksheet: 
    Dim PremiumRange As Range:
    Dim lastRow As Long, ColPremium As Long: 
    
    Set QTSheet = ThisWorkbook.Worksheets("QT_อยู่ดีมีสุข")'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Set CFSheet = ThisWorkbook.Worksheets("CF_อยู่ดีมีสุข") '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    ColPremium = 1 ' สมมติว่าตารางทุนประกันภัยอยู่ในคอลัมน์ A (ปรับตามจริงได้)'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

    ' กำหนดช่วงข้อมูลของตารางทุนประกันภัย (สมมติอยู่ในคอลัมน์ A-B ของ CF_อยู่ดีมีสุข)
    lastRow = CFSheet.Cells(CFSheet.Rows.count, ColPremium).End(xlUp).Row
    Set PremiumRange = CFSheet.Range(CFSheet.Cells(2, ColPremium), CFSheet.Cells(lastRow, ColPremium))

    ' เรียกใช้ Logic กลาง
    valResult = GetPackageValidation(totalVal, PremiumRange)
    
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
    Call SetSheetProtection(QTSheet, SheetLockSetting)
End Sub

' ############################################################################################################
' Sub สำหรับจัดการชีทและเขียนลิสต์อำเภอ,ตำบล (Action)
' ############################################################################################################
Public Sub UpdateLocationList1(ByVal Mode As String, ByVal Prov As String, Optional ByVal Amp As String = "")
    Dim wsTarget As Worksheet
    Dim arrResults As Variant
    Dim targetCol As Integer, colAmphoe As Integer, colTambon As Integer
    Dim targetLastRow As Long
    
    ' --- CONFIG: กำหนดเป้าหมาย ---
    Set wsTarget = ThisWorkbook.Worksheets("CF_อยู่ดีมีสุข") '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    colAmphoe = FindHeaderColumn(wsTarget, 1, "Amphoe1") ' คอลัมน์สำหรับ LIST_อำเภอ'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    colTambon = FindHeaderColumn(wsTarget, 1, "Tambon1") ' คอลัมน์สำหรับ LIST_ตำบล'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    targetCol = IIf(Mode = "Amphoe", colAmphoe, colTambon)
    
    Application.ScreenUpdating = False
    Call SetSheetProtection(wsTarget, False)
    
    ' 1. ล้างข้อมูลเก่า (Cleanup)
    targetLastRow = wsTarget.Cells(wsTarget.Rows.count, targetCol).End(xlUp).Row
    If targetLastRow >= 2 Then
        wsTarget.Range(wsTarget.Cells(2, targetCol), wsTarget.Cells(targetLastRow, targetCol)).ClearContents
    End If
    
    ' พิเศษ: หากเลือกจังหวัดใหม่ ให้ล้างรายการตำบลเดิมในชีทปลายทางทิ้งด้วย
    If Mode = "Amphoe" Then
        Dim lastRowV As Long
        lastRowV = wsTarget.Cells(wsTarget.Rows.count, colTambon).End(xlUp).Row
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
    Call SetSheetProtection(wsTarget, SheetLockSetting)
    Application.ScreenUpdating = True
End Sub

' ############################################################################################################
' Sub สำหรับจัดการชีทและเขียนลิสต์อำเภอ,ตำบล (Action) ชุดที่ 2
' ############################################################################################################
Public Sub UpdateLocationList2(ByVal Mode As String, ByVal Prov As String, Optional ByVal Amp As String = "")
    Dim wsTarget As Worksheet
    Dim arrResults As Variant
    Dim targetCol As Integer, colAmphoe As Integer, colTambon As Integer
    Dim targetLastRow As Long
    
    ' --- CONFIG: กำหนดเป้าหมาย ---
    Set wsTarget = ThisWorkbook.Worksheets("CF_อยู่ดีมีสุข") '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    colAmphoe = FindHeaderColumn(wsTarget, 1, "Amphoe2") ' คอลัมน์สำหรับ LIST_อำเภอ'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    colTambon = FindHeaderColumn(wsTarget, 1, "Tambon2") ' คอลัมน์สำหรับ LIST_ตำบล'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    targetCol = IIf(Mode = "Amphoe", colAmphoe, colTambon)
    
    Application.ScreenUpdating = False
    Call SetSheetProtection(wsTarget, False)
    
    ' 1. ล้างข้อมูลเก่า (Cleanup)
    targetLastRow = wsTarget.Cells(wsTarget.Rows.count, targetCol).End(xlUp).Row
    If targetLastRow >= 2 Then
        wsTarget.Range(wsTarget.Cells(2, targetCol), wsTarget.Cells(targetLastRow, targetCol)).ClearContents
    End If
    
    ' พิเศษ: หากเลือกจังหวัดใหม่ ให้ล้างรายการตำบลเดิมในชีทปลายทางทิ้งด้วย
    If Mode = "Amphoe" Then
        Dim lastRowV As Long
        lastRowV = wsTarget.Cells(wsTarget.Rows.count, colTambon).End(xlUp).Row
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
    Call SetSheetProtection(wsTarget, SheetLockSetting)
    Application.ScreenUpdating = True
End Sub

' ############################################################################################################
' Sub สำหรับจัดการชีทและเขียนลิสต์อำเภอ,ตำบล (Action) ชุดที่ 3
' ############################################################################################################
Public Sub UpdateLocationList3(ByVal Mode As String, ByVal Prov As String, Optional ByVal Amp As String = "")
    Dim wsTarget As Worksheet
    Dim arrResults As Variant
    Dim targetCol As Integer, colAmphoe As Integer, colTambon As Integer
    Dim targetLastRow As Long
    
    ' --- CONFIG: กำหนดเป้าหมาย ---
    Set wsTarget = ThisWorkbook.Worksheets("CF_อยู่ดีมีสุข") '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    colAmphoe = FindHeaderColumn(wsTarget, 1, "Amphoe3") ' คอลัมน์สำหรับ LIST_อำเภอ'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    colTambon = FindHeaderColumn(wsTarget, 1, "Tambon3") ' คอลัมน์สำหรับ LIST_ตำบล'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    targetCol = IIf(Mode = "Amphoe", colAmphoe, colTambon)
    
    Application.ScreenUpdating = False
    Call SetSheetProtection(wsTarget, False)
    
    ' 1. ล้างข้อมูลเก่า (Cleanup)
    targetLastRow = wsTarget.Cells(wsTarget.Rows.count, targetCol).End(xlUp).Row
    If targetLastRow >= 2 Then
        wsTarget.Range(wsTarget.Cells(2, targetCol), wsTarget.Cells(targetLastRow, targetCol)).ClearContents
    End If
    
    ' พิเศษ: หากเลือกจังหวัดใหม่ ให้ล้างรายการตำบลเดิมในชีทปลายทางทิ้งด้วย
    If Mode = "Amphoe" Then
        Dim lastRowV As Long
        lastRowV = wsTarget.Cells(wsTarget.Rows.count, colTambon).End(xlUp).Row
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
    Call SetSheetProtection(wsTarget, SheetLockSetting)
    Application.ScreenUpdating = True
End Sub