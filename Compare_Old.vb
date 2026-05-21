Public Sub อยู่ดีมีสุข_Clear_Input()
    Dim QTSheet As Worksheet
    Dim i As Long
    
    On Error GoTo ClearErrorHandler
    Set QTSheet = ThisWorkbook.Worksheets("QT_อยู่ดีมีสุข") '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    ' ปิดระบบ Event ชั่วคราวป้องกันการขัดจังหวะขณะล้างข้อมูลหลายๆ เซลล์พร้อมกัน
    Application.EnableEvents = False
    
    With QTSheet
        ' ปลดล็อกแผ่นงาน
        .Unprotect Password:=myPassword
        
        ' แก้ไขปัญหา Merged Cells พังด้วยการรวบรวม Address และล้างทีละส่วนอย่างเป็นระบบ
        ' แบ่งชุดเซลล์ที่ไม่มีปัญหาในการล้างด้วย ClearContents ออกเป็นกลุ่มๆ
        ' (แนะนำให้เขียนแยกกลุ่มแบบนี้ เพื่อให้อ่านและปรับปรุงโค้ดได้ง่าย ไม่เกิด Syntax Error)
        
        ' กลุ่ม 1: ข้อมูลผู้เอาประกัน และ ที่อยู่จัดส่ง
        .Range("G24:M26,H28,J28,L28").MergeArea.ClearContents
        
        ' กลุ่ม 2: ข้อมูลสิ่งปลูกสร้าง/ทรัพย์สิน
        .Range("G41:I41,G42:H42,L41:N41").MergeArea.ClearContents
        
        ' กลุ่ม 3: ข้อมูลรายละเอียดและส่วนควบอื่นๆ (ล้างข้อมูลที่ระบุในโค้ดต้นฉบับเดิมของคุณ)
        Dim targetRanges As Variant
        targetRanges = Array("G31:H33", "H35:H36", "L35:L38", "G45:J45", "G49:M49", _
                             "H51", "J51", "L51", "H53", "J53", "L53", "G57:M57", _
                             "H59", "J59", "L59", "H61", "G64:I64", "L64:M64", _
                             "G66:I66", "G68:H68")
                             
        ' วนลูปเพื่อใช้คำสั่งล้างข้อมูลผ่าน .MergeArea เพื่อความปลอดภัยกับ Merged Cells 100%
        For i = LBound(targetRanges) To UBound(targetRanges)
            .Range(targetRanges(i)).MergeArea.ClearContents
        Next i
        
        ' ล็อกแผ่นงานคืนกลับ
        .Protect Password:=myPassword
    End With
    
    MsgBox "ล้างข้อมูลหน้าแบบฟอร์มเรียบร้อยแล้ว", vbInformation, "ล้างข้อมูล"

ClearSafeExit:
    Application.EnableEvents = True
        ' เรียกฟังก์ชันรีเซ็ตค่าระบบเสริม เพื่อให้มั่นใจว่า Excel กลับมาอยู่ในสถานะพร้อมทำงานปกติ
    Call ResetExcelEvents
    
    ' จบการทำงานของ Sub หลัก (ป้องกันไม่ให้โค้ดไหลไปทำงานใน ErrorHandler)
    Exit Sub

    

ClearErrorHandler:
    MsgBox "เกิดข้อผิดพลาดขณะล้างข้อมูล: " & Err.Description, vbCritical, "Error"
    ' เปิดระบบคืนทุกครั้งแม้โค้ดจะทำงานผิดพลาด
    If Not QTSheet Is Nothing Then QTSheet.Protect Password:=myPassword
    Application.EnableEvents = True
End Sub