Public Sub อยู่ดีมีสุข_Clear_Input()
    Dim QTSheet As Worksheet
    
    On Error GoTo ClearErrorHandler
    Set QTSheet = ThisWorkbook.Worksheets("QT_อยู่ดีมีสุข")
    
    ' ปิดระบบ Event ชั่วคราวป้องกันการขัดจังหวะขณะล้างข้อมูลหลายๆ เซลล์พร้อมกัน
    Application.EnableEvents = False
    
   With QTSheet
        ' ปลดล็อกแผ่นงาน
        .Unprotect Password:=myPassword
        
        ' 1. ประกาศตัวแปร Array เพื่อเก็บเฉพาะ "ชื่อเซลล์แรกสุด (Top-Left Cell)" ของแต่ละช่อง
        Dim cleanRanges As Variant
        Dim cellName As Variant
        
        ' 2. นำชื่อเซลล์ทั้งหมดมารวมกันไว้ใน Array เดียว (เรียงลำดับตามต้องการได้เลย)
        cleanRanges = Array( _
            "G24", _
            "G26" _
        )
        
        ' 3. ใช้ For Each เพื่อ Loop ดึงชื่อเซลล์ออกมา Set ค่าเป็นว่างทีละช่อง
        For Each cellName In cleanRanges
            ' ใช้ .Value = "" ตรงไปที่เซลล์นั้นๆ ปลอดภัยจาก Error 1004 แน่นอน
            .Range(cellName).Value = ""
        Next cellName

        ' ล็อกแผ่นงานคืนกลับ
        .Protect Password:=myPassword
    End With
    
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