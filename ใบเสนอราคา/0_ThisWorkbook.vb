Private Sub Workbook_Open()
    ' ทุกครั้งที่เปิดไฟล์ ให้สั่งล็อกชีทด้วยฟังก์ชันกลางของคุณทันทีเพื่ออัปเดตค่า UserInterfaceOnly
    ' สมมติว่าสั่งล็อกชีทชื่อ QT_อยู่ดีมีสุข
    On Error Resume Next ' ป้องกันไว้ก่อนเผื่อชื่อชีทมีการเปลี่ยนแปลง
    Call SetSheetProtection(ThisWorkbook.Worksheets(Sheet1), True)
    msgbox "ยินดีต้อนรับสู่ระบบคำนวณเบี้ยประกันภัย Non-Motor Quotation!" & vbNewLine & _
           "โปรดตรวจสอบข้อมูลในชีท 'QT_อยู่ดีมีสุข' และกรอกข้อมูลให้ครบถ้วนเพื่อการคำนวณที่ถูกต้อง", _
           vbInformation, "Welcome"
End Sub