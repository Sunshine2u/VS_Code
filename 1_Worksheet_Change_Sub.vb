'==============================================
' (A) แจ้งเตือนจังหวัดกลุ่มเสี่ยงภัยน้ำท่วมที่ H28
' (B) ตรวจสอบทุนรวมที่ G43 ให้ตรงตามตาราง Premium Table แบบเป๊ะๆ
'     - หากไม่ตรง จะแสดงทุนแนะนำที่ใกล้เคียงที่สุด (Floor/Ceiling) ที่ L43
'     - ตรวจสอบช่วงทุนให้อยู่ระหว่าง 500,000 - 10,000,000
'     - ทำงานทั้งเมื่อมีการแก้ไข (Change) และเมื่อสูตรคำนวณใหม่ (Calculate)
'===================================================================


Private Sub Worksheet_Change(ByVal Target As Range)
    ' STEP 1: ตรวจสอบขอบเขตเซลล์ที่ต้องการดักจับ (จังหวัด H28, อำเภอ J28,จังหวัด H51, อำเภอ J51, ทุนประกัน G41-G42)
    If Intersect(Target, Me.Range("H28,J28,H51,J51,G41:H41,G42:H42")) Is Nothing Then Exit Sub

    On Error GoTo ErrorHandler
    
    ' STEP 2: เตรียมระบบก่อนเริ่มทำงาน
    Application.EnableEvents = False ' ปิด Event เพื่อป้องกัน Code รันซ้อนกันเอง
    Call SetSheetProtection(Me, False) ' ปลดล็อก Sheet ชั่วคราว

    Dim provName1 As String:provName1 = Trim$(CStr(Me.Range("H28").Value))'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Dim ampName1 As String :ampName1 = Trim$(CStr(Me.Range("J28").Value))'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Dim provName2 As String:provName2 = Trim$(CStr(Me.Range("H51").Value))'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Dim ampName2 As String :ampName2 = Trim$(CStr(Me.Range("J51").Value))'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

    ' ---------- (A1) กรณีเปลี่ยน "จังหวัด" เอาประกัน(H28) ----------
    If Not Intersect(Target, Me.Range("H28")) Is Nothing Then
        ' 1. ตรวจสอบพื้นที่เสี่ยงภัยน้ำท่วม
        Dim riskList As Variant: riskList = GetListRange("CF_Common", 1, "จังหวัดยกเว้นน้ำท่วม1")
        If Not IsError(Application.Match(provName1, riskList, 0)) Then
            MsgBox "พบว่าจังหวัด " & provName & " เป็นพื้นที่เสี่ยงภัยน้ำท่วม" & vbCrLf & _
                   "โปรดติดต่อเจ้าหน้าที่ MTI ผู้ดูแลตัวแทน ในการออกใบเสนอราคา", vbExclamation, "แจ้งเตือนความเสี่ยง"
        End If

        ' 2. อัปเดตรายชื่อ "อำเภอ" ลงในฐานข้อมูล (คอลัมน์ Z ใน CF_อยู่ดีมีสุข)
        ' และล้างค่าอำเภอ/ตำบลเดิมที่หน้าจอออกเพื่อให้เลือกใหม่
        Me.Range("J28,L28").ClearContents

        Call UpdateLocationList1("Amphoe", provName1)

        If len(Trim(Me.Range("G26").Text)) = 0 Then
            Me.Range("G26").Value = "     บ้านเลขที่.....หมู่ที่....อาคาร/หมู่บ้าน..... ซอย.... ถนน...."
            Me.Range("G26").Font.Color = RGB(166, 166, 166)
        End If

        If len(Trim(Me.Range("G49").Text)) = 0 Then
            Me.Range("G49").Value = "     บ้านเลขที่.....หมู่ที่....อาคาร/หมู่บ้าน..... ซอย.... ถนน...."
            Me.Range("G49").Font.Color = RGB(166, 166, 166)
        End If
        Me.Range("G26,G49").Font.Color = RGB(0, 0, 0)

        
    End If

    ' ---------- (B1) กรณีเปลี่ยน "อำเภอ" (J28) ----------
    If Not Intersect(Target, Me.Range("J28")) Is Nothing Then

        ' 1. ตรวจสอบว่ามีการเลือกจังหวัดไว้ก่อนหรือไม่
        If provName1 <> "" And ampName1 <> "" Then
            ' 2. อัปเดตรายชื่อ "ตำบล" ลงในฐานข้อมูล (คอลัมน์ AA ใน CF_อยู่ดีมีสุข)
            ' และล้างค่าตำบลเดิมที่หน้าจอ (L28) ออก
            Me.Range("L28").ClearContents
            Call UpdateLocationList1("Tambon", provName1, ampName1)
        End If
        
    End If


        ' ---------- (A2) กรณีเปลี่ยน "จังหวัด" เอาประกัน(H51) ----------
    If Not Intersect(Target, Me.Range("H51")) Is Nothing Then
        ' 2. อัปเดตรายชื่อ "อำเภอ" ลงในฐานข้อมูล (คอลัมน์  "W" ใน CF_อยู่ดีมีสุข)
        ' และล้างค่าอำเภอ/ตำบลเดิมที่หน้าจอออกเพื่อให้เลือกใหม่
        Me.Range("J51,L51:M51").ClearContents
        Call UpdateLocationList2("Amphoe", provName2)
        
    End If

    ' ---------- (B2) กรณีเปลี่ยน "อำเภอ" (J51) ----------
    If Not Intersect(Target, Me.Range("J51")) Is Nothing Then
        ' 1. ตรวจสอบว่ามีการเลือกจังหวัดไว้ก่อนหรือไม่
        If provName2 <> "" And ampName2 <> "" Then
            ' 2. อัปเดตรายชื่อ "ตำบล" ลงในฐานข้อมูล (คอลัมน์ AA ใน CF_อยู่ดีมีสุข)
            ' และล้างค่าตำบลเดิมที่หน้าจอ (L51) ออก
            Me.Range("L51:M51").ClearContents
            Call UpdateLocationList2("Tambon", provName2, ampName2)
        End If
        
    End If

        ' ---------- (C) กรณีเปลี่ยน "ทุนประกัน" (G41, G42) ----------
    
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

    If Not Intersect(Target, Me.Range("H28")) Is Nothing Then 'เมื่อกรอกชื่อจังหวัด
        If len(Trim(Me.Range("G49").Text)) = 0 Then
            Me.Range("G49").Value = "     บ้านเลขที่.....หมู่ที่....อาคาร/หมู่บ้าน..... ซอย.... ถนน...."
            Me.Range("G49").Font.Color = RGB(166, 166, 166)
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

