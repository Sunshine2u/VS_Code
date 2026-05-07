'==== Sheet: Cal_premium ==========================================
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
    Me.Unprotect Password:=myPassword
    
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
        list = GetList("CF_อยู่ดีมีสุข", 1, "จังหวัดยกเว้นน้ำท่วม")
        
        ' ใช้ Match ค้นหาชื่อจังหวัดใน Array ได้เลยไม่ต้องวน Loop
        result = Application.Match(prov, list, 0)
        
        If Not IsError(result) Then
            MsgBox "พบว่าจังหวัด " & prov & " เป็นพื้นที่เสี่ยงภัยน้ำท่วม" & vbCrLf & _
                           "โปรดติดต่อเจ้าหน้าที่ MTI ผู้ดูแลตัวแทน ในการออกใบเสนอราคา", _
                           vbExclamation, "แจ้งเตือนความเสี่ยง"
        End If
        call c
    End If
    
    ' ล็อกชีทคืน
    Me.Protect Password:=myPassword
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    MsgBox "เกิดข้อผิดพลาด: " & Err.Description, vbCritical, "Error"
    Me.Protect Password:=myPassword
    Application.EnableEvents = True
End Sub

' ---------- Sub Routine สำหรับรีเซ็ตระบบด้วยตนเอง ----------
Public Sub ResetExcelEvents()
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "ระบบ Event และ Calculation ถูกรีเซ็ตเรียบร้อยแล้ว", vbInformation
End Sub




Public Function IsFloodRisk(ByVal provinceName As String) As Boolean
    Dim list As Variant
    Dim result As Variant
    Dim cleanProv As String
    
    ' 1. ทำความสะอาดข้อมูลเบื้องต้น
    cleanProv = Trim$(CStr(provinceName))
    
    ' กรณีค่าว่าง ให้ถือว่าไม่พบความเสี่ยง
    If cleanProv = "" Then
        IsFloodRisk = False
        Exit Function
    End If
    
    ' 2. ดึงรายชื่อจังหวัดจากฐานข้อมูล (เรียกใช้ GetProvinceList ที่คุณมีอยู่)
    ' หมายเหตุ: ตรวจสอบให้แน่ใจว่าชื่อ Sheet และหัวตารางถูกต้องตามหน้างานจริง
    list = GetProvinceList("CF_อยู่ดีมีสุข", 1, "จังหวัดยกเว้นน้ำท่วม")
    
    ' 3. ตรวจสอบด้วย Match
    ' หาก list ว่างเปล่า (หาคอลัมน์ไม่เจอ) จะไม่เกิด Error แต่จะข้ามไป False
    If IsArray(list) Then
        result = Application.Match(cleanProv, list, 0)
        
        If Not IsError(result) Then
            IsFloodRisk = True
            Exit Function
        End If
    End If
    
    IsFloodRisk = False
End Function