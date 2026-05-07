' ############################################################################################################
' ส่วนที่ 1: Worksheet Event (สำหรับชีท QT_อยู่ดีมีสุข)
' ############################################################################################################

Private Sub Worksheet_Change(ByVal Target As Range)
    ' STEP 1: ตรวจสอบขอบเขตเซลล์ที่ต้องการดักจับ (จังหวัด H28, อำเภอ J28, ทุนประกัน G41-G42)
    If Intersect(Target, Me.Range("H28,J28,G41:H41,G42:H42")) Is Nothing Then Exit Sub

    On Error GoTo ErrorHandler
    
    ' STEP 2: เตรียมระบบก่อนเริ่มทำงาน
    Application.EnableEvents = False ' ปิด Event เพื่อป้องกัน Code รันซ้อนกันเอง
    Call SetSheetProtection(Me, False) ' ปลดล็อก Sheet ชั่วคราว

    Dim provName As String: provName = Trim$(CStr(Me.Range("H28").Value))
    Dim ampName As String: ampName = Trim$(CStr(Me.Range("J28").Value))

    ' ---------- (A) กรณีเปลี่ยน "จังหวัด" (H28) ----------
    If Not Intersect(Target, Me.Range("H28")) Is Nothing Then
        
        ' 1. ตรวจสอบพื้นที่เสี่ยงภัยน้ำท่วม
        Dim riskList As Variant: riskList = GetListRange("CF_Common", 1, "จังหวัดยกเว้นน้ำท่วม1")
        If Not IsError(Application.Match(provName, riskList, 0)) Then
            MsgBox "พบว่าจังหวัด " & provName & " เป็นพื้นที่เสี่ยงภัยน้ำท่วม" & vbCrLf & _
                   "โปรดติดต่อเจ้าหน้าที่ MTI ผู้ดูแลตัวแทน ในการออกใบเสนอราคา", vbExclamation, "แจ้งเตือนความเสี่ยง"
        End If

        ' 2. อัปเดตรายชื่อ "อำเภอ" ลงในฐานข้อมูล (คอลัมน์ Z ใน CF_อยู่ดีมีสุข)
        ' และล้างค่าอำเภอ/ตำบลเดิมที่หน้าจอออกเพื่อให้เลือกใหม่
        Me.Range("J28,L28").ClearContents
        Call UpdateLocationList("Amphoe", provName)
        
    End If

    ' ---------- (B) กรณีเปลี่ยน "อำเภอ" (J28) ----------
    If Not Intersect(Target, Me.Range("J28")) Is Nothing Then
        
        ' 1. ตรวจสอบว่ามีการเลือกจังหวัดไว้ก่อนหรือไม่
        If provName <> "" And ampName <> "" Then
            ' 2. อัปเดตรายชื่อ "ตำบล" ลงในฐานข้อมูล (คอลัมน์ AA ใน CF_อยู่ดีมีสุข)
            ' และล้างค่าตำบลเดิมที่หน้าจอ (L28) ออก
            Me.Range("L28").ClearContents
            Call UpdateLocationList("Tambon", provName, ampName)
        End If
        
    End If

    ' ---------- (C) กรณีเปลี่ยน "ทุนประกัน" (G41, G42) ----------
    If Not Intersect(Target, Me.Range("G41,G42")) Is Nothing Then
        
        ' 1. ตรวจสอบช่องเฟอร์นิเจอร์ (G42)
        If Len(Trim(Me.Range("G42").Text)) = 0 Then
            Me.Range("J42").Value = "ถ้าไม่มีให้กรอกเลข 0"
            Me.Range("J42").Font.Color = RGB(255, 0, 0)
        Else
            Me.Range("J42").ClearContents
        End If
        
        ' 2. คำนวณทุนรวม (G43) และแสดงคำแนะนำ Package (J43)
        If IsNumeric(Me.Range("G41").Value) And IsNumeric(Me.Range("G42").Value) Then
            Me.Range("G43").Value = Me.Range("G41").Value + Me.Range("G42").Value
            Call CheckAndSuggestPremium(Me.Range("G43").Value)
        Else
            Me.Range("G43").ClearContents
            Me.Range("J43:L43").ClearContents
        End If
        
    End If

    ' STEP 3: จบการทำงานและคืนค่าระบบ
    Call SetSheetProtection(Me, FileLockSetting)
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    MsgBox "เกิดข้อผิดพลาด: " & Err.Description, vbCritical, "System Error"
    Call SetSheetProtection(Me, FileLockSetting)
    Application.EnableEvents = True
End Sub

' ############################################################################################################
' ส่วนที่ 2: ฟังก์ชันสำหรับอัปเดตรายการ Dropdown (UpdateLocationList)
' ############################################################################################################

Public Sub UpdateLocationList(ByVal Mode As String, ByVal Prov As String, Optional ByVal Amp As String = "")
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("CF_อยู่ดีมีสุข")
    Dim targetCol As String
    Dim rawData As Variant
    Dim resultData() As String
    Dim i As Long, count As Long, lastRow As Long
    
    ' STEP 1: กำหนดเป้าหมาย (Mode "Amphoe" -> คอลัมน์ Z, Mode "Tambon" -> คอลัมน์ AA)
    targetCol = IIf(Mode = "Amphoe", "Z", "AA")
    Application.ScreenUpdating = False
    ws.Unprotect Password:=myPassword

    ' STEP 2: ล้างข้อมูลเก่าในคอลัมน์เป้าหมาย
    Dim targetLastRow As Long: targetLastRow = ws.Cells(ws.Rows.count, targetCol).End(xlUp).Row
    If targetLastRow >= 2 Then
        ws.Range(ws.Cells(2, targetCol), ws.Cells(targetLastRow, targetCol)).ClearContents
    End If

    ' พิเศษ: ถ้าเลือกจังหวัดใหม่ ต้องล้างตำบล (AA) ทิ้งด้วยเสมอ
    If Mode = "Amphoe" Then
        Dim lastRowAA As Long: lastRowAA = ws.Cells(ws.Rows.count, "AA").End(xlUp).Row
        If lastRowAA >= 2 Then ws.Range("AA2:AA" & lastRowAA).ClearContents
    End If

    ' STEP 3: คัดกรองข้อมูลจาก Master Data (คอลัมน์ T, U, V)
    lastRow = ws.Cells(ws.Rows.count, "T").End(xlUp).Row
    If lastRow < 2 Then GoTo CleanUp
    
    rawData = ws.Range("T2:V" & lastRow).Value
    ReDim resultData(1 To UBound(rawData, 1), 1 To 1)
    count = 0

    For i = 1 To UBound(rawData, 1)
        If Mode = "Amphoe" Then
            ' กรองเฉพาะอำเภอที่อยู่ในจังหวัดที่เลือก
            If rawData(i, 1) = Prov And rawData(i, 2) <> "" Then
                If Not IsInArray(CStr(rawData(i, 2)), resultData, count) Then
                    count = count + 1: resultData(count, 1) = rawData(i, 2)
                End If
            End If
        ElseIf Mode = "Tambon" Then
            ' กรองเฉพาะตำบลที่อยู่ในจังหวัด และอำเภอที่เลือก
            If rawData(i, 1) = Prov And rawData(i, 2) = Amp And rawData(i, 3) <> "" Then
                If Not IsInArray(CStr(rawData(i, 3)), resultData, count) Then
                    count = count + 1: resultData(count, 1) = rawData(i, 3)
                End If
            End If
        End If
    Next i

    ' STEP 4: เขียนผลลัพธ์ลงในตาราง List เพื่อให้ Data Validation เรียกใช้
    If count > 0 Then
        ws.Cells(2, targetCol).Resize(count, 1).Value = resultData
    End If

CleanUp:
    ws.Protect Password:=myPassword
    Application.ScreenUpdating = True
End Sub