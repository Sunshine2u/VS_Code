Private Sub Worksheet_Change(ByVal Target As Range)

    ' 1. ตรวจสอบเบื้องต้น
    If Intersect(Target, Me.Range("H28,G41:H41,G42:H42")) Is Nothing Then Exit Sub

    On Error GoTo ErrorHandler
    
    ' ปิดการแจ้งเตือนชั่วคราวเพื่อป้องกันการรันวนซ้ำ
    Application.EnableEvents = False
    
    ' ปลดล็อกชีทก่อนเริ่มทำงาน
    Call SetSheetProtection(Me, False)

    
    ' ---------- (A) ส่วนแจ้งเตือนจังหวัดที่ H28 ----------
    If Not Intersect(Target, Me.Range("H28")) Is Nothing Then
        Dim list As Variant
        Dim result As Variant
        Dim Prov As String
        
        Prov = Trim$(CStr(Me.Range("H28").Value))
        list = GetListRange("CF_Common", 1, "จังหวัดยกเว้นน้ำท่วม1")
        
        ' ใช้ Match ค้นหาชื่อจังหวัดใน Array ได้เลยไม่ต้องวน Loop
        result = Application.Match(Prov, list, 0)
        
        If Not IsError(result) Then
            MsgBox "พบว่าจังหวัด " & Prov & " เป็นพื้นที่เสี่ยงภัยน้ำท่วม" & vbCrLf & _
                           "โปรดติดต่อเจ้าหน้าที่ MTI ผู้ดูแลตัวแทน ในการออกใบเสนอราคา", _
                           vbExclamation, "แจ้งเตือนความเสี่ยง"
        End If
        Call UpdateLocationList("Amphoe", Prov) ' อัปเดตอำเภอที่เกี่ยวข้องกับจังหวัดนี้
        
    End If

    IF Not intersect(Target, Me.Range("J28")) Is Nothing Then
        Dim Ampor As String
        Ampor = Trim$(CStr(Me.Range("J28").Value))
            call UpdateLocationList("Tambon", Prov, Ampor) ' ล้างตำบลทิ้งด้วยเพราะอำเภอเปลี่ยนแล้ว
            MsgBox "Run ตำบลแล้ว"
        End if 
    
    ' ล็อกชีทคืน
    Call SetSheetProtection(Me, FileLockSetting) ' ใช้ค่าจาก Const ที่ตั้งไว้ใน 2_Product1_Sub.vb
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    MsgBox "เกิดข้อผิดพลาด: " & Err.Description, vbCritical, "Error"
    Call SetSheetProtection(Me, FileLockSetting)
    Application.EnableEvents = True
End Sub



Public Sub UpdateLocationList(ByVal Mode As String, ByVal Prov As String, Optional ByVal Amp As String = "")
    Dim ws As Worksheet
    Dim rawData As Variant
    Dim resultData() As String
    Dim lastRow As Long, i As Long, count As Long
    Dim targetCol As String
    
' --- STEP 1: ตั้งค่าเริ่มต้น และกำหนดเป้าหมาย ---
Set ws = ThisWorkbook.Worksheets("CF_อยู่ดีมีสุข")

' กำหนดคอลัมน์ปลายทาง: ถ้าหาอำเภอเขียนลง Z, ถ้าหาตำบลเขียนลง AA
targetCol = IIf(Mode = "Amphoe", "Z", "AA")

' ปิดการประมวลผลหน้าจอ (ทำให้โค้ดวิ่งเร็วขึ้น) และปลดล็อคชีท
Application.ScreenUpdating = False
ws.Unprotect Password:=myPassword

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
    ws.Protect Password:=myPassword
    Application.ScreenUpdating = True
End Sub