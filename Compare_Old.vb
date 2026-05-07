' ############################################################################################################
' ส่วนที่ 2: ฟังก์ชันสำหรับอัปเดตรายการ Dropdown (UpdateLocationList)
' อ่านจาก: CF_Common (C=จังหวัด, D=อำเภอ, E=ตำบล)
' เขียนลง: CF_อยู่ดีมีสุข (U=LIST_อำเภอ, V=LIST_ตำบล)
' ############################################################################################################

Public Sub UpdateLocationList(ByVal Mode As String, ByVal Prov As String, Optional ByVal Amp As String = "")
    Dim ws As Worksheet
    Dim wsCommon As Worksheet
    Dim rawData As Variant
    Dim resultData() As String
    Dim lastRow As Long, i As Long, count As Long
    Dim targetCol As String
    
    ' --- STEP 1: ตั้งค่าเริ่มต้น และกำหนดเป้าหมาย ---
    Set ws = ThisWorkbook.Worksheets("CF_อยู่ดีมีสุข")
    Set wsCommon = ThisWorkbook.Worksheets("CF_Common")

    ' กำหนดคอลัมน์ปลายทางในชีท CF_อยู่ดีมีสุข: ถ้าหาอำเภอเขียนลง U, ถ้าหาตำบลเขียนลง V
    targetCol = IIf(Mode = "Amphoe", "U", "V")

    ' เตรียมระบบ
    Application.ScreenUpdating = False
    Call SetSheetProtection(ws, False) ' ปลดล็อกชีทปลายทางชั่วคราว

    ' --- STEP 2: ล้างข้อมูลเก่า (Cleanup) ในชีท CF_อยู่ดีมีสุข ---
    ' หาบรรทัดสุดท้ายของคอลัมน์ที่จะเขียน เพื่อลบข้อมูลเดิมออกก่อน
    Dim targetLastRow As Long
    targetLastRow = ws.Cells(ws.Rows.count, targetCol).End(xlUp).Row

    If targetLastRow >= 2 Then
        ws.Range(ws.Cells(2, targetCol), ws.Cells(targetLastRow, targetCol)).ClearContents
    End If

    ' พิเศษ: หากเป็นการเลือกจังหวัดใหม่ (Mode = Amphoe) ให้ล้าง List ตำบล (V) ทิ้งด้วย
    If Mode = "Amphoe" Then
        Dim lastRowV As Long
        lastRowV = ws.Cells(ws.Rows.count, "V").End(xlUp).Row
        If lastRowV >= 2 Then ws.Range("V2:V" & lastRowV).ClearContents
    End If

    ' --- STEP 3: ดึงข้อมูลจากฐานข้อมูล CF_Common เข้า Array ---
    ' หาบรรทัดสุดท้ายของฐานข้อมูลใน CF_Common (คอลัมน์ C คือจังหวัด)
    lastRow = wsCommon.Cells(wsCommon.Rows.count, "C").End(xlUp).Row '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    If lastRow < 2 Then GoTo CleanUp  ' ถ้าไม่มีข้อมูลเลย ให้ข้ามไปขั้นตอนสุดท้าย

    ' ดึงข้อมูล จังหวัด/อำเภอ/ตำบล (C:E) มาเก็บไว้ในตัวแปร Array
    rawData = wsCommon.Range("C2:E" & lastRow).Value '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

    ' เตรียมพื้นที่เก็บผลลัพธ์ (Array) ขนาดสูงสุดเท่ากับจำนวนข้อมูลที่มี
    ReDim resultData(1 To UBound(rawData, 1), 1 To 1)
    count = 0

    ' --- STEP 4: วนลูปคัดกรองข้อมูลตามเงื่อนไข ---
    For i = 1 To UBound(rawData, 1)
        
        ' กรณีที่ 1: หา "อำเภอ" ของจังหวัดที่เลือก
        If Mode = "Amphoe" Then
            If CStr(rawData(i, 1)) = Prov And rawData(i, 2) <> "" Then
                ' ตรวจสอบว่าชื่ออำเภอนี้ถูกเพิ่มไปหรือยัง (ป้องกันชื่อซ้ำ)
                If Not IsInArray(CStr(rawData(i, 2)), resultData, count) Then
                    count = count + 1
                    resultData(count, 1) = rawData(i, 2)
                End If
            End If
            
        ' กรณีที่ 2: หา "ตำบล" ของจังหวัดและอำเภอที่เลือก
        ElseIf Mode = "Tambon" Then
            If CStr(rawData(i, 1)) = Prov And CStr(rawData(i, 2)) = Amp And rawData(i, 3) <> "" Then
                ' ตรวจสอบชื่อซ้ำก่อนเพิ่มลงรายการ
                If Not IsInArray(CStr(rawData(i, 3)), resultData, count) Then
                    count = count + 1
                    resultData(count, 1) = rawData(i, 3)
                End If
            End If
        End If
    Next i

    ' --- STEP 5: เขียนผลลัพธ์ลงในชีท CF_อยู่ดีมีสุข ---
    If count > 0 Then
        ws.Cells(2, targetCol).Resize(count, 1).Value = resultData
    End If

CleanUp:
    Call SetSheetProtection(ws, SheetLockSetting) ' ล็อกชีทคืนหลังทำงานเสร็จ
    Application.ScreenUpdating = True
End Sub