' ======================================================================================
' GetPackageValidation: ฟังก์ชันตัวกลางสำหรับดึงข้อมูลและคำนวณหน้างาน
' ทำหน้าที่: ตรวจสอบว่าทุนที่กรอกมา "ผ่านเกณฑ์" หรือ "ควรแนะนำค่าไหน"
' ======================================================================================
Private Function GetPackageValidation(ByVal totalVal As Double) As Variant
    Dim tableRange As Range
    Dim result(1 To 5) As Variant ' เตรียมที่เก็บข้อมูล 5 ช่อง (สถานะ, ค่าต่ำกว่า, ค่าสูงกว่า, ต่ำสุด, สูงสุด)
    Dim matchIdx As Variant
    
    ' --- ขั้นตอนที่ 1: กำหนดตารางอ้างอิง ---
    Set tableRange = ThisWorkbook.Worksheets("CF_อยู่ดีมีสุข").Range("A2:A46")
    
    ' --- ขั้นตอนที่ 2: หาขอบเขต ต่ำสุด-สูงสุด จากตารางจริง ---
    result(4) = Application.WorksheetFunction.Min(tableRange) ' เก็บค่าต่ำสุดไว้ช่อง 4
    result(5) = Application.WorksheetFunction.Max(tableRange) ' เก็บค่าสูงสุดไว้ช่อง 5

    ' --- ขั้นตอนที่ 3: เช็คว่าทุน "น้อยไป" หรือ "มากเกินไป" หรือไม่ ---
    If totalVal < result(4) Or totalVal > result(5) Then
        result(1) = "OutOfRange" ' ระบุสถานะว่า: นอกขอบเขต
        GetPackageValidation = result
        Exit Function ' จบการทำงานทันที
    End If

    ' --- ขั้นตอนที่ 4: ตรวจสอบว่า "ตรงกับแผนเป๊ะๆ" หรือไม่ ---
    matchIdx = Application.Match(totalVal, tableRange, 0)
    
    If Not IsError(matchIdx) Then
        ' ถ้าหาเจอ (ไม่ Error)
        result(1) = "Valid" ' ระบุสถานะว่า: ถูกต้อง
    Else
        ' ถ้าหาไม่เจอ (ไม่ตรงแผนเป๊ะ แต่ยังอยู่ในช่วงที่รับได้)
        result(1) = "Invalid" ' ระบุสถานะว่า: ไม่ตรงแผน (แต่มีค่าแนะนำ)
        
        On Error Resume Next
        ' หาค่าที่ "น้อยกว่าและใกล้ที่สุด" (Floor)
        result(2) = Application.WorksheetFunction.Lookup(totalVal, tableRange)
        ' หาตำแหน่งลำดับของค่าที่ใกล้เคียง
        matchIdx = Application.Match(totalVal, tableRange, 1)
        
        ' หาค่าที่ "มากกว่าและใกล้ที่สุด" (Ceiling)
        If Not IsError(matchIdx) And matchIdx < tableRange.Rows.count Then
            result(3) = tableRange.Cells(matchIdx + 1, 1).Value
        Else
            result(3) = result(2) ' กรณีสูงสุดแล้วให้ใช้ค่าเดิม
        End If
        On Error GoTo 0
    End If
    
    ' ส่งผลลัพธ์ทั้งหมดกลับไปให้คนเรียกใช้
    GetPackageValidation = result
End Function