' ======================================================================================
' GetPackageValidation: ฟังก์ชันคำนวณหน้างาน (เรียกใช้ภายใน)
' รับค่า: ทุนประกัน (Double) และ ช่วงเซลล์ตารางอ้างอิง (Range)
' คืนค่า: Array 5 ช่อง [สถานะ, ค่าต่ำกว่า, ค่าสูงกว่า, ต่ำสุด, สูงสุด]
' ======================================================================================
Public Function GetPackageValidation(ByVal totalVal As Double, ByVal tableRef As Range) As Variant
    Dim result(1 To 5) As Variant
    Dim matchIdx As Variant
    
    ' STEP 1: ตรวจสอบความปลอดภัย (หากส่ง Range มาว่างเปล่าให้เด้งออก)
    If tableRef Is Nothing Then
        result(1) = "Error: No Table Reference"
        GetPackageValidation = result
        Exit Function
    End If
    
    ' STEP 2: คำนวณขอบเขตล่าง-บนของตารางจริง (Dynamic Min-Max)
    result(4) = Application.WorksheetFunction.Min(tableRef)
    result(5) = Application.WorksheetFunction.Max(tableRef)

    ' STEP 3: ตรวจสอบว่าทุน "หลุดช่วง" ที่รับประกันหรือไม่
    If totalVal < result(4) Or totalVal > result(5) Then
        result(1) = "OutOfRange"
        GetPackageValidation = result
        Exit Function
    End If

    ' STEP 4: ตรวจสอบแบบตรงตัว (Exact Match)
    matchIdx = Application.Match(totalVal, tableRef, 0)
    
    If Not IsError(matchIdx) Then
        ' พบแผนที่ตรงเป๊ะ
        result(1) = "Valid" 
    Else
        ' ไม่ตรงแผนเป๊ะ -> เริ่มกระบวนการหาแผนใกล้เคียง (Suggestion)
        result(1) = "Invalid" 
        On Error Resume Next
        ' หาแผนที่ทุนน้อยกว่าที่ใกล้ที่สุด
        result(2) = Application.WorksheetFunction.Lookup(totalVal, tableRef)
        ' หาแผนที่ทุนมากกว่าที่ใกล้ที่สุด
        matchIdx = Application.Match(totalVal, tableRef, 1)
        If Not IsError(matchIdx) And matchIdx < tableRef.Rows.Count Then
            result(3) = tableRef.Cells(matchIdx + 1, 1).Value
        Else
            result(3) = result(2) 
        End If
        On Error GoTo 0
    End If
    
    GetPackageValidation = result
End Function