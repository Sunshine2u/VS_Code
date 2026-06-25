Sub SplitThaiAddressFinalWhitespaceFix()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' สร้าง Header ผลลัพธ์ตามโครงสร้างที่ต้องการ
    ws.Range("B1:L1").Value = Array("รหัสไปรษณีย์", "จังหวัด", "อำเภอ/เขต", "ตำบล/แขวง", "ถนน", "ซอย", "หมู่บ้าน/อาคาร/ตึก/คอนโด", "หมู่", "บ้านเลขที่", "เลขที่โฉนด", "ข้อความที่เหลือ")
    
    Dim i As Long
    Dim tempAddr As String
    Dim regEx As Object, matches As Object
    Dim startPos As Long, endPos As Long
    
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = False
    regEx.IgnoreCase = True
    
    For i = 2 To lastRow
        tempAddr = Trim(ws.Cells(i, 1).Value)
        
        ' ----------------------------------------------------
        ' 1. ล้าง โดยใช้ฟังก์ชัน String ป้องกัน Error 5017
        Do While InStr(tempAddr, "[") > 0 And InStr(tempAddr, "]") > 0
            startPos = InStr(tempAddr, "[")
            endPos = InStr(tempAddr, "]")
            If endPos > startPos Then
                tempAddr = Left(tempAddr, startPos - 1) & Mid(tempAddr, endPos + 1)
            Else
                Exit Do
            End If
        Loop
        tempAddr = Trim(tempAddr)
        ' ----------------------------------------------------
        
        ' 2. แยกดึงรหัสไปรษณีย์ (ตัวเลข 5 หลัก)
        Dim zip As String: zip = ""
        regEx.Pattern = "\b\d{5}\b"
        If regEx.Test(tempAddr) Then
            zip = regEx.Execute(tempAddr)(0).Value
            tempAddr = Trim(regEx.Replace(tempAddr, ""))
        End If
        
        ' 3. แยกดึงจังหวัด (รองรับ กรุงเทพ, กทม., กรุงเทพฯ, กรุงเทพมหานคร และ จ.ต่างๆ)
        Dim province As String: province = ""
        regEx.Pattern = "(จังหวัด|จ\.)\s*([^\s]+)|(กรุงเทพมหานคร|กรุงเทพฯ|กรุงเทพ|กทม\.)"
        If regEx.Test(tempAddr) Then
            Set matches = regEx.Execute(tempAddr)
            If matches(0).SubMatches.Count > 1 Then
                province = IIf(matches(0).SubMatches(1) <> "", matches(0).SubMatches(1), matches(0).SubMatches(2))
            Else
                province = matches(0).Value
            End If
            province = Replace(province, "จังหวัด", "")
            province = Replace(province, "จ.", "")
            province = Trim(province)
            If province = "กรุงเทพ" Or province = "กทม." Or province = "กรุงเทพฯ" Then
                province = "กรุงเทพมหานคร"
            End If
            tempAddr = Trim(regEx.Replace(tempAddr, ""))
        End If
        
        ' 4. แยกดึงอำเภอ / เขต
        Dim amphur As String: amphur = ""
        regEx.Pattern = "(อำเภอ|อ\.|เขต)\s*([^\s]+)"
        If regEx.Test(tempAddr) Then
            amphur = regEx.Execute(tempAddr)(0).Value
            tempAddr = Trim(regEx.Replace(tempAddr, ""))
        End If
        
        ' 5. แยกดึงตำบล / แขวง / ตรอก
        Dim tumbon As String: tumbon = ""
        regEx.Pattern = "(ตำบล|ต\.|แขวง|ตรอก)\s*([^\s]+)"
        If regEx.Test(tempAddr) Then
            tumbon = regEx.Execute(tempAddr)(0).Value
            tempAddr = Trim(regEx.Replace(tempAddr, ""))
        End If
        
        ' 6. แยกดึงเลขที่โฉนด (โฉนด, ฉ., ฉ นำหน้าตัวเลข)
        Dim titleDeed As String: titleDeed = ""
        regEx.Pattern = "(โฉนด|ฉ\.|ฉ)\s*(\d+)"
        If regEx.Test(tempAddr) Then
            titleDeed = regEx.Execute(tempAddr)(0).Value
            tempAddr = Trim(regEx.Replace(tempAddr, ""))
        End If
        
        ' 7. แยกดึงซอย (รองรับการเว้นวรรคคั่นระหว่างชื่อและเลขซอย เช่น ซ.ลาดพร้าว  101)
        Dim soi As String: soi = ""
        regEx.Pattern = "(ซอย|ซ\.)\s*([^\s\d]+)?\s*([\d\-\/]+)?(\s*\(.*?\))?(\s+แยก\s*[\d\-\/]+)?"
        If regEx.Test(tempAddr) Then
            Dim matchVal As String
            matchVal = regEx.Execute(tempAddr)(0).Value
            If Trim(matchVal) <> "ซอย" And Trim(matchVal) <> "ซ." Then
                soi = matchVal
                tempAddr = Trim(regEx.Replace(tempAddr, ""))
            End If
        End If
        
        ' 8. [จุดที่แก้ไขใหม่] แยกดึงถนน (รองรับการข้ามเว้นวรรคเพื่อไปเก็บตัวเลขสาย/ชื่อถนนด้านหลัง)
        ' แพทเทิร์นจับกลุ่มคำว่า ถนน/ถ./ถ.ถ. -> ชื่อถนน -> ข้ามเว้นวรรคกี่ช่องก็ได้ -> ตัวเลขหรือคำระบุสายต่อท้าย เช่น "พุทธมณฑลสาย   1"
        Dim road As String: road = ""
        regEx.Pattern = "(ถนน|ถ\.|ถ\.ถ\.)\s*([^\s\d]+)?\s*([\d\-\/]+|[^\s]+)?"
        If regEx.Test(tempAddr) Then
            Dim matchRoad As String
            matchRoad = regEx.Execute(tempAddr)(0).Value
            If Trim(matchRoad) <> "ถนน" And Trim(matchRoad) <> "ถ." And Trim(matchRoad) <> "ถ.ถ." Then
                road = matchRoad
                tempAddr = Trim(regEx.Replace(tempAddr, ""))
            End If
        End If
        
        ' 9. แยกดึง หมู่บ้าน / อาคาร / ตึก / คอนโด 
        Dim propertyGroup As String: propertyGroup = ""
        regEx.Pattern = "(หมู่บ้าน|มบ\.)\s*([^\s]+(\s+[^\s]+)?)|(อาคาร|ตึก|คอนโด)\s*([^\s]+(\s+[^\s]+)?)"
        If regEx.Test(tempAddr) Then
            propertyGroup = regEx.Execute(tempAddr)(0).Value
            tempAddr = Trim(regEx.Replace(tempAddr, ""))
        End If
        
        ' 10. แยกดึง "หมู่" ออกมาต่างหาก (หมู่ 3, ม.5, หมู่ที่ 12)
        Dim moo As String: moo = ""
        regEx.Pattern = "(หมู่ที่|หมู่|ม\.)\s*\d+"
        If regEx.Test(tempAddr) Then
            moo = regEx.Execute(tempAddr)(0).Value
            tempAddr = Trim(regEx.Replace(tempAddr, ""))
        End If
        
        ' 11. แยกดึงบ้านเลขที่ (ส่วนที่อยู่ด้านหน้าสุด)
        Dim houseNo As String: houseNo = ""
        regEx.Pattern = "^(เลขที่\s*)?(\d+[\/\d\-\.\(\)\s,]+)"
        If regEx.Test(tempAddr) Then
            houseNo = regEx.Execute(tempAddr)(0).Value
            tempAddr = Trim(regEx.Replace(tempAddr, ""))
            houseNo = Trim(Replace(houseNo, "เลขที่", ""))
        End If
        
        ' บันทึกข้อมูลพิมพ์ลงใน Worksheet (คอลัมน์ B - L)
        ws.Cells(i, 2).Value = zip
        ws.Cells(i, 3).Value = province
        ws.Cells(i, 4).Value = amphur
        ws.Cells(i, 5).Value = tumbon
        ws.Cells(i, 6).Value = road           ' คอลัมน์ถนน (แก้ไขให้เก็บเลขที่เว้นวรรคคั่นแล้ว)
        ws.Cells(i, 7).Value = soi            
        ws.Cells(i, 8).Value = propertyGroup  
        ws.Cells(i, 9).Value = moo            
        ws.Cells(i, 10).Value = houseNo       
        ws.Cells(i, 11).Value = titleDeed     
        ws.Cells(i, 12).Value = Trim(tempAddr) 
    Next i
    
    ws.Columns("B:L").AutoFit
    MsgBox "ปรับปรุงระบบดึงเลขถนนที่มีการเว้นวรรคคั่น เรียบร้อยแล้วครับ!", vbInformation
End Sub