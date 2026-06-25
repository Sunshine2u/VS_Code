Sub SplitThaiAddressWithProgressAndAbort()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' หากไม่มีข้อมูลให้หยุดทำงาน
    If lastRow < 2 Then Exit Sub
    
    ' 1. ตั้งค่าเริ่มต้นและเปิดฟอร์ม
    With frmProgress
        .lblBar.Width = 0
        .lblPct.Caption = "0%"
        .Tag = "" ' ล้างค่าสัญญาณ Abort เก่าออกก่อน
        .Show vbModeless
    End With
    
    ' สร้าง Header ผลลัพธ์
    ws.Range("B1:L1").Value = Array("รหัสไปรษณีย์", "จังหวัด", "อำเภอ/เขต", "ตำบล/แขวง", "ถนน", "ซอย", "หมู่บ้าน/อาคาร/ตึก/คอนโด", "หมู่", "บ้านเลขที่", "เลขที่โฉนด", "ข้อความที่เหลือ")
    
    Dim i As Long
    Dim tempAddr As String
    Dim regEx As Object, matches As Object
    Dim startPos As Long, endPos As Long
    
    Dim totalRows As Long: totalRows = lastRow - 1
    Dim currentPct As Double
    Dim maxBarWidth As Single: maxBarWidth = frmProgress.FrameProgress.Width
    
    ' ตัวแปรตรวจสอบว่าหลุดลูปเพราะกด Abort หรือไม่
    Dim isAborted As Boolean: isAborted = False
    
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = False
    regEx.IgnoreCase = True
    
    For i = 2 To lastRow
        ' ----------------------------------------------------
        ' ตรวจสอบว่าผู้ใช้กดปุ่ม Abort หรือไม่ก่อนประมวลผลแถวถัดไป
        ' ----------------------------------------------------
        If frmProgress.Tag = "ABORTED" Then
            isAborted = True
            Exit For ' สั่งดีดออกจาก Loop ทันที
        End If
        ' ----------------------------------------------------
        
        tempAddr = Trim(ws.Cells(i, 1).Value)
        
        ' อัปเดตสถานะ Progress Bar
        currentPct = (i - 1) / totalRows
        frmProgress.lblPct.Caption = Format(currentPct, "0%")
        frmProgress.lblBar.Width = currentPct * maxBarWidth
        
        ' DoEvents ทำหน้าที่เปิดช่องให้ Windows รับอีเวนต์การคลิกปุ่ม Abort บนฟอร์มด้วย
        DoEvents 
        
        ' --- เริ่มกระบวนการแยกคำ (Logic เดิมทั้งหมด) ---
        ' 1. ล้าง 
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
        
        ' 2. รหัสไปรษณีย์
        Dim zip As String: zip = ""
        regEx.Pattern = "\b\d{5}\b"
        If regEx.Test(tempAddr) Then
            zip = regEx.Execute(tempAddr)(0).Value
            tempAddr = Trim(regEx.Replace(tempAddr, ""))
        End If
        
        ' 3. จังหวัด
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
        
        ' 4. อำเภอ / เขต (Strict District Matching)
        Dim amphur As String: amphur = ""
        regEx.Pattern = "(อำเภอ|อ\.|เขต)\s*([^\s]+)(?=\s*(กรุงเทพมหานคร|จังหวัด|$))"
        If regEx.Test(tempAddr) Then
            amphur = regEx.Execute(tempAddr)(0).Value
            tempAddr = Trim(regEx.Replace(tempAddr, ""))
        End If
        
        ' 5. ตำบล / แขวง / ตรอก
        Dim tumbon As String: tumbon = ""
        regEx.Pattern = "(ตำบล|ต\.|แขวง|ตรอก)\s*([^\s]+)"
        If regEx.Test(tempAddr) Then
            tumbon = regEx.Execute(tempAddr)(0).Value
            tempAddr = Trim(regEx.Replace(tempAddr, ""))
        End If
        
        ' 6. เลขที่โฉนด
        Dim titleDeed As String: titleDeed = ""
        regEx.Pattern = "(โฉนด|ฉ\.|ฉ)\s*(\d+)"
        If regEx.Test(tempAddr) Then
            titleDeed = regEx.Execute(tempAddr)(0).Value
            tempAddr = Trim(regEx.Replace(tempAddr, ""))
        End If
        
        ' 7. ซอย (Lookahead รูดชนฟิลด์ถัดไป)
        Dim soi As String: soi = ""
        regEx.Pattern = "(ซอย|ซ\.)\s*(.*?)(?=(ถนน|ถ\.|ถ\.ถ\.|ตำบล|ต\.|แขวง|อำเภอ|อ\.|เขต|$))"
        If regEx.Test(tempAddr) Then
            Dim matchVal As String
            matchVal = regEx.Execute(tempAddr)(0).Value
            If Trim(matchVal) <> "ซอย" And Trim(matchVal) <> "ซ." Then
                soi = matchVal
                tempAddr = Trim(regEx.Replace(tempAddr, ""))
            End If
        End If
        
        ' 8. ถนน (Lookahead รูดชนฟิลด์ถัดไป)
        Dim road As String: road = ""
        regEx.Pattern = "(ถนน|ถ\.|ถ\.ถ\.)\s*(.*?)(?=(ตำบล|ต\.|แขวง|อำเภอ|อ\.|เขต|$))"
        If regEx.Test(tempAddr) Then
            Dim matchRoad As String
            matchRoad = regEx.Execute(tempAddr)(0).Value
            If Trim(matchRoad) <> "ถนน" And Trim(matchRoad) <> "ถ." And Trim(matchRoad) <> "ถ.ถ." Then
                road = matchRoad
                tempAddr = Trim(regEx.Replace(tempAddr, ""))
            End If
        End If
        
        ' 9. หมู่บ้าน / อาคาร / ตึก / คอนโด
        Dim propertyGroup As String: propertyGroup = ""
        regEx.Pattern = "(หมู่บ้าน|มบ\.)\s*([^\s]+(\s+[^\s]+)?)|(อาคาร|ตึก|คอนโด)\s*([^\s]+(\s+[^\s]+)?)"
        If regEx.Test(tempAddr) Then
            propertyGroup = regEx.Execute(tempAddr)(0).Value
            tempAddr = Trim(regEx.Replace(tempAddr, ""))
        End If
        
        ' 10. หมู่
        Dim moo As String: moo = ""
        regEx.Pattern = "(หมู่ที่|หมู่|ม\.)\s*\d+"
        If regEx.Test(tempAddr) Then
            moo = regEx.Execute(tempAddr)(0).Value
            tempAddr = Trim(regEx.Replace(tempAddr, ""))
        End If
        
        ' 11. บ้านเลขที่
        Dim houseNo As String: houseNo = ""
        regEx.Pattern = "^(เลขที่\s*)?(\d+[\/\d\-\.\(\)\s,]+)"
        If regEx.Test(tempAddr) Then
            houseNo = regEx.Execute(tempAddr)(0).Value
            tempAddr = Trim(regEx.Replace(tempAddr, ""))
            houseNo = Trim(Replace(houseNo, "เลขที่", ""))
        End If
        
        ' ส่งข้อมูลลงตาราง
        ws.Cells(i, 2).Value = zip
        ws.Cells(i, 3).Value = province
        ws.Cells(i, 4).Value = amphur
        ws.Cells(i, 5).Value = tumbon
        ws.Cells(i, 6).Value = road           
        ws.Cells(i, 7).Value = soi            
        ws.Cells(i, 8).Value = propertyGroup  
        ws.Cells(i, 9).Value = moo            
        ws.Cells(i, 10).Value = houseNo       
        ws.Cells(i, 11).Value = titleDeed     
        ws.Cells(i, 12).Value = Trim(tempAddr) 
    Next i
    
    ' ปิดฟอร์มเมื่อจบกระบวนการ (ไม่ว่าจะรันเสร็จหรือกด Abort)
    Unload frmProgress
    ws.Columns("B:L").AutoFit
    
    ' แจ้งเตือนผู้ใช้แยกตามสถานะที่จบการทำงาน
    If isAborted Then
        MsgBox "หยุดการทำงานกลางคันแล้ว ข้อมูลได้รับการบันทึกถึงแถวที่ " & (i - 1), vbExclamation, "Aborted"
    Else
        MsgBox "แยกข้อมูลที่อยู่เสร็จสมบูรณ์ 100% แล้วครับ!", vbInformation, "Success"
    End If
End Sub