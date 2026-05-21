' ############################################################################################################
' Sub ที่ดักจับการเปลี่ยนแปลงข้อมูลในเซลล์ที่เกี่ยวข้องกับจังหวัด/อำเภอ และทุนประกันภัย แล้วเปลี่ยนแปลงข้อมูลใน Worksheet ตามเงื่อนไขที่กำหนดไว้
' Worksheet_Change Event สำหรับหน้าจอออกใบเสนอราคา (Quotation)
' โดยมีการตรวจสอบและจัดการข้อมูลที่เกี่ยวข้องกับจังหวัด/อำเภอ และทุนประกันภัย
' โดยมีการแจ้งเตือนและแนะนำข้อมูลที่จำเป็นเพื่อให้การกรอกข้อมูลเป็นไปอย่างถูกต้องและครบถ้วน
' #############################################################################################################

Private Sub Worksheet_Change(ByVal Target As Range)
    
    ' =========================================================================
    ' 📍 [ส่วนแก้ไขพิกัด] ย้ายหน้าจอใหม่ ให้มาแก้ไขตำแหน่งเซลล์ตรงนี้ที่เดียวเท่านั้น! 
    ' กรอกพิกัดเซล์ให้ครอบคลุม Merge Cell ด้วยนะครับ เช่น ถ้าเซลล์ที่ต้องการดักจับเป็น A1 แต่ A1:C1 มีการ Merge Cell อยู่ ก็ให้กรอกพิกัดเป็น A1:C1 แบบนี้ครับ
    ' =========================================================================
    Dim Input_Prov1 As Range: Set Input_Prov1 = Me.Range("H28")     ' จังหวัด ชุดที่ 1
    Dim Input_Amp1 As Range:  Set Input_Amp1 = Me.Range("J28")      ' อำเภอ ชุดที่ 1
    Dim Target_Clear1 As Range: Set Target_Clear1 = Me.Range("J28,L28") ' เซลล์ที่ต้องล้างค่าเมื่อเปลี่ยนจังหวัด 1
    
    Dim Input_Prov2 As Range: Set Input_Prov2 = Me.Range("H51")     ' จังหวัด ชุดที่ 2
    Dim Input_Amp2 As Range:  Set Input_Amp2 = Me.Range("J51")      ' อำเภอ ชุดที่ 2
    Dim Target_Clear2 As Range: Set Target_Clear2 = Me.Range("J51,L51") ' เซลล์ที่ต้องล้างค่าเมื่อเปลี่ยนจังหวัด 2
    
    Dim Input_Prov3 As Range: Set Input_Prov3 = Me.Range("H59")     ' จังหวัด ชุดที่ 3
    Dim Input_Amp3 As Range:  Set Input_Amp3 = Me.Range("J59")      ' อำเภอ ชุดที่ 3
    Dim Target_Clear3 As Range: Set Target_Clear3 = Me.Range("J59,L59") ' เซลล์ที่ต้องล้างค่าเมื่อเปลี่ยนจังหวัด 3
    
    Dim Input_BuildingPrice As Range:   Set Input_BuildingPrice = Me.Range("G41")       ' ทุนอาคาร
    Dim Input_FurniturePrice As Range:   Set Input_FurniturePrice = Me.Range("G42")       ' ทุนเฟอร์นิเจอร์
    Dim Input_Notification As Range:   Set Input_Notification = Me.Range("J42")       ' ช่องแจ้งเตือนเลข 0
    Dim Input_InsuranceSum As Range:   Set Input_InsuranceSum = Me.Range("G43")       ' ช่องรวมทุนประกัน
    
    Dim Noti_Addr1 As Range: Set Noti_Addr1 = Me.Range("G26")     ' ช่องแนะนำการเขียนที่อยู่ 1
    Dim Noti_Addr2 As Range: Set Noti_Addr2 = Me.Range("G49")     ' ช่องแนะนำการเขียนที่อยู่ 2
    Dim Noti_Addr3 As Range: Set Noti_Addr3 = Me.Range("G57")     ' ช่องแนะนำการเขียนที่อยู่ 3
    
    Dim InsureName As Range: Set InsureName = Me.Range("G24")     ' ช่องชื่อผู้เอาประกันภัย (ใช้ตรวจสอบว่ามีการกรอกข้อมูลในส่วนนี้หรือยัง เพื่อเป็นเงื่อนไขในการแสดงผลแจ้งเตือนที่อยู่)
    Dim Input_BuildType As Range: Set Input_BuildType = Me.Range("G31") ' ประเภทสิ่งปลูกสร้าง
    Dim Noti_Floor As Range:     Set Noti_Floor = Me.Range("H36")     ' ช่องจำนวนชั้น
    Dim SamePostOption as Range: Set SamePostOption = Me.Range("F55") ' ตัวเลือกที่อยู่เดียวกัน (Yes/No)
    
    ' รวมเซลล์ทั้งหมดที่ต้องการให้เกิด Event (ดักจับการเปลี่ยนแปลง)
    ' *หมายเหตุ: ตรง Me.Range ข้างล่างนี้ ระบบจะอ้างอิงจากตัวแปรด้านบนให้อัตโนมัติ ไม่ต้องตามแก้แล้วครับ*
    If Intersect(Target, Union(Input_Prov1, Input_Amp1, _
                               Input_Prov2, Input_Amp2, _
                               Input_Prov3, Input_Amp3, _
                               Input_BuildingPrice, Input_FurniturePrice, Noti_Addr1, Noti_Addr2, Noti_Addr3, _
                               Input_BuildType, SamePostOption, InsureName)) Is Nothing Then Exit Sub
    ' =========================================================================

    On Error GoTo ErrorHandler
    
    ' เตรียมระบบก่อนเริ่มทำงาน
    Application.EnableEvents = False
    Call SetSheetProtection(Me, False)

    ' -----------------------------------------------------------------
    ' 🔥 ส่งค่าเข้า Sub ย่อย (ห้ามแก้โค้ดส่วนนี้ ระบบจะดึงค่าจากพิกัดด้านบนมาทำงานเอง)
    ' -----------------------------------------------------------------
    Call อยู่ดีมีสุข_Hide_Address_Rows() ' เปิดแถวที่อยู่ก่อนกรอกข้อมูล เพื่อความสวยงามของหน้าจอขณะทำงาน
    
    ' 1. จัดการเรื่องสถานที่และพื้นที่เสี่ยงภัย (ชุดที่ 1, 2, 3)
    Call HandleLocationChange(Target, Input_Prov1, Input_Amp1, Target_Clear1, 1)
    Call HandleLocationChange(Target, Input_Prov2, Input_Amp2, Target_Clear2, 2)
    Call HandleLocationChange(Target, Input_Prov3, Input_Amp3, Target_Clear3, 3)

    ' 2. ตรวจสอบและคำนวณทุนประกัน
    Call HandleInsuranceSum(Target, Input_BuildingPrice, Input_FurniturePrice, Input_Notification, Input_InsuranceSum)

    ' 3. ตรวจสอบและแนะนำการเขียนที่อยู่ (Placeholder)
    ' ถ้าระบบตรวจพบว่าช่องที่อยู่ช่องใดช่องหนึ่ง หรือจังหวัด/อำเภอเปลี่ยน จะทำการตรวจสอบความสะอาดของข้อมูลที่อยู่ทันที
    Call ApplyAddressPlaceholder(InsureName) ' ถ้าเริ่มกรอก InsureName แล้ว ถึงจะเริ่มตรวจสอบที่อยู่และแสดง Placeholder แนะนำการกรอกที่อยู่
    Call ApplyAddressPlaceholder(Noti_Addr1)
    Call ApplyAddressPlaceholder(Noti_Addr2)
    Call ApplyAddressPlaceholder(Noti_Addr3)

    ' 4. เติมจำนวนชั้นอัตโนมัติ ตามลักษณะสิ่งปลูกสร้าง
    Call AutoFillBuildingFloor(Target, Input_BuildType, Noti_Floor)

    ' -----------------------------------------------------------------
    ' ล็อกชีทคืนและเปิดระบบ Event
    Call SetSheetProtection(Me, SheetLockSetting)
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    MsgBox "เกิดข้อผิดพลาด: " & Err.Description, vbCritical, "Error"
    Call SetSheetProtection(Me, SheetLockSetting)
    Application.EnableEvents = True
End Sub


' ===================================================================
' Sub ย่อยสำหรับจัดการสถานที่ และอัปเดตข้อมูลจังหวัด/อำเภอ
' ===================================================================
Private Sub HandleLocationChange(ByVal Target As Range, ByVal ProvCell As Range, ByVal AmpCell As Range, ByVal ClearRange As Range, ByVal GroupNumber As Integer)
    Dim provName As String: provName = Trim$(CStr(ProvCell.Value))
    Dim ampName As String: ampName = Trim$(CStr(AmpCell.Value))
    Dim i As Long

    ' กรณีเปลี่ยนจังหวัด
    If Not Intersect(Target, ProvCell) Is Nothing Then
        ' ตรวจสอบพื้นที่เสี่ยงภัยน้ำท่วม (เฉพาะกลุ่ม 1 ตามโค้ดเดิม)
        If GroupNumber = 1 Then
            Dim riskList As Variant: riskList = GetListRange(Sheet6, 1, "จังหวัดยกเว้นน้ำท่วม1")
            If Not IsError(Application.Match(provName, riskList, 0)) Then
                MsgBox "พบว่าจังหวัด " & provName & " เป็นพื้นที่เสี่ยงภัยน้ำท่วม" & vbCrLf & _
                       "โปรดติดต่อเจ้าหน้าที่ MTI ผู้ดูแลตัวแทน ในการออกใบเสนอราคา", vbExclamation, "แจ้งเตือนความเสี่ยง"
            End If
        End If

        ' ล้างค่าเก่าที่หน้าจอ
        For i = 1 To ClearRange.Areas.Count ' Loop ตามจำนวนเซลล์ใน ClearRange
        ClearRange.Areas(i).MergeArea.ClearContents ' ล้างค่าแบบนี้จะเคลียร์ได้หมดจด ไม่ว่าจะเป็น Merge Cell หรือไม่ก็ตาม
        Next i  ' จบ Loop ด้วย Next i (ลบ End For ออกไปแล้ว)

        ' อัปเดตรายชื่ออำเภอเข้าคลังข้อมูลตามกลุ่ม
        Select Case GroupNumber
            Case 1: Call UpdateLocationList1("Amphoe", provName)
            Case 2: Call UpdateLocationList2("Amphoe", provName)
            Case 3: Call UpdateLocationList3("Amphoe", provName)
        End Select
    End If

    ' กรณีเปลี่ยนอำเภอ
    If Not Intersect(Target, AmpCell) Is Nothing Then
        If provName <> "" And ampName <> "" Then
            ' ล้างค่าตำบลที่หน้าจอ (ดึงเอาคอลัมน์สุดท้ายของ ClearRange มาล้างค่าตำบล)
            ClearRange.Areas(2).MergeArea.ClearContents
            
            Select Case GroupNumber
                Case 1: Call UpdateLocationList1("Tambon", provName, ampName)
                Case 2: Call UpdateLocationList2("Tambon", provName, ampName)
                Case 3: Call UpdateLocationList3("Tambon", provName, ampName)
            End Select
        End If
    End If
End Sub

' ===================================================================
' Sub ย่อยสำหรับคำนวณและตรวจสอบทุนประกัน
' ===================================================================
Private Sub HandleInsuranceSum(ByVal Target As Range, ByVal BuildingPrice_Cell As Range, ByVal FurniturePrice_Cell As Range, ByVal Notification_Cell As Range, ByVal InsuranceSum_Cell As Range)
    Dim i As Long

    If Intersect(Target, Union(BuildingPrice_Cell, FurniturePrice_Cell)) Is Nothing Then Exit Sub

    ' ตรวจสอบการกรอกช่องทุนเฟอร์นิเจอร์
    If Len(Trim(FurniturePrice_Cell.Text)) = 0 Then
        Notification_Cell.Value = "ถ้าไม่มีให้กรอกเลข 0"
        InsuranceSum_Cell.Font.Color = RGB(0, 0, 0)
    Else
        InsuranceSum_Cell.MergeArea.ClearContents ' ล้างค่าแบบนี้จะเคลียร์ได้หมดจด ไม่ว่าจะเป็น Merge Cell หรือไม่ก็ตาม
        Notification_Cell.MergeArea.ClearContents ' 
    End If
    
    ' คำนวณผลรวมทุนประกัน
    If IsNumeric(BuildingPrice_Cell.Value) And IsNumeric(FurniturePrice_Cell.Value) Then
        InsuranceSum_Cell.Value = BuildingPrice_Cell.Value + FurniturePrice_Cell.Value
        
        With Worksheets("CF_อยู่ดีมีสุข")'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            .Range("P12").Value = BuildingPrice_Cell.Value'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            .Range("P13").Value = FurniturePrice_Cell.Value'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        End With
        Call CheckAndSuggestPremium(InsuranceSum_Cell.Value)
    Else
        InsuranceSum_Cell.MergeArea.ClearContents
    End If
End Sub

' ===================================================================
' Sub ย่อยสำหรับคุมข้อความแนะนำ (Placeholder) ในช่องที่อยู่
' ===================================================================
Private Sub ApplyAddressPlaceholder(ByVal AddrCell As Range)
    Const PLACEHOLDER_TEXT As String = "     บ้านเลขที่.....หมู่ที่....อาคาร/หมู่บ้าน..... ซอย.... ถนน...."
    
    If Len(Trim(AddrCell.Text)) = 0 Then
        AddrCell.Value = PLACEHOLDER_TEXT
        AddrCell.Font.Color = RGB(166, 166, 166)
    Else
        If AddrCell.Value = PLACEHOLDER_TEXT Then
            AddrCell.Font.Color = RGB(166, 166, 166)
        Else
            AddrCell.Font.Color = RGB(0, 0, 0)
        End If
    End If
End Sub


' ===================================================================
' Sub ย่อยสำหรับใส่จำนวนชั้นสิ่งปลูกสร้างอัตโนมัติ (เวอร์ชันแก้ Bug Merge/Clear)
' ===================================================================
Private Sub AutoFillBuildingFloor(ByVal Target As Range, ByVal TypeCell As Range, ByVal FloorCell As Range)
    If Intersect(Target, TypeCell) Is Nothing Then Exit Sub
    
    Dim BuildingType As String
    
    ' แก้ตรงนี้: ใช้ .Cells(1, 1) เพื่อป้องกัน Type mismatch เวลาเจอกลุ่มเซลล์ หรือ Merged Cells
    BuildingType = Trim$(CStr(TypeCell.Cells(1, 1).Value))
    
    Select Case BuildingType
        Case "บ้านเดี่ยว - ตึก 1 ชั้น", _
             "ทาวส์เฮ้าส์,ทาวน์โฮม - 1 ชั้น", _
             "ตึกแถว,อาคารพาณิชย์ - 1 ชั้น", _
             "คอนโดมิเนียม"
             
            ' บังคับใส่ค่าที่เซลล์แรกป้องกัน Error จากการ Merge
            FloorCell.Cells(1, 1).Value = "1"
            
        Case Else
            FloorCell.MergeArea.ClearContents ' ล้างค่าแบบนี้จะเคลียร์ได้หมดจด ไม่ว่าจะเป็น Merge Cell หรือไม่ก็ตาม
    End Select
End Sub