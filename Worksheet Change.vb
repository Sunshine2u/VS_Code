Option Explicit
Private Const myPassword As String = "QTMTI"
Private Const MIN_PREM As Long = 500000
Private Const MAX_PREM As Long = 10000000

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrHandler
    If Target Is Nothing Then Exit Sub
    If Target.CountLarge > 1 Then Exit Sub

    ' ตรวจสอบช่วงที่สนใจ
    If Intersect(Target, Me.Range("H28,G41:H41,G42:H42")) Is Nothing Then Exit Sub

    Application.EnableEvents = False
    Me.Unprotect Password:=myPassword

    ' ถ้าแก้ G41 หรือ G42 ให้คำนวณ G43 และตรวจพรีเมียม
    If Not Intersect(Target, Me.Range("G41,G42")) Is Nothing Then
        If Len(Trim(Me.Range("G42").Text)) = 0 Then
            Me.Range("J42").Value = "กรุณากรอกข้อมูล"
            Me.Range("J42").Font.Color = RGB(255, 0, 0)
        Else
            Me.Range("J42").ClearContents
        End If

        If IsNumeric(Me.Range("G41").Value) And IsNumeric(Me.Range("G42").Value) Then
            Me.Range("G43").Value = CLng(Me.Range("G41").Value) + CLng(Me.Range("G42").Value)
            Call CheckAndSuggestPremium(Me.Range("G43").Value)
        Else
            Me.Range("G43").ClearContents
        End If
    End If

    ' ถ้าแก้ H28 (provider) ให้ตรวจสอบค่า
    If Not Intersect(Target, Me.Range("H28")) Is Nothing Then
        Dim prov As String
        prov = Trim$(CStr(Me.Range("H28").Value))
        Dim provList As Variant
        ' ปรับรายการ provider ให้ตรงรายการจริง หรือใช้ Named Range/List ในชีท
        provList = Array("ProviderA", "ProviderB", "ProviderC")

        If IsError(Application.Match(prov, provList, 0)) Then
            MsgBox "ค่าผู้ให้บริการไม่ถูกต้อง: " & prov, vbExclamation, "ตรวจสอบผู้ให้บริการ"
            ' ถ้าต้องการ เคลียร์หรือทำอย่างอื่นที่นี่
        End If
    End If

Cleanup:
    Me.Protect Password:=myPassword
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    MsgBox "เกิดข้อผิดพลาด: " & Err.Number & " - " & Err.Description, vbCritical, "Error"
    Resume Cleanup
End Sub

Private Sub CheckAndSuggestPremium(ByVal totalVal As Variant)
    On Error GoTo ErrHandler
    If Not IsNumeric(totalVal) Then Exit Sub
    totalVal = CLng(totalVal)

    ' ตรวจช่วงพรีเมียม
    If totalVal < MIN_PREM Or totalVal > MAX_PREM Then
        Me.Range("J43").Value = "จำนวนต้องอยู่ระหว่าง " & Format(MIN_PREM, "#,##0") & " - " & Format(MAX_PREM, "#,##0")
        Exit Sub
    End If

    ' หา Premium Table (ปรับชื่อชีท/ช่วงให้ตรงของท่าน)
    Dim tableRange As Range
    On Error Resume Next
    Set tableRange = ThisWorkbook.Worksheets("CF_PremiumTable").Range("A2:A46")
    On Error GoTo ErrHandler

    If tableRange Is Nothing Then
        Me.Range("J43").Value = "ไม่พบตารางพรีเมียม (ตรวจชื่อชีท/ช่วง)"
        Exit Sub
    End If

    Dim m As Variant
    m = Application.Match(totalVal, tableRange, 0)

    If Not IsError(m) Then
        ' พบค่าเท่ากันในตาราง
        Me.Range("J43").Value = "ตรงกับตาราง: " & tableRange.Cells(m, 1).Value
        Exit Sub
    Else
        ' หา floor/ceiling โดยใช้ Match(..., 1)
        Dim idx As Variant, floorVal As Variant, ceilVal As Variant
        idx = Application.Match(totalVal, tableRange, 1) ' index ของ floor (ตำแหน่ง)
        If Not IsError(idx) Then
            floorVal = tableRange.Cells(idx, 1).Value
            If idx < tableRange.Rows.Count Then
                ceilVal = tableRange.Cells(idx + 1, 1).Value
                Me.Range("J43").Value = "แนะนำระหว่าง " & floorVal & " (floor) และ " & ceilVal & " (ceiling)"
            Else
                Me.Range("J43").Value = "ค่าอยู่เหนือช่วงตาราง: แนะนำใช้ค่า " & floorVal
            End If
        Else
            Me.Range("J43").Value = "ไม่พบค่าที่ใกล้เคียงในตาราง"
        End If
    End If

Cleanup:
    Exit Sub

ErrHandler:
    Me.Range("J43").Value = "Error: " & Err.Number
    Resume Cleanup
End Sub

' Optional: ช่วยเปิด events และการคำนวณกลับมา (ใช้เมื่อ debug)
Public Sub ResetExcelEvents()
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "เปิด Events และ Calculation แล้ว", vbInformation
End Sub
