Attribute VB_Name = "XL_2013"
Option Explicit

Sub FileUpgrader(ByVal Path As String)
    Application.ScreenUpdating = False
    
    Dim OldWkBk As Workbook
    
    Workbooks.Open Filename:=Path
    Set OldWkBk = ActiveWorkbook
    OldWkBk.Save
    OldWkBk.Close

End Sub

Sub MarkBadData(ByVal Path As String, ByVal Valid_Date As Boolean, ByVal Valid_Clock_Number As Boolean)
    Application.ScreenUpdating = False
    
    Dim BadWkBk As Workbook
    
    Workbooks.Open Filename:=Path
    Set BadWkBk = ActiveWorkbook
    
    Call Lock_Unlock_WkBk("Unlock")
    
    If Valid_Date = False Then
        Call MarkCell("J2")
    End If
    
    If Valid_Clock_Number = False Then
        Call MarkCell("J3")
    End If
    
    Call Lock_Unlock_WkBk("Lock")
    Range("A1").Select
    
    BadWkBk.Save
    BadWkBk.Close
    
End Sub

Private Sub MarkCell(ByVal CellRange As String)
    Application.ScreenUpdating = False
    
    Range(CellRange).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
End Sub

Private Sub Lock_Unlock_WkBk(ByVal Action)
    Application.ScreenUpdating = False
    
    Const ActivePW As String = "81643"
    If Action = "Unlock" Then
        ActiveSheet.Unprotect Password:=ActivePW
    ElseIf Action = "Lock" Then
        ActiveSheet.Protect Password:=ActivePW, DrawingObjects:=True, Contents:=True, Scenarios:=True
    End If

End Sub

Sub TransferTo_Version_2(ByVal Path As String)
    Application.ScreenUpdating = False
    
    Dim CodeWkBk As Workbook
    Dim OldWkBk As Workbook
    Dim NewWkBk As Workbook
    
    Set CodeWkBk = ActiveWorkbook
    
    Sheets("Weekly Time Log").Select
    Sheets("Weekly Time Log").Copy
    Set NewWkBk = ActiveWorkbook
    
    Workbooks.Open Filename:=Path
    Set OldWkBk = ActiveWorkbook
    
    Dim Source_Range(14) As String
    
    Source_Range(0) = "J2" 'Mondays Date
    Source_Range(1) = "J3" 'Clock Number
    Source_Range(2) = "D10:J16" 'Installations
    Source_Range(3) = "D19:J25" 'Preventative Maintenance Site Visits
    Source_Range(4) = "D28:J34" 'Instrument Repair or Instrument Troubleshooting at a Customer Site
    Source_Range(5) = "D37:J42" 'Remote Hardware Support
    Source_Range(6) = "D45:J50" 'Remote Software Support
    Source_Range(7) = "D53:J56" 'Hardware Repair, Upgrade, or Refurbish (In-House)
    Source_Range(8) = "D59:J61" 'Miscellaneous
    Source_Range(9) = "D65:J72" 'Document Generation
    Source_Range(10) = "D74:J74" 'R&D Support: Total Hours
    Source_Range(11) = "D80:J88" 'Online Training
    Source_Range(12) = "D91:J103" 'Onsite Training
    Source_Range(13) = "D106:J119" 'In-house Training
    Source_Range(14) = "D122:J130" 'Validation Duties
    
    Dim Destination_Range(14) As String
    
    Destination_Range(0) = "J2" 'Mondays Date
    Destination_Range(1) = "J3" 'Clock Number
    Destination_Range(2) = "D10:J16" 'Sep Sci Installations
    Destination_Range(3) = "D28:J34" 'Sep Sci Preventative Maintenance Site Visits
    Destination_Range(4) = "D46:J52" 'Sep Sci Instrument Repair or Instrument Troubleshooting at a Customer Site
    Destination_Range(5) = "D64:J69" 'Sep Sci Remote Hardware Support
    Destination_Range(6) = "D80:J85" 'Sep Sci Remote Software Support
    Destination_Range(7) = "D96:J99" 'Sep Sci Hardware Repair, Upgrade, or Refurbish (In-House)
    Destination_Range(8) = "D108:J110" 'Miscellaneous
    Destination_Range(9) = "D113:J120" 'Sep Sci Document Generation
    Destination_Range(10) = "D133:J133" 'Sep Sci Interdepartmental Support: R&D
    Destination_Range(11) = "D149:J157" 'Sep Sci Online Training
    Destination_Range(12) = "D167:J179" 'Sep Sci Onsite Training
    Destination_Range(13) = "D196:J209" 'Sep Sci In-house Training
    Destination_Range(14) = "D227:J235" 'Validation Duties
    
    Dim Rng As Byte
    
    For Rng = 0 To 14
        OldWkBk.Activate
        Range(Source_Range(Rng)).Copy
        NewWkBk.Activate
        Range(Destination_Range(Rng)).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False
    Next Rng
    
    OldWkBk.Close
    
    NewWkBk.Activate
    Call Lock_Unlock_WkBk("Lock")
    Application.DisplayAlerts = False
    NewWkBk.SaveAs Path
    NewWkBk.Close
    
End Sub

Sub TransferTo_Version_2_From_1(ByVal Path As String)
    Application.ScreenUpdating = False
    
    Dim CodeWkBk As Workbook
    Dim OldWkBk As Workbook
    Dim NewWkBk As Workbook
    
    Set CodeWkBk = ActiveWorkbook
    
    Sheets("Weekly Time Log").Select
    Sheets("Weekly Time Log").Copy
    Set NewWkBk = ActiveWorkbook
    
    Workbooks.Open Filename:=Path
    Set OldWkBk = ActiveWorkbook
    
    Dim Source_Range(13) As String
    
    Source_Range(0) = "J2" 'Mondays Date
    Source_Range(1) = "J3" 'Clock Number
    Source_Range(2) = "D8:J13" 'Installations
    Source_Range(3) = "D15:J21" 'Preventative Maintenance Site Visits
    Source_Range(4) = "D23:J29" 'Instrument Repair or Instrument Troubleshooting at a Customer Site
    Source_Range(5) = "D31:J36" 'Remote Hardware Support
    Source_Range(6) = "D38:J43" 'Remote Software Support
    Source_Range(7) = "D45:J48" 'Hardware Repair, Upgrade, or Refurbish (In-House)
    Source_Range(8) = "D50:J53" 'Miscellaneous
    Source_Range(9) = "D55:J62" 'Document Generation
    'Source_Range(10) = "Dxx:Jxx" : Total Hours
    Source_Range(10) = "D68:J76" 'Online Training
    Source_Range(11) = "D78:J90" 'Onsite Training
    Source_Range(12) = "D92:J105" 'In-house Training
    Source_Range(13) = "D107:J115" 'Validation Duties
    
    Dim Destination_Range(14) As String
    
    Destination_Range(0) = "J2" 'Mondays Date
    Destination_Range(1) = "J3" 'Clock Number
    Destination_Range(2) = "D10:J16" 'Sep Sci Installations
    Destination_Range(3) = "D28:J34" 'Sep Sci Preventative Maintenance Site Visits
    Destination_Range(4) = "D46:J52" 'Sep Sci Instrument Repair or Instrument Troubleshooting at a Customer Site
    Destination_Range(5) = "D64:J69" 'Sep Sci Remote Hardware Support
    Destination_Range(6) = "D80:J85" 'Sep Sci Remote Software Support
    Destination_Range(7) = "D96:J99" 'Sep Sci Hardware Repair, Upgrade, or Refurbish (In-House)
    Destination_Range(8) = "D108:J110" 'Miscellaneous
    Destination_Range(9) = "D113:J120" 'Sep Sci Document Generation
    'Destination_Range(10) = "D133:J133" 'Sep Sci Interdepartmental Support: R&D
    Destination_Range(10) = "D149:J157" 'Sep Sci Online Training
    Destination_Range(11) = "D167:J179" 'Sep Sci Onsite Training
    Destination_Range(12) = "D196:J209" 'Sep Sci In-house Training
    Destination_Range(13) = "D227:J235" 'Validation Duties
    
    Dim Rng As Byte
    
    For Rng = 0 To 13
        OldWkBk.Activate
        Range(Source_Range(Rng)).Copy
        NewWkBk.Activate
        Range(Destination_Range(Rng)).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False
    Next Rng
    
    'Transferring R&D Support to Sep Sci Interdepartmental Support: R&D
    Dim Column As Integer
    Dim Row As Integer
    Dim Sum_Value As Double
    
    For Column = 4 To 10
        OldWkBk.Activate
        Sum_Value = 0
        For Row = 64 To 66
            Sum_Value = Sum_Value + Cells(Row, Column).Value
        Next Row
        NewWkBk.Activate
        Cells(133, Column).Value = Sum_Value
    Next Column
    
    
    OldWkBk.Close
    
    NewWkBk.Activate
    Call Lock_Unlock_WkBk("Lock")
    Application.DisplayAlerts = False
    NewWkBk.SaveAs Path
    NewWkBk.Close
    
End Sub

