Sub AttendanceReport()

Dim Day As Integer
Dim Week As String
Dim Month As String
Dim NumToMonth As String
Dim Year As String

Dim NameSheetAttendance As String
Dim ReportDate As String
Dim headoffice As Variant
Dim branch As Variant

'////////////////////////////////////////
'/////////Fill This Information//////////
Date1 = "08/09/2021"
Date2 = "08/10/2021"
Date3 = "08/11/2021"
Date4 = "08/12/2021"
Date5 = "08/13/2021"
Day = 9
Week = "W2"
Month = 8
Year = 2021
'/////////Fill This Information//////////
'////////////////////////////////////////

NameSheetAttendance = ActiveSheet.Name
NumToMonth = VBA.Format(Month * 29, "mmmm")
ReportDate = Year & "_" & NumToMonth & "_" & Week

'///////////Prepare New Sheet////////////
'////////////////////////////////////////
Columns(3).EntireColumn.Delete
Columns(4).EntireColumn.Delete

headoffice = Array("Accounting", "Assistant Director", "BDO", "Branch Controller", "Business Process Improvement", "CAD", "CCD", "Collection", "Compliance", "DPC", "Factoring", "FBD", "GAD", "HRD", "HX", "Internal Audit", "Internal Control", "ITD", "Legal", "MCAI", "MEO", "MMC Business", "OPL", "Portfolio Management", "Tax", "Treasury", "xx1", "xx2", "xx3", "xx4", "xx5", "xx6", "xx7")
branch = Array("BDG", "BJM", "BKS", "BKT", "BLP", "CLG", "CRB", "DPS", "DRI", "JBR", "JKC", "JKN", "JKS", "JMB", "KRW", "LMP", "MBO", "MDN", "MDN2", "MKS", "MLG", "PDG", "PDS", "PKB", "PLG", "PTK", "PWK", "RPT", "SBY", "SKU", "SLO", "SMG", "TNG")
For i = 0 To 65
  If i < 33 Then
    Sheets.Add
    ActiveSheet.Name = "Branch " + branch(i)
    Worksheets("Branch " + branch(i)).Activate
    Range("A1") = branch(i) & " Branch Attendance Summary " & Week & " " & NumToMonth & " " & Year
    With Range("A1")
        .Font.Size = 16
        .Font.Bold = True
        .Font.Name = "Calibri"
    End With

    With Range("A1:K2")
        .Merge
        .Select
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
    End With
  Else
    Sheets.Add
    ActiveSheet.Name = "HO " + headoffice(i-33)
    Worksheets("HO " + headoffice(i-33)).Activate
    Range("A1") = headoffice(i-33) & " Division Attendance Summary " & Week & " " & NumToMonth & " " & Year
    With Range("A1")
        .Font.Size = 16
        .Font.Bold = True
        .Font.Name = "Calibri"
    End With

    With Range("A1:K2")
        .Merge
        .Select
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
    End With
  End If
    Range("A4") = "Employee Name"
    Range("B4") = Month & "/" & Day & "/" & Year
    Range("D4") = Range("B4") + 1
    Range("F4") = Range("B4") + 2
    Range("H4") = Range("B4") + 3
    Range("J4") = Range("B4") + 4
    Range("B4:J4").NumberFormat = "d-mmm-yy"
    Range("B5:J5") = "Start"
    Range("C5:K5") = "End"

    Range("A4:K5").HorizontalAlignment = xlCenterAcrossSelection
    Range("A4:K5").VerticalAlignment = xlCenter
    Range("A4:K5").Interior.Color = RGB(242, 242, 242)
    Range("A4:K5").Font.Bold = True

    Range("A4:A5").Merge
    Range("B4:C4").Merge
    Range("D4:E4").Merge
    Range("F4:G4").Merge
    Range("H4:I4").Merge
    Range("J4:K4").Merge
Next


'////////////Fill New Sheet/////////////
'////////////////////////////////////////
Worksheets(NameSheetAttendance).Activate    
lRow = Sheets(NameSheetAttendance).Range("A50000").End(xlUp).Row
For j = 1 To lRow Step 1
  For i = 0 To 32
'////////////Fill Branch\ Sheet/////////////
'////////////Cmpr Date > LogonName/////////////
    If Sheets(NameSheetAttendance).Range("B" & j) = "Branch\" & branch(i) Then
      bRow = Sheets("Branch " & branch(i)).Range("B50000").End(xlUp).Row
      If Sheets(NameSheetAttendance).Range("A" & j) = "08/09/2021" Then
        Sheets(NameSheetAttendance).Range("D" & j & ":F" & j).Cut Sheets("Branch " & branch(i)).Range("A" & bRow + 1)
      ElseIf Sheets(NameSheetAttendance).Range("A" & j) = "08/10/2021" Then
        Found = "no"
        For k = bRow To 5 Step -1
          If Sheets(NameSheetAttendance).Range("D" & j) = Sheets("Branch " & branch(i)).Range("A" & k) Then
            Sheets(NameSheetAttendance).Range("E" & j & ":F" & j).Cut Sheets("Branch " & branch(i)).Range("D" & k)
            Found = "yes"
          End If
        Next
        If Found = "no" Then
          Sheets(NameSheetAttendance).Range("D" & j).Cut Sheets("Branch " & branch(i)).Range("A" & bRow + 1)
          Sheets(NameSheetAttendance).Range("E" & j & ":F" & j).Cut Sheets("Branch " & branch(i)).Range("D" & bRow + 1)
        End If
      ElseIf Sheets(NameSheetAttendance).Range("A" & j) = "08/11/2021" Then
        Found = "no"
        For k = bRow To 5 Step -1
          If Sheets(NameSheetAttendance).Range("D" & j) = Sheets("Branch " & branch(i)).Range("A" & k) Then
            Sheets(NameSheetAttendance).Range("E" & j & ":F" & j).Cut Sheets("Branch " & branch(i)).Range("F" & k)
            Found = "yes"
          End If
        Next
        If Found = "no" Then
          Sheets(NameSheetAttendance).Range("D" & j).Cut Sheets("Branch " & branch(i)).Range("A" & bRow + 1)
          Sheets(NameSheetAttendance).Range("E" & j & ":F" & j).Cut Sheets("Branch " & branch(i)).Range("F" & bRow + 1)
        End If
      ElseIf Sheets(NameSheetAttendance).Range("A" & j) = "08/12/2021" Then
        Found = "no"
        For k = bRow To 5 Step -1
          If Sheets(NameSheetAttendance).Range("D" & j) = Sheets("Branch " & branch(i)).Range("A" & k) Then
            Sheets(NameSheetAttendance).Range("E" & j & ":F" & j).Cut Sheets("Branch " & branch(i)).Range("H" & k)
            Found = "yes"
          End If
        Next
        If Found = "no" Then
          Sheets(NameSheetAttendance).Range("D" & j).Cut Sheets("Branch " & branch(i)).Range("A" & bRow + 1)
          Sheets(NameSheetAttendance).Range("E" & j & ":F" & j).Cut Sheets("Branch " & branch(i)).Range("H" & bRow + 1)
        End If
      ElseIf Sheets(NameSheetAttendance).Range("A" & j) = "08/13/2021" Then
        Found = "no"
        For k = bRow To 5 Step -1
          If Sheets(NameSheetAttendance).Range("D" & j) = Sheets("Branch " & branch(i)).Range("A" & k) Then
            Sheets(NameSheetAttendance).Range("E" & j & ":F" & j).Cut Sheets("Branch " & branch(i)).Range("J" & k)
            Found = "yes"
          End If
        Next
        If Found = "no" Then
          Sheets(NameSheetAttendance).Range("D" & j).Cut Sheets("Branch " & branch(i)).Range("A" & bRow + 1)
          Sheets(NameSheetAttendance).Range("E" & j & ":F" & j).Cut Sheets("Branch " & branch(i)).Range("J" & bRow + 1)
        End If
      End If
'////////////Fill HO\ Sheet/////////////
'//////////Cmpr Date > LogonName///////////
    ElseIf Sheets(NameSheetAttendance).Range("B" & j) = "HO\" & headoffice(i) Then
      bRow = Sheets("HO " & headoffice(i)).Range("B50000").End(xlUp).Row
      If Sheets(NameSheetAttendance).Range("A" & j) = "08/09/2021" Then
        Sheets(NameSheetAttendance).Range("D" & j & ":F" & j).Cut Sheets("HO " & headoffice(i)).Range("A" & bRow + 1)
      ElseIf Sheets(NameSheetAttendance).Range("A" & j) = "08/10/2021" Then
        Found = "no"
        For k = bRow To 5 Step -1
          If Sheets(NameSheetAttendance).Range("D" & j) = Sheets("HO " & headoffice(i)).Range("A" & k) Then
            Sheets(NameSheetAttendance).Range("E" & j & ":F" & j).Cut Sheets("HO " & headoffice(i)).Range("D" & k)
            Found = "yes"
          End If
        Next
        If Found = "no" Then
          Sheets(NameSheetAttendance).Range("D" & j).Cut Sheets("HO " & headoffice(i)).Range("A" & bRow + 1)
          Sheets(NameSheetAttendance).Range("E" & j & ":F" & j).Cut Sheets("HO " & headoffice(i)).Range("D" & bRow + 1)
        End If
      ElseIf Sheets(NameSheetAttendance).Range("A" & j) = "08/11/2021" Then
        Found = "no"
        For k = bRow To 5 Step -1
          If Sheets(NameSheetAttendance).Range("D" & j) = Sheets("HO " & headoffice(i)).Range("A" & k) Then
            Sheets(NameSheetAttendance).Range("E" & j & ":F" & j).Cut Sheets("HO " & headoffice(i)).Range("F" & k)
            Found = "yes"
          End If
        Next
        If Found = "no" Then
          Sheets(NameSheetAttendance).Range("D" & j).Cut Sheets("HO " & headoffice(i)).Range("A" & bRow + 1)
          Sheets(NameSheetAttendance).Range("E" & j & ":F" & j).Cut Sheets("HO " & headoffice(i)).Range("F" & bRow + 1)
        End If
      ElseIf Sheets(NameSheetAttendance).Range("A" & j) = "08/12/2021" Then
        Found = "no"
        For k = bRow To 5 Step -1
          If Sheets(NameSheetAttendance).Range("D" & j) = Sheets("HO " & headoffice(i)).Range("A" & k) Then
            Sheets(NameSheetAttendance).Range("E" & j & ":F" & j).Cut Sheets("HO " & headoffice(i)).Range("H" & k)
            Found = "yes"
          End If
        Next
        If Found = "no" Then
          Sheets(NameSheetAttendance).Range("D" & j).Cut Sheets("HO " & headoffice(i)).Range("A" & bRow + 1)
          Sheets(NameSheetAttendance).Range("E" & j & ":F" & j).Cut Sheets("HO " & headoffice(i)).Range("H" & bRow + 1)
        End If
      ElseIf Sheets(NameSheetAttendance).Range("A" & j) = "08/13/2021" Then
        Found = "no"
        For k = bRow To 5 Step -1
          If Sheets(NameSheetAttendance).Range("D" & j) = Sheets("HO " & headoffice(i)).Range("A" & k) Then
            Sheets(NameSheetAttendance).Range("E" & j & ":F" & j).Cut Sheets("HO " & headoffice(i)).Range("J" & k)
            Found = "yes"
          End If
        Next
        If Found = "no" Then
          Sheets(NameSheetAttendance).Range("D" & j).Cut Sheets("HO " & headoffice(i)).Range("A" & bRow + 1)
          Sheets(NameSheetAttendance).Range("E" & j & ":F" & j).Cut Sheets("HO " & headoffice(i)).Range("J" & bRow + 1)
        End If
      End If
    End If
  Next
Next

'////////////Format New Sheet/////////////
'////////////////////////////////////////
For i = 0 To 32 Step 1
  aRow = Sheets("HO " & headoffice(i)).Range("A50000").End(xlUp).Row
  Worksheets("HO " & headoffice(i)).Range("B4").Font.Color = vbRed
  For j = 66 To 75 Step 1
    For k = 6 To aRow Step 1
      If j Mod 2 = 1 Then
      If Worksheets("HO " & headoffice(i)).Range(Chr(j) & k).Value > TimeValue("17:00:00") Then
        Worksheets("HO " & headoffice(i)).Range(Chr(j) & k).Font.Color = vbBlue
      End If
      Else
      If Worksheets("HO " & headoffice(i)).Range(Chr(j) & k).Value > TimeValue("08:00:00") Then
        Worksheets("HO " & headoffice(i)).Range(Chr(j) & k).Font.Color = vbRed
      End If
      End If
    Next
  Next
  With Worksheets("HO " & headoffice(i)).Range("A4:K" & aRow).Borders
    .LineStyle = xlContinuous
    .Weight = xlThin
  End With
  With Worksheets("HO " & headoffice(i)).Columns("A")
    .ColumnWidth = 20
  End With
Next
For i = 0 To 32 Step 1
  aRow = Sheets("Branch " & branch(i)).Range("A50000").End(xlUp).Row
  Worksheets("Branch " & branch(i)).Range("B4").Font.Color = vbRed
  For j = 66 To 75 Step 1
    For k = 6 To aRow Step 1
      If j Mod 2 = 1 Then
      If Worksheets("Branch " & branch(i)).Range(Chr(j) & k).Value > TimeValue("17:00:00") Then
        Worksheets("Branch " & branch(i)).Range(Chr(j) & k).Font.Color = vbBlue
      End If
      Else
      If Worksheets("Branch " & branch(i)).Range(Chr(j) & k).Value > TimeValue("08:00:00") Then
        Worksheets("Branch " & branch(i)).Range(Chr(j) & k).Font.Color = vbRed
      End If
      End If
    Next
  Next
  With Worksheets("Branch " & branch(i)).Range("A4:K" & aRow).Borders
    .LineStyle = xlContinuous
    .Weight = xlThin
  End With
  With Worksheets("Branch " & branch(i)).Columns("A")
    .ColumnWidth = 20
  End With
Next

'/////////////Export New xls//////////////
'////////////////////////////////////////
Worksheets(NameSheetAttendance).Activate

'Export the attandance report
Dim FileExtStr As String
Dim FileFormatNum As Long
Dim xWs As Worksheet
Dim xWb As Workbook
Dim FolderName As String
Application.ScreenUpdating = False
Set xWb = Application.ThisWorkbook
DateString = Format(Now, "yyyy-mm-dd hh-mm-ss")
FolderName = xWb.Path & "\" & xWb.Name & " " & DateString
MkDir FolderName
For Each xWs In xWb.Worksheets
    xWs.Copy
    If Val(Application.Version) < 12 Then
        FileExtStr = ".xls": FileFormatNum = -4143
    Else
        Select Case xWb.FileFormat
            Case 51:
                FileExtStr = ".xlsx": FileFormatNum = 51
            Case 52:
                If Application.ActiveWorkbook.HasVBProject Then
                    FileExtStr = ".xlsm": FileFormatNum = 52
                Else
                    FileExtStr = ".xlsx": FileFormatNum = 51
                End If
            Case 56:
                FileExtStr = ".xls": FileFormatNum = 56
            Case Else:
                FileExtStr = ".xlsb": FileFormatNum = 50
        End Select
    End If
    xFile = FolderName & "\" & ReportDate & "_Attendance Summary " & Application.ActiveWorkbook.Sheets(1).Name & FileExtStr
    Application.ActiveWorkbook.SaveAs xFile, FileFormat:=FileFormatNum
    Application.ActiveWorkbook.Close False
Next
MsgBox "This script is created by Ghifari. You can find the files in " & FolderName
Application.ScreenUpdating = True

End Sub
