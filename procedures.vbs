Option Explicit

'========================== Global variable declaration
'Sheets
Public ws_interface As Worksheet


'Ranges
' Public rb_import As Range


'========================== Interface Level

Sub Initialize_vars()
    On Error GoTo ErrHandler:
    'Initialize sheets
    Set ws_interface = ThisWorkbook.Sheets("Interface")

    'Initialize Ranges
    ' Set rb_import = ws_dataimport.Range("A11")

    Exit Sub

ErrHandler:
    MsgBox "Sheet failed to initialize properly." & vbCrLf & "Please enable Macros and click the Reset button."
End Sub

Sub Reset()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    Call reset_silent

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True

    MsgBox "Reset Complete!"
End Sub

Private Sub Test()
    'Call Initialize_vars
End Sub

Sub GetFiles()
    Dim files As Variant
    Dim entry As Long, x As Long
    Dim line As Variant
    Dim linepair() As String
    Dim spectrum As Variant
    Dim Delimiter As String
    Dim N As Long
    Dim FileContent As String
    Dim TextFile As Integer
    Dim LineArray() As String
    Dim datas() As String
    Dim prec As Long
    Dim linechose As String
    Dim filterv As Double
    Dim matchv As Double
    
    Call reset_silent


    files = file_choose()
    On Error GoTo ErrHandler:
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    entry = 0
    Delimiter = vbTab
    filterv = CDbl(ws_interface.Range("filter_value").Value)

    If IsArray(files) Then

        For Each spectrum In files
            TextFile = FreeFile
            Open spectrum For Input As TextFile

            FileContent = Input(LOF(TextFile), TextFile)

            Close TextFile

            LineArray() = Split(FileContent, vbCrLf)

            ws_interface.Range("A10").Offset(entry, 0).Value = Replace(UCase(LineArray(2)), ".TXT", "")

            datas = Split(Split(FileContent, "#DATA" & vbCrLf)(1), vbCrLf)

           For Each line In datas
              If InStr(line, Delimiter) <> 0 Then
                  linepair() = Split(line, Delimiter)
                    
                  If CDbl(linepair(0)) = filterv Then
                      ws_interface.Range("A10").Offset(entry, 1).Value = linepair(1)
                      Exit For
                  End If
              End If
                                   
                ws_interface.Range("A10").Offset(entry, 1).Value = "value not found"
            Next line

            entry = entry + 1

        Next spectrum
    Else
        If files = "False" Then
              Exit Sub
        End If
    End If

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True

    Exit Sub
    ErrHandler:
    MsgBox "Incorrect file format."
End Sub

