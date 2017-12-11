Option Explicit

Function file_choose() As Variant

	file_choose = Application.GetOpenFilename( _
	           FileFilter:="Text Files (*.txt), *.txt", _
	           Title:="Select files to process", _
	           MultiSelect:=True)

End Function


Sub reset_silent()
	Dim cur_ws As Worksheet
	Set cur_ws = ActiveSheet 
	Call Initialize_vars

	Call clear_right_down(ws_interface.Range("A10"))

	Worksheets(cur_ws.Name).Activate
End Sub