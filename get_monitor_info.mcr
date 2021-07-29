' get_monitor_info

Sub Main ()

	'SelectTreeItem (Resulttree.GetFirstChildName("Probes"))
	'MsgBox GetSelectedTreeItem

	Dim ItemPath As String
	Dim numOfMonitors As Integer

'	numOfMonitors = 0

	ItemPath = Resulttree.GetFirstChildName("Probes")
	numOfMonitors = 1
'	MsgBox cstr(numOfMonitors) & vbCrLf & ItemPath

'	Do While ItemPath <> ""
	Do While numOfMonitors < 3
		ItemPath = Resulttree.GetNextItemName(ItemPath)
		numOfMonitors = numOfMonitors + 1
'		MsgBox cstr(numOfMonitors) & vbCrLf & ItemPath
	Loop
'	MsgBox cstr(numOfMonitors)

	Dim buf1 As String, buf2 As String, buf3 As String
	Dim buf4 As String, buf5 As String

	With Probe
'		.SetCoordinateSystemType "Spherical"
		.GetFirst
		Do
'			MsgBox .GetCaption & vbCrLf & cstr(.GetTheta)
			buf1 = buf1 & vbCrLf & .GetCaption & " -> " & cstr(.GetPosition1)
			buf2 = buf2 & vbCrLf & .GetCaption & " -> " & cstr(.GetPosition2)
			buf3 = buf3 & vbCrLf & .GetCaption & " -> " & cstr(.GetPosition3)
			buf4 = buf4 & vbCrLf & "r = " & Format(.GetPosition3 * sind(.GetPosition1) , "0000") & ", h = " & Format(.GetPosition3 *cosd(.GetPosition1) , "+0000;-0000")
			buf5 = buf5 & vbCrLf & .GetCaption & " -> " & .GetOrientation & ", phi = " & Format(.GetPosition2, "000") & " , h = " & Format(.GetPosition3 *cosd(.GetPosition1) , "+0000;-0000")
		Loop Until .GetNext = 0
	End With
'	MsgBox buf1
'	MsgBox buf2
'	MsgBox buf3
'	MsgBox buf4
	MsgBox buf5


End Sub
