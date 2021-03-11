#SingleInstance, force
#NoEnv
SetTitleMatchMode 1
DetectHiddenWindows, On
SetBatchLines -1
SetControlDelay, -1
SetWinDelay, -1
SetMouseDelay, -1
SendMode Input
ListLines Off

;The next line makes the hotkeys active only when PowerPoint is the active window
#IfWinActive ahk_class PPTFrameClass

; F3 to resize selected object to 65%
F3::
Resize_Object(65)
Return

; Ctrl + Shift + 8 to select the % to resize the selected object
^+8::
Resize_InputBox:
Inputbox, Size_InputBox, Resize factor, What resize factor for the object?,,250, 100,,,,,80
If Errorlevel
	Return
Resize_Object(Size_InputBox)
Return

Resize_Object(PercentSize)
	{
	WinActivate, ahk_class PPTFrameClass
	Try
		{
		ppt := ComObjActive("PowerPoint.Application")
		ppt.ActiveWindow.Selection.ShapeRange.ScaleHeight(PercentSize/100,1)
		}
	}


; Ctrl + Shift + r to change the color of selected objects or text to red
^+r::
WinActivate, ahk_class screenClass ahk_exe POWERPNT.EXE
Try
	{
	ppt := ComObjActive("PowerPoint.Application")
	;Msgbox, % ppt.ActiveWindow.Selection.Shaperange.Count
	If (ppt.ActiveWindow.Selection.Type = 2)
		{
		Try
			ppt.ActiveWindow.Selection.ShapeRange.TextFrame.TextRange.Font.Color.RGB:=0x0000FF
		Try
			ppt.ActiveWindow.Selection.ShapeRange.Line.ForeColor.RGB:=0x0000FF
		}
	If (ppt.ActiveWindow.Selection.Type = 3)
		ppt.ActiveWindow.Selection.TextRange.Font.Color.RGB:=0x0000FF
	}
Return

; Ctrl + Shift + u to change the color of selected objects or text to blue
^+u::
WinActivate, ahk_class screenClass ahk_exe POWERPNT.EXE
Try
	{
	ppt := ComObjActive("PowerPoint.Application")
	If (ppt.ActiveWindow.Selection.Type = 2)
		{
		Try
			ppt.ActiveWindow.Selection.ShapeRange.TextFrame.TextRange.Font.Color.RGB:=0xFF0000
		Try
			ppt.ActiveWindow.Selection.ShapeRange.Line.ForeColor.RGB:=0xFF0000
		}
	If (ppt.ActiveWindow.Selection.Type = 3)
		ppt.ActiveWindow.Selection.TextRange.Font.Color.RGB:=0xFF0000
	}
Return

; Ctrl + Shift + k to change the color of selected objects or text to black
^+k::
WinActivate, ahk_class screenClass ahk_exe POWERPNT.EXE
Try
	{
	ppt := ComObjActive("PowerPoint.Application")
	If (ppt.ActiveWindow.Selection.Type = 2)
		{
		Try
			ppt.ActiveWindow.Selection.ShapeRange.TextFrame.TextRange.Font.Color.RGB:=0x000000
		Try
			ppt.ActiveWindow.Selection.ShapeRange.Line.ForeColor.RGB:=0x000000
		}
	If (ppt.ActiveWindow.Selection.Type = 3)
		ppt.ActiveWindow.Selection.TextRange.Font.Color.RGB:=0x000000
	}
Return

; Ctrl + 9 to decrease line thickness of selected object by 0.25
^9::
WinActivate, ahk_class screenClass ahk_exe POWERPNT.EXE
Try
	{
	Tooltip
	ppt := ComObjActive("PowerPoint.Application")
	SetFormat, FloatFast, 0.2
	Current_Thickness := ppt.ActiveWindow.Selection.ShapeRange.Line.Weight
	ppt.ActiveWindow.Selection.ShapeRange.Line.Weight:= Current_Thickness - 0.25
	Tooltip, % Current_Thickness - 0.25
	SetTimer, RemoveToolTip, 1000
	}
Return

; Ctrl + 0 to increase line thickness of selected object by 0.25
^0::
WinActivate, ahk_class screenClass ahk_exe POWERPNT.EXE
Try
	{
	Tooltip
	ppt := ComObjActive("PowerPoint.Application")
	SetFormat, FloatFast, 0.2
	Current_Thickness := ppt.ActiveWindow.Selection.ShapeRange.Line.Weight
	ppt.ActiveWindow.Selection.ShapeRange.Line.Weight:= Current_Thickness + 0.25
	Tooltip, % Current_Thickness + 0.25
	SetTimer, RemoveToolTip, 1000
	}
Return

RemoveToolTip: 
SetTimer, RemoveToolTip, Off
ToolTip
Return

; Ctrl + Shift + f to insert a formatted textbox with predefined and editable text
^+f::
WinActivate, ahk_class screenClass ahk_exe POWERPNT.EXE
Try
    {
    Inputbox, Slide_Title, Slide title, What is the number of the figure?,,250, 130,,,,,1
	If Errorlevel
		Return
	Text = Fig. %Slide_Title%
	ppt := ComObjActive("PowerPoint.Application")
	gog := ppt.ActiveWindow.View.Slide.Shapes.AddTextBox(1,50,40,100,0)
	gog.TextFrame.TextRange.Font.Color:=0x0000FF
    gog.TextFrame.TextRange.Font.Size:=14
    gog.TextFrame.TextRange.Font.Name:="Arial"
    gog.TextFrame.TextRange.Font.Bold:=0
	gog.TextFrame.TextRange.Font.Italic:=1
	gog.TextFrame.TextRange.Font.Underline:=0
	gog.TextFrame.TextRange.ParagraphFormat.Alignment:=2
	gog.Line.Visible:=1
	gog.Fill.ForeColor.RGB:=0xFF0000
	gog.Fill.Transparency:=0.8
	gog.Line.ForeColor.RGB:=0xFFFF33
    gog.TextFrame.TextRange.Characters.Text:=Text
    }
Return

; Ctrl + Shift + y to add all selected objects to the animation cue with "appear" entrance animation, with the first object set to "start" on click and others set to start with previous
^+y::
WinActivate, ahk_class screenClass ahk_exe POWERPNT.EXE
Try
	{
	ppt := ComObjActive("PowerPoint.Application")
	Loop % ppt.ActiveWindow.Selection.ShapeRange.Count
		{
		shpSelected:=ppt.ActiveWindow.Selection.ShapeRange(A_Index)
		If A_Index = 1
			ppt.ActiveWindow.Selection.SlideRange.Timeline.MainSequence.AddEffect(shpSelected,1)
		Else
			ppt.ActiveWindow.Selection.SlideRange.Timeline.MainSequence.AddEffect(shpSelected,1,,2)
		}
	}
Return

; Ctrl + Shift + z to add all selected objects to the animation cue with "appear" entrance animation, with all objects set to start with previous
^+z::
WinActivate, ahk_class screenClass ahk_exe POWERPNT.EXE
Try
	{
	ppt := ComObjActive("PowerPoint.Application")
	Loop % ppt.ActiveWindow.Selection.ShapeRange.Count
		{
		shpSelected:=ppt.ActiveWindow.Selection.ShapeRange(A_Index)
		ppt.ActiveWindow.Selection.SlideRange.Timeline.MainSequence.AddEffect(shpSelected,1,,2)
		}
	}
Return

; Ctrl + Alt + c to align selected objects to horizontal center of slide
^!c::
WinActivate, ahk_class screenClass ahk_exe POWERPNT.EXE
Try
	{
	ppt := ComObjActive("PowerPoint.Application")
	ppt.ActiveWindow.Selection.ShapeRange.Align(1,1)
	}
Return