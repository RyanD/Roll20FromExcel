#NoEnv  
#Include Chrome.ahk
; #Warn
SendMode Input 
SetWorkingDir %A_ScriptDir% 

FileCreateDir, ChromeProfile
ChromeInst := new Chrome("ChromeProfile")
PageInst := ChromeInst.GetPage()
 
PageInst.Call("Page.navigate", {"url": "https://app.roll20.net/editor/"})
PageInst.WaitForLoad()

TwitchWindowName := "Twitch"
TwitchWindowClass := 
TwitchEXE := "TwitchUI.exe"
ExcelWindowName := "Hype.xlsb.xlsx"
ExcelWindowClass := 
ExcelEXE := "EXCEL.EXE"

return

+^f1::
SetTitleMatchMode, 2
GroupAdd, rpg, Twitch
GroupAdd, rpg, Hype.xlsb.xlsx
GroupActivate, rpg, R
if ErrorLevel
   MsgBox, No window was found.
Return
return

+^f2::
if WinActive("ahk_class XLMAIN")
	XL := ComObjActive("Excel.Application")
	Cell := XL.ActiveCell
	Label := % XL.Cells(Cell.Row,Cell.Column-3).Value
	Roll := % XL.Cells(Cell.Row,Cell.Column+1).Value
	PageInst.Evaluate("document.getElementById('textchat-input').getElementsByTagName('textarea')[0].value='Rolling " Label "'")
	PageInst.Evaluate("document.getElementById('textchat-input').getElementsByTagName('button')[0].click()")
	PageInst.Evaluate("document.getElementById('textchat-input').getElementsByTagName('textarea')[0].value='" Roll "'")
	PageInst.Evaluate("document.getElementById('textchat-input').getElementsByTagName('button')[0].click()")
return

+^f3::
PageInst.Evaluate("document.getElementById('textchat-input').getElementsByTagName('textarea')[0].value='/roll 1d20'")
PageInst.Evaluate("document.getElementById('textchat-input').getElementsByTagName('button')[0].click()")
return

+^f4::

return

+^f5::
PageInst.Evaluate("document.getElementById('textchat-input').getElementsByTagName('textarea')[0].value='/roll 1d12'")
PageInst.Evaluate("document.getElementById('textchat-input').getElementsByTagName('button')[0].click()")
return

+^f6::
PageInst.Evaluate("document.getElementById('textchat-input').getElementsByTagName('textarea')[0].value='/roll 1d8'")
PageInst.Evaluate("document.getElementById('textchat-input').getElementsByTagName('button')[0].click()")
return

+^f7::
PageInst.Evaluate("document.getElementById('textchat-input').getElementsByTagName('textarea')[0].value='/roll 1d4'")
PageInst.Evaluate("document.getElementById('textchat-input').getElementsByTagName('button')[0].click()")
return

+^f8::
PageInst.Evaluate("document.getElementById('textchat-input').getElementsByTagName('textarea')[0].value='/roll 1d6'")
PageInst.Evaluate("document.getElementById('textchat-input').getElementsByTagName('button')[0].click()")
return

+^f9::
PageInst.Evaluate("document.getElementById('textchat-input').getElementsByTagName('textarea')[0].value='/roll 2d6'")
PageInst.Evaluate("document.getElementById('textchat-input').getElementsByTagName('button')[0].click()")
return

+^f10::
PageInst.Evaluate("document.getElementById('textchat-input').getElementsByTagName('textarea')[0].value='/roll 3d6'")
PageInst.Evaluate("document.getElementById('textchat-input').getElementsByTagName('button')[0].click()")
return

+^f11::
PageInst.Evaluate("document.getElementById('textchat-input').getElementsByTagName('textarea')[0].value='/roll 4d6'")
PageInst.Evaluate("document.getElementById('textchat-input').getElementsByTagName('button')[0].click()")
return

+^f12::
PageInst.Evaluate("document.getElementById('textchat-input').getElementsByTagName('textarea')[0].value='/roll 5d6'")
PageInst.Evaluate("document.getElementById('textchat-input').getElementsByTagName('button')[0].click()")
return

^f1::
PageInst.Evaluate("document.getElementById('textchat-input').getElementsByTagName('textarea')[0].value='/roll 6d6'")
PageInst.Evaluate("document.getElementById('textchat-input').getElementsByTagName('button')[0].click()")
return

^f2::
PageInst.Evaluate("document.getElementById('textchat-input').getElementsByTagName('textarea')[0].value='/roll 5d8'")
PageInst.Evaluate("document.getElementById('textchat-input').getElementsByTagName('button')[0].click()")
return