#SingleInstance, Force
SendMode, InputThenPlay
CoordMode, Mouse, Window

main()
Return

main()
{
	Global Bar, currentCode, codeWin
	
	codes := SetupCodes()

	if(!webBrowser)
	{
		webBrowser := ComObjCreate("InternetExplorer.Application")
		webBrowser.Visible := True
	}

	webBrowser.Navigate("https://www.defiance.com/en/my-ego/")

	WaitForPage(webBrowser)

	MsgBox, 4, Easy Defiance Codes, Are you already logged into the Defiance page?
	IfMsgBox No
	{
		MsgBox, Please login to the page now.`nDO NOT click OK below until you are logged in.
		webBrowser.Navigate("https://www.defiance.com/en/my-ego/")
		WaitForPage(webBrowser)
	}
	
	if((buttonID := FindID(webBrowser)) < 1)
	{
		MsgBox, Couldn't locate the code input page.`nNow exiting.
		ExitApp
	}
	
	if(!InStr(webBrowser.LocationURL, "www.defiance.com/en/my-ego") || webBrowser.Document.All[buttonID].InnerText != "Submit Code")
	{
		MsgBox, Couldn't locate the code input page.`nNow exiting.
		ExitApp
	}
	
	MsgBox, Press OK, then click on the inputbox for the codes.
	KeyWait, LButton, D
	MouseGetPos, codeX, codeY, codeWin
	Sleep, 500
	MsgBox, Code position set. Please try not to close the window or click elsewhere while the process is running.`n`nIf the program gets stuck, press Escape.
	
	max := codes.MaxIndex()
	xPos := (A_ScreenWidth/2)-300
	yPos := A_ScreenHeight*.9
	
	Gui, +AlwaysOnTop
	Gui, Add, Progress, x0 y2 w600 h20 Range0-%max% vBar, 0
	Gui, Add, Text, x250 y4 w100 h20 +BackgroundTrans +Center vcurrentCode, None
	Gui, Show, x%xPos% y%yPos% h24 w600, Code Entry Progress
	
	Loop, % codes.MaxIndex()
	{
		GuiControl,, Bar, %A_Index%
		GuiControl, Text, currentCode, %A_Index%
		WinActivate, ahk_id %codeWin%
		
		Click, %codeX%, %codeY%
		Send % codes[A_Index] "{Enter}"
		
		Loop
		{
			if(!webBrowser)
			{
				MsgBox, You appear to have closed the window.`nNow exiting.
				ExitApp
			}
			Sleep 200
		}
		Until (webBrowser.Document.All.arkfallcode.Value == "")
	}
	
	MsgBox, All done!
	
	ExitApp
}

Escape::ExitApp

OnExit:
{
	If(codeWin)
		WinClose, %codeWin%
	Return
}
	
WaitForPage(session, timeout = 10)
{
	timeout *= 1000
	startWaitTime := A_TickCount
	
	Loop																					; Wait until page starts to load
	{
		if(timeout && (timeout) <= (A_TickCount - startWaitTime))
			Return
	}
	Until (session.busy)
	Loop																					; Wait until it finishes loading
	{
		if(timeout && (timeout) <= (A_TickCount - startWaitTime))
			Return
	}
	Until (!session.busy)
	Sleep 50
	Loop																					; Optional: Wait for the page to "completely" load
	{
		if(timeout && (timeout) <= (A_TickCount - startWaitTime))
			Return
	}
	Until (session.Document.Readystate = "Complete")
	
	Return
}

FindID(webBrowser)
{
	id := 0
	index := 0
	exitLoop := false
	
	While(!exitLoop)
	{
		Try
		{
			Loop
			{
				index := mod(index, 250) + 1
				
				Sleep 10

				if(A_Index >= 30000)
					Return false
			}
			Until (webBrowser.Document.All[id := index].InnerText == "Submit Code")

			exitLoop := true
		}
		Catch
		{
			Sleep 1000
		}
	}
	
	Return id
}

SetupCodes()
{
	codes := Array()
	codes.Insert("E2DFMO")
	codes.Insert("kbZmVz")
	codes.Insert("VO5DHS")
	codes.Insert("6CIBJT")
	codes.Insert("fYkqzX")
	codes.Insert("RpR9Zs")
	codes.Insert("Tfm2iH")
	codes.Insert("fccePO")
	codes.Insert("vpZk7b")
	codes.Insert("XBNBPE")
	codes.Insert("DKCJK3")
	codes.Insert("CWAPDX")
	codes.Insert("ETCTFJ")
	codes.Insert("TLK9F1")
	codes.Insert("AMY2ZH")
	codes.Insert("WA94BD")
	codes.Insert("PVXDTU")
	codes.Insert("9ELBUP")
	codes.Insert("YJuicp")
	codes.Insert("SCsd5t")
	codes.Insert("hvN5yD")
	codes.Insert("8MKLUL")
	codes.Insert("v2227Q")
	codes.Insert("XCTFGY")
	codes.Insert("TWIK9S")
	codes.Insert("VGFSW9")
	codes.Insert("4OSBUA")
	codes.Insert("BJT9iD")
	codes.Insert("LDBAMY")
	codes.Insert("RQXPGX")
	codes.Insert("Z5OP9L")
	codes.Insert("CH8FJB")
	codes.Insert("I71ADU")
	codes.Insert("GBIXKI")
	codes.Insert("GT1P3H")
	codes.Insert("W0UODG")
	codes.Insert("JJD6N5")
	codes.Insert("3Q6VZN")
	codes.Insert("mvV8sl")
	codes.Insert("vW4LQ0")
	codes.Insert("WKJIXT")
	codes.Insert("U1T8PG")
	codes.Insert("CEOzCG")
	codes.Insert("6zDGAR")
	codes.Insert("yOt4zA")
	codes.Insert("LQ7WVN")
	codes.Insert("7VLERN")
	codes.Insert("o2am1s")
	codes.Insert("J71AAM")
	codes.Insert("IS3KAF")
	codes.Insert("htg7bf")
	codes.Insert("PSZ3iE")
	codes.Insert("LFM6RV")
	codes.Insert("UUXTGI")
	codes.Insert("OTIINU")
	codes.Insert("34TPus")
	codes.Insert("FBwtVC")
	codes.Insert("6f9kiv")
	codes.Insert("3wwjkk")
	codes.Insert("8MOMX2")
	codes.Insert("A1aqCa")
	codes.Insert("JIMOFE")
	codes.Insert("2A9WZT")
	codes.Insert("KHIBBT")
	codes.Insert("M2BDRP")
	codes.Insert("HMDQRN")
	codes.Insert("EXA6RK")
	codes.Insert("WOMCCQ")
	codes.Insert("YW6ZQ8")
	codes.Insert("LK47XN")
	codes.Insert("JQ9ZTM")
	codes.Insert("R1T3EB")
	codes.Insert("WBIEos")
	codes.Insert("JIGR3J")
	codes.Insert("RXVI6J")
	codes.Insert("ISHSQF")
	codes.Insert("CHUROO")
	codes.Insert("mewfwy")
	codes.Insert("vmnnre")
	codes.Insert("Dt0chz")
	codes.Insert("K880P4")
	codes.Insert("RFtrHs")
	codes.Insert("OiUQQ5")
	codes.Insert("FZY8CB")
	codes.Insert("53VIZG")
	codes.Insert("OLXQJ6")
	codes.Insert("V5WKFT")
	codes.Insert("OBWTZL")
	codes.Insert("hC1gN0")
	codes.Insert("ZFR5TA")
	codes.Insert("D91AQK")
	codes.Insert("W72URO")
	codes.Insert("41US8B")
	codes.Insert("YDBF4L")
	codes.Insert("6JOMJO")
	codes.Insert("SZUZSY")
	codes.Insert("CNABZX")
	codes.Insert("XT4PMX")
	codes.Insert("PKYLKN")
	codes.Insert("8KBMUT")
	codes.Insert("JBS7QX")
	codes.Insert("TU3RJF")
	codes.Insert("7BRHNJ")
	codes.Insert("MEBTWK")
	codes.Insert("F3LZLK")
	codes.Insert("TA6NCJ")
	codes.Insert("PTZ4QR")
	codes.Insert("N4S9EK")
	codes.Insert("5X4WJ4")
	codes.Insert("UT7DE0")
	codes.Insert("6VWFCI")
	codes.Insert("81HABS")
	codes.Insert("HUBAZV")
	codes.Insert("GEPKRJ")
	codes.Insert("Q4DUBL")
	codes.Insert("NC9VYK")
	codes.Insert("TTDGX4")
	codes.Insert("9C9N87")
	codes.Insert("ZZKJ2G")
	codes.Insert("T76SN6")
	codes.Insert("XWRN9M")
	codes.Insert("88UPWS")
	codes.Insert("NFQ7PG")
	codes.Insert("35P72U")
	codes.Insert("8UNKAD")
	codes.Insert("YRE5SN")
	codes.Insert("HP9X6X")
	codes.Insert("XQBG4H")
	codes.Insert("4HT2G9")
	codes.Insert("LYCEHY")
	codes.Insert("7P8NDC")
	codes.Insert("C2VNDF")
	Return codes
}