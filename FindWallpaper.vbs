' Find Wallpaper for Windows 8.1
' Author: Jason Ku
' Version: 1.0.0
' Description: Finds the current desktop wallpaper(s), and selects the file in a new Explorer window.
'
' Notes: Only tested on Windows 8.1.
'        May open several other unknown images since the TranscodedImageCache keys are sometimes not deleted.

set WshShell = CreateObject("WScript.Shell")

maxMonitors = 20
regKeyEnding = ""

For monitorIndex = -1 To maxMonitors	'TranscodedImageCache, TranscodedImageCache_000, TranscodedImageCache_001, ...
	If monitorIndex = -1 Then
		regKeyEnding = ""
	ElseIf monitorIndex < 10 Then
		regKeyEnding = "_00" & CInt(monitorIndex)
	Else
		regKeyEnding = "_0" & CInt(monitorIndex)
	End If

	On Error Resume Next
	imgInfoArray = WshShell.RegRead("HKCU\Control Panel\Desktop\TranscodedImageCache" & regKeyEnding)

	If Err.number = 0 Then
		intArray = imgInfoArray
		imgChr = " "
		imgPath = ""
		endOfName = 256
		
		For I = 24 To Ubound(imgInfoArray)		'Ignore first 23 characters
			imgChr = " "
			intArray(I) = Cint(imgInfoArray(I))

			If intArray(I) > 0 And I < endOfName Then 'Check if chr is valid
				imgChr = Chr(intArray(I))	'Convert to chr
				imgPath = imgPath & imgChr
				If intArray(I) = 46 And intArray(I+10) = 0 Or intArray(I+12) = 0 then		'Look for end of file name (.jpg, .jpeg, .png, etc.)
					endOfName = (I+8)
				End If
			End If

		Next
		
		'Open the file in explorer
		filePath = """" & imgPath & """"
		return = WshShell.Run("explorer.exe /select," & filePath,,true)
	End If
Next

WScript.Quit