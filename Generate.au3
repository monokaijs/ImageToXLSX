#cs
###
###		Author: @MonokaiJs | https://nstudio.pw
###
#ce

#include <GDIPlus.au3>
#include <Excel.au3>
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <ProgressConstants.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#Region ### START Koda GUI section ### Form=
$Form1 = GUICreate("Image to Excel", 342, 86)
$Label1 = GUICtrlCreateLabel("Import Image", 8, 8, 65, 17)
$Input1 = GUICtrlCreateInput("", 80, 8, 201, 21, $ES_READONLY)
$Button1 = GUICtrlCreateButton("...", 288, 8, 49, 25)
$Button2 = GUICtrlCreateButton("Start", 288, 40, 49, 33)
$Progress1 = GUICtrlCreateProgress(8, 40, 273, 33)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit
		Case $Button1
			$fh = FileOpenDialog("Choose Image file", @ScriptDir, "Image File(*.*)", 1)
			if $fh <> "" Then
				GUICtrlSetData($Input1, $fh)
			EndIf
		Case $Button2
			if (GUICtrlRead($Input1) <> "") Then
			$sImage = GUICtrlRead($Input1)
			_GDIPlus_Startup()
			$hImage = _GDIPlus_ImageLoadFromFile($sImage)
			$iWidth = _GDIPlus_ImageGetWidth($hImage)
			$iHeight = _GDIPlus_ImageGetHeight($hImage)
			$amount = $iWidth * $iHeight
			$fh = FileOpen("runtime.au3", 8+2)
			FileWriteLine($fh, "#include <Excel.au3>")
			FileWriteLine($fh, "Global $oExcel = _Excel_Open()")
			FileWriteLine($fh, "Global $oWorkbook = _Excel_BookNew($oExcel, 5)")
			FileWriteLine($fh, "Global $sWorkbook = @ScriptDir & '\test.xlsx'")
			FileWriteLine($fh, "_Excel_BookSaveAs($oWorkbook, $sWorkbook, $xlWorkbookDefault, True)")
			FileWriteLine($fh, "With $oExcel.ActiveWorkbook.Sheets(1)")
			$old_name = ""
			GUICtrlSetData($Progress1, 0)
			$done = 0
			For $x = 1 to $iWidth
				For $y = 1 to $iHeight
					$col = _ColorConvert("0x" & string(Hex(_GDIPlus_BitmapGetPixel($hImage, $x, $y), 6)))
					if $old_name <> Test_Name($x) Then
						FileWriteLine($fh, ".Range('" & Test_Name($x) & $y & ":" & Test_Name($x) & $y & "').Columns.ColumnWidth = 2")
						$old_name = Test_Name($x)
					EndIf
					FileWriteLine($fh, ".Range('" & Test_Name($x) & $y & ":" & Test_Name($x) & $y & "').Interior.Color =" & $col)
					$done += 100/$amount
					GUICtrlSetData($Progress1, $done)
				Next
			Next
			FileWriteLine($fh, "EndWith")
			FileClose($fh)
			_GDIPlus_ImageDispose($hImage)
			_GDIPlus_ShutDown()
			MsgBox(64, "Success", "Done Generation!")
			Else
			MsgBox(16, "Error", "You must choose one image!")
			EndIf
	EndSwitch
WEnd

Func _ColorConvert($nColor);RGB to BGR or BGR to RGB
    Return Hex( _
        BitOR(BitShift(BitAND($nColor, 0x000000FF), -16), _
        BitAND($nColor, 0x0000FF00), _
        BitShift(BitAND($nColor, 0x00FF0000), 16)), 6)
EndFunc

Func Test_Name($i)
	$dividend = $i;
    $columnName = "";
    while ($dividend > 0)
        $modulo = Mod(($dividend - 1), 26);
        $columnName = Chr(65 + $modulo) & $columnName;
        $dividend = int(($dividend - $modulo) / 26);
    WEnd
	Return $columnName
EndFunc
