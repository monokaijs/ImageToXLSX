<?php
$target_dir = "images/";
$uploadOk = 1;
$imageFileType = strtolower(pathinfo($_FILES['fileToUpload']['name'],PATHINFO_EXTENSION));

if(isset($_POST["submit"])) {
    $check = getimagesize($_FILES["fileToUpload"]["tmp_name"]);
    if($check !== false) {
		if($imageFileType != "jpg" && $imageFileType != "jpeg" ) {
			echo "Sorry, only JPG, JPEG files are allowed.";
		} else {
			if ($_FILES["fileToUpload"]["size"] > 100) {
				$filename = auth_str(25);
				$target_file = $target_dir . $filename . '.' . $imageFileType;
				if (move_uploaded_file($_FILES["fileToUpload"]["tmp_name"], $target_file)) {
					create_code($target_file, "sources/" . $filename . ".au3");
					echo "<a href='sources/" . $filename . ".au3'>Download Source</a><br/>";
				} else {
					echo "Unexpected error!";
				}
			} else {
				echo "File too large.";
			}
		}
    } else {
        echo "File is not an image.";
    }
}

function auth_str($length = 10) {
    $characters = '0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ';
    $charactersLength = strlen($characters);
    $randomString = '';
    for ($i = 0; $i < $length; $i++) {
        $randomString .= $characters[rand(0, $charactersLength - 1)];
    }
    return $randomString . time();
}
function create_code($sImage, $sOutput) {
	$hImage = new imagick($sImage);
	$ImageSize = getimagesize($sImage);
	$iWidth = $ImageSize[0];
	$iHeight= $ImageSize[1];
	$hFile = fopen($sOutput, "w+");
	fwrite($hFile, '#include <Excel.au3>
	Global $oExcel = _Excel_Open()
	Global $oWorkbook = _Excel_BookNew($oExcel, 5)
	Global $sWorkbook = @ScriptDir & "\test.xlsx"
	_Excel_BookSaveAs($oWorkbook, $sWorkbook, $xlWorkbookDefault, True)
	With $oExcel.ActiveWorkbook.Sheets(1)'.PHP_EOL);
	$x = 0;
	$y = 0;
	For ($x = 0; $x< $iWidth; $x++) {
		fwrite($hFile, ".Range('" . num2alpha($x) . ($y+1) . ":" . num2alpha($x) . ($y + 1) . "').Columns.ColumnWidth = 2".PHP_EOL);
		For ($y = 0; $y< $iHeight; $y++) {
			$pixel = $hImage->getimagepixelcolor($x, $y);
	        $color = $pixel->getColor();
			$color = sprintf("%02x%02x%02x", $color['b'], $color['g'], $color['r']);
			fwrite($hFile, ".Range('" . num2alpha($x) . ($y+1) . ":" . num2alpha($x) . ($y + 1) . "').Interior.Color = 0x" . $color. PHP_EOL);
		}
	}
	fwrite($hFile, 'EndWith');
	fclose($hFile);
}
function num2alpha($n)
{
    for($r = ""; $n >= 0; $n = intval($n / 26) - 1)
        $r = chr($n%26 + 0x41) . $r;
    return $r;
}
?>
