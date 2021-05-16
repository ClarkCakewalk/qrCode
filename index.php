<?php
//ini_set('display_errors', 1);
set_time_limit(0);
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

use Endroid\QrCode\Color\Color;
use Endroid\QrCode\Encoding\Encoding;
use Endroid\QrCode\ErrorCorrectionLevel\ErrorCorrectionLevelLow;
use Endroid\QrCode\QrCode;
use Endroid\QrCode\Label\Label;
use Endroid\QrCode\Logo\Logo;
use Endroid\QrCode\RoundBlockSizeMode\RoundBlockSizeModeMargin;
use Endroid\QrCode\Writer\PngWriter;

 function num2alpha($n)  //數字轉英文(0=>A、1=>B、26=>AA...以此類推)
{
    for($r = ""; $n >= 0; $n = intval($n / 26) - 1)
        $r = chr($n%26 + 0x41) . $r; 
    return $r; 
}

function logoSelect($logoImg) {
	switch ($logoImg) {
		case 'PSFD':
			$logopath='./img/PSFD.jpg';
			break;
		case 'CSR':
			$logopath='./img/CSR.jpg';
			break;
		case 'SRDA':
			$logopath='./img/SRDA.png';
			break;
		case 'SROD':
			$logopath='./img/SROD.jpg';	
			break;	
		case 'RCHSS':
			$logopath='./img/RCHSS.png';
			break;
		case 'AS':
			$logopath='./img/AS_logo-1.png';
			break;
		default:
			$logopath='';
			break;
	}
	return $logopath;
}

function createQRCode ($value, $logoImg, $mode, $name='default') {
	$writer = new PngWriter();

	// Create QR code
	$qrCode = QrCode::create($value)
	    ->setEncoding(new Encoding('UTF-8'))
	    ->setErrorCorrectionLevel(new ErrorCorrectionLevelLow())
	    ->setSize(200)
	    ->setMargin(10)
	    ->setRoundBlockSizeMode(new RoundBlockSizeModeMargin())
	    ->setForegroundColor(new Color(0, 0, 0))
	    ->setBackgroundColor(new Color(255, 255, 255));

	// Create generic logo
	if (empty(logoSelect($logoImg))) {
		$logo=null;
	}
	else {
		$logo = Logo::create(logoSelect($logoImg))
	    ->setResizeToWidth(40);
	}


	// Create generic label
	//$label = Label::create('Label')
	//    ->setTextColor(new Color(255, 0, 0));
	//    ->setBackgroundColor(new Color(0, 0, 0));

	$result = $writer->write($qrCode, $logo);

	if($mode==1) {
		$result->saveToFile(__DIR__.'/filetmp/'.$name.'.png');
	}
	if($mode==2) {
		$dataUri = $result->getDataUri();
		return $dataUri;
	}	
}
?>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>QR Code產生器</title>
</head>

<body>
<script type="text/javascript" language="javascript">
function checkfile(sender) {
    var validExts = new Array(".xlsx", ".xls");
    var fileExt = sender.value;
    fileExt = fileExt.substring(fileExt.lastIndexOf('.'));
    if (validExts.indexOf(fileExt) < 0) {
      alert("檔案格式錯誤，請選擇" +
               validExts.toString() + "格式的檔案。");
      return false;
    }
    else return true;
}
</script>
<h4>批次產生QR code</h4>
<form id="upload" name="upload" method="post" action="" enctype="multipart/form-data">
  <p>
    <label>Step1: 請選擇上傳檔案
      <input type="file" name="file" id="file" onchange="checkfile(this);" accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel"/>
    </label>
  </p>
    <p><label>
  	<table>
  		<tr>
  			<td colspan="2">Step2: 請選擇logo</td>
  		</tr>
  		<tr>
  			<td colspan="2"><input type="radio" name="logoImg" value="none" checked="checked">不加logo</td>
  		</tr>
  		<tr>
  			<td><input type="radio" name="logoImg" value="PSFD">PSFD</td>
  			<td><input type="radio" name="logoImg" value="CSR">CSR</td>
  		</tr>
  		<tr>
  			<td><input type="radio" name="logoImg" value="SRDA">SRDA</td>
  			<td><input type="radio" name="logoImg" value="SROD">SROD</td>
  		</tr>
  		<tr>
  			<td><input type="radio" name="logoImg" value="RCHSS">RCHSS</td>
  			<td><input type="radio" name="logoImg" value="AS">Academia Sinica</td>
  		</tr>
<!--  		<tr>
  			<td colspan="2"><input type="radio" name="logoImg" value="custom">自行上傳logo<input type="file" name="clogo" id="clogo" accept="image/png"/><br />自行上傳logo限png格式圖檔。</td>
  		</tr>
-->  	  		  		
  	</table>
  </label></p>
  <p>
    <label>
      <input type="submit" name="button" id="button" value="開始轉換"/>
    </label>
  </p>
</form>
<p>上傳檔案說明：</p>
<p>1. 上傳檔案限制為excel格式檔案。</p>
<p>2. A欄為QR code檔名，B欄為QR code內容（例如網址），第一列為標題列，系統將從第二列開始進行轉換。</p>
<?php
if (!empty($_FILES)) {
$filetype=strrchr($_FILES["file"]["name"], ".");
$filename=substr($_FILES["file"]["name"],0, strripos($_FILES["file"]["name"], $filetype));
if ($filetype==".xls" or $filetype==".xlsx") {
	move_uploaded_file($_FILES["file"]["tmp_name"],"./filetmp/".$_FILES["file"]["name"]);
	$loadFile=PhpOffice\PhpSpreadsheet\IOFactory::load('./filetmp/'.$_FILES["file"]["name"]);
//	$loadFile=PhpOffice\PhpSpreadsheet\IOFactory::load($_FILES);
	$sheetData=$loadFile->getSheet(0)->toArray();
	$zip = new ZipArchive();
	$zipfile=$filename.'.zip';
	$createZip=$zip->open(__DIR__.'/filetmp/'.$zipfile, ZipArchive::CREATE);
	if($createZip===true) {
		for ($i=1; $i<count($sheetData); $i++) {
			createQRCode($sheetData[$i][1], $_POST["logoImg"], 1, $sheetData[$i][0]);
			$filepath='/filetmp/';
			$qrfile[]=$sheetData[$i][0].'.png';
			$zip->addFile(__DIR__.$filepath.$sheetData[$i][0].'.png', $sheetData[$i][0].'.png');
		}
	}
	$zip->close();
	unlink(__DIR__.'/filetmp/'.$_FILES["file"]["name"]);
	foreach ($qrfile as $delfile) {
		unlink(__DIR__.'/filetmp/'.$delfile);
	}
	header("Content-type:application/zip");
	header("Content-Disposition:filename=".$zipfile);
	ob_clean();
	flush();
	readfile(__DIR__.'/filetmp/'.$zipfile);
	unlink(__DIR__.'/filetmp/'.$zipfile);
}
else { 
	echo "上傳檔案格式錯誤！！";
	unlink($_FILES["file"]["tmp_name"]);
}
}
?>
<hr />
<h4>產生單筆QR Code</h4>
<form id="form1" name="form1" method="post" action="">
  <p><label>Step1: 請輸QR Code內容：
    <input type="text" name="name" id="name" />
  </label></p>
  <p><label>
  	<table>
  		<tr>
  			<td colspan="2">Step2: 請選擇logo</td>
  		</tr>
  		<tr>
  			<td colspan="2"><input type="radio" name="logoImg" value="none" checked="checked">不加logo</td>
  		</tr>
  		<tr>
  			<td><input type="radio" name="logoImg" value="PSFD">PSFD</td>
  			<td><input type="radio" name="logoImg" value="CSR">CSR</td>
  		</tr>
  		<tr>
  			<td><input type="radio" name="logoImg" value="SRDA">SRDA</td>
  			<td><input type="radio" name="logoImg" value="SROD">SROD</td>
  		</tr>
  		<tr>
  			<td><input type="radio" name="logoImg" value="RCHSS">RCHSS</td>
  			<td><input type="radio" name="logoImg" value="AS">Academia Sinica</td>
  		</tr>
 <!-- 		<tr>
  			<td colspan="2"><input type="radio" name="logoImg" value="custom">自行上傳logo<input type="file" name="clogo" id="clogo" accept="image/png"/><br />自行上傳logo限png格式圖檔。</td>
  		</tr>
-->  		  	  		  		
  	</table>
  </label></p>
  <label>
    <input type="submit" name="button2" id="button2" value="查詢" />
  </label>
</form>
<?php
if (!empty($_POST["name"])) {
	echo "<img src=\"".createQRCode($_POST["name"], $_POST["logoImg"], 2)."\"/>";
}
?>
</body>
</html>
