<?php

include('Classes/PHPExcel.php');
include('connect/connect.php');

if(isset($_POST['btn'])){
	// $f = $_FILES['file'];
	// print_r($f); die;
	$file = $_FILES['file']['tmp_name'];
	$result = $mysqli->query("SELECT * FROM lop");
	
	$objPHPExcel = new PHPExcel(); 
	$numSheet = 0;
	while($row = mysqli_fetch_array($result)){ 
		$numSheet++;
		$objPHPExcel->createSheet();
		$objPHPExcel->setActiveSheetIndex($numSheet-1);
		$sheet = $objPHPExcel->getActiveSheet()->setTitle($row[1]);
		
		$rowCount = 1;
		$sheet->SetCellValue('A'.$rowCount,'Firstname');
		$sheet->SetCellValue('B'.$rowCount,'Toán');
		$sheet->SetCellValue('C'.$rowCount,'Lý');

		$ds = $mysqli->query("SELECT * FROM diem WHERE id_lop=$row[0]");
		while($hs = mysqli_fetch_array($ds)){
			$rowCount++;
		    $sheet->SetCellValue('A'.$rowCount, $hs['2']);
		    $sheet->SetCellValue('B'.$rowCount, $hs['3']);
		    $sheet->SetCellValue('C'.$rowCount, $hs['4']);
		}
	    
	} 
	    
	//die;
	$objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
	$filename = "4.xlsx"; 
	$objWriter->save($filename);
	header('Content-Disposition: attachment; filename="' . $filename . '"');
	header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
	header('Content-Length: ' . filesize($filename));
	header('Content-Transfer-Encoding: binary');
	header('Cache-Control: must-revalidate');
	header('Pragma: no-cache');
	readfile($filename);
	return;

}

?>

<!DOCTYPE html>
<html>
<head>
	<meta charset="utf-8">
	<meta http-equiv="X-UA-Compatible" content="IE=edge">
	<title>Index</title>
	<link rel="stylesheet" href="">
</head>
<body>
	<form action="" method="POST" accept-charset="utf-8" enctype="multipart/form-data">
		Chọn file<input type="file" name="file">
		<button type="submit" name="btn">Gửi</button>
	</form>
</body>
</html>