<?php

include('Classes/PHPExcel.php');
include('connect/connect.php');

if(isset($_POST['btnGui'])){
	// $f = $_FILES['file'];
	// print_r($f); die;
	$file = $_FILES['file']['tmp_name'];
	$result = $mysqli->query("SELECT * FROM diem");
	
	$objPHPExcel = new PHPExcel(); 
	$objPHPExcel->setActiveSheetIndex(0);
	$sheet = $objPHPExcel->getActiveSheet()->setTitle("f");
	//$objWorkSheet->setName("Bảng điểm");
	 
	//print_r($sheet); die;
	$rowCount = 1;
	$sheet->SetCellValue('A'.$rowCount,'Firstname');
	$sheet->SetCellValue('B'.$rowCount,'Toán');
	$sheet->SetCellValue('C'.$rowCount,'Lý');
	$sheet->SetCellValue('D'.$rowCount,'Hóa');
	while($row = mysqli_fetch_array($result)){ 
		
	    $rowCount++;
	    $sheet->SetCellValue('A'.$rowCount, $row['2']);
	    $sheet->SetCellValue('B'.$rowCount, $row['3']);
	    $sheet->SetCellValue('C'.$rowCount, $row['4']);
	    $sheet->SetCellValue('D'.$rowCount, $row['5']);
	} 

	$objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
	$filename = "n.xlsx"; 
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
		<button type="submit" name="btnGui">Gửi</button>
	</form>
</body>
</html>