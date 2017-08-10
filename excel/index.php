<?php

include('Classes/PHPExcel.php');

include('connect/connect.php');

if(isset($_POST['btnGui'])){
	// $f = $_FILES['file'];
	// print_r($f); die;
	$file = $_FILES['file']['tmp_name'];
	//----------------------video 1: read only sheet------------------
	/*$objReader = PHPExcel_IOFactory::createReaderForFile($file);
	$objReader->setLoadSheetsOnly('10A1');
	$objPHPExcel = $objReader->load($file);
	$sheetData = $objPHPExcel->getActiveSheet()->toArray(null,true,true,true);
	echo ' Highest Column '. $getHighestColumn 
								= $objPHPExcel->setActiveSheetIndex()->getHighestColumn();
	echo ' Get Highest Row '. $getHighestRow 
								= $objPHPExcel->setActiveSheetIndex()->getHighestRow(); 
	
	for($row = 2; $row<=$getHighestRow; $row++){
		$name = ($sheetData[$row]['A']);
		$toan = ($sheetData[$row]['B']);
		$ly = ($sheetData[$row]['C']);
		$hoa = ($sheetData[$row]['D']);

		$sql = "INSERT INTO diem (name,toan,ly,hoa) VALUES ('$name',$toan,$ly, $hoa)";
		$mysqli->query($sql);

	}
	echo 'thành công';*/
	
	//-----------------------video 2: read all sheeet---------------------
	$objReader = PHPExcel_IOFactory::createReaderForFile($file);
	$loadedSheetNames = $objReader->listWorksheetNames($file);
	
    foreach($loadedSheetNames as $sheetIndex => $loadedSheetName) {
    	$result = $mysqli->query("SELECT id FROM lop WHERE name='$loadedSheetName'");
		$id = mysqli_fetch_assoc($result); print_r($id);;die;
    	$objReader->setLoadSheetsOnly($loadedSheetName);
		$objPHPExcel = $objReader->load($file);
		$sheetData = $objPHPExcel->getActiveSheet()->toArray(null,true,true,true);
		$getHighestColumn = $objPHPExcel->setActiveSheetIndex()->getHighestColumn();
		$getHighestRow = $objPHPExcel->setActiveSheetIndex()->getHighestRow(); 
		for($row = 2; $row<=$getHighestRow; $row++){
			$name = ($sheetData[$row]['A']);
			$toan = ($sheetData[$row]['B']);
			$ly = ($sheetData[$row]['C']);
			$hoa = ($sheetData[$row]['D']);

			$sql = "INSERT INTO diem (name,toan,ly,hoa) VALUES ('$name',$toan,$ly, $hoa)";
			$mysqli->query($sql);

		}
    }

    die;
    //------------------video 3: export only sheet----------------------

    $result = $mysqli->query("SELECT * FROM diem");
	
	$objPHPExcel = new PHPExcel(); 
	$objPHPExcel->getProperties()->setCreator("Huong1212")
                                 ->setTitle("Office 2016 XLSX Test Document")
                                 ->setSubject("Office 2016 XLSX Test Document")
                                 ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
                                 ->setKeywords("PHP Excel KhoaPham")
                                 ->setCategory("Course PHP");

	


	$objPHPExcel->setActiveSheetIndex(0);
	$sheet = $objPHPExcel->getActiveSheet()->setTitle("f");	
	$sheet->getColumnDimension('A')->setAutoSize(true);

	$rowCount = 1;
	$sheet->SetCellValue('A'.$rowCount,'Firstname');
	$sheet->SetCellValue('B'.$rowCount,'Toán');
	$sheet->SetCellValue('C'.$rowCount,'Lý');
	$sheet->SetCellValue('D'.$rowCount,'Hóa');
	$sheet->getStyle('A1:D1')->getFill()->setFillType(\PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('00ffff00');
	$sheet->getStyle('A1:D1')->getAlignment()
    		->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

	while($row = mysqli_fetch_array($result)){ 
		
	    $rowCount++;
	    $sheet->SetCellValue('A'.$rowCount, $row['2']);
	    $sheet->SetCellValue('B'.$rowCount, $row['3']);
	    $sheet->SetCellValue('C'.$rowCount, $row['4']);
	    $sheet->SetCellValue('D'.$rowCount, $row['5']);
	} 
	$sheet->setCellValue('D'.($rowCount+1), "=SUM(D2:D$rowCount)/COUNT(D2:D$rowCount)");
	$sheet->getStyle('D'.($rowCount+1))->getFont()->setBold(true);
	$sheet->mergeCells("A".($rowCount+1).":C".($rowCount+1));
	$sheet->SetCellValue('A'.($rowCount+1), 'Điểm trung bình:');
	$sheet->getStyle('A'.($rowCount+1))->getAlignment()
    		->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

    $styleArray = array(
      	'borders' => array(
          	'allborders' => array(
              	'style' => PHPExcel_Style_Border::BORDER_THIN
          	)
      	)
  	);
	$sheet->getStyle(
	    'A1:' . 'D'.($rowCount+1)
	)->applyFromArray($styleArray);


	$objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);

	$filename = "excelfilename.xlsx"; 
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