<?php
require_once 'connect/db_connect.php';
require_once 'Classes/PHPExcel.php';

if (isset($_POST['btnSubmit'])) {
    $objExcel = new PHPExcel;
    $objExcel->setActiveSheetIndex(0);

    // name of sheet
    $sheet = $objExcel->getActiveSheet()->setTitle('10A1');

    $rowCount = 1;
    $sheet->setCellValue('A'.$rowCount, 'Ho ten');
    $sheet->setCellValue('B'.$rowCount, 'Toan');
    $sheet->setCellValue('C'.$rowCount, 'Ly');
    $sheet->setCellValue('D'.$rowCount, 'Hoa');

    $sheet->getColumnDimension("A")->setAutoSize(true);
    $sheet->getStyle('A1:D1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    $sheet->getStyle('A1:D1')->getFill()->setFillType(\PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('00ffff00');

    // get data from db
    $conn = connection();
    $data = array();
    $sql = "SELECT hoten, toan, ly, hoa FROM point WHERE class_id = 1";
    $stmt = $conn->prepare($sql);
    if ($stmt) {
        if ($stmt->execute()) {
            if ($stmt->rowCount()>0) {
                $data = $stmt->fetchAll(PDO::FETCH_ASSOC);
            }
        }
        $stmt->closeCursor();
    }
    disconnection($conn);

    foreach ($data as $value) {
        $rowCount++;
        $sheet->setCellValue('A'.$rowCount, $value['hoten']);
        $sheet->setCellValue('B'.$rowCount, $value['toan']);
        $sheet->setCellValue('C'.$rowCount, $value['ly']);
        $sheet->setCellValue('D'.$rowCount, $value['hoa']);
    }
    // tinh diem trung binh
    $sheet->setCellValue('B'.($rowCount+1), "=AVERAGE(B2:B$rowCount)");
    $sheet->setCellValue('C'.($rowCount+1), "=AVERAGE(C2:C$rowCount)");
    $sheet->setCellValue('D'.($rowCount+1), "=AVERAGE(D2:D$rowCount)");
    $sheet->setCellValue('A'.($rowCount+1), "DTB:");

    $sheet->getStyle('A'.($rowCount + 1))->getFont()->setBold(true);
    $styleArray = array(
        'borders' => array(
            'allborders' => array(
                'style' => PHPExcel_Style_Border::BORDER_THIN
            )
        )
    );
    $sheet->getStyle('A1:' . 'D'.($rowCount))->applyFromArray($styleArray);

    $objWriter = new PHPExcel_Writer_Excel2007($objExcel);
    $filename = "export.xlsx";
    $objWriter->save($filename);

    header('Content-Disposition: attachment; filename="'.$filename.'"');
    header('Content-Type: application/vnd.openxmlformatsofficedocument.spreadsheetml.sheet');
    header('Content-Length:' . filesize($filename));
    header('Content-Transfer-Endcoding: binary');
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
    <title>Export data 1 file excel</title>
    <!-- Latest compiled and minified CSS & JS -->
    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/css/bootstrap.min.css" integrity="sha384-1q8mTJOASx8j1Au+a5WDVnPi2lkFfwwEAa8hDDdjZlpLegxhjVME1fgjWPGmkzs7" crossorigin="anonymous">
    <script src="//code.jquery.com/jquery.js"></script>
    <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/js/bootstrap.min.js" integrity="sha384-0mSbJDEHialfmuBBQP6A4Qrprq5OVfW37PRR3j5ELqxss1yVqOtnepnHVP9aJ7xS" crossorigin="anonymous"></script>
</head>
<body style="margin: 20px auto;">
    <div class="container">
        <form action="" method="post" class="form-horizontal" role="form">
            <div class="form-group">
                <div class="col-sm-4 col-sm-offset-4">
                    <button type="submit" class="btn btn-primary" name="btnSubmit">Export</button>
                </div>
            </div>
        </form>
    </div>
</body>
</html>
