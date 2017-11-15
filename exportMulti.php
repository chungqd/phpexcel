<?php
require_once 'connect/db_connect.php';
require_once 'Classes/PHPExcel.php';

if (isset($_POST['btnSubmit'])) {
    $objExcel = new PHPExcel;
    $numSheet = 0;

    $conn = connection();
    $data = array();
    $sql = "SELECT class.*, GROUP_CONCAT(DISTINCT point.hoten, '|', point.toan, '|', point.ly, '|', point.hoa) AS student
            FROM class
            INNER JOIN point ON point.class_id = class.id
            GROUP BY class.id";
    $stmt = $conn->prepare($sql);
    if ($stmt) {
        if ($stmt->execute()) {
            if ($stmt->rowCount() > 0) {
                $data = $stmt->fetchAll(PDO::FETCH_ASSOC);
            }
        }
        $stmt->closeCursor();
    }
    disconnection($conn);
//    print_r($data); die();
    foreach ($data as $value) {
        $objExcel->createSheet();
        $objExcel->setActiveSheetIndex($numSheet);
        $sheet = $objExcel->getActiveSheet()->setTitle($value['name']);

        $rowCount = 1;
        $sheet->setCellValue('A'.$rowCount, 'Ho ten');
        $sheet->setCellValue('B'.$rowCount, 'Toan');
        $sheet->setCellValue('C'.$rowCount, 'Ly');
        $sheet->setCellValue('D'.$rowCount, 'Hoa');

        $sheet->getColumnDimension("A")->setAutoSize(true);
        $sheet->getStyle('A1:D1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('A1:D1')->getFill()->setFillType(\PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('00ffff00');
        $sheet->getStyle('A1:D1')->getFont()->setBold(true);

        $student = explode(",", $value['student']);
        foreach ($student as $st) {
            list($name, $toan, $ly, $hoa) = explode("|", $st);
            $rowCount++;
            $sheet->setCellValue('A'.$rowCount, $name);
            $sheet->setCellValue('B'.$rowCount, $toan);
            $sheet->setCellValue('C'.$rowCount, $ly);
            $sheet->setCellValue('D'.$rowCount, $hoa);
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
        $sheet->getStyle('A1:' . 'D'.($rowCount + 1))->applyFromArray($styleArray);
        $numSheet++;
    }
    $objWriter = new PHPExcel_Writer_Excel2007($objExcel);
    $filename = "exportMulti.xlsx";
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

<!doctype html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport"
          content="width=device-width, user-scalable=no, initial-scale=1.0, maximum-scale=1.0, minimum-scale=1.0">
    <meta http-equiv="X-UA-Compatible" content="ie=edge">
    <title>Export multi sheet</title>
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
                <button type="submit" class="btn btn-primary" name="btnSubmit">Export Multi</button>
            </div>
        </div>
    </form>
</div>
</body>
</html>