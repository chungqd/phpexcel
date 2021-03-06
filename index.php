<?php
require_once 'connect/db_connect.php';
require_once 'Classes/PHPExcel.php';

if (isset($_POST['btnSubmit'])) {
    $file = $_FILES['file']['tmp_name'];

    $objReader = PHPExcel_IOFactory::createReaderForFile($file);
    $objReader->setLoadSheetsOnly('10A1');

    $objExcel = $objReader->load($file);
    $sheetData = $objExcel->getActiveSheet()->toArray('null', true, true, true);
    $getHighestRow = $objExcel->setActiveSheetIndex()->getHighestRow();

    $conn = connection();
    for($row = 2; $row <= $getHighestRow; $row++) {
        $class_id = 1;
        $name = $sheetData[$row]['A'];
        $toan = $sheetData[$row]['B'];
        $ly = $sheetData[$row]['C'];
        $hoa = $sheetData[$row]['D'];

        // insert to db
        $sql = "INSERT INTO point(class_id, hoten, toan, ly, hoa) VALUES (:class_id, :hoten, :toan, :ly, :hoa)";
        $stmt = $conn->prepare($sql);
        if ($stmt) {
            $stmt->bindPARAM(":class_id", $class_id, PDO::PARAM_INT);
            $stmt->bindPARAM(":hoten", $name, PDO::PARAM_STR);
            $stmt->bindPARAM(":toan", $toan, PDO::PARAM_INT);
            $stmt->bindPARAM(":ly", $ly, PDO::PARAM_INT);
            $stmt->bindPARAM(":hoa", $hoa, PDO::PARAM_INT);
            $stmt->execute();
            $stmt->closeCursor();
        }
    }
    disconnection($conn);
    echo "Insert success";
}
?>
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <title>Import data from file sheet</title>
    <!-- Latest compiled and minified CSS & JS -->
    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/css/bootstrap.min.css" integrity="sha384-1q8mTJOASx8j1Au+a5WDVnPi2lkFfwwEAa8hDDdjZlpLegxhjVME1fgjWPGmkzs7" crossorigin="anonymous">
    <script src="//code.jquery.com/jquery.js"></script>
    <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/js/bootstrap.min.js" integrity="sha384-0mSbJDEHialfmuBBQP6A4Qrprq5OVfW37PRR3j5ELqxss1yVqOtnepnHVP9aJ7xS" crossorigin="anonymous"></script>
</head>
<body style="margin: 20px auto;">
    <div class="container">
        <div class="row">
            <form action="" method="POST" class="col-md-4 col-md-offset-4 form-inline" enctype="multipart/form-data" role="form">
                    <div class="form-group">
                        <input type="file" name="file" class="form-control">
                        <button type="submit" class="btn btn-primary" name="btnSubmit">Send</button>
                    </div>
            </form>
        </div>
    </div>
</body>
</html>
