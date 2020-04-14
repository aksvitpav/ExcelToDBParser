<?php

use Phppot\DataSource;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx as Reader;

require_once 'db.php';
require_once ('./vendor/autoload.php');

if (isset($_POST['import'])) {
    $allowedFileTypes = [
        'application/vnd.ms-excel',
        'text/xls',
        'text/xlsx',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    ];

    if (in_array($_FILES["file"]["type"], $allowedFileTypes)) {
        $targetPath = 'uploads/' . $_FILES['file']['name'];
        move_uploaded_file($_FILES['file']['tmp_name'], $targetPath);
        $Reader = new Reader();
        $spreadSheet = $Reader->load($targetPath);
        $excelSheet = $spreadSheet->getActiveSheet();
        $spreadSheetArray = $excelSheet->toArray();
        $sheetCount = count($spreadSheetArray);
        $insertCount = 0;
        for ($i = 1; $i <= $sheetCount; $i ++) {
            $name=null;
            if (isset($spreadSheetArray[$i][0])) {
                $name = $spreadSheetArray[$i][0];
            }
            $count=null;
            if (isset($spreadSheetArray[$i][1])) {
                $count = $spreadSheetArray[$i][1];
            }
            $price=null;
            if (isset($spreadSheetArray[$i][2])) {
                $price = $spreadSheetArray[$i][2];
            }
            if (isset($name) && isset($count) && isset($price)) {
                $sql = "INSERT INTO products (name, count, price) VALUES (?,?,?)";
                $request = $pdo->prepare($sql);
                $request->execute([$name, $count, $price]);
                $insertCount++;
            }
        }
        if ($insertCount == $sheetCount-1) {
            $type = 'success';
            $msg = 'Все записи из файла добавлены в базу данных!';
        } 
        else {
            $type = 'danger';
            $msg = "Импортировано $insertCount записей из ".($sheetCount-1); 
        }
    }
    else {
        $type ='danger';
        $msg = 'Некорректный тип файла! Загрузите файл Excel!';
    }
}
?>

<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Парсер Excel - Результаты</title>
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.4.1/css/bootstrap.min.css" integrity="sha384-Vkoo8x4CGsO3+Hhxv8T/Q5PaXtkKtu6ug5TOeNV6gBiFeWPGFN9MuhOf23Q9Ifjh" crossorigin="anonymous">
</head>
<body>
<div class="container">
        <div class="row">
            <div class="col-6 offset-3">
                <h1>Импорт данных из Excel</h1>
                <div class="alert alert-<?php echo($type) ?>" role="alert">
                    <?php echo($msg) ?>
                </div>
            </div>
        </div>
    </div>
    <script src="https://code.jquery.com/jquery-3.4.1.slim.min.js" integrity="sha384-J6qa4849blE2+poT4WnyKhv5vZF5SrPo0iEjwBvKU7imGFAV0wwj1yYfoRSJoZ+n" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/popper.js@1.16.0/dist/umd/popper.min.js" integrity="sha384-Q6E9RHvbIyZFJoft+2mJbHaEWldlvI9IOYy5n3zV9zzTtmI3UksdQRVvoxMfooAo" crossorigin="anonymous"></script>
    <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.4.1/js/bootstrap.min.js" integrity="sha384-wfSDF2E50Y2D1uUdj0O3uMBJnjuUD4Ih7YwaYd1iqfktj0Uod8GCExl3Og8ifwB6" crossorigin="anonymous"></script> 
</body>
</html>

