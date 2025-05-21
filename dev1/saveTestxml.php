<?php
require 'vendor/autoload.php'; // Composerでライブラリを読み込む
 
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx as XlsxWriter;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx as XlsxReader;
 
// 新しいスプレッドシートオブジェクトを作成
//$spreadsheet = new Spreadsheet();
//$spreadsheet = PhpOffice\PhpSpreadsheet\IOFactory::load('TemplateData/PDFTest.xlsx');

$reader = new XlsxReader();
$spreadsheet = $reader->load('TemplateData/PDFTest.xlsx');

//// 保存先のディレクトリ
//$saveDirectory = "uploads/";
//
//// ディレクトリが存在しない場合は作成
//if (!is_dir($saveDirectory))
//{
//    mkdir($saveDirectory, 0777, true);
//
//}

//// POSTデータを受け取る
//$filename = $_POST['name'] ?? 'default.xml';
//$content = $_POST['content'] ?? '';

//// ファイルパスを設定
//$filePath = $saveDirectory . basename($filename);

// 受信したものを取り出す
$rawData = file_get_contents("php://input");

// JSONをデコード
$data = json_decode($rawData, true);

if ($data !== null) {
    $date_time = new DateTime();

    $testStr1 = $data['testStr1'];
    $testStr2 = $data['testStr2'];
    $testStr2 = $testStr2.$date_time->format("Y年n月d日");

    $testStr3 = $data['testStr3'];
    $testStr4 = $data['testStr4'];

    //$testStr3Count = count($);
    //$testStr4Count = count($data['testStr4']);

    //$testStr3 = $data['testStr3'][0];
    //$testStr4 = $data['testStr4'][0];

    $sheet = $spreadsheet->getActiveSheet();
    $sheet->setCellValue('B2', $testStr1);
    $sheet->setCellValue('G3', $testStr2);

    $startCell = 'C5';
    foreach ($testStr3 as $value) {
        if ($startCell == 'D0') {
            $startCell = 'C10';
        }
        $sheet->setCellValue($startCell, $value);
        ++$startCell;
    }

    $startCell = 'G5';
    foreach ($testStr4 as $value) {
        if ($startCell == 'H0') {
            $startCell = 'G10';
        }
        $sheet->setCellValue($startCell, $value);
        ++$startCell;
    }

    //$sheet->fromArray($data, null, 'A1');
        //$sheet->setCellValue("A", $data[0]['name']);
        //$sheet->setCellValue("B", $data[0]['score']);

    // セルの横幅をテキストに合わせて自動調整
    //$sheet->getColumnDimension('A')->setAutoSize(true);
    $sheet->getColumnDimension('C')->setAutoSize(true);
    $sheet->getColumnDimension('G')->setAutoSize(true);

    // カーソルをA1セルに設定
    $sheet->setSelectedCell('A1');

    // サーバーに保存するファイル名
//    $filePath = 'example.xlsx';

        //$fileName = '請求書_20200412.xlsx';
        //    
        //header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;');
        //header("Content-Disposition: attachment; filename=\"{$fileName}\"");
        //header('Cache-Control: max-age=0');

    //// Excelファイルにデータを書き出し
    //$writer = new Xlsx($spreadsheet);
    //$writer->save($filePath/*'example.xlsx'*/);

//    $fileName = $filePath;
//        
//    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
//    header("Content-Disposition: attachment; filename=\"{$fileName}\"");
//    header('Cache-Control: max-age=0');
//
//    $writer2 = IOFactory::createWriter($spreadsheet, 'Xlsx');
//    //$writer2 = PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xlsx');
//    $writer2->save('php://output');
//    exit;



    // ファイルの出力設定
    $filename = 'exsample.xlsx';
    
    // ヘッダーの設定（ブラウザにダウンロードさせるため）
    //header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;');
    header('Content-Disposition: attachment; filename="example.xlsx"');
    header('Cache-Control: max-age=0');
    
    // Excelファイルを書き出し
    $writer = new XlsxWriter($spreadsheet);
    $writer->save('php://output');
    exit;
}
?>

