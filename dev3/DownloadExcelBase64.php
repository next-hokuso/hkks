<?php
require 'vendor/autoload.php'; // Composerでライブラリを読み込む
 
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx as XlsxWriter;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx as XlsxReader;
 
$reader = new XlsxReader();
$spreadsheet = $reader->load('TemplateData/PDFTest.xlsx');

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

    $sheet = $spreadsheet->getActiveSheet();
    $sheet->setCellValue('B2', $testStr1);
    $sheet->setCellValue('G3', $testStr2);

    $startCell = 'C5';
    foreach ($testStr3 as $value) {
        // インクリメントだとC9の次はD0になってしまうため、C10に変更させる
        if ($startCell == 'D0') {
            $startCell = 'C10';
        }
        $sheet->setCellValue($startCell, $value);
        ++$startCell;
    }

    $startCell = 'G5';
    foreach ($testStr4 as $value) {
        // インクリメントだとG9の次はH0になってしまうため、G10に変更させる
        if ($startCell == 'H0') {
            $startCell = 'G10';
        }
        $sheet->setCellValue($startCell, $value);
        ++$startCell;
    }

    // セルの横幅をテキストに合わせて自動調整
    //$sheet->getColumnDimension('A')->setAutoSize(true);
    $sheet->getColumnDimension('C')->setAutoSize(true);
    //$sheet->getColumnDimension('G')->setAutoSize(true);

    // カーソルをA1セルに設定
    $sheet->setSelectedCell('A1');

    // ローカルサーバー以外だとCORS設定を弄らないとエラーになってセーブ出来ない
    //// サーバーに保存するファイル名
    //$filePath = 'example.xlsx';
    //// Excelファイルにデータを書き出し
    //$writer = new XlsxWriter($spreadsheet);
    //$writer->save($filePath);

    $writer = new XlsxWriter($spreadsheet);
    // 一旦メモリに保存
    ob_start();
    $writer->save('php://output');
    $excelData = ob_get_clean();
    
    // Base64エンコード
    $base64 = base64_encode($excelData);
    
    // JSONで返す
    header('Content-Type: application/json');
    echo json_encode(['base64' => $base64]);
    exit();
}
?>
