<?php
ini_set("memory_limit", "-1");
set_time_limit(0);

require 'vendor/autoload.php';
require 'Banks/N43.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

function parseN43File($filePath) {

    $fileContents = "";
    $handle = fopen($filePath, "r");
    if ($handle) {
        while (($line = fgets($handle, 81)) !== false) {
            $fileContents .= $line.PHP_EOL;
        }
        fclose($handle);
    } else {
        echo "Error opening the file.";
    }

    $data = array();

    $file = new Banks_N43();
    $file->parse($fileContents);

    foreach ($file->accounts as $account) {
        foreach ($account->entries as $entry) {
            $data[$account->number][] = $entry;
        }
    }

    return $data;
}

function createExcelFile($data, $excelFilePath) {
    
    $spreadsheet = new Spreadsheet();
    $i = 0;

    foreach($data as $account => $lines) {
        if ($i != 1) {
            $spreadsheet->createSheet();
        }
        $spreadsheet->setActiveSheetIndex($i);
        $spreadsheet->getActiveSheet()->setTitle($account);
        $spreadsheet->getActiveSheet()->freezePane('A2');
        $spreadsheet->getActiveSheet()->setAutoFilter('A1:M1');
        $sheet = $spreadsheet->getActiveSheet();
        
        $sheet->setCellValue('A1', 'Oficina');
        $sheet->setCellValue('B1', 'Fecha');
        $sheet->setCellValue('C1', 'Fecha Valor');
        $sheet->setCellValue('D1', 'Concepto Común');
        $sheet->setCellValue('E1', 'Concepto Propio');
        $sheet->setCellValue('F1', 'Tipo');
        $sheet->setCellValue('G1', 'Importe');
        $sheet->setCellValue('H1', 'Documento');
        $sheet->setCellValue('I1', 'Referencia 1');
        $sheet->setCellValue('J1', 'Referencia 2');
        $sheet->setCellValue('K1', 'Concepto 1');
        $sheet->setCellValue('L1', 'Concepto 2');
        $sheet->setCellValue('M1', 'Concepto 3');
        
        $row = 2;
        foreach ($lines as $transaction) {

            $sheet->setCellValue('A'.$row, $transaction->office);
            $sheet->setCellValue('B'.$row, date("d/m/Y", $transaction->date));
            $sheet->setCellValue('C'.$row, date("d/m/Y", $transaction->date_value));
            $sheet->setCellValue('D'.$row, $transaction->concept_common);
            $sheet->setCellValue('E'.$row, $transaction->concept_own);
            $sheet->setCellValue('F'.$row, $transaction->type);
            
            $transaction->amount = $transaction->type == Banks_N43::TYPE_DEBIT ? (-1 * $transaction->amount) : $transaction->amount;
            $sheet->setCellValue('G'.$row, $transaction->amount);
            $sheet->setCellValue('H'.$row, $transaction->document);
            $sheet->setCellValue('I'.$row, $transaction->refererence_1);
            $sheet->setCellValue('J'.$row, $transaction->refererence_2);
            if (is_array($transaction->concepts)) {
                if (isset($transaction->concepts['01'])) { $sheet->setCellValue('K'.$row, $transaction->concepts['01']); }
                if (isset($transaction->concepts['02'])) { $sheet->setCellValue('L'.$row, $transaction->concepts['02']); }
                if (isset($transaction->concepts['03'])) { $sheet->setCellValue('M'.$row, $transaction->concepts['03']); }
            }

            $row++;
        }
        
        foreach ($sheet->getColumnIterator() as $column) {
            $sheet->getColumnDimension($column->getColumnIndex())->setAutoSize(true);
        }

        $i++;
    }

    $writer = new Xlsx($spreadsheet);
    $writer->save($excelFilePath);
}

function displayHelp() {
    echo "Converts a Norma 43 (N43) file to an Excel file.\n";
    echo "Usage: php n43_to_excel.php <input_n43_file> <output_excel_file>\n";
}

if ($argc < 3 || $argv[1] == '--help') {
    displayHelp();
    exit(1);
}

$n43FilePath = $argv[1];
$excelFilePath = $argv[2];

if (!file_exists($n43FilePath)) {
    echo "Error: Input N43 file not found.\n";
    exit(1);
}

$data = parseN43File($n43FilePath);

createExcelFile($data, $excelFilePath);

echo "Excel file generated successfully!\n";