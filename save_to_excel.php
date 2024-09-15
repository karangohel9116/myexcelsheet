<?php
require 'vendor/autoload.php'; // PhpSpreadsheet autoloader

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

if ($_SERVER["REQUEST_METHOD"] == "POST") {
    // Get form data
    $name = $_POST['name'];
    $email = $_POST['email'];
    $age = $_POST['age'];

    // Load or create the Excel file
    $fileName = 'data.xlsx';
    if (file_exists($fileName)) {
        $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($fileName);
        $worksheet = $spreadsheet->getActiveSheet();
    } else {
        $spreadsheet = new Spreadsheet();
        $worksheet = $spreadsheet->getActiveSheet();
        // Add headers if it's a new file
        $worksheet->setCellValue('A1', 'Name')
            ->setCellValue('B1', 'Email')
            ->setCellValue('C1', 'Age');
    }

    // Find the next empty row
    $row = $worksheet->getHighestRow() + 1;

    // Write data to Excel
    $worksheet->setCellValue("A$row", $name)
        ->setCellValue("B$row", $email)
        ->setCellValue("C$row", $age);

    // Save the file
    $writer = new Xlsx($spreadsheet);
    $writer->save($fileName);

    echo "Data has been saved to Excel!<br>";
    echo "<a href='data.xlsx' download>Click here to download the Excel file</a>";
}
