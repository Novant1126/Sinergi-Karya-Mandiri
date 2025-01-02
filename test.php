<?php
require __DIR__ . '/vendor/autoload.php'; // Autoload library dari Composer

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

try {
  // 1. Buat spreadsheet baru
  $spreadsheet = new Spreadsheet();
  $sheet = $spreadsheet->getActiveSheet();

  // 2. Tambahkan data ke spreadsheet
  $sheet->setCellValue('A1', 'No');
  $sheet->setCellValue('B1', 'Nama');
  $sheet->setCellValue('C1', 'Email');

  $dummyData = [
    ['1', 'John Doe', 'john.doe@example.com'],
    ['2', 'Jane Smith', 'jane.smith@example.com'],
    ['3', 'Alice Brown', 'alice.brown@example.com']
  ];

  $row = 2; // Mulai dari baris kedua
  foreach ($dummyData as $data) {
    $sheet->setCellValue('A' . $row, $data[0]);
    $sheet->setCellValue('B' . $row, $data[1]);
    $sheet->setCellValue('C' . $row, $data[2]);
    $row++;
  }

  // 3. Export data ke file Excel
  $fileName = 'dummy-data.xlsx';
  $writer = new Xlsx($spreadsheet);
  $writer->save($fileName);

  echo "File Excel berhasil dibuat: <a href='$fileName' download>Download</a>";
} catch (Exception $e) {
  echo "Terjadi kesalahan: " . $e->getMessage();
}
