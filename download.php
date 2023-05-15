<?php
require 'path_to_phpspreadsheet_library/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;

// MySQL database connection details
$servername = "localhost";
$username = "your_username";
$password = "your_password";
$dbname = "your_database";

// Create a new Spreadsheet object
$spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();

// Establish MySQL database connection
$conn = new mysqli($servername, $username, $password, $dbname);
if ($conn->connect_error) {
    die("Connection failed: " . $conn->connect_error);
}

// Retrieve data from the MySQL table
$sql = "SELECT * FROM your_table_name";
$result = $conn->query($sql);

// Set the column headers in the Excel file
$columnHeaders = array('Column 1', 'Column 2', 'Column 3');
$sheet->fromArray($columnHeaders, NULL, 'A1');

// Set the data rows in the Excel file
if ($result->num_rows > 0) {
    $row = 2;
    while ($data = $result->fetch_assoc()) {
        $rowData = array($data['column1'], $data['column2'], $data['column3']);
        $sheet->fromArray($rowData, NULL, 'A' . $row);
        $row++;
    }
}

// Close the MySQL connection
$conn->close();

// Set the appropriate headers for the download
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="table_data.xlsx"');
header('Cache-Control: max-age=0');

// Save the Excel file to the output
$writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
$writer->save('php://output');
