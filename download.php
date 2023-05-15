<?php
require 'path_to_phpspreadsheet_library/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;

// MySQL database connection details
$servername = "localhost";
$username = "your_username";
$password = "your_password";
$dbname = "your_database";

// Establish MySQL database connection
$conn = new mysqli($servername, $username, $password, $dbname);
if ($conn->connect_error) {
    die("Connection failed: " . $conn->connect_error);
}

// Retrieve filter options from the form
$startDate = $_POST['start_date'];
$endDate = $_POST['end_date'];
$employeeID = $_POST['employee_id'];

// Build the SQL query based on the filter options
$sql = "SELECT * FROM your_table_name WHERE 1=1";

if (!empty($startDate)) {
    $sql .= " AND date_column >= '{$startDate}'";
}

if (!empty($endDate)) {
    $sql .= " AND date_column <= '{$endDate}'";
}

if (!empty($employeeID)) {
    $sql .= " AND employee_id_column = '{$employeeID}'";
}

// Execute the query
$result = $conn->query($sql);

// Create a new Spreadsheet object
$spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();

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
    header('Content-Disposition: attachment;filename="filtered_data.xlsx"');
    header('Cache-Control: max-age=0');

    // Save the Excel file to the output
    $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
    $writer->save('php://output');
}
?>
