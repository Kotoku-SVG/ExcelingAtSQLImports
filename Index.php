<!DOCTYPE html>
<html>
<head>
    <title>Excel Upload and Download</title>
</head>
<body>
    <h2>Excel Upload and Download</h2>

    <!-- Form for uploading Excel file -->
    <form method="POST" enctype="multipart/form-data">
        <input type="file" name="excel_file" required>
        <input type="submit" name="upload" value="Upload">
    </form>

    <?php
    // Check if the form is submitted
    if (isset($_POST['upload'])) {
        $targetDir = "uploads/"; // Directory to store uploaded files

        // Check if the uploads directory exists, if not, create it
        if (!file_exists($targetDir)) {
            mkdir($targetDir, 0777, true);
        }

        // Generate a unique filename
        $targetFile = $targetDir . uniqid() . '_' . basename($_FILES['excel_file']['name']);
        $uploadOk = 1;
        $fileType = strtolower(pathinfo($targetFile, PATHINFO_EXTENSION));

        // Check if the uploaded file is a valid Excel file
        if ($fileType != "xls" && $fileType != "xlsx") {
            echo "<p>Only Excel files are allowed.</p>";
            $uploadOk = 0;
        }

        // Move the uploaded file to the target directory
        if ($uploadOk) {
            if (move_uploaded_file($_FILES['excel_file']['tmp_name'], $targetFile)) {
                echo "<p>File uploaded successfully.</p>";

                // Load the Excel file
                require 'path_to_phpspreadsheet_library/vendor/autoload.php';
                use PhpOffice\PhpSpreadsheet\IOFactory;

                $spreadsheet = IOFactory::load($targetFile);
                $worksheet = $spreadsheet->getActiveSheet();

                // Establish MySQL database connection
                $servername = "localhost";
                $username = "your_username";
                $password = "your_password";
                $dbname = "your_database";
                $conn = new mysqli($servername, $username, $password, $dbname);
                if ($conn->connect_error) {
                    die("Connection failed: " . $conn->connect_error);
                }

                // Retrieve the highest row and column numbers
                $highestRow = $worksheet->getHighestRow();
                $highestColumn = $worksheet->getHighestColumn();

                // Loop through each row of the worksheet
                for ($row = 1; $row <= $highestRow; $row++) {
                    $rowData = $worksheet->rangeToArray('A' . $row . ':' . $highestColumn . $row, null, true, false);

                    // Prepare the values for MySQL insertion
                    $values = implode("','", $rowData[0]);
                    $values = "'" . $values . "'";

                    // Build the MySQL query to check if the row exists
                    $checkQuery = "SELECT * FROM your_table_name WHERE column1 = '{$rowData[0][0]}'";

                    // Execute the query
                    $result = $conn->query($checkQuery);

                    // If the row exists, update the existing row
                    if ($result->num_rows > 0) {
                        $updateQuery = "UPDATE your_table_name SET column2 = '{$rowData[0][1]}', column3 = '{$rowData[0][2]}' WHERE column1 = '{$rowData[0][0]}'";
                        $conn->query($updateQuery);
                    } else {
                        // If the row doesn't exist, insert a new row
$insertQuery = "INSERT INTO your_table_name (column1, column2, column3) VALUES (" . $values . ")";
$conn->query($insertQuery);
}
}

// Close the MySQL connection
$conn->close();

echo "<p>Excel data appended successfully to the MySQL table.</p>";

// Provide a download link to retrieve the table data in Excel format
echo "<p><a href='download.php'>Download Excel File</a></p>";
}
?>

</body>
</html>

