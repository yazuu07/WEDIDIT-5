<?php
session_start();
require 'db.php';

// Include PhpSpreadsheet library
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Redirect if user is not authenticated
if (!isset($_SESSION['user_id'])) {
    header("Location: login.php");
    exit();
}

// Get user parameter
if (!isset($_GET['user'])) {
    echo "User not specified.";
    exit();
}

$username = $_GET['user'];

// Fetch user details
$stmt = $pdo->prepare("SELECT * FROM admin1 WHERE user = ?");
$stmt->execute([$username]);
$user = $stmt->fetch(PDO::FETCH_ASSOC);

if (!$user) {
    echo "User not found.";
    exit();
}

$user_id = $user['id'];

// Fetch user's images and timestamps
$stmt = $pdo->prepare("SELECT image_path, location, uploaded_at FROM uploads WHERE user_id = ? ORDER BY uploaded_at ASC");
$stmt->execute([$user_id]);
$images = $stmt->fetchAll(PDO::FETCH_ASSOC);

if (empty($images)) {
    echo "No records to export.";
    exit();
}

// Create a new Spreadsheet
$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();

// Set the header row
$headers = ['Name', 'Upload_at', 'Time In', 'Time Out', 'Remark 1', 'Location'];
$sheet->fromArray($headers, NULL, 'A1');

// Prepare data rows
$row = 2;
$lastTimeIn = null;

foreach ($images as $image) {
    $name = $username;
    $location = $image['location'];
    $uploadedAt = $image['uploaded_at'];

    // Determine Time In, Time Out, and calculate remarks
    if ($location === 'In') {
        $timeIn = $uploadedAt;
        $timeOut = '';
        $lastTimeIn = $timeIn; // Save the last "In" time
        $remark1 = '';
    } elseif ($location === 'Out' && $lastTimeIn) {
        $timeIn = $lastTimeIn;
        $timeOut = $uploadedAt;
        $lastTimeIn = null; // Reset after pairing with "Out"

        // Calculate hours worked
        $hoursWorked = (strtotime($timeOut) - strtotime($timeIn)) / 3600;

        // Assign Remark 1 based on hours worked
        if ($hoursWorked < 9) {
            $remark1 = 'Undertime';
        } elseif ($hoursWorked >= 9 && $hoursWorked <= 10) {
            $remark1 = 'On Time';
        } else {
            $remark1 = 'Overtime';
        }
    } else {
        $timeIn = '';
        $timeOut = '';
        $remark1 = '';
    }

    // Write data to the spreadsheet
    $sheet->setCellValue("A$row", $name);
    $sheet->setCellValue("B$row", $uploadedAt);
    $sheet->setCellValue("C$row", $timeIn);
    $sheet->setCellValue("D$row", $timeOut);
    $sheet->setCellValue("E$row", $remark1);
    $sheet->setCellValue("F$row", $location);

    $row++;
}

// Set column widths for better visibility
foreach (range('A', 'F') as $column) {
    $sheet->getColumnDimension($column)->setAutoSize(true);
}

// Set the filename
$filename = "Export_" . $username . "_" . date("Y-m-d") . ".xlsx";

// Send the file as a download
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="' . $filename . '"');
header('Cache-Control: max-age=0');

// Save the file to output
$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
exit();
?>
