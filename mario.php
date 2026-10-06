<?php

require "vendor/autoload.php";

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Style\Fill;

$spreadsheet = new Spreadsheet();
$activeWorksheet = $spreadsheet->getActiveSheet();

$activeWorksheet->getStyle('A1')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('B1')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('C1')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('D1')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('E1')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE52521');
$activeWorksheet->getStyle('F1')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE52521');
$activeWorksheet->getStyle('G1')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE52521');
$activeWorksheet->getStyle('H1')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE52521');
$activeWorksheet->getStyle('I1')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE52521');
$activeWorksheet->getStyle('J1')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('K1')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('L1')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('M1')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');

$activeWorksheet->getStyle('A2')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('B2')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('C2')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('D2')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE52521');
$activeWorksheet->getStyle('E2')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE52521');
$activeWorksheet->getStyle('F2')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE52521');
$activeWorksheet->getStyle('G2')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE52521');
$activeWorksheet->getStyle('H2')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE52521');
$activeWorksheet->getStyle('I2')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE52521');
$activeWorksheet->getStyle('J2')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE52521');
$activeWorksheet->getStyle('K2')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE52521');
$activeWorksheet->getStyle('L2')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE52521');
$activeWorksheet->getStyle('M2')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');

$activeWorksheet->getStyle('A3')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('B3')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('C3')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('D3')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF6B3B0A');
$activeWorksheet->getStyle('E3')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF6B3B0A');
$activeWorksheet->getStyle('F3')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF6B3B0A');
$activeWorksheet->getStyle('G3')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF8B878');
$activeWorksheet->getStyle('H3')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF8B878');
$activeWorksheet->getStyle('I3')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF6B3B0A');
$activeWorksheet->getStyle('J3')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF8B878');
$activeWorksheet->getStyle('K3')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('L3')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('M3')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');

$activeWorksheet->getStyle('A4')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('B4')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('C4')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF6B3B0A');
$activeWorksheet->getStyle('D4')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF8B878');
$activeWorksheet->getStyle('E4')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF6B3B0A');
$activeWorksheet->getStyle('F4')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF8B878');
$activeWorksheet->getStyle('G4')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF8B878');
$activeWorksheet->getStyle('H4')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF8B878');
$activeWorksheet->getStyle('I4')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF6B3B0A');
$activeWorksheet->getStyle('J4')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF8B878');
$activeWorksheet->getStyle('K4')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF8B878');
$activeWorksheet->getStyle('L4')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF8B878');
$activeWorksheet->getStyle('M4')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');

$activeWorksheet->getStyle('A5')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('B5')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('C5')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF6B3B0A');
$activeWorksheet->getStyle('D5')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF8B878');
$activeWorksheet->getStyle('E5')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF6B3B0A');
$activeWorksheet->getStyle('F5')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF6B3B0A');
$activeWorksheet->getStyle('G5')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF8B878');
$activeWorksheet->getStyle('H5')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF8B878');
$activeWorksheet->getStyle('I5')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF8B878');
$activeWorksheet->getStyle('J5')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF6B3B0A');
$activeWorksheet->getStyle('K5')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF8B878');
$activeWorksheet->getStyle('L5')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF8B878');
$activeWorksheet->getStyle('M5')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF8B878');

$activeWorksheet->getStyle('A6')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('B6')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('C6')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF6B3B0A');
$activeWorksheet->getStyle('D6')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF6B3B0A');
$activeWorksheet->getStyle('E6')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF8B878');
$activeWorksheet->getStyle('F6')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF8B878');
$activeWorksheet->getStyle('G6')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF8B878');
$activeWorksheet->getStyle('H6')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF8B878');
$activeWorksheet->getStyle('I6')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF6B3B0A');
$activeWorksheet->getStyle('J6')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF6B3B0A');
$activeWorksheet->getStyle('K6')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF6B3B0A');
$activeWorksheet->getStyle('L6')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF6B3B0A');
$activeWorksheet->getStyle('M6')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');

$activeWorksheet->getStyle('A7')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('B7')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('C7')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('D7')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('E7')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF8B878');
$activeWorksheet->getStyle('F7')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF8B878');
$activeWorksheet->getStyle('G7')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF8B878');
$activeWorksheet->getStyle('H7')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF8B878');
$activeWorksheet->getStyle('I7')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF8B878');
$activeWorksheet->getStyle('J7')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF8B878');
$activeWorksheet->getStyle('K7')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF8B878');
$activeWorksheet->getStyle('L7')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('M7')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');

$activeWorksheet->getStyle('A8')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('B8')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('C8')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('D8')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE52521');
$activeWorksheet->getStyle('E8')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE52521');
$activeWorksheet->getStyle('F8')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF0050C8');
$activeWorksheet->getStyle('G8')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE52521');
$activeWorksheet->getStyle('H8')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE52521');
$activeWorksheet->getStyle('I8')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE52521');
$activeWorksheet->getStyle('J8')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('K8')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('L8')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('M8')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');

$activeWorksheet->getStyle('A9')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('B9')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('C9')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE52521');
$activeWorksheet->getStyle('D9')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE52521');
$activeWorksheet->getStyle('E9')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE52521');
$activeWorksheet->getStyle('F9')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF0050C8');
$activeWorksheet->getStyle('G9')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE52521');
$activeWorksheet->getStyle('H9')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE52521');
$activeWorksheet->getStyle('I9')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF0050C8');
$activeWorksheet->getStyle('J9')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE52521');
$activeWorksheet->getStyle('K9')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE52521');
$activeWorksheet->getStyle('L9')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE52521');
$activeWorksheet->getStyle('M9')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');

$activeWorksheet->getStyle('A10')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('B10')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE52521');
$activeWorksheet->getStyle('C10')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE52521');
$activeWorksheet->getStyle('D10')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE52521');
$activeWorksheet->getStyle('E10')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE52521');
$activeWorksheet->getStyle('F10')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF0050C8');
$activeWorksheet->getStyle('G10')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF0050C8');
$activeWorksheet->getStyle('H10')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF0050C8');
$activeWorksheet->getStyle('I10')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF0050C8');
$activeWorksheet->getStyle('J10')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE52521');
$activeWorksheet->getStyle('K10')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE52521');
$activeWorksheet->getStyle('L10')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE52521');
$activeWorksheet->getStyle('M10')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE52521');

$activeWorksheet->getStyle('A11')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('B11')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF8B878');
$activeWorksheet->getStyle('C11')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF8B878');
$activeWorksheet->getStyle('D11')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE52521');
$activeWorksheet->getStyle('E11')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF0050C8');
$activeWorksheet->getStyle('F11')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFD700');
$activeWorksheet->getStyle('G11')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF0050C8');
$activeWorksheet->getStyle('H11')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF0050C8');
$activeWorksheet->getStyle('I11')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFD700');
$activeWorksheet->getStyle('J11')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF0050C8');
$activeWorksheet->getStyle('K11')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE52521');
$activeWorksheet->getStyle('L11')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF8B878');
$activeWorksheet->getStyle('M11')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF8B878');

$activeWorksheet->getStyle('A12')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('B12')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF8B878');
$activeWorksheet->getStyle('C12')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF8B878');
$activeWorksheet->getStyle('D12')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF8B878');
$activeWorksheet->getStyle('E12')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF0050C8');
$activeWorksheet->getStyle('F12')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF0050C8');
$activeWorksheet->getStyle('G12')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF0050C8');
$activeWorksheet->getStyle('H12')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF0050C8');
$activeWorksheet->getStyle('I12')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF0050C8');
$activeWorksheet->getStyle('J12')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF0050C8');
$activeWorksheet->getStyle('K12')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF8B878');
$activeWorksheet->getStyle('L12')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF8B878');
$activeWorksheet->getStyle('M12')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF8B878');

$activeWorksheet->getStyle('A13')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('B13')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF8B878');
$activeWorksheet->getStyle('C13')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF8B878');
$activeWorksheet->getStyle('D13')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF0050C8');
$activeWorksheet->getStyle('E13')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF0050C8');
$activeWorksheet->getStyle('F13')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF0050C8');
$activeWorksheet->getStyle('G13')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF0050C8');
$activeWorksheet->getStyle('H13')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF0050C8');
$activeWorksheet->getStyle('I13')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF0050C8');
$activeWorksheet->getStyle('J13')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF0050C8');
$activeWorksheet->getStyle('K13')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF0050C8');
$activeWorksheet->getStyle('L13')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF8B878');
$activeWorksheet->getStyle('M13')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF8B878');

$activeWorksheet->getStyle('A14')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('B14')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('C14')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('D14')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF0050C8');
$activeWorksheet->getStyle('E14')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF0050C8');
$activeWorksheet->getStyle('F14')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF0050C8');
$activeWorksheet->getStyle('G14')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('H14')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('I14')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF0050C8');
$activeWorksheet->getStyle('J14')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF0050C8');
$activeWorksheet->getStyle('K14')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF0050C8');
$activeWorksheet->getStyle('L14')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('M14')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');

$activeWorksheet->getStyle('A15')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('B15')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('C15')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF6B3B0A');
$activeWorksheet->getStyle('D15')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF6B3B0A');
$activeWorksheet->getStyle('E15')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF6B3B0A');
$activeWorksheet->getStyle('F15')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('G15')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('H15')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('I15')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('J15')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF6B3B0A');
$activeWorksheet->getStyle('K15')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF6B3B0A');
$activeWorksheet->getStyle('L15')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF6B3B0A');
$activeWorksheet->getStyle('M15')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');

$activeWorksheet->getStyle('A16')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('B16')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF6B3B0A');
$activeWorksheet->getStyle('C16')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF6B3B0A');
$activeWorksheet->getStyle('D16')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF6B3B0A');
$activeWorksheet->getStyle('E16')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF6B3B0A');
$activeWorksheet->getStyle('F16')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('G16')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('H16')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('I16')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFFFF');
$activeWorksheet->getStyle('J16')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF6B3B0A');
$activeWorksheet->getStyle('K16')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF6B3B0A');
$activeWorksheet->getStyle('L16')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF6B3B0A');
$activeWorksheet->getStyle('M16')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF6B3B0A');

for ($column = 1; $column <= 13; $column++) {
    $activeWorksheet->getColumnDimension(Coordinate::stringFromColumnIndex($column))->setWidth(4);
}

for ($row = 1; $row <= 16; $row++) {
    $activeWorksheet->getRowDimension($row)->setRowHeight(24);
}

$activeWorksheet->setShowGridlines(false);

if (ob_get_length()) { ob_end_clean(); }

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment; filename="mario_pixel_art.xlsx"');
header('Cache-Control: max-age=0');

$writer = new Xlsx($spreadsheet);
$writer->save('php://output');

exit; 