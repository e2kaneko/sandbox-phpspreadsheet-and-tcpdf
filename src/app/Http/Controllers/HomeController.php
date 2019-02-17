<?php

namespace App\Http\Controllers;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Shared\File;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class HomeController extends Controller
{
    public function index()
    {
        $spreadsheet = IOFactory::load(public_path() . '/excel/template.xlsx');
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setCellValue('A3', 'サンプル株式会社');

        File::setUseUploadTempDirectory(public_path());

        $writer = IOFactory::createWriter($spreadsheet, 'Tcpdf');
        $writer->setFont('ipaexm');
        $writer->save(public_path() . '/excel/output.pdf');

        $writer = new Xlsx($spreadsheet);
        $writer->save(public_path() . '/excel/output.xlsx');
    }
}
