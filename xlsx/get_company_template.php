<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
require_once('/var/www/classes/class.XRDB.php');

$spreadsheet = new Spreadsheet();
$spreadsheet->getActiveSheet()->setTitle('Company');
$worksheet1 = $spreadsheet->createSheet();
$worksheet1->setTitle('Plans');
$worksheet2 = $spreadsheet->createSheet();
$worksheet2->setTitle('Enrollment');
$worksheet3 = $spreadsheet->createSheet();
$worksheet3->setTitle('Additions');
$worksheet4 = $spreadsheet->createSheet();
$worksheet4->setTitle('Terminations');
$sheet = $spreadsheet->getActiveSheet();

ini_set('display_errors',1);
ini_set('display_startup_errors',1);
//-------------------------------------------------------------------------------
//  Styles
//-------------------------------------------------------------------------------
//
$styleArrayEnrollment = [
    'font' => [
        'bold' => true,
    ],
    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT,
        'height' => '10',
    ],
    'borders' => [
        'top' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],
        'bottom' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],
        'left' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],
        'right' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],
    ],
    'fill' => [
        'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
    ],
];

$styleArrayBold = [
    'font' => [
        'bold' => true,
    ],
    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT,
    ],
    'borders' => [
        'top' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],
        'bottom' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],
        'left' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],
        'right' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],
    ],
    'fill' => [
        'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
    ],
];

$styleArrayHead = [
    'font' => [
        'bold' => true,
        'size' => '24',
    ],
];

$styleArray = [
    'font' => [
        'bold' => true,
        'size' => '12',
    ],
    'borders' => [
        'top' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],
        'bottom' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],
        'left' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],
        'right' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],
    ],
    'fill' => [
        'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
    ],
];
$X=new XRDB();

if (isset($argv[1])) { $_GET['id']=$argv[1]; $_GET['display']="F"; }
if (isset($argv[2])) { $_GET['month_id']=$argv[2]; $_GET['display']="F"; }
if (isset($argv[3])) { $_GET['display']=$argv[3]; }

if (isset($_GET['id'])) {
     $id = $_GET['id'];
     $company_id = $_GET['id'];
} else {
    die();
}
if (isset($_GET['month_id'])) {
     $month_id = $_GET['month_id'];
} else {
     $month_id = "2022-04";
}
if (isset($_GET['display'])) {
     $display = $_GET['display'];
} else {
     $display = "B";
}

        if ($company_id=="ALL") {
                $sql="select * from nua_company order by id";
                if ($display=='B') $display='F';
        } else {
                $sql="select * from nua_company where id = " . $company_id;
        }
        $list=$X->sql($sql);
        foreach($list as $bbb) {
                 $company = $bbb;
                 $company_id = $bbb['id'];

                 //-- Column Widths
                 $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(32);
                 $spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(42);
                 $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(8);
                 $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(42);
                 $spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(42);
                 //-- Company Name Title
                 //
                 $spreadsheet->getActiveSheet()->getStyle('A1')->applyFromArray($styleArrayHead);
                 $spreadsheet->getActiveSheet()->setCellValue('A1', $company['company_name']);

                 //-- Company ID
                 $sheet->setCellValue('A3', "ID");
                 $spreadsheet->getActiveSheet()->getStyle('B3')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('B3')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 $spreadsheet->getActiveSheet()->setCellValue('B3', $company['id']);

                //-- Company Name
                 $sheet->setCellValue('A4', "Company Name");
                 $spreadsheet->getActiveSheet()->getStyle('B4')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('B4')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 $spreadsheet->getActiveSheet()->setCellValue('B4', $company['company_name']);

                 //-- Broker Name
                 $sheet->setCellValue('A5', "Broker Name");
                 $spreadsheet->getActiveSheet()->getStyle('B5')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('B5')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 $spreadsheet->getActiveSheet()->setCellValue('B5', $company['broker_name']);

                 //-- Broker Email
                 $sheet->setCellValue('A6', "Broker Email");
                 $spreadsheet->getActiveSheet()->getStyle('B6')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('B6')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 $spreadsheet->getActiveSheet()->setCellValue('B6', $company['broker_email']);

                 //-- Company Type
                 $sheet->setCellValue('A7', "Company Type");
                 $spreadsheet->getActiveSheet()->getStyle('B7')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('B7')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 $spreadsheet->getActiveSheet()->setCellValue('B7', $company['company_type']);

                 //-- Tax ID
                 $sheet->setCellValue('A8', "Tax ID");
                 $spreadsheet->getActiveSheet()->getStyle('B8')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('B8')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 $sheet->setCellValue('B8', $company['tax_id']);

                 //-- Contact Name
                 $sheet->setCellValue('A9', "Contact Name");
                 $spreadsheet->getActiveSheet()->getStyle('B9')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('B9')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 $sheet->setCellValue('B9', $company['contact_phone']);

                 //-- Contact Email
                 $sheet->setCellValue('A10', "Contact Email");
                 $spreadsheet->getActiveSheet()->getStyle('B10')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('B10')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 $sheet->setCellValue('B10', $company['contact_email']);

                 //-- Medical
                 $sheet->setCellValue('A12', "Medical (Yes/No)");
                 $spreadsheet->getActiveSheet()->getStyle('B12')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('B12')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 $sheet->setCellValue('B12', $company['medical']);
		 //
                 //-- Dental
                 $sheet->setCellValue('A13', "Dental (Yes/No)");
                 $spreadsheet->getActiveSheet()->getStyle('B13')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('B13')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 $sheet->setCellValue('B13', $company['dental']);
		 //
                 //-- Vision
                 $sheet->setCellValue('A14', "Vision (Yes/No)");
                 $spreadsheet->getActiveSheet()->getStyle('B14')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('B14')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 $sheet->setCellValue('B14', $company['vision']);
		 //
                 //-- Provider
                 $sheet->setCellValue('A16', "Previous Insurance Provider");
                 $spreadsheet->getActiveSheet()->getStyle('B16')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('B16')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 $sheet->setCellValue('B16', $company['current_insurance_provider']);
		 //
                 //-- Contact Email
                 $sheet->setCellValue('D3', "Company Address");
                 $spreadsheet->getActiveSheet()->getStyle('E3')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E3')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 $sheet->setCellValue('E3', $company['address']);

                 //-- Suite
                 $sheet->setCellValue('D4', "Line 2");
                 $spreadsheet->getActiveSheet()->getStyle('E4')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E4')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 $sheet->setCellValue('E4', $company['suite']);

                 //-- City
                 $sheet->setCellValue('D5', "City");
                 $spreadsheet->getActiveSheet()->getStyle('E5')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E5')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 $sheet->setCellValue('E5', $company['city']);

                 //-- State
                 $sheet->setCellValue('D6', "State");
                 $spreadsheet->getActiveSheet()->getStyle('E6')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E6')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 $sheet->setCellValue('E6', $company['state']);

                 //-- Zip
                 $sheet->setCellValue('D7', "Zip");
                 $spreadsheet->getActiveSheet()->getStyle('E7')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E7')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 $sheet->setCellValue('E7', $company['zip']);

                 //-- Website
                 $sheet->setCellValue('D8', "Website");
                 $spreadsheet->getActiveSheet()->getStyle('E8')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E8')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 $sheet->setCellValue('E8', $company['website']);

                 //-- Billing Address
                 $sheet->setCellValue('D10', "Billing Address");

                 //-- Company Name
                 $sheet->setCellValue('D11', "Company Name");
                 $spreadsheet->getActiveSheet()->getStyle('E11')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E11')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 $sheet->setCellValue('E11', $company['billing_company_name']);

                 //-- Mailing Address
                 $sheet->setCellValue('D12', "Address");
                 $spreadsheet->getActiveSheet()->getStyle('E12')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E12')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 $sheet->setCellValue('E12', $company['invoice_mailing_address']);

                 //-- Invoice Suite
                 $sheet->setCellValue('D12', "Suite");
                 $spreadsheet->getActiveSheet()->getStyle('E12')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E12')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 $sheet->setCellValue('E12', $company['invoice_suite']);


                 //-- Invoice City
                 $sheet->setCellValue('D12', "City");
                 $spreadsheet->getActiveSheet()->getStyle('E13')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E13')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 $sheet->setCellValue('E13', $company['invoice_city']);

                 //-- Invoice State
                 $sheet->setCellValue('D14', "State");
                 $spreadsheet->getActiveSheet()->getStyle('E14')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E14')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 $sheet->setCellValue('E14', $company['invoice_state']);

                 //-- Invoice Zip
                 //
                 $sheet->setCellValue('D15', "Zip");
                 $spreadsheet->getActiveSheet()->getStyle('E15')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E15')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 $sheet->setCellValue('E15', $company['invoice_zip']);

                 //-- Invoice Contact Name
                 $sheet->setCellValue('D17', "Billing Contact Name");
                 $spreadsheet->getActiveSheet()->getStyle('E17')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E17')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 $spreadsheet->getActiveSheet()->setCellValue('E17', $company['billing_contact_name']);

                 //-- Invoice Contact Phone
                 $sheet->setCellValue('D18', "Billing Contact Phone");
                 $spreadsheet->getActiveSheet()->getStyle('E18')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E18')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 $sheet->setCellValue('E18', $company['billing_contact_name']);

                 //-- Billing Contact Email
                 $sheet->setCellValue('D19', "Billing Contact Email");
                 $spreadsheet->getActiveSheet()->getStyle('E19')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E19')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 $sheet->setCellValue('E19', $company['billing_contact_email']);

                 //-- Billing Contact Email
                 $sheet->setCellValue('D20', "Billing Contact Email 2");
                 $spreadsheet->getActiveSheet()->getStyle('E20')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E20')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 $sheet->setCellValue('E20', $company['billing_contact_email2']);

                 //-- Billing Contact Email
                 $sheet->setCellValue('D21', "Billing Contact Email 3");
                 $spreadsheet->getActiveSheet()->getStyle('E21')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E21')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 $sheet->setCellValue('E21', $company['billing_contact_email3']);

                 //-- Billing Contact Email
                 $sheet->setCellValue('D22', "Billing Contact Email 4");
                 $spreadsheet->getActiveSheet()->getStyle('E22')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E22')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 $sheet->setCellValue('E22', $company['billing_contact_email4']);


                 //-- Billing Contact Email
                 $sheet->setCellValue('D23', "Billing Contact Email 5");
                 $spreadsheet->getActiveSheet()->getStyle('E23')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E23')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 $sheet->setCellValue('E23', $company['billing_contact_email5']);
                $spreadsheet->setActiveSheetIndex(1);
                 //-- Column Widths
                 $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(32);
                 $spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(15);
                 $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(20);
                 $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(20);
                 $spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(20);
                 $spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(20);
                 //-- Company Name Title
                 //
                 $spreadsheet->getActiveSheet()->getStyle('A1')->applyFromArray($styleArrayHead);
                 $spreadsheet->getActiveSheet()->setCellValue('A1', "Company Plans");
                 $spreadsheet->getActiveSheet()->getStyle('A2')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('A2')->getFill()->getStartColor()->setARGB('FFFFCC66');
                 $spreadsheet->getActiveSheet()->setCellValue('A2', "Plan Name");
                 $spreadsheet->getActiveSheet()->getStyle('B2')->applyFromArray($styleArrayBold);
		$spreadsheet->getActiveSheet()->getStyle('B2')->getFill()->getStartColor()->setARGB('FFFFCC66');
                 $spreadsheet->getActiveSheet()->setCellValue('B2', "APA");
                 $spreadsheet->getActiveSheet()->getStyle('C2')->applyFromArray($styleArrayBold);
		$spreadsheet->getActiveSheet()->getStyle('C2')->getFill()->getStartColor()->setARGB('FFFFCC66');
                 $spreadsheet->getActiveSheet()->setCellValue('C2', "EE Price");
                 $spreadsheet->getActiveSheet()->getStyle('D2')->applyFromArray($styleArrayBold);
		$spreadsheet->getActiveSheet()->getStyle('D2')->getFill()->getStartColor()->setARGB('FFFFCC66');
                 $spreadsheet->getActiveSheet()->setCellValue('D2', "EES Price");
                 $spreadsheet->getActiveSheet()->getStyle('E2')->applyFromArray($styleArrayBold);
		$spreadsheet->getActiveSheet()->getStyle('E2')->getFill()->getStartColor()->setARGB('FFFFCC66');
                 $spreadsheet->getActiveSheet()->setCellValue('E2', "EEC Price");
                 $spreadsheet->getActiveSheet()->getStyle('F2')->applyFromArray($styleArrayBold);
		$spreadsheet->getActiveSheet()->getStyle('F2')->getFill()->getStartColor()->setARGB('FFFFCC66');
                 $spreadsheet->getActiveSheet()->setCellValue('F2', "FAM Price");
$sql="select * from nua_company_plan where company_id = " . $company_id . " order by plan_type, plan_code";
$p=$X->sql($sql);
$row=2;
foreach ($p as $q) {
      $row++;
      $cell="A".$row;
      $spreadsheet->getActiveSheet()->setCellValue($cell, $q['plan_code']);
      $cell="B".$row;
      $spreadsheet->getActiveSheet()->setCellValue($cell, $q['APA_CODE']);
      $cell="C".$row;
      $spreadsheet->getActiveSheet()->setCellValue($cell, $q['ee_price']);
      $cell="D".$row;
      $spreadsheet->getActiveSheet()->setCellValue($cell, $q['ees_price']);
      $cell="E".$row;
      $spreadsheet->getActiveSheet()->setCellValue($cell, $q['eec_price']);
      $cell="F".$row;
      $spreadsheet->getActiveSheet()->setCellValue($cell, $q['fam_price']);


}

                $spreadsheet->setActiveSheetIndex(2);

				$styleArrayEnrollment = [
				'font' => [
				'bold' => true,
				],
				'alignment' => [
				'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
				],
				'borders' => [
					'top' => [
						'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
				],
				'bottom' => [
					'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
				],
				'left' => [
					'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
				],
				'right' => [
					'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
				],
			],
			'fill' => [
				'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
			],
		];
		$spreadsheet->getActiveSheet()->getRowDimension('1')->setRowHeight(100);
        $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('A1')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('A1')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('A1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('A1', "Eff Date");
        $spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('B1')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('B1')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('B1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('B1', "Employee Code");
        $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(15);
        $spreadsheet->getActiveSheet()->getStyle('C1')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('C1')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('C1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('C1', "Dependent Code");
        $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('D1')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('D1')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('D1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('D1', "Social Security Number");
        $spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('E1')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('E1')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('E1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('E1', "Last Name");
        $spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('F1')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('F1')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('F1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('F1', "First Name");
        $spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('G1')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('G1')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('G1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('G1', "Middle Initial");
        $spreadsheet->getActiveSheet()->getColumnDimension('H')->setWidth(15);
        $spreadsheet->getActiveSheet()->getStyle('H1')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('H1')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('H1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('H1', "Relationship (Employee, Husband, Wife Son, Daughter)");
        $spreadsheet->getActiveSheet()->getColumnDimension('I')->setWidth(15);
        $spreadsheet->getActiveSheet()->getStyle('I1')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('I1')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('I1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('I1', "Date of Birth");
        $spreadsheet->getActiveSheet()->getColumnDimension('J')->setWidth(10);
        $spreadsheet->getActiveSheet()->getStyle('J1')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('J1')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('J1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('J1', "Gender (M/F)");
        $spreadsheet->getActiveSheet()->getColumnDimension('K')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('K1')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('K1')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('K1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('K1', "Marital Status");
        $spreadsheet->getActiveSheet()->getColumnDimension('L')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('L1')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('L1')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('L1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('L1', "Address");
        $spreadsheet->getActiveSheet()->getColumnDimension('M')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('M1')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('M1')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('M1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('M1', "Suite / Apt");
        $spreadsheet->getActiveSheet()->getColumnDimension('N')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('N1')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('N1')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('N1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('N1', "City");
        $spreadsheet->getActiveSheet()->getColumnDimension('O')->setWidth(10);
        $spreadsheet->getActiveSheet()->getStyle('O1')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('O1')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('O1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('O1', "State");
        $spreadsheet->getActiveSheet()->getColumnDimension('P')->setWidth(10);
        $spreadsheet->getActiveSheet()->getStyle('P1')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('P1')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('P1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('P1', "Zipcode");
        $spreadsheet->getActiveSheet()->getColumnDimension('Q')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('Q1')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('Q1')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('Q1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('Q1', "Email (Provide a personal email)");
        $spreadsheet->getActiveSheet()->getColumnDimension('R')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('R1')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('R1')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('R1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('R1', "Contact Phone (Required if no email)");
        $spreadsheet->getActiveSheet()->getColumnDimension('S')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('S1')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('S1')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('S1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('S1', "Date of Hire");
        $spreadsheet->getActiveSheet()->getColumnDimension('T')->setWidth(10);
        $spreadsheet->getActiveSheet()->getStyle('T1')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('T1')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('T1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('T1', "Work Status");
        $spreadsheet->getActiveSheet()->getColumnDimension('U')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('U1')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('U1')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('U1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('U1', "Effective Date");


// MEDICAL

	$spreadsheet->getActiveSheet()->getColumnDimension('V')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('V1')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('V1')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('V1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('V1', "Medical Plan");

        $spreadsheet->getActiveSheet()->getColumnDimension('W')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('W1')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('W1')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('W1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('W1', "Coverage Election (EE, ES, EC, FAM)");

        $spreadsheet->getActiveSheet()->getColumnDimension('X')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('X1')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('X1')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('X1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('X1', "APA Member ID");

        $spreadsheet->getActiveSheet()->getColumnDimension('Y')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('Y1')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('Y1')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('Y1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('Y1', "Medical Price");

// DENTAL

		$spreadsheet->getActiveSheet()->getColumnDimension('Z')->setWidth(20);
                $spreadsheet->getActiveSheet()->getStyle('Z1')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('Z1')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('Z1')->getAlignment()->setWrapText(true);
		$spreadsheet->getActiveSheet()->setCellValue('Z1', "Dental Plan (Blank if none)");

                $spreadsheet->getActiveSheet()->getColumnDimension('AA')->setWidth(20);
                $spreadsheet->getActiveSheet()->getStyle('AA1')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('AA1')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('AA1')->getAlignment()->setWrapText(true);
                $spreadsheet->getActiveSheet()->setCellValue('AA1', "Dental Coverage (EE,ES,EC,FAM)");

                $spreadsheet->getActiveSheet()->getColumnDimension('AB')->setWidth(20);
                $spreadsheet->getActiveSheet()->getStyle('AB1')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('AB1')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('AB1')->getAlignment()->setWrapText(true);
                $spreadsheet->getActiveSheet()->setCellValue('AB1', "Guardian Member ID");

                $spreadsheet->getActiveSheet()->getColumnDimension('AC')->setWidth(20);
                $spreadsheet->getActiveSheet()->getStyle('AC1')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('AC1')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('AC1')->getAlignment()->setWrapText(true);
                $spreadsheet->getActiveSheet()->setCellValue('AC1', "Dental Price");

// VISION
//
                $spreadsheet->getActiveSheet()->getColumnDimension('AD')->setWidth(20);
                $spreadsheet->getActiveSheet()->getStyle('AD1')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('AD1')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('AD1')->getAlignment()->setWrapText(true);
                $spreadsheet->getActiveSheet()->setCellValue('AD1', "Vision Plan");

                $spreadsheet->getActiveSheet()->getColumnDimension('AE')->setWidth(20);
                $spreadsheet->getActiveSheet()->getStyle('AE1')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('AE1')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('AE1')->getAlignment()->setWrapText(true);
                $spreadsheet->getActiveSheet()->setCellValue('AE1', "Vision Coverage");

                $spreadsheet->getActiveSheet()->getColumnDimension('AF')->setWidth(20);
                $spreadsheet->getActiveSheet()->getStyle('AF1')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('AF1')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('AF1')->getAlignment()->setWrapText(true);
                $spreadsheet->getActiveSheet()->setCellValue('AF1', "Vision Price");

                $spreadsheet->getActiveSheet()->getColumnDimension('AG')->setWidth(20);
                $spreadsheet->getActiveSheet()->getStyle('AG1')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('AG1')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('AG1')->getAlignment()->setWrapText(true);
                $spreadsheet->getActiveSheet()->setCellValue('AG1', "Action (ADMIN)");
        $sql="select distinct employee_code, employee_id, eff_dt, term_dt, last_name, first_name from nua_monthly_member_census where company_id = " . $company_id . " and month_id = '" . $month_id . "' and dependent_code = '' order by last_name, first_name";
        $d=$X->sql($sql);
        $row=1;
        foreach($d as $e) {
            $sql="select * from nua_employee where id = " . $e['employee_id'];
            $x=$X->sql($sql);
	    if (sizeof($x)==0) {
		  print_r($e);
                  echo $sql;
	     }

            $employee=$x[0];
            $row++;
            $cell="A".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $month_id);     
            $cell="B".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $e['employee_code']);
            $cell="D".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['social_security_number']);
            $cell="E".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['last_name']);
            $cell="F".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['first_name']);
            $cell="G".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['middle_name']);
            $cell="H".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, "Employee");
            $cell="I".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['date_of_birth']);
            $cell="J".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['gender']);
            $cell="K".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['marital_status']);
            $cell="L".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['address']);
            $cell="M".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['suite']);
            $cell="N".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['city']);
            $cell="O".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['state']);
            $cell="P".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['zip']);
            $cell="Q".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['email']);
            $cell="R".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['phone']);
            $cell="S".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['date_hired']);
            $cell="T".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['work_status']);
            $cell="U".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $e['eff_dt']);
            $sql="select * from nua_monthly_member_census where company_id = " . $company_id . " and employee_code = '" . $e['employee_code'] . "' and ";
            $sql.=" dependent_code = ''";
            $hh=$X->sql($sql);
            $medical_plan="";
            $medical_coverage="";
            $dental_plan="";
            $dental_coverage="";
            $vision_plan="";
            $vision_coverage="";
$medical_price="";
$dental_price="";
$vision_price="";
            foreach($hh as $ii) {
                if ($ii['plan_type']=="*MEDICAL*" ) {
                     $medical_plan=$ii['client_plan'];
                     $medical_coverage=$ii['coverage_level'];
                     $medical_price=$ii['coverage_price'];
                }
                if ($ii['plan_type']=="*DENTAL*" ) {
                     $dental_plan=$ii['client_plan'];
                     $dental_coverage=$ii['coverage_level'];
                     $dental_price=$ii['coverage_price'];
                }
                if ($ii['plan_type']=="*VISION*" ) {
                     $vision_plan=$ii['client_plan'];
                     $vision_coverage=$ii['coverage_level'];
                     $vision_price=$ii['coverage_price'];
                }
            }
            $cell="V".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $medical_plan);
            $cell="W".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $medical_coverage);
            $cell="X".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['apa_member_id']);
            $cell="Y".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $medical_price);
            $cell="Z".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $dental_plan);
            $cell="AA".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $dental_coverage);
            $cell="AB".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['guardian_member_id']);
            $cell="AC".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $dental_price);
            $cell="AD".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $vision_plan);
            $cell="AE".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $vision_coverage);
            $cell="AF".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $vision_price);

            $sql="select * from nua_employee_dependent where employee_id = " . $e['employee_id'] . " order by dependent_id";
            $z=$X->sql($sql);
            foreach($z as $ff) {
                $row++;
                $cell="A".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $month_id);     
                $cell="B".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $e['employee_code']);
                $cell="C".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $ff['dependent_id']);
                $cell="D".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $ff['social_security_number']);
                $cell="E".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $ff['last_name']);
                $cell="F".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $ff['first_name']);
                $cell="G".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $ff['middle_name']);
                  $cell="H".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $ff['relationship']);
                $cell="I".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $ff['date_of_birth']);
                $cell="J".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $ff['gender']);
            }

	}
        $spreadsheet->setActiveSheetIndex(3);

        $spreadsheet->getActiveSheet()->getStyle('A1')->applyFromArray($styleArrayHead);
        $spreadsheet->getActiveSheet()->setCellValue('A1', "Additions");
	$spreadsheet->getActiveSheet()->getRowDimension('2')->setRowHeight(100);
        $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('A2')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('A2')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('A2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('A2', "Eff Date");
        $spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('B2')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('B2')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('B2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('B2', "ACTION (ADD / DELETE / CHANGE)");
        $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(15);
        $spreadsheet->getActiveSheet()->getStyle('C2')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('C2')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('C2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('C2', "Employee ID (if set)");
        $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('D2')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('D2')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('D2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('D2', "Social Security Number");
        $spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('E2')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('E2')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('E2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('E2', "Last Name");
        $spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('F2')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('F2')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('F2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('F2', "First Name");
        $spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('G2')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('G2')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('G2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('G2', "Middle Initial");
        $spreadsheet->getActiveSheet()->getColumnDimension('H')->setWidth(15);
        $spreadsheet->getActiveSheet()->getStyle('H2')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('H2')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('H2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('H2', "Relationship (Employee, Husband, Wife Son, Daughter)");
        $spreadsheet->getActiveSheet()->getColumnDimension('I')->setWidth(15);
        $spreadsheet->getActiveSheet()->getStyle('I2')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('I2')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('I2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('I2', "Date of Birth");
        $spreadsheet->getActiveSheet()->getColumnDimension('J')->setWidth(10);
        $spreadsheet->getActiveSheet()->getStyle('J2')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('J2')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('J2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('J2', "Gender (M/F)");
        $spreadsheet->getActiveSheet()->getColumnDimension('K')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('K2')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('K2')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('K2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('K2', "Marital Status");
        $spreadsheet->getActiveSheet()->getColumnDimension('L')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('L2')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('L2')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('L2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('L2', "Address");
        $spreadsheet->getActiveSheet()->getColumnDimension('M')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('M2')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('M2')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('M2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('M2', "Suite / Apt");
        $spreadsheet->getActiveSheet()->getColumnDimension('N')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('N2')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('N2')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('N2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('N2', "City");
        $spreadsheet->getActiveSheet()->getColumnDimension('O')->setWidth(10);
        $spreadsheet->getActiveSheet()->getStyle('O2')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('O2')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('O2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('O2', "State");
        $spreadsheet->getActiveSheet()->getColumnDimension('P')->setWidth(10);
        $spreadsheet->getActiveSheet()->getStyle('P2')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('P2')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('P2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('P2', "Zipcode");
        $spreadsheet->getActiveSheet()->getColumnDimension('Q')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('Q2')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('Q2')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('Q2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('Q2', "Email (Provide a personal email)");
        $spreadsheet->getActiveSheet()->getColumnDimension('R')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('R2')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('R2')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('R2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('R2', "Contact Phone (Required if no email)");
        $spreadsheet->getActiveSheet()->getColumnDimension('S')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('S2')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('S2')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('S2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('S2', "Date of Hire");
        $spreadsheet->getActiveSheet()->getColumnDimension('T')->setWidth(10);
        $spreadsheet->getActiveSheet()->getStyle('T2')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('T2')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('T2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('T2', "Work Status");
        $spreadsheet->getActiveSheet()->getColumnDimension('U')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('U2')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('U2')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('U2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('U2', "Effective Date");
        $spreadsheet->getActiveSheet()->getColumnDimension('V')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('V2')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('V2')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('V2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('V2', "Medical Plan");
        $spreadsheet->getActiveSheet()->getColumnDimension('W')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('W2')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('W2')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('W2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('W2', "Coverage Election (EE, ES, EC, FAM)");
        $spreadsheet->getActiveSheet()->getColumnDimension('X')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('X2')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('X2')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('X2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('X2', "Dental Plan (Blank if none)");
        $spreadsheet->getActiveSheet()->getColumnDimension('Y')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('Y2')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('Y2')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('Y2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('Y2', "Dental Coverage (EE,ES,EC,FAM)");
        $spreadsheet->getActiveSheet()->getColumnDimension('Z')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('Z2')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('Z2')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->setCellValue('Z2', "Vision Plan");
		$spreadsheet->getActiveSheet()->getStyle('Z2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('AA')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('AA2')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('AA2')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('AA2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('AA2', "Vision Coverage");

        $sql="select distinct employee_code, employee_id, eff_dt, term_dt, last_name, first_name, month_id from nua_monthly_member_additions where company_id = " . $company_id . " and month_id >= '" . $month_id . "' and dependent_code = '' order by last_name, first_name";
        $d=$X->sql($sql);
        $row=2;
        foreach($d as $e) {
            $sql="select * from nua_employee where id = " . $e['employee_id'];
            $x=$X->sql($sql);
            $employee=$x[0];
            $row++;
            $cell="A".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $e['month_id']);     
//            $cell="B".$row;
//			$spreadsheet->getActiveSheet()->setCellValue($cell, $e['employee_code']);
            $cell="C".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $e['employee_id']);
            $cell="D".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['social_security_number']);
            $cell="E".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['last_name']);
            $cell="F".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['first_name']);
            $cell="G".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['middle_name']);
            $cell="H".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, "Employee");
            $cell="I".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['date_of_birth']);
            $cell="J".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['gender']);
            $cell="K".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['marital_status']);
            $cell="L".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['address']);
            $cell="M".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['suite']);
            $cell="N".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['city']);
            $cell="O".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['state']);
            $cell="P".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['zip']);
            $cell="Q".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['email']);
            $cell="R".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['phone']);
            $cell="S".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['date_hired']);
            $cell="T".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['work_status']);
        $sql="select distinct employee_code, employee_id, eff_dt, term_dt, last_name, first_name, month_id from nua_monthly_member_terminations where company_id = " . $company_id . " and month_id >= '" . $month_id . "' and dependent_code = '' order by last_name, first_name";
            $hh=$X->sql($sql);
            $medical_plan="";
            $medical_coverage="";
            $dental_plan="";
            $dental_coverage="";
            $vision_plan="";
            $vision_coverage="";
            foreach($hh as $ii) {
                if ($ii['plan_type']=="*MEDICAL*" ) {
                     $medical_plan=$ii['client_plan'];
                     $medical_coverage=$ii['coverage_level'];
                }
                if ($ii['plan_type']=="*DENTAL*" ) {
                     $dental_plan=$ii['client_plan'];
                     $dental_coverage=$ii['coverage_level'];
                }
                if ($ii['plan_type']=="*VISION*" ) {
                     $vision_plan=$ii['client_plan'];
                     $vision_coverage=$ii['coverage_level'];
                }
            }
            $cell="U".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $medical_plan);
            $cell="V".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $medical_coverage);
            $cell="W".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $dental_plan);
            $cell="X".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $dental_coverage);
            $cell="Y".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $vision_plan);
            $cell="Z".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $vision_coverage);

          /*  $sql="select * from nua_employee_dependent where employee_id = " . $e['employee_id'] . " order by dependent_id";
            $z=$X->sql($sql);
            foreach($z as $ff) {
                $row++;
                $cell="A".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $e['month_id']);     
                $cell="C".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $ff['employee_id']);
                $cell="D".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $ff['social_security_number']);
                $cell="E".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $ff['last_name']);
                $cell="F".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $ff['first_name']);
                $cell="G".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $ff['middle_name']);
                  $cell="H".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $ff['relationship']);
                $cell="I".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $ff['date_of_birth']);
                $cell="J".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $ff['gender']);
            }
*/
	}

        $spreadsheet->setActiveSheetIndex(4);

        $spreadsheet->getActiveSheet()->getStyle('A1')->applyFromArray($styleArrayHead);
        $spreadsheet->getActiveSheet()->setCellValue('A1', "Terminations");
	$spreadsheet->getActiveSheet()->getRowDimension('2')->setRowHeight(100);
        $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('A2')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('A2')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('A2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('A2', "Last Month");
        $spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('B2')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('B2')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('B2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('B2', "ACTION (ADD / DELETE / CHANGE)");
        $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(15);
        $spreadsheet->getActiveSheet()->getStyle('C2')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('C2')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('C2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('C2', "Employee ID (if set)");
        $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('D2')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('D2')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('D2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('D2', "Social Security Number");
        $spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('E2')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('E2')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('E2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('E2', "Last Name");
        $spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('F2')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('F2')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('F2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('F2', "First Name");
        $spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('G2')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('G2')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('G2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('G2', "Middle Initial");
        $spreadsheet->getActiveSheet()->getColumnDimension('H')->setWidth(15);
        $spreadsheet->getActiveSheet()->getStyle('H2')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('H2')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('H2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('H2', "Relationship (Employee, Husband, Wife Son, Daughter)");
        $spreadsheet->getActiveSheet()->getColumnDimension('I')->setWidth(15);
        $spreadsheet->getActiveSheet()->getStyle('I2')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('I2')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('I2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('I2', "Date of Birth");
        $spreadsheet->getActiveSheet()->getColumnDimension('J')->setWidth(10);
        $spreadsheet->getActiveSheet()->getStyle('J2')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('J2')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('J2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('J2', "Gender (M/F)");
        $spreadsheet->getActiveSheet()->getColumnDimension('K')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('K2')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('K2')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('K2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('K2', "Marital Status");
        $spreadsheet->getActiveSheet()->getColumnDimension('L')->setWidth(20);
                $spreadsheet->getActiveSheet()->getStyle('L2')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('L2')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('L2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('L2', "Effective Date");
        $spreadsheet->getActiveSheet()->getColumnDimension('L')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('M2')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('M2')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('M2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('M2', "Term Date");
        $spreadsheet->getActiveSheet()->getColumnDimension('N')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('N2')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('N2')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('N2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('N2', "Medical Plan");
        $spreadsheet->getActiveSheet()->getColumnDimension('O')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('O2')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('O2')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('O2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('O2', "Coverage Election (EE, ES, EC, FAM)");
        $spreadsheet->getActiveSheet()->getColumnDimension('P')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('P2')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('P2')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('P2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('P2', "Dental Plan (Blank if none)");
        $spreadsheet->getActiveSheet()->getColumnDimension('Q')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('Q2')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('Q2')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('Q2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('Q2', "Dental Coverage (EE,ES,EC,FAM)");
        $spreadsheet->getActiveSheet()->getColumnDimension('R')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('R2')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('R2')->getFill()->getStartColor()->setARGB('FFFFCC66');
        $spreadsheet->getActiveSheet()->setCellValue('R2', "Vision Plan");
		$spreadsheet->getActiveSheet()->getStyle('S2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('S')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('S2')->applyFromArray($styleArrayEnrollment);
		$spreadsheet->getActiveSheet()->getStyle('S2')->getFill()->getStartColor()->setARGB('FFFFCC66');
		$spreadsheet->getActiveSheet()->getStyle('S2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('S2', "Vision Coverage");

        $sql="select distinct employee_code, employee_id, eff_dt, term_dt, last_name, first_name from nua_monthly_member_terminations where company_id = " . $company_id . " and month_id = '2022-04' and dependent_code = '' order by last_name, first_name";
        $d=$X->sql($sql);
        $row=2;
        foreach($d as $e) {
            $sql="select * from nua_employee where id = " . $e['employee_id'];
            $x=$X->sql($sql);
            $employee=$x[0];
            $row++;
            $cell="A".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, "2022-04");     
//            $cell="B".$row;
//			$spreadsheet->getActiveSheet()->setCellValue($cell, $e['employee_code']);
            $cell="C".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $e['employee_id']);
            $cell="D".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['social_security_number']);
            $cell="E".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['last_name']);
            $cell="F".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['first_name']);
            $cell="G".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['middle_name']);
            $cell="H".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, "Employee");
            $cell="I".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['date_of_birth']);
            $cell="J".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['gender']);
            $cell="K".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['marital_status']);
            $cell="L".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $e['eff_dt']);
            $cell="M".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $e['term_dt']);
            $sql="select * from nua_monthly_member_census where company_id = " . $company_id . " and employee_code = '" . $e['employee_code'] . "' and ";
            $sql.=" dependent_code = ''";
            $hh=$X->sql($sql);
            $medical_plan="";
            $medical_coverage="";
            $dental_plan="";
            $dental_coverage="";
            $vision_plan="";
            $vision_coverage="";
            foreach($hh as $ii) {
                if ($ii['plan_type']=="*MEDICAL*" ) {
                     $medical_plan=$ii['client_plan'];
                     $medical_coverage=$ii['coverage_level'];
                }
                if ($ii['plan_type']=="*DENTAL*" ) {
                     $dental_plan=$ii['client_plan'];
                     $dental_coverage=$ii['coverage_level'];
                }
                if ($ii['plan_type']=="*VISION*" ) {
                     $vision_plan=$ii['client_plan'];
                     $vision_coverage=$ii['coverage_level'];
                }
            }
            $cell="N".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $medical_plan);
            $cell="O".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $medical_coverage);
            $cell="P".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $dental_plan);
            $cell="Q".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $dental_coverage);
            $cell="R".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $vision_plan);
            $cell="S".$row;
			$spreadsheet->getActiveSheet()->setCellValue($cell, $vision_coverage);

	}



                $spreadsheet->setActiveSheetIndex(0);
$writer = new Xlsx($spreadsheet);
//$writer->save('sheets/hell.xlsx');
$filename=$company_id . "_" . str_replace(' ','_',str_replace('#',' ',$company['company_name'])) . ".xlsx";

}
if ($display=="B") {
    header('Content-disposition: attachment; filename=' . $filename);
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet; charset=utf-8');
    header('Cache-Control: must-revalidate');
    sleep(1);
    $writer->save("php://output");
    die();
} 
if ($display=="F") {
    $file="sheets/" . $file;
    $writer->save($file);
}
?>


