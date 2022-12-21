<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
require_once('/var/www/classes/class.XRDB.php');


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

$styleArrayLine = [
    'font' => [
        'bold' => true,
    ],
    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
        'height' => '10',
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

$month_id="2022-04";

$X=new XRDB();

if (isset($argv[1])) { $_GET['id']=$argv[1]; $_GET['display']="F"; }
if (isset($argv[2])) { $_GET['display']=$argv[2]; }

if (isset($_GET['id'])) {
     $id = $_GET['id'];
     $company_id = $_GET['id'];
} else {
    die();
}
if (isset($_GET['display'])) {
     $display = $_GET['display'];
} else {
     $display = "B";
}

if (isset($_GET['uid'])) {
     $uid = $_GET['uid'];
     $sql="select * from nua_user where id = " . $uid;
     $u=$X->sql($sql);
     $role=$u[0]['role'];
} else {
     $uid = "0";
     $role = "broker";
}
$role="sadmin";
if ($company_id==0) {
    $company=array();   
} else {
    $sql="select * from nua_company where id = " . $company_id;
    $t=$X->sql($sql);
    $company=$t[0];
}

//
// 0
//
//
$spreadsheet = new Spreadsheet();
$spreadsheet->getActiveSheet()->setTitle('Company');

if ($company_id!='0') {
//
// 1
//
//
   $worksheet1 = $spreadsheet->createSheet();
   $worksheet1->setTitle('Preenrollment-Census');
}

$quoted_plan="N";
$company_plan="N";

if ($company_id!='0') {
  
    //--
    //-- Only Admins and Quuoting Plan Team gets page unless its filled out.
    //--

    if ($role=="sadmin"||$role=="quoting") { 
        $quoted_plan="Y";
    } else {
        $sql="select * from nua_quoted_plan where company_id = " . $company_id;
        $z=$X->sql($sql);
        $sql="select * from nua_company_plan where company_id = " . $company_id;
        $y=$X->sql($sql);
        if (sizeof($z)>0||sizeof($y>0))  $quoted_plan="Y"; 
    }
    if ($quoted_plan=="Y") {
        //
	// 2
	// 
	//
        $worksheet2 = $spreadsheet->createSheet();
        $worksheet2->setTitle('Quoted-Plans');
    }

    $sql="select * from nua_company_plan where company_id = " . $company_id;
    $z=$X->sql($sql);
    if (sizeof($z)==0) {
        $company_plan="N";
    } else {
        $company_plan="Y";
        $worksheet3 = $spreadsheet->createSheet();
        $worksheet3->setTitle('Accepted-Plans');
    }
    if ($company_plan=="Y") {
       $worksheet4 = $spreadsheet->createSheet();
       $worksheet4->setTitle('Enrollment');
    }
}

//================================================================================
//  COMPANY TEMPLATE
//================================================================================

      $sheet = $spreadsheet->getActiveSheet();
      //-- Column Widths
      $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(32);
      $spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(42);
      $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(8);
      $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(42);
      $spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(42);
      //-- Company Name Title
      //
      $spreadsheet->getActiveSheet()->getStyle('A1')->applyFromArray($styleArrayHead);
      if ($company_id=='0') {
	      $spreadsheet->getActiveSheet()->setCellValue('A1', "NEW PROSPECT COMPANY DATA");
      }  else {
              $spreadsheet->getActiveSheet()->setCellValue('A1', $company['company_name']);
      }

      //-- Company ID
      $sheet->setCellValue('A3', "ID");
      $spreadsheet->getActiveSheet()->getStyle('B3')->applyFromArray($styleArrayBold);
      $spreadsheet->getActiveSheet()->getStyle('B3')->getFill()->getStartColor()->setARGB('DDDDDDDD');
      if ($company_id!='0') { 
            $spreadsheet->getActiveSheet()->setCellValue('B3', $company['id']);
      } else {
             $spreadsheet->getActiveSheet()->setCellValue('B3', "(NuAxess Use Only)");
      }
                //-- Company Name
                 $sheet->setCellValue('A4', "Company Name");
                 $spreadsheet->getActiveSheet()->getStyle('B4')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('B4')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $spreadsheet->getActiveSheet()->setCellValue('B4', $company['company_name']);

                 //-- Broker Name
                 $sheet->setCellValue('A5', "Broker Name");
                 $spreadsheet->getActiveSheet()->getStyle('B5')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('B5')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $spreadsheet->getActiveSheet()->setCellValue('B5', $company['broker_name']);

                 //-- Broker Email
                 $sheet->setCellValue('A6', "Broker Email");
                 $spreadsheet->getActiveSheet()->getStyle('B6')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('B6')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $spreadsheet->getActiveSheet()->setCellValue('B6', $company['broker_email']);

                 //-- Company Type
                 $sheet->setCellValue('A7', "Company Type");
                 $spreadsheet->getActiveSheet()->getStyle('B7')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('B7')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $spreadsheet->getActiveSheet()->setCellValue('B7', $company['company_type']);

                 //-- Tax ID
                 $sheet->setCellValue('A8', "Tax ID");
                 $spreadsheet->getActiveSheet()->getStyle('B8')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('B8')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('B8', $company['tax_id']);

                 //-- Contact Name
                 $sheet->setCellValue('A9', "Contact Name");
                 $spreadsheet->getActiveSheet()->getStyle('B9')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('B9')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('B9', $company['contact_name']);

		 //-- Contact Phone
                 $sheet->setCellValue('A10', "Contact Phone");
                 $spreadsheet->getActiveSheet()->getStyle('B10')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('B10')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('B10', $company['contact_phone']);

                 //-- Contact Email
                 $sheet->setCellValue('A11', "Contact Email");
                 $spreadsheet->getActiveSheet()->getStyle('B11')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('B11')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('B11', $company['contact_email']);

                 //-- Contact Email
                 $sheet->setCellValue('A12', "Employee Count (est)");
                 $spreadsheet->getActiveSheet()->getStyle('B12')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('B12')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('B12', $company['employee_count']);

                 //-- Medical
                 $sheet->setCellValue('A14', "Medical (Yes/No)");
                 $spreadsheet->getActiveSheet()->getStyle('B14')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('B14')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('B14', $company['medical']);
		 //
                 //-- Dental
                 $sheet->setCellValue('A15', "Dental (Yes/No)");
                 $spreadsheet->getActiveSheet()->getStyle('B15')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('B15')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('B15', $company['dental']);
		 //
                 //-- Vision
                 $sheet->setCellValue('A16', "Vision (Yes/No)");
                 $spreadsheet->getActiveSheet()->getStyle('B16')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('B16')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('B16', $company['vision']);
		 //
                 //-- Provider
                 $sheet->setCellValue('A18', "Previous Insurance Provider");
                 $spreadsheet->getActiveSheet()->getStyle('B18')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('B18')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('B18', $company['current_provider']);
		 //
                 //-- Contact Email
                 $sheet->setCellValue('D3', "Company Address");
                 $spreadsheet->getActiveSheet()->getStyle('E3')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E3')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('E3', $company['address']);

                 //-- Suite
                 $sheet->setCellValue('D4', "Line 2");
                 $spreadsheet->getActiveSheet()->getStyle('E4')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E4')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('E4', $company['suite']);

                 //-- City
                 $sheet->setCellValue('D5', "City");
                 $spreadsheet->getActiveSheet()->getStyle('E5')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E5')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('E5', $company['city']);

                 //-- State
                 $sheet->setCellValue('D6', "State");
                 $spreadsheet->getActiveSheet()->getStyle('E6')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E6')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('E6', $company['state']);

                 //-- Zip
                 $sheet->setCellValue('D7', "Zip");
                 $spreadsheet->getActiveSheet()->getStyle('E7')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E7')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('E7', $company['zip']);

                 //-- Website
                 $sheet->setCellValue('D8', "Website");
                 $spreadsheet->getActiveSheet()->getStyle('E8')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E8')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('E8', $company['website']);

                 //-- Billing Address
                 $sheet->setCellValue('D10', "Billing Address");

                 //-- Company Name
                 $sheet->setCellValue('D11', "Company Name");
                 $spreadsheet->getActiveSheet()->getStyle('E11')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E11')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('E11', $company['billing_company_name']);

                 //-- Mailing Address
                 $sheet->setCellValue('D12', "Address");
                 $spreadsheet->getActiveSheet()->getStyle('E12')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E12')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('E12', $company['invoice_mailing_address']);

                 //-- Invoice Suite
                 $sheet->setCellValue('D13', "Suite");
                 $spreadsheet->getActiveSheet()->getStyle('E13')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E13')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('E13', $company['invoice_suite']);

                 //-- Invoice City
                 $sheet->setCellValue('D14', "City");
                 $spreadsheet->getActiveSheet()->getStyle('E14')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E14')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('E14', $company['invoice_city']);

                 //-- Invoice State
                 $sheet->setCellValue('D15', "State");
                 $spreadsheet->getActiveSheet()->getStyle('E15')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E15')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('E15', $company['invoice_state']);

                 //-- Invoice Zip
                 //
                 $sheet->setCellValue('D16', "Zip");
                 $spreadsheet->getActiveSheet()->getStyle('E16')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E16')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('E16', $company['invoice_zip']);

                 //-- Invoice Contact Name
                 $sheet->setCellValue('D18', "Billing Contact Name");
                 $spreadsheet->getActiveSheet()->getStyle('E18')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E18')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $spreadsheet->getActiveSheet()->setCellValue('E18', $company['billing_contact_name']);

                 //-- Invoice Contact Phone
                 $sheet->setCellValue('D19', "Billing Contact Phone");
                 $spreadsheet->getActiveSheet()->getStyle('E19')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E19')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('E19', $company['billing_contact_name']);

                 //-- Billing Contact Email
                 $sheet->setCellValue('D20', "Billing Contact Email");
                 $spreadsheet->getActiveSheet()->getStyle('E20')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E20')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('E20', $company['billing_contact_email']);

                 //-- Billing Contact Email
                 $sheet->setCellValue('D21', "Billing Contact Email 2");
                 $spreadsheet->getActiveSheet()->getStyle('E21')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E21')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('E21', $company['billing_contact_email2']);

                 //-- Billing Contact Email
                 $sheet->setCellValue('D22', "Billing Contact Email 3");
                 $spreadsheet->getActiveSheet()->getStyle('E22')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E22')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('E22', $company['billing_contact_email3']);

                 //-- Billing Contact Email
                 $sheet->setCellValue('D23', "Billing Contact Email 4");
                 $spreadsheet->getActiveSheet()->getStyle('E23')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E23')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('E23', $company['billing_contact_email4']);


                 //-- Billing Contact Email
                 $sheet->setCellValue('D24', "Billing Contact Email 5");
                 $spreadsheet->getActiveSheet()->getStyle('E24')->applyFromArray($styleArrayBold);
                 $spreadsheet->getActiveSheet()->getStyle('E24')->getFill()->getStartColor()->setARGB('DDDDDDDD');
                 if ($company_id!='0') $sheet->setCellValue('E24', $company['billing_contact_email5']);

//================================================================================
//  PREENROLLMENT CENSUS
//================================================================================

       if ($company_id!=0) {
        $spreadsheet->setActiveSheetIndex(1);

        $spreadsheet->getActiveSheet()->getStyle('A1')->applyFromArray($styleArrayHead);
        $spreadsheet->getActiveSheet()->setCellValue('A1', "PRE-ENROLLMENT CENSUS");

        //-- Company ID
	$spreadsheet->getActiveSheet()->getRowDimension('2')->setRowHeight(100);
        $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('A2')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('A2')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('A2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('A2', "Member ID (Social Security Number)");
        $spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('B2')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('B2')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('B2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('B2', "Dependent ID (Social Security Number)");
        $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(15);
        $spreadsheet->getActiveSheet()->getStyle('C2')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('C2')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('C2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('C2', "Last Name");
        $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('D2')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('D2')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('D2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('D2', "First Name");
        $spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('E2')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('E2')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('E2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('E2', "Middle Initial");

        $spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('F2')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('F2')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('F2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('F2', "Relationship (Employee, Husband, Wife Son, Daughter)");

        $spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('G2')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('G2')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('G2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('G2', "Date of Birth");

        $spreadsheet->getActiveSheet()->getColumnDimension('H')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('H2')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('H2')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('H2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('H2', "Gender (M/F)");

        $spreadsheet->getActiveSheet()->getColumnDimension('I')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('I2')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('I2')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('I2')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('I2', "Marital Status - S/M");

        $sql="select * from nua_preenrollment_census where company_id = " . $company_id . " order by last_name, first_name";
        $x=$X->sql($sql);
        $row=2;
	foreach($x as $q) {
      $row++;
      $cell="A".$row;
      $spreadsheet->getActiveSheet()->setCellValue($cell, $q['social_security_number']);
      $cell="B".$row;
      $spreadsheet->getActiveSheet()->setCellValue($cell, $q['dependent_social_security_number']);
      $cell="C".$row;
      $spreadsheet->getActiveSheet()->setCellValue($cell, $q['last_name']);
      $cell="D".$row;
      $spreadsheet->getActiveSheet()->setCellValue($cell, $q['first_name']);
      $cell="E".$row;
      $spreadsheet->getActiveSheet()->setCellValue($cell, $q['middle_name']);
      $cell="F".$row;
      $spreadsheet->getActiveSheet()->setCellValue($cell, $q['relationship']);
      $cell="G".$row;
      $spreadsheet->getActiveSheet()->setCellValue($cell, $q['date_of_birth']);
      $cell="H".$row;
      $spreadsheet->getActiveSheet()->setCellValue($cell, $q['gender']);
      $cell="I".$row;
      $spreadsheet->getActiveSheet()->setCellValue($cell, $q['marital_status']);

        }
}

if ($company_id!=0) {
if ($quoted_plan=="Y") {
                $spreadsheet->setActiveSheetIndex(2);
                 //-- Column Widths
                 $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(32);
                 $spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(15);
                 $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(20);
                 $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(20);
                 $spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(20);
                 $spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(20);
                 $spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(20);
                 //-- Company Name Title
                 //
                 $spreadsheet->getActiveSheet()->getStyle('A1')->applyFromArray($styleArrayHead);
                 $spreadsheet->getActiveSheet()->setCellValue('A1', "Quoted Plans");
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
                 $spreadsheet->getActiveSheet()->getStyle('G2')->applyFromArray($styleArrayBold);
		$spreadsheet->getActiveSheet()->getStyle('G2')->getFill()->getStartColor()->setARGB('FFFFCC66');
                 $spreadsheet->getActiveSheet()->setCellValue('G2', "Accepted");

$sql="select * from nua_quoted_plan where company_id = " . $company_id . " order by plan_type, plan_code";
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
      $cell="G".$row;
      $spreadsheet->getActiveSheet()->setCellValue($cell, $q['accepted']);
}
} // QUOTED PLAN
} // COPMANY !=0

if ($company_id!=0) {
if ($company_plan=="Y") {

                $spreadsheet->setActiveSheetIndex(3);
                 //-- Column Widths
                 $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(32);
                 $spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(15);
                 $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(20);
                 $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(20);
                 $spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(20);
                 $spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(20);
                 $spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(20);
                 //-- Company Name Title
                 //
                 $spreadsheet->getActiveSheet()->getStyle('A1')->applyFromArray($styleArrayHead);
                 $spreadsheet->getActiveSheet()->setCellValue('A1', "Active Plans");
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
                 $spreadsheet->getActiveSheet()->getStyle('G2')->applyFromArray($styleArrayBold);
		$spreadsheet->getActiveSheet()->getStyle('G2')->getFill()->getStartColor()->setARGB('FFFFCC66');
                 $spreadsheet->getActiveSheet()->setCellValue('G2', "Plan Type");

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
      $cell="G".$row;
      $spreadsheet->getActiveSheet()->setCellValue($cell, $q['plan_type']);
}
}
}

if ($company_id!=0) {
if ($company_plan=="Y") {

        $spreadsheet->setActiveSheetIndex(4);
	$spreadsheet->getActiveSheet()->getRowDimension('1')->setRowHeight(100);
        $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('A1')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('A1')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('A1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('A1', "Month Effective");

        $spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('B1')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('B1')->getFill()->getStartColor()->setARGB('DDDDDDDD');
	$spreadsheet->getActiveSheet()->getStyle('B1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('B1', "Latest Month");

        $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('C1')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('C1')->getFill()->getStartColor()->setARGB('DDDDDDDD');
	$spreadsheet->getActiveSheet()->getStyle('C1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('C1', "Employee Code (Leave Blank)");

        $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(15);
        $spreadsheet->getActiveSheet()->getStyle('D1')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('D1')->getFill()->getStartColor()->setARGB('DDDDDDDD');
	$spreadsheet->getActiveSheet()->getStyle('D1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('D1', "Dependent Code (Leave Blank)");

        $spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('E1')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('E1')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('E1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('E1', "Social Security Number");

        $spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('F1')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('F1')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('F1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('F1', "Dependent SSN");

        $spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('G1')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('G1')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('G1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('G1', "Last Name");

        $spreadsheet->getActiveSheet()->getColumnDimension('H')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('H1')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('H1')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('H1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('H1', "First Name");

        $spreadsheet->getActiveSheet()->getColumnDimension('I')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('I1')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('I1')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('I1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('I1', "Middle Initial");

        $spreadsheet->getActiveSheet()->getColumnDimension('J')->setWidth(15);
        $spreadsheet->getActiveSheet()->getStyle('J1')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('J1')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('J1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('J1', "Relationship (Employee, Husband, Wife Son, Daughter)");

        $spreadsheet->getActiveSheet()->getColumnDimension('K')->setWidth(15);
        $spreadsheet->getActiveSheet()->getStyle('K1')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('K1')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('K1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('K1', "Date of Birth");

        $spreadsheet->getActiveSheet()->getColumnDimension('L')->setWidth(10);
        $spreadsheet->getActiveSheet()->getStyle('L1')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('L1')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('L1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('L1', "Gender (M/F)");

        $spreadsheet->getActiveSheet()->getColumnDimension('M')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('M1')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('M1')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('M1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('M1', "Marital Status");

        $spreadsheet->getActiveSheet()->getColumnDimension('N')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('N1')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('N1')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('N1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('N1', "Address");

        $spreadsheet->getActiveSheet()->getColumnDimension('O')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('O1')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('O1')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('O1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('O1', "Suite / Apt");

        $spreadsheet->getActiveSheet()->getColumnDimension('P')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('P1')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('P1')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('P1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('P1', "City");

        $spreadsheet->getActiveSheet()->getColumnDimension('Q')->setWidth(10);
        $spreadsheet->getActiveSheet()->getStyle('Q1')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('Q1')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('Q1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('Q1', "State");

        $spreadsheet->getActiveSheet()->getColumnDimension('R')->setWidth(10);
        $spreadsheet->getActiveSheet()->getStyle('R1')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('R1')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('R1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('R1', "Zipcode");

        $spreadsheet->getActiveSheet()->getColumnDimension('S')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('S1')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('S1')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('S1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('S1', "Email (Provide a personal email)");

        $spreadsheet->getActiveSheet()->getColumnDimension('T')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('T1')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('T1')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('T1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('T1', "Contact Phone (Required if no email)");

        $spreadsheet->getActiveSheet()->getColumnDimension('U')->setWidth(25);
        $spreadsheet->getActiveSheet()->getStyle('U1')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('U1')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('U1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('U1', "Date of Hire");

        $spreadsheet->getActiveSheet()->getColumnDimension('V')->setWidth(10);
        $spreadsheet->getActiveSheet()->getStyle('V1')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('V1')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('V1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('V1', "Work Status (A (Full Time) / PT (Part Time)");

        $spreadsheet->getActiveSheet()->getColumnDimension('W')->setWidth(30);
        $spreadsheet->getActiveSheet()->getStyle('W1')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('W1')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('W1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('W1', "Effective Date (If blank defaults to first day of month in column A)");

        $spreadsheet->getActiveSheet()->getColumnDimension('X')->setWidth(30);
        $spreadsheet->getActiveSheet()->getStyle('X1')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('X1')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('X1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('X1', "Term Date");


// MEDICAL

	$spreadsheet->getActiveSheet()->getColumnDimension('Y')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('Y1')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('Y1')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('Y1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('Y1', "Medical Plan");

        $spreadsheet->getActiveSheet()->getColumnDimension('Z')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('Z1')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('Z1')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('Z1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('Z1', "Coverage Election (EE, ES, EC, FAM)");

        $spreadsheet->getActiveSheet()->getColumnDimension('AA')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('AA1')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('AA1')->getFill()->getStartColor()->setARGB('DDDDDDDD');
	$spreadsheet->getActiveSheet()->getStyle('AA1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('AA1', "Medical Price (NuAxess Use Only)");

// DENTAL

	$spreadsheet->getActiveSheet()->getColumnDimension('AB')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('AB1')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('AB1')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('AB1')->getAlignment()->setWrapText(true);
	$spreadsheet->getActiveSheet()->setCellValue('AB1', "Dental Plan (Blank if none)");

        $spreadsheet->getActiveSheet()->getColumnDimension('AC')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('AC1')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('AC1')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('AC1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('ACA1', "Dental Coverage (EE,ES,EC,FAM)");

        $spreadsheet->getActiveSheet()->getColumnDimension('AD')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('AD1')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('AD1')->getFill()->getStartColor()->setARGB('DDDDDDDD');
	$spreadsheet->getActiveSheet()->getStyle('AD1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('AD1', "Dental Price (NuAxess Use Only)");

// VISION
//
        $spreadsheet->getActiveSheet()->getColumnDimension('AE')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('AE1')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('AE1')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('AE1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('AE1', "Vision Plan");

        $spreadsheet->getActiveSheet()->getColumnDimension('AF')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('AF1')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('AF1')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('AF1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('AF1', "Vision Coverage");

        $spreadsheet->getActiveSheet()->getColumnDimension('AG')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('AG1')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('AG1')->getFill()->getStartColor()->setARGB('DDDDDDDD');
	$spreadsheet->getActiveSheet()->getStyle('AG1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('AG1', "VISION Price (NuAxess Use Only)");

        $spreadsheet->getActiveSheet()->getColumnDimension('AH')->setWidth(20);
        $spreadsheet->getActiveSheet()->getStyle('AH1')->applyFromArray($styleArrayEnrollment);
	$spreadsheet->getActiveSheet()->getStyle('AH1')->getFill()->getStartColor()->setARGB('FFFFCC66');
	$spreadsheet->getActiveSheet()->getStyle('AH1')->getAlignment()->setWrapText(true);
        $spreadsheet->getActiveSheet()->setCellValue('AH1', "Action (ADMIN)");

//--
//-- This gets Everyone 
//-- 
        $sql="select employee_id, employee_code, min(month_id) as min_month_id, max(month_id) as max_month_id, last_name, first_name from "; 
        $sql.=" nua_monthly_member_census where company_id = " . $company_id;
        $sql.=" and dependent_code = '' group by employee_id, employee_code, last_name, first_name order by last_name, first_name";

        $d=$X->sql($sql);
        $row=1;

        foreach($d as $e) {
            $sql="select * from nua_employee where id = " . $e['employee_id'];
            $x=$X->sql($sql);
            $employee=$x[0];
            $sql="select * from nua_monthly_member_census where company_id = " . $company_id . " and employee_id = " . $e['employee_id'] . " and ";
            $sql.=" dependent_code = '' and month_id = '" . $e['max_month_id'] . "'";
	    $hh=$X->sql($sql);
            $census=$hh[0];

            $sql="select * from nua_monthly_member_terminations where company_id = " . $company_id . " and employee_id = " . $e['employee_id'] . " and ";
            $sql.=" dependent_code = '' order by month_id";
	    $ii=$X->sql($sql);
            $term_dt="";
            foreach($ii as $jj) $term_dt=$jj['term_dt'];

            $row++;
            $cell="A".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if ($e['max_month_id']!=$month_id) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
		if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
		$spreadsheet->getActiveSheet()->setCellValue($cell, $e['min_month_id']);     
            $cell="B".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if ($e['max_month_id']!=$month_id) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
		if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
		$spreadsheet->getActiveSheet()->setCellValue($cell, $e['max_month_id']);     
            $cell="C".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if ($e['max_month_id']!=$month_id) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
		if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['employee_code']);
            $cell="D".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if ($e['max_month_id']!=$month_id) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
		if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['dependent_code']);
            $cell="E".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if ($e['max_month_id']!=$month_id) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
		if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['social_security_number']);
            $cell="F".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if ($e['max_month_id']!=$month_id) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
		if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
            $cell="G".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if ($e['max_month_id']!=$month_id) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
		if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['last_name']);
            $cell="H".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if ($e['max_month_id']!=$month_id) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
		if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['first_name']);
            $cell="I".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if ($e['max_month_id']!=$month_id) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
		if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['middle_name']);
            $cell="J".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if ($e['max_month_id']!=$month_id) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
		if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
			$spreadsheet->getActiveSheet()->setCellValue($cell, "Employee");
            $cell="K".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if ($e['max_month_id']!=$month_id) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
		if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['date_of_birth']);
            $cell="L".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if ($e['max_month_id']!=$month_id) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
		if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['gender']);
            $cell="M".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if ($e['max_month_id']!=$month_id) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
		if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['marital_status']);
            $cell="N".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if ($e['max_month_id']!=$month_id) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
		if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['address']);
            $cell="O".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if ($e['max_month_id']!=$month_id) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
		if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['suite']);
            $cell="P".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if ($e['max_month_id']!=$month_id) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
		if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['city']);
            $cell="Q".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if ($e['max_month_id']!=$month_id) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
		if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['state']);
            $cell="R".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if ($e['max_month_id']!=$month_id) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
		if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['zip']);
            $cell="S".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if ($e['max_month_id']!=$month_id) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
		if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['email']);
            $cell="T".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if ($e['max_month_id']!=$month_id) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
		if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['phone']);
            $cell="U".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if ($e['max_month_id']!=$month_id) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
		if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['date_hired']);
            $cell="V".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if ($e['max_month_id']!=$month_id) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
		if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['work_status']);
            $cell="W".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if ($e['max_month_id']!=$month_id) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
		if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['eff_dt']);

            $cell="X".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if ($e['max_month_id']!=$month_id) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
		if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
			$spreadsheet->getActiveSheet()->setCellValue($cell, $term_dt);

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
            $cell="Y".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if ($e['max_month_id']!=$month_id) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
		if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
			$spreadsheet->getActiveSheet()->setCellValue($cell, $medical_plan);
            $cell="Z".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if ($e['max_month_id']!=$month_id) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
		if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
			$spreadsheet->getActiveSheet()->setCellValue($cell, $medical_coverage);
            $cell="AA".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if ($e['max_month_id']!=$month_id) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
		if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
			$spreadsheet->getActiveSheet()->setCellValue($cell, $medical_price);
            $cell="AB".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if ($e['max_month_id']!=$month_id) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
		if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
			$spreadsheet->getActiveSheet()->setCellValue($cell, $dental_plan);
            $cell="AC".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if ($e['max_month_id']!=$month_id) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
		if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
			$spreadsheet->getActiveSheet()->setCellValue($cell, $dental_coverage);
            $cell="AD".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if ($e['max_month_id']!=$month_id) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
		if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
			$spreadsheet->getActiveSheet()->setCellValue($cell, $dental_price);
            $cell="AE".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if ($e['max_month_id']!=$month_id) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
		if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
			$spreadsheet->getActiveSheet()->setCellValue($cell, $vision_plan);
            $cell="AF".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if ($e['max_month_id']!=$month_id) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
		if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
			$spreadsheet->getActiveSheet()->setCellValue($cell, $vision_coverage);
            $cell="AG".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if ($e['max_month_id']!=$month_id) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
		if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
			$spreadsheet->getActiveSheet()->setCellValue($cell, $vision_price);
            $cell="AH".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
		if ($e['max_month_id']!=$month_id) {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FF00CCFF');
                }
		if ($census['error_msg']!="") {
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
                     $spreadsheet->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setARGB('FFFF8080');
                }
			$spreadsheet->getActiveSheet()->setCellValue($cell, $census['error_msg']);

            $sql="select * from nua_employee_dependent where employee_id = " . $e['employee_id'] . " order by dependent_id";
            $z=$X->sql($sql);
            foreach($z as $ff) {
                $row++;
                $cell="C".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
			$spreadsheet->getActiveSheet()->setCellValue($cell, $e['employee_code']);
                $cell="D".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
			$spreadsheet->getActiveSheet()->setCellValue($cell, $ff['dependent_id']);
                $cell="E".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
			$spreadsheet->getActiveSheet()->setCellValue($cell, $employee['social_security_number']);
                $cell="F".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
			$spreadsheet->getActiveSheet()->setCellValue($cell, $ff['social_security_number']);
                $cell="G".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
			$spreadsheet->getActiveSheet()->setCellValue($cell, $ff['last_name']);
                $cell="H".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
			$spreadsheet->getActiveSheet()->setCellValue($cell, $ff['first_name']);
                $cell="I".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
			$spreadsheet->getActiveSheet()->setCellValue($cell, $ff['middle_name']);
                $cell="J".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
			$spreadsheet->getActiveSheet()->setCellValue($cell, $ff['relationship']);
                $cell="K".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
			$spreadsheet->getActiveSheet()->setCellValue($cell, $ff['date_of_birth']);
                $cell="L".$row;
                     $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleArrayLine);
			$spreadsheet->getActiveSheet()->setCellValue($cell, $ff['gender']);
            }

	}
}
}


$spreadsheet->setActiveSheetIndex(0);
$writer = new Xlsx($spreadsheet);
//$writer->save('sheets/hell.xlsx');
if ($company_id==0) {
   $filename="NuAxess_Quoting_Template.xlsx";
} else {
   $filename="QUOTE_TEMPLATE_" . $company_id . "_" . str_replace(' ','_',str_replace('#',' ',$company['company_name'])) . ".xlsx";
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


