<?php 
error_reporting(0);
define('DB_NAME', 'samaj-demo');
define('DB_USER', 'root');
define('DB_PASSWORD', '');
define('DB_HOST', 'localhost');
  
// Create connection
$db     =   new mysqli(DB_HOST, DB_USER, DB_PASSWORD, DB_NAME);
// Check connection
if ($db->connect_error) {
    die("Connection failed: " . $db->connect_error);
}


include('PHPExcel/Classes/PHPExcel.php');
$objPHPExcel = new PHPExcel();

$result         =   $db->query("SELECT * FROM emp_mst") or die(mysql_error());
$objPHPExcel->setActiveSheetIndex(0);
$style = array(
        'alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
        )
    );

$objPHPExcel->getActiveSheet()->mergeCells('A1:G1');
$objPHPExcel->getActiveSheet()->setCellValue('A1', 'THE SAMAJ');
$objPHPExcel->getActiveSheet()->getStyle('A1:G1')->getFont()->setBold(true)->setSize(25);
$objPHPExcel->getActiveSheet()->getStyle("A1:G1")->applyFromArray($style);


$objPHPExcel->getActiveSheet()->mergeCells('A2:G2');
$objPHPExcel->getActiveSheet()->setCellValue('A2', 'EMPLOYEES LIST(Category Wise)');
$objPHPExcel->getActiveSheet()->getStyle('A2:G2')->getFont()->setBold(true)->setSize(17);
$objPHPExcel->getActiveSheet()->getStyle("A2:G2")->applyFromArray($style);


$objPHPExcel->getActiveSheet()->getStyle('A3:G3')->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle("A3:G3")->applyFromArray($style);
$objPHPExcel->getActiveSheet()->getStyle('A3:G3')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('7bd2fd');

$objPHPExcel->getActiveSheet()->SetCellValue('A3', 'EMP TYPE');
$objPHPExcel->getActiveSheet()->SetCellValue('B3', 'CODE');
$objPHPExcel->getActiveSheet()->SetCellValue('C3', 'EMPLOYEE NAME');
$objPHPExcel->getActiveSheet()->SetCellValue('D3', 'DESIGNATION');
$objPHPExcel->getActiveSheet()->SetCellValue('E3', 'DEPARTMENT');
$objPHPExcel->getActiveSheet()->SetCellValue('F3', 'ACTIVE TYPE');
$objPHPExcel->getActiveSheet()->SetCellValue('G3', 'JOINING DATE');


$i   =   4;
while($row  =   $result->fetch_assoc()){
$objPHPExcel->getActiveSheet()->SetCellValue('A'.$i, mb_strtoupper($row['emp_type'],'UTF-8'));
$objPHPExcel->getActiveSheet()->SetCellValue('B'.$i, mb_strtoupper($row['employee_code'],'UTF-8'));
$objPHPExcel->getActiveSheet()->SetCellValue('C'.$i, mb_strtoupper($row['emp_name'],'UTF-8'));
$objPHPExcel->getActiveSheet()->SetCellValue('D'.$i, mb_strtoupper($row['desg_code'],'UTF-8'));
$objPHPExcel->getActiveSheet()->SetCellValue('E'.$i, mb_strtoupper($row['dept_no'],'UTF-8'));
$objPHPExcel->getActiveSheet()->SetCellValue('F'.$i, mb_strtoupper($row['active_type'],'UTF-8'));
$objPHPExcel->getActiveSheet()->SetCellValue('G'.$i, mb_strtoupper($row['DOB'],'UTF-8'));
$objPHPExcel->getActiveSheet()->getStyle("A".$i.":G".$i)->applyFromArray($style);
$i++;
}
 
 
$objWriter  =   new PHPExcel_Writer_Excel2007($objPHPExcel);
header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename="you-file-name.xlsx"');
header('Cache-Control: max-age=0');
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');  
$objWriter->save('php://output');
?>