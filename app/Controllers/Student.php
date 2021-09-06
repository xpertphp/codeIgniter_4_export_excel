<?php 

namespace App\Controllers;
 
use CodeIgniter\Controller;
use App\Models\StudentModel;
use CodeIgniter\HTTP\RequestInterface;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
 
class Student extends Controller
{
 
	public function __construct()
    {
         helper(['form', 'url']);
    }
    public function index()
    {   
        $model = new StudentModel();
		$data_result = $model->orderBy('id', 'DESC')->findAll();	
		
		$fileName = 'students.xlsx';  
		$spreadsheet = new Spreadsheet();

		$sheet = $spreadsheet->getActiveSheet();
		$sheet->setCellValue('A1', 'Id');
		$sheet->setCellValue('B1', 'First Name');
		$sheet->setCellValue('C1', 'Last Name');
		$sheet->setCellValue('D1', 'Email');
		$sheet->setCellValue('E1', 'Address');
		$sheet->setCellValue('F1', 'Mobile');       
		$rows = 2;

		foreach ($data_result as $val){
		  $sheet->setCellValue('A' . $rows, $val['id']);
		  $sheet->setCellValue('B' . $rows, $val['first_name']);
		  $sheet->setCellValue('C' . $rows, $val['last_name']);
		  $sheet->setCellValue('D' . $rows, $val['email']);
		  $sheet->setCellValue('E' . $rows, $val['address']);
		  $sheet->setCellValue('F' . $rows, $val['mobile']);
		  $rows++;
		} 
		$writer = new Xlsx($spreadsheet);
		$writer->save("upload/".$fileName);
		header("Content-Type: application/vnd.ms-excel");
		redirect(base_url()."/upload/".$fileName);
    }
 
}

?>