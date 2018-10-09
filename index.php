<?php
require "vendor/autoload.php";
include 'Curl/CaseInsensitiveArray.php';
include 'Curl/Curl.php';
include 'Curl/MultiCurl.php';

include 'DiDom/Document.php';
include 'DiDom/Element.php';
include 'DiDom/Query.php';



use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;

use Curl\Curl;
use DiDom\Document;
use DiDom\Element;


$url = "https://diemthi.tuyensinh247.com/diem-chuan.html";
$allLink;
$data=[];
$postURL = "https://diemthi.tuyensinh247.com/diem-chuan/dai-hoc-bach-khoa-ha-noi-BKA.html?y=2013";
$current=0;
$index=0;
get_link($url, $allLink);
get_detail($allLink,array('y'=>'2013'),$data);
export_excel($data);
function export_excel($data)
{
    $spreadsheet = new Spreadsheet();
//Specify the properties for this document
    $spreadsheet->getProperties()
        ->setTitle('Sheet');
//Adding data to the excel sheet
    for ($i = 0; $i < count($data); $i++) {
        $spreadsheet->setActiveSheetIndex(0)
            ->setCellValue('A' . ($i + 1), $data[$i]["stt"]);
        $spreadsheet->getActiveSheet()
            ->setCellValue('B' . ($i + 1), $data[$i]["idDepartment"]);
        $spreadsheet->getActiveSheet()
            ->setCellValue('C' . ($i + 1), $data[$i]["name"]);
        $spreadsheet->getActiveSheet()
            ->setCellValue('D' . ($i + 1), $data[$i]["subjects"]);
        $spreadsheet->getActiveSheet()
            ->setCellValue('E' . ($i + 1), $data[$i]["point"]);
        if(trim($data[$i]["stt"])=='Đại Học Hà Tĩnh - 2013'){
            $spreadsheet->getActiveSheet()->getStyle('A'. ($i + 1))
                ->getFont()->getColor()->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_RED);
            $spreadsheet->getActiveSheet()->getStyle('B'. ($i + 1))
                ->getFont()->getColor()->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_RED);
            $spreadsheet->getActiveSheet()->getStyle('C'. ($i + 1))
                ->getFont()->getColor()->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_RED);
            $spreadsheet->getActiveSheet()->getStyle('D'. ($i + 1))
                ->getFont()->getColor()->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_RED);
            $spreadsheet->getActiveSheet()->getStyle('E'. ($i + 1))
                ->getFont()->getColor()->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_RED);
        }
    }
    $writer = IOFactory::createWriter($spreadsheet, "Xlsx"); //Xls is also possible
    $filename = "dai_hoc_ha_tinh.xlsx";
    $path = __DIR__ . '/files/' . $filename;
    $writer->save($path);
}


function get_detail($link,$parameter,&$data)
{
    foreach ($link as $detail) {
        $baseURL = "https://diemthi.tuyensinh247.com";
        $strUrl = $baseURL . $detail;
        get_data($strUrl, $parameter,$data);
    }
}

function get_data($url,$parameter, &$data)
{
    $curl = new Curl();
    $curl->setOpt(CURLOPT_ENCODING, '');
    $html = $curl->get($url,$parameter);
    $curl->close();
    $doc = new Document();
    $doc->loadHtml($html);
    $NameOfUniversity = $doc->find('.clblue2')[0]->text();

    echo $NameOfUniversity."\n";
    $elements = $doc->find('div[class=resul-seah] .bg_white');
    for ($i = 0; $i < count($elements); ++$i) {
        $e = $elements[$i];
        $td = $e->find('td');
        if (($idDepartment = trim($td[1]->text())) == '7480201') {
            $result["stt"] = $NameOfUniversity;
            $result["idDepartment"] = $td[1]->text();
            $result["name"] = $td[2]->text();
            $result["subjects"] = $td[3]->text();
            $result["point"] = $td[4]->text();
            $data[] = $result;
        }
    }
}

function get_link($url, &$data)
{
    $curl = new Curl();
    $curl->setOpt(CURLOPT_ENCODING, '');
    $html = $curl->get($url);
    $curl->close();
    $doc = new Document();
    $doc->loadHtml($html);

    $elements = $doc->find('ul[id=benchmarking]');
    $child = $elements[0]->find('li');
    for ($i = 0; $i < count($child); ++$i) {
        $e = $child[$i];
        $href = $e->find('a')[0]->href;
        $data[] = $href;
    }
}

?>
