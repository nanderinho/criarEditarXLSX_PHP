<?php

/* ========================================= */
//INSTRUÇOES

//Caso não execute, no console do projeto, execute:
//composer require phpoffice/phpspreadsheet

//Para retornar um valor
//$cellB2 = $sheet->getCell('B2')->getValue();

//Para Setar um valor
//$sheet->setCellValue('D1', 'Value'); 

/* ========================================= */

require 'vendor/autoload.php'; //autoload do projeto


use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

//Instancia nova planilha
$spreadsheet = new Spreadsheet(); 

//Aba ativa
$sheet = $spreadsheet->getActiveSheet(); 

$sheet->setCellValue('A1', '@nanderinho');
$sheet->setCellValue('B1', 'Games');
$sheet->setCellValue('C1', 'Dicas');
$sheet->setCellValue('D1', 'Tecnologia');


$sheet->setCellValue('A3', 'Siga-me nas redes sociais @nanderinho');
$sheet->setCellValue('B3', 'Instagram');
$sheet->setCellValue('C3', 'Youtube');
$sheet->setCellValue('D3', 'TikTok');


$sheet->setCellValue('F1', 'Repósitorio GITHUB: ');
$sheet->setCellValue('G1', 'https://github.com/nanderinho?tab=repositories');

$sheet->setCellValue('H1', 'nanderinho.com.br');

//Instaciar planilha
$writer = new Xlsx($spreadsheet);

$writer->save('arquivo/Planilha.xlsx');
