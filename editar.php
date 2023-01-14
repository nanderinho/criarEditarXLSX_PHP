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


use PhpOffice\PhpSpreadsheet\IOFactory; //classe responsável por ler uma planilha

use PhpOffice\PhpSpreadsheet\Spreadsheet; //classe responsável pela manipulação da planilha

use PhpOffice\PhpSpreadsheet\Writer\Xlsx; //classe que salvará a planilha em .xlsx

//
$clientes = [
    "Cliente1",
    "cliente2",
    "Cliente3",
    "Cliente4",
    "Cliente5",
    "Cliente6",
    "Cliente7",
    "Cliente8",
    "Cliente9",
    "Cliente10",
    "Cliente11",
    "Cliente12",
    "Cliente13",
    "Cliente14",
    "Cliente15",
    "Cliente16",
    "Cliente17",
    "Cliente18",
    "Cliente19",
    "Cliente20",
    "Cliente21",
    "Cliente22",
    "Cliente23",
    "Cliente24",
    "Cliente25",
    "Cliente26",
    "Cliente27",
    "Cliente28",
    "Cliente29",
    "Cliente30",
    "Cliente31",
    "Cliente32",
    "Cliente33",
    "Cliente34",
    "Cliente35",
    "Cliente36",
    "Cliente37",
    "Cliente38",
    "Cliente39",
    "Cliente40",
    "Cliente41",
    "Cliente42",
    "Cliente43",
    "Cliente44",
    "Cliente45",
    "Cliente46",
    "Cliente47",
    "Cliente48",
    "Cliente49",
    "Cliente50",
    "Cliente51",
    "Cliente52",
    "Cliente53",
    "Cliente54",
    "Cliente55",
    "Cliente56",
    "Cliente57",
    "Cliente58"
];

$spreadsheet = IOFactory::load('arquivo/ControledeRota.xlsx'); //tornando a planilha em um objeto

$sheet = $spreadsheet->getActiveSheet(); //retornando a aba ativa

//PLACA
$sheet->setCellValue('F1', 'QTD1234');

//Turno
$sheet->setCellValue('I1', 'T');

//Motorista
$sheet->setCellValue('L1', 'Jade');

//Ajudante
$sheet->setCellValue('Q1', 'Fernando');

//Data
$sheet->setCellValue('T1', '14/01/2022');


if (count($clientes) <= 58) {
    //Posição coluna
    $col = 10;

    for ($i = 0; $i < count($clientes); $i++) {
        if($i == 30){
            $col++;
        }
        $sheet->setCellValue("A".$col, $clientes[$i]);

        $col++;
    }
}



$writer = new Xlsx($spreadsheet); //Instanciando uma nova planilha

$writer->save('arquivo/ControledeRotaOK.xlsx'); //salvando a planilha na extensão definida