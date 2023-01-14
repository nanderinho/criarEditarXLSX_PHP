# criarEditarXLSX_PHP

/* ========================================= */

INSTRUÇOES

Caso não execute, no console do projeto, execute:
composer require phpoffice/phpspreadsheet

Para retornar um valor
$cellB2 = $sheet->getCell('B2')->getValue();

Para Setar um valor
$sheet->setCellValue('D1', 'Value'); 

Para teste rápidos sem criação de novas páginas. Execute o projeto no browser com um um servidor local (laragon, apache).
Irá varia de acordo com suas configuração, mas segue um exemplo abaixo.

CRIAR
Exemplo: http://localhost:8082/criarEditarXLSX_PHP/criar.php

EDITAR
Exemplo: http://localhost:8082/criarEditarXLSX_PHP/criar.php

