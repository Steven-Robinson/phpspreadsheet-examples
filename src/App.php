<?php

namespace App;

use App\Component\Spreadsheet\Service\SpreadsheetService;
use App\Component\Spreadsheet\Controller\SpreadsheetController;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Pimple\Container;

class App extends Container
{
    public function __construct()
    {
        $this['spreadsheet'] = function () {
            return new Spreadsheet();
        };

        $this['xlsx.writer'] = function ($c) {
            return new Xlsx(
                $c['spreadsheet']
            );
        };

        $this['spreadsheet.service'] = function ($c) {
            return new SpreadsheetService(
                $c['spreadsheet'],
                $c['xlsx.writer']
            );
        };

        $this['spreadsheet.controller'] = function ($c) {
            return new SpreadsheetController(
                $c['spreadsheet.service']
            );
        };
    }
}
