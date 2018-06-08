<?php

namespace App\Component\Spreadsheet\Controller;

use App\Component\Spreadsheet\Service\SpreadsheetService;

class SpreadsheetController
{
    /**
     * @var SpreadsheetService
     */
    private $spreadsheet;

    public function __construct(SpreadsheetService $spreadsheetService)
    {
        $this->spreadsheet = $spreadsheetService;
    }

    public function index()
    {
        $data = [
            ['one', 'two', 'three', 'four', 'five'],
            ['six', 'seven', 'eight', 'nine', 'ten'],
            ['eleven', 'twelve', 'thirteen', 'fourteen', 'fifteen'],
            ['sixteen', 'seventeen', 'eighteen', 'nineteen', 'twenty'],
            ['twenty one', 'twenty two', 'twenty three', 'twenty four', 'twenty five'],
        ];

        $this->spreadsheet->initializeWithData($data);
        $this->spreadsheet->setHeaderRowBold();
        $this->spreadsheet->styleBorders('00000000');
        $this->spreadsheet->colourRowConditionallyFromCellValue('C', 'thirteen', '00CCCCCC');
        $this->spreadsheet->outputDownloadHeaders($filename = 'report.xlsx');
        $this->spreadsheet->setAutoColumnWidths();

        $this->spreadsheet->getWriter()->save('php://output');
    }
}
