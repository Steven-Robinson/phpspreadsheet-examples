<?php

namespace App\Component\Controller\Test;

use App\Component\Spreadsheet\Controller\SpreadsheetController;
use App\Component\Spreadsheet\Service\SpreadsheetService;
use Mockery;
use PHPUnit\Framework\TestCase;
use PhpOffice\PhpSpreadsheet\Writer\Exception as WriterException;

class SpreadsheetControllerTest extends TestCase
{
    private $data;
    private $spreadsheetServiceMock;
    private $sut;

    public function setUp()
    {
        parent::setUp();

        $this->spreadsheetServiceMock = Mockery::mock(SpreadsheetService::class);

        $this->sut = new SpreadsheetController(
            $this->spreadsheetServiceMock
        );

        $this->data = [
            ['one', 'two', 'three', 'four', 'five'],
            ['six', 'seven', 'eight', 'nine', 'ten'],
            ['eleven', 'twelve', 'thirteen', 'fourteen', 'fifteen'],
            ['sixteen', 'seventeen', 'eighteen', 'nineteen', 'twenty'],
            ['twenty one', 'twenty two', 'twenty three', 'twenty four', 'twenty five'],
        ];
    }

    public function testSuccessfulIndexWorkflow()
    {
        $this->createAndFormatSpreadsheet();

        $this->spreadsheetServiceMock
            ->shouldReceive('getWriter->save')
            ->once();

        self::assertTrue($this->sut->index());
    }

    public function testExceptionIsThrownOnError()
    {
        $this->createAndFormatSpreadsheet();

        $this->spreadsheetServiceMock
            ->shouldReceive('getWriter->save')
            ->once()
            ->andThrow(WriterException::class);

        self::assertFalse($this->sut->index());
    }

    private function createAndFormatSpreadsheet(): void
    {
        $this->spreadsheetServiceMock
            ->shouldReceive('initializeWithData')
            ->once()
            ->with($this->data)
            ->shouldReceive('setHeaderRowBold')
            ->once()
            ->shouldReceive('styleBorders')
            ->once()
            ->with('00000000')
            ->shouldReceive('colourRowConditionallyFromCellValue')
            ->once()
            ->with(
                'C',
                'thirteen',
                '00CCCCCC'
            )
            ->shouldReceive('setAutoColumnWidths')
            ->once()
            ->shouldReceive('freezeHeaders')
            ->once()
            ->shouldReceive('outputDownloadHeaders')
            ->once();
    }
}