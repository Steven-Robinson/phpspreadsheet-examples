<?php

namespace App\Component\Spreadsheet\Service;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class SpreadsheetService
{
    const HEADER_COLUMN_RANGE_FORMAT = 'A1:%s1';
    const ROW_CELL_RANGE_FORMAT = 'A%d:%s%d';
    const FULL_SPREADSHEET_RANGE_FORMAT = 'A1:%s%d';

    private $spreadsheet;
    private $writer;
    private $sheet;

    public function __construct(Spreadsheet $spreadsheet, Xlsx $writer)
    {
        $this->spreadsheet = $spreadsheet;
        $this->writer = $writer;
    }

    public function initializeWithData(array $data): void
    {
        $this->sheet = $this->spreadsheet->getActiveSheet();
        $this->sheet->fromArray($data);
    }

    public function getWriter()
    {
        return $this->writer;
    }

    public function setHeaderRowBold(): void
    {
        $headerCellRange = sprintf(
            self::HEADER_COLUMN_RANGE_FORMAT,
            $this->sheet->getHighestDataColumn()
        );

        $this->sheet
            ->getStyle($headerCellRange)
            ->getFont()
            ->setBold(true);
    }

    public function colourRowConditionallyFromCellValue($columnIdentifier, $checkForCellValue, $rowBgColor): void
    {
        foreach ($this->sheet->getRowIterator() as $rowKey => $row) {
            $cellToCheck = $columnIdentifier . $rowKey;
            $rowCellRange = sprintf(
                self::ROW_CELL_RANGE_FORMAT,
                $rowKey,
                $this->sheet->getHighestDataColumn(),
                $rowKey
            );

            if ($this->sheet->getCell($cellToCheck)->getValue() === $checkForCellValue) {
                $this->setRowColour($rowCellRange, $rowBgColor);
            }
        }
    }

    public function styleBorders($borderColor = '00FFFFFF', $borderStyle = Border::BORDER_THIN): void
    {
        $borderStyles = [
            'borders' => [
                'allBorders' => [
                    'borderStyle' => $borderStyle,
                    'color' => ['argb' => $borderColor],
                ],
            ],
        ];

        $cellRange = sprintf(
            self::FULL_SPREADSHEET_RANGE_FORMAT,
            $this->sheet->getHighestDataColumn(),
            $this->sheet->getHighestDataRow()
        );

        $this->sheet
            ->getStyle($cellRange)
            ->applyFromArray($borderStyles);
    }

    public function outputDownloadHeaders(string $filename)
    {
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="' . $filename . '"');
        header('Cache-Control: max-age=0');
    }

    public function setRowColour($cellRange, $colour): void
    {
        $this->sheet
            ->getStyle($cellRange)
            ->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->getStartColor()
            ->setARGB($colour);
    }
}