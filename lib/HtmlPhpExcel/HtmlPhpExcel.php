<?php

namespace Ticketpark\HtmlPhpExcel;

use PhpOffice\PhpSpreadsheet\Cell\Cell;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Row;
use Ticketpark\HtmlPhpExcel\Elements as HtmlPhpExcelElement;
use Ticketpark\HtmlPhpExcel\Exception\HtmlPhpExcelException;
use Ticketpark\HtmlPhpExcel\Parser\Parser;

class HtmlPhpExcel
{
    /**
     * The string or file containing the html to be parsed
     *
     * @var string
     */
    private $htmlStringOrFile;

    /**
     * The class attribute the tables must have to be parsed
     * Default is none, which results in all tables.
     *
     * @var string
     */
    private $tableClass;

    /**
     * The class attribute the rows (<tr>) must have to be parsed.
     * Default is none, which results in all rows of a parsed table.
     *
     * @var string
     */
    private $rowClass;

    /**
     * The class attribute the rows (<td> or <th>) must have to be parsed.
     * Default is none, which results in all cells of a parsed row.
     *
     * @var string
     */
    private $cellClass;

    /**
     * The Spreadsheet instance generated with this class
     *
     * @var Spreadsheet
     */
    private $spreadsheet;

    /**
     * The document instance which contains the parsed html elements
     *
     * @var \Ticketpark\HtmlPhpExcel\Elements\Document
     */
    private $document;

    /**
     * Determines if the values should be encoded in some way before writing to the excel cell
     *
     * @var null|string
     */
    private $changeEncoding;

    public function __construct(string $htmlStringOrFile = null)
    {
        $this->htmlStringOrFile = $htmlStringOrFile;
    }

    public function setTableClass(string $class = null): self
    {
        $this->tableClass = $class;

        return $this;
    }

    public function setRowClass(string $class = null): self
    {
        $this->rowClass = $class;

        return $this;
    }

    public function setCellClass(string $class = null): self
    {
        $this->cellClass = $class;

        return $this;
    }

    public function process(Spreadsheet $spreadsheet = null): self
    {
        if ($spreadsheet) {
            $this->spreadsheet = $spreadsheet;
        } else {
            $this->spreadsheet = new Spreadsheet();
        }

        $this->parseHtml();
        $this->createExcel();

        return $this;
    }

    public function getExcelObject(): Spreadsheet
    {
        if (!$this->spreadsheet instanceof Spreadsheet) {
            throw new HtmlPhpExcelException('You must run process() first to create a PhpSpreadsheet instance');
        }

        return $this->spreadsheet;
    }

    public function output(string $filename = 'excel.xlsx', string $excelWriterType = 'xlsx'): void
    {
        if (!$this->spreadsheet instanceof Spreadsheet) {
            throw new HtmlPhpExcelException('You must run process() first to create a PhpSpreadsheet instance');
        }

        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="'.$filename.'"');
        header('Cache-Control: max-age=1');

        $writer = IOFactory::createWriter($this->spreadsheet, ucfirst($excelWriterType));
        $writer->save('php://output');
    }

    public function save(string $file, string $excelWriterType = 'xlsx'): self
    {
        if (!$this->spreadsheet instanceof Spreadsheet) {
            throw new HtmlPhpExcelException('You must run process() first to create a PhpSpreadsheet instance');
        }

        $writer = IOFactory::createWriter($this->spreadsheet, ucfirst($excelWriterType));
        $writer->save($file);

        return $this;
    }

    public function utf8EncodeValues(): self
    {
        $this->changeEncoding = 'utf8_encode';

        return $this;
    }

    public function utf8DecodeValues(): self
    {
        $this->changeEncoding = 'utf8_decode';

        return $this;
    }

    private function parseHtml(): void
    {
        $parser = new Parser($this->htmlStringOrFile);
        $document = $parser->setTableClass($this->tableClass)
            ->setRowClass($this->rowClass)
            ->setCellClass($this->cellClass)
            ->parse();

        $this->document = $document;
    }

    private function createExcel(): void
    {
        // Loop over all tables in document
        $tableNumber = 0;
        foreach($this->document->getTables() as $table){

            // Handle worksheets
            if ($tableNumber > 0) {
                $this->spreadsheet->createSheet();
            }
            $excelWorksheet = $this->spreadsheet->setActiveSheetIndex($tableNumber);
            if ($sheetTitle = $table->getAttribute('_excel-name')) {
                $excelWorksheet->setTitle($sheetTitle);
            }

            // Loop over all rows
            $rowNumber = $this->getHighestRow($excelWorksheet);
            foreach($table->getRows() as $row){

                $excelWorksheet->getStyle($rowNumber.':'.$rowNumber)->applyFromArray($this->getRowStylesArray($row));
                $this->setDimensionsForRow($excelWorksheet, $excelWorksheet->getRowIterator($rowNumber)->current(), $row);

                // Loop over all cells in row
                $cellNumber = 1;
                foreach($row->getCells() as $cell){
                    $excelCellIndex = Coordinate::stringFromColumnIndex($cellNumber).$rowNumber;

                    // Skip cells withing merge range
                    while ($excelWorksheet->getCell($excelCellIndex)->isInMergeRange()) {
                        $cellNumber++;
                        $excelCellIndex = Coordinate::stringFromColumnIndex($cellNumber).$rowNumber;
                    }

                    // Set value
                    $explicitCellType = $cell->getAttribute('_excel-explicit');
                    if (!$explicitCellType) {
                        $explicitCellType = $row->getAttribute('_excel-explicit');
                    }

                    if ($explicitCellType) {
                        $excelWorksheet->setCellValueExplicit(
                            $excelCellIndex,
                            $this->changeValueEncoding($cell->getValue()),
                            $this->convertStaticPhpSpreadsheetConstantsFromStringsToConstants($explicitCellType)
                        );
                    } else {
                        $excelWorksheet->setCellValue(
                            $excelCellIndex,
                            $this->changeValueEncoding($cell->getValue())
                        );
                    }

                    // Merge cells
                    $colspan = $cell->getAttribute('colspan');
                    $rowspan = $cell->getAttribute('rowspan');

                    if ($colspan || $rowspan) {
                        if ($colspan) {$colspan = $colspan - 1;}
                        if ($rowspan) {$rowspan = $rowspan - 1;}
                        $mergeCellsTargetCellIndex = Coordinate::stringFromColumnIndex($cellNumber + $colspan).($rowNumber + $rowspan);
                        $excelWorksheet->mergeCells($excelCellIndex.':'.$mergeCellsTargetCellIndex);
                    }

                    // Set styles
                    $excelWorksheet->getStyle($excelCellIndex)->applyFromArray($this->getCellStylesArray($cell));
                    $this->setDimensionsForCell($excelWorksheet, $excelWorksheet->getCell($excelCellIndex), $cell);

                    $cellNumber++;
                }

                $rowNumber++;
            }

            $tableNumber++;
        }
    }

    private function setDimensionsForRow(Worksheet $excelWorksheet, Row $excelElement, HtmlPhpExcelElement\Row $row): void
    {
        $dimensions = $this->getDimensionsArray($row);

        if (isset($dimensions['row'])) {
            foreach($dimensions['row'] as $rowKey => $rowValue) {
                $method = 'set'.ucfirst($rowKey);
                if ($excelElement instanceof Cell) {
                    $excelWorksheet->getRowDimension($excelElement->getRow())->$method($rowValue);
                } elseif ($excelElement instanceof Row) {
                    $excelWorksheet->getRowDimension($excelElement->getRowIndex())->$method($rowValue);
                }

            }
        }
    }

    private function setDimensionsForCell(Worksheet $excelWorksheet, Cell $excelElement, HtmlPhpExcelElement\Cell $cell): void
    {
        $dimensions = $this->getDimensionsArray($cell);

        if (isset($dimensions['column'])) {
            foreach($dimensions['column'] as $columnKey => $columnValue) {
                $method = 'set'.ucfirst($columnKey);
                $excelWorksheet->getColumnDimension($excelElement->getColumn())->$method($columnValue);
            }
        }
    }

    private function getRowStylesArray(HtmlPhpExcelElement\Row $row): array
    {
        return $this->getStylesArray($row);
    }

    private function getCellStylesArray(HtmlPhpExcelElement\Cell $cell): array
    {
        $styles = $this->getStylesArray($cell);

        if ($cell->isHeader()) {
            $styles['font']['bold'] = true;
        }

        return $styles;
    }

    private function getStylesArray(HtmlPhpExcelElement\Element $documentElement): array
    {
        $styles = array();

        if ($attributeStyles = $documentElement->getAttribute('_excel-styles')) {
            if (!is_array($attributeStyles)) {
                $decodedJson = json_decode($attributeStyles, true);
                if (null !== $decodedJson) {
                    $attributeStyles = $decodedJson;
                }
            }
        }

        if (is_array($attributeStyles)) {
            $styles = $attributeStyles;
        }

        $styles = $this->sanitizeArray($styles);

        return $styles;
    }

    private function getDimensionsArray(HtmlPhpExcelElement\Element $documentElement): array
    {
        $dimensions = array();

        if ($attributeDimensions= $documentElement->getAttribute('_excel-dimensions')) {
            if (!is_array($attributeDimensions)) {
                $decodedJson = json_decode($attributeDimensions, true);
                if (null !== $decodedJson) {
                    $attributeDimensions = $decodedJson;
                }
            }
        }

        if (is_array($attributeDimensions)) {
            $dimensions = $attributeDimensions;
        }

        $dimensions = $this->sanitizeArray($dimensions);

        return $dimensions;
    }

    private function getHighestRow(Worksheet $excelWorksheet): int
    {
        $highestRow = $excelWorksheet->getHighestRow();

        return $highestRow + ($highestRow > 1);
    }

    private function sanitizeArray(array $array): array
    {
        foreach($array as $key => $value){
            if(is_array($value)){
                $array[$key] = $this->sanitizeArray($value);
            } else {
                $array[$key] = $this->convertStaticPhpSpreadsheetConstantsFromStringsToConstants($value);
            }
        }

        return $array;
    }

    /**
     * Turn Spreadsheet constants into actual constants
     *
     * Example:
     * If the html element contains a _excel-styles attribute with the json-encoded version of the array below,
     * the value PhpSpreadsheet_Style_Fill::FILL_SOLID would be treated as a string.
     * We need to treat it as a static class constant and also apply the correct namespace.
     *
     * array (
     *   'fill' => array (
     *     'type' => 'PhpSpreadsheet_Style_Fill::FILL_SOLID',
     *     'color' => array (
     *       'rgb' => '4F4F4F',
     *     ),
     *   ),
     * )
     *
     * @param string $value
     * @return string
     */
    private function convertStaticPhpSpreadsheetConstantsFromStringsToConstants(string $value)
    {
        if (strpos($value, 'PHPExcel_') === 0 || strpos($value, 'PhpSpreadsheet_') === 0) {
            $parts = explode('::', $value);

            $namespaceParts = explode('_', $parts[0]);
            $fqns = 'PhpOffice\\PhpSpreadsheet';
            unset($namespaceParts[0]);
            foreach($namespaceParts as $namespacePart) {
                $fqns .= '\\' . $namespacePart;
            }

            $class = new \ReflectionClass($fqns);
            $value = $class->getConstant($parts[1]);
        }

        return $value;
    }

    private function changeValueEncoding(string $value): string
    {
        if (null !== $this->changeEncoding) {
            $value = call_user_func($this->changeEncoding, $value);
        }

        return $value;
    }
}
