<?php

namespace Ticketpark\HtmlPhpExcel;

use Ticketpark\HtmlPhpExcel\Elements\Document;
use Ticketpark\HtmlPhpExcel\Elements\Cell;
use Ticketpark\HtmlPhpExcel\Elements\Row;
use Ticketpark\HtmlPhpExcel\Exception\HtmlPhpExcelException;
use Ticketpark\HtmlPhpExcel\Parser\Parser;

/**
 * HtmlPhpExcel
 *
 * @author Manuel Reinhard <manu@sprain.ch>
 */
class HtmlPhpExcel
{
    /**
     * The string or file containing the html to be parsed
     *
     * @var string
     */
    protected $htmlStringOrFile;

    /**
     * The class attribute the tables must have to be parsed
     * Default is none, which results in all tables.
     *
     * @var string
     */
    protected $tableClass;

    /**
     * The class attribute the rows (<tr>) must have to be parsed.
     * Default is none, which results in all rows of a parsed table.
     *
     * @var string
     */
    protected $rowClass;

    /**
     * The class attribute the rows (<td> or <th>) must have to be parsed.
     * Default is none, which results in all cells of a parsed row.
     *
     * @var string
     */
    protected $cellClass;

    /**
     * The PHPExcel instance generated with this class
     *
     * @var \PHPExcel
     */
    protected $phpexcel;

    /**
     * The document instance which contains the parsed html elements
     *
     * @var \Ticketpark\HtmlPhpExcel\Elements\Document
     */
    protected $document;

    /**
     * Determines if the values should be encoded in some way before writing to the excel cell
     *
     * @var null|string
     */
    protected $changeEncoding;

    /**
     * Constructor
     *
     * @param string|null $htmlStringOrFile
     */
    public function __construct($htmlStringOrFile)
    {
        $this->htmlStringOrFile = $htmlStringOrFile;
    }

    /**
     * Set html class of tables (<table>) to be parsed
     *
     * @param string $class
     * @return $this
     */
    public function setTableClass($class)
    {
        $this->tableClass = $class;

        return $this;
    }

    /**
     * Set html class of rows (<tr>) within tables to be parsed
     *
     * @param $class
     * @return $this
     */
    public function setRowClass($class)
    {
        $this->rowClass = $class;

        return $this;
    }

    /**
     * Set html class of cells (<td> or <th>) within rows to be parsed
     *
     * @param string $class
     * @return $this
     */
    public function setCellClass($class)
    {
        $this->cellClass = $class;

        return $this;
    }

    /**
     * Let's put things together!
     *
     * @return $this
     */
    public function process()
    {
        $this->parseHtml();
        $this->createExcel();

        return $this;
    }

    /**
     * Get the PHPExcel object
     *
     * @return \PHPExcel
     */
    public function getExcelObject()
    {
        if (!$this->phpexcel instanceof \PHPExcel) {
            throw new HtmlPhpExcelException('You must run process() first to create a phpexcel instance');
        }

        return $this->phpexcel;
    }

    /**
     * Output the created excel file
     *
     * @param string $filename The name of the output file
     * @param string $excelWriterType Excel file type
     * @throws Exception\HtmlPhpExcelException
     */
    public function output($filename = 'excel.xls', $excelWriterType = 'Excel2007')
    {
        if (!$this->phpexcel instanceof \PHPExcel) {
            throw new HtmlPhpExcelException('You must run process() first to create a phpexcel instance');
        }

        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="'.$filename.'"');
        header('Cache-Control: max-age=1');

        $writer = \PHPExcel_IOFactory::createWriter($this->phpexcel, $excelWriterType);
        $writer->save('php://output');
    }

    /**
     * Save the created excel file
     *
     * @param string $excelWriterType
     * @throws Exception\HtmlPhpExcelException
     */
    public function save($file, $excelWriterType = 'Excel2007')
    {
        if (!$this->phpexcel instanceof \PHPExcel) {
            throw new HtmlPhpExcelException('You must run process() first to create a phpexcel instance');
        }

        $writer = \PHPExcel_IOFactory::createWriter($this->phpexcel, $excelWriterType);
        $writer->save($file);

        return $this;
    }

    /**
     * Get the Document instance
     *
     * @return \Ticketpark\HtmlPhpExcel\Elements\Document
     */
    public function getDocument()
    {
        if (!$this->phpexcel instanceof \PHPExcel) {
            throw new HtmlPhpExcelException('You must run process() first to get ');
        }

        return $this->document;
    }

    /**
     * UTF8-encode values before writing to excel cell
     */
    public function utf8EncodeValues()
    {
        $this->changeEncoding = 'utf8_encode';
    }

    /**
     * UTF8-decode values before writing to excel cell
     */
    public function utf8DecodeValues()
    {
        $this->changeEncoding = 'utf8_decode';
    }

    /**
     * Parse the html and return document
     *
     * @return \Ticketpark\HtmlPhpExcel\Elements\Document
     */
    protected function parseHtml()
    {
        $parser = new Parser($this->htmlStringOrFile);
        $document = $parser->setTableClass($this->tableClass)
            ->setRowClass($this->rowClass)
            ->setCellClass($this->cellClass)
            ->parse();

        $this->document = $document;

        return $document;
    }

    /**
     * Create excel from document
     *
     * @return \PHPExcel
     */
    protected function createExcel()
    {
        $this->phpexcel = new \PHPExcel();
        $tableNumber = 0;

        // Loop over all tables in document
        foreach($this->document->getTables() as $table){

            // Handle worksheets
            if ($tableNumber > 0) {
                $this->phpexcel->createSheet();
            }
            $excelWorksheet = $this->phpexcel->setActiveSheetIndex($tableNumber);
            if ($sheetTitle = $table->getAttribute('_excel-name')) {
                $excelWorksheet->setTitle($sheetTitle);
            }

            // Loop over all rows
            $rowNumber = 1;
            foreach($table->getRows() as $row){

                $excelWorksheet->getStyle($rowNumber.':'.$rowNumber)->applyFromArray($this->getRowStylesArray($row));
                $this->setDimensions($excelWorksheet, $excelWorksheet->getRowIterator($rowNumber)->current(), $row);

                // Loop over all cells in row
                $cellNumber = 0;
                foreach($row->getCells() as $cell){
                    $excelCellIndex = \PHPExcel_Cell::stringFromColumnIndex($cellNumber).$rowNumber;

                    // Set value
                    if ($explicitCellType = $cell->getAttribute('_excel-explicit') || $explicitCellType = $row->getAttribute('_excel-explicit')) {
                        $excelWorksheet->setCellValueExplicit(
                            $excelCellIndex,
                            $this->changeValueEncoding($cell->getValue()),
                            $this->convertStaticPhpExcelConstantsFromStringsToConstants($explicitCellType)
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
                        $mergeCellsTargetCellIndex = \PHPExcel_Cell::stringFromColumnIndex($cellNumber + $colspan).($rowNumber + $rowspan);
                        $excelWorksheet->mergeCells($excelCellIndex.':'.$mergeCellsTargetCellIndex);
                    }

                    // Set styles
                    $excelWorksheet->getStyle($excelCellIndex)->applyFromArray($this->getCellStylesArray($cell));
                    $this->setDimensions($excelWorksheet, $excelWorksheet->getCell($excelCellIndex), $cell);

                    $cellNumber++;
                }

                $rowNumber++;
            }

            $tableNumber++;
        }

        return $this->phpexcel;
    }

    /**
     * Set dimensions of row or column
     *
     * @param \PHPExcel_Worksheet $excelWorksheet
     * @param $excelElement
     * @param $documentElement
     */
    protected function setDimensions(\PHPExcel_Worksheet $excelWorksheet, $excelElement, $documentElement)
    {
        $dimensions = $this->getDimensionsArray($documentElement);

        if (isset($dimensions['column']) && $excelElement instanceof \PHPExcel_Cell) {
            foreach($dimensions['column'] as $columnKey => $columnValue) {
                $method = 'set'.ucfirst($columnKey);
                $excelWorksheet->getColumnDimension($excelElement->getColumn())->$method($columnValue);
            }
        }

        if (isset($dimensions['row'])) {
            foreach($dimensions['row'] as $rowKey => $rowValue) {
                $method = 'set'.ucfirst($rowKey);
                if ($excelElement instanceof \PHPExcel_Cell) {
                    $excelWorksheet->getRowDimension($excelElement->getRow())->$method($rowValue);
                } elseif ($excelElement instanceof \PHPExcel_Worksheet_Row) {
                    $excelWorksheet->getRowDimension($excelElement->getRowIndex())->$method($rowValue);
                }

            }
        }
    }

    /**
     * Prepare styles array for a cell
     *
     * @param Cell $cell
     * @param \PHPExcel_Cell $cell
     */
    protected function getRowStylesArray(Row $row)
    {
        return $this->getStylesArray($row);
    }

    /**
     * Prepare styles array for a cell
     *
     * @param Cell $cell
     * @param \PHPExcel_Cell $cell
     */
    protected function getCellStylesArray(Cell $cell)
    {
        $styles = $this->getStylesArray($cell);

        if ($cell->isHeader()) {
            $styles['font']['bold'] = true;
        }

        return $styles;
    }

    /**
     * Get the styles array for any element
     *
     * @param $documentElement
     * @return array
     */
    protected function getStylesArray($documentElement)
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

    /**
     * Get the styles array for any element
     *
     * @param $documentElement
     * @return array
     */
    protected function getDimensionsArray($documentElement)
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

    /**
     * Sanitize styles array
     *
     * @param array $styles
     * @return array
     */
    protected function sanitizeArray($array)
    {
        foreach($array as $key => $value){
            if(is_array($value)){
                $array[$key] = $this->sanitizeArray($value);
            } else {
                $array[$key] = $this->convertStaticPhpExcelConstantsFromStringsToConstants($value);
            }
        }

        return $array;
    }

    /**
     * Turn PHPExcel constants into actual constants
     *
     * Example:
     * If the html element contains a _excel-styles attribute with the json-encoded version of the array below,
     * the value PHPExcel_Style_Fill::FILL_SOLID would be treated as a string.
     * We need to treat it as a static class constant and also apply the correct namespace.
     *
     * array (
     *   'fill' => array (
     *     'type' => 'PHPExcel_Style_Fill::FILL_SOLID',
     *     'color' => array (
     *       'rgb' => '4F4F4F',
     *     ),
     *   ),
     * )
     *
     * @param string $value
     * @return string
     */
    protected function convertStaticPhpExcelConstantsFromStringsToConstants($value)
    {
        if (strpos($value, 'PHPExcel_') === 0) {
            $parts = explode('::', $value);
            $class = new \ReflectionClass($parts[0]);
            $value = $class->getConstant($parts[1]);
        }

        return $value;
    }

    /**
     * Apply modifications to value before writing to excel cell
     *
     * @param mixed $value
     */
    protected function changeValueEncoding($value)
    {
        if (null !== $this->changeEncoding) {
            $value = call_user_func($this->changeEncoding, $value);
        }

        return $value;
    }
}
