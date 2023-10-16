<?php

namespace Ticketpark\HtmlPhpExcel;

use avadim\FastExcelWriter\Excel;
use Ticketpark\HtmlPhpExcel\Elements as HtmlPhpExcelElement;
use Ticketpark\HtmlPhpExcel\Exception\HtmlPhpExcelException;
use Ticketpark\HtmlPhpExcel\Parser\Parser;

class HtmlPhpExcel
{
    /**
     * The class attribute the tables must have to be parsed
     * Default is none, which results in all tables.
     */
    private ?string $tableClass = null;

    /**
     * The class attribute the rows (<tr>) must have to be parsed.
     * Default is none, which results in all rows of a parsed table.
     */
    private ?string $rowClass = null;

    /**
     * The class attribute the rows (<td> or <th>) must have to be parsed.
     * Default is none, which results in all cells of a parsed row.
     */
    private ?string $cellClass = null;

    /**
     * The document instance which contains the parsed html elements
     */
    private HtmlPhpExcelElement\Document $document;

    /**
     * The instance of the Excel creator used within this library.
     */
    private Excel $excel;

    public function __construct(
        private string $htmlStringOrFile
    ) {
        $this->excel = Excel::create();
    }

    public function setTableClass(?string $class): self
    {
        $this->tableClass = $class;

        return $this;
    }

    public function setRowClass(?string $class): self
    {
        $this->rowClass = $class;

        return $this;
    }

    public function setCellClass(?string $class): self
    {
        $this->cellClass = $class;

        return $this;
    }

    public function process(): self
    {
        $this->parseHtml();
        $this->createExcel();

        return $this;
    }

    public function download(string $filename): void
    {
        $filename = str_ireplace('.xlsx', '', $filename);
        $this->excel->download($filename . '.xlsx');
    }

    public function save(string $filename): void
    {
        $filename = str_ireplace('.xlsx', '', $filename);
        $this->excel->save($filename . '.xlsx');
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
        foreach($this->document->getTables() as $table) {

            // Handle worksheets
            $this->excel->makeSheet($table->getAttribute('_excel-name'));
            $sheet = $this->excel->sheet();

            // Loop over all rows
            $rowIndex = 1;
            foreach($table->getRows() as $row) {

                // Loop over all cells in a row
                $colIndex = 1;
                foreach($row->getCells() as $cell) {
                    $styles = $this->getStyles($cell);

                    $sheet->writeCell(
                        trim($cell->getValue()),
                        $styles
                    );

                    if (isset($styles['width'])) {
                        $sheet->setColWidth($colIndex, $styles['width']);
                    }

                    if (isset($styles['height'])) {
                        $sheet->setRowHeight($rowIndex, $styles['height']);
                    }

                    $colIndex++;
                }

                $sheet->setRowStyles($rowIndex, $this->getStyles($row));
                $sheet->nextRow();
                $rowIndex++;
            }
        }
    }

    private function getStyles(HtmlPhpExcelElement\Element $documentElement): array
    {
        $styles = [];

        if ($attributeStyles = $documentElement->getAttribute('_excel-styles')) {
            if (!is_array($attributeStyles)) {
                $decodedJson = json_decode($attributeStyles, true, 512, JSON_THROW_ON_ERROR);
                if (null !== $decodedJson) {
                    $attributeStyles = $decodedJson;
                }
            }
        }

        if (is_array($attributeStyles)) {
            $styles = $attributeStyles;
        }

        return $styles;
    }
}
