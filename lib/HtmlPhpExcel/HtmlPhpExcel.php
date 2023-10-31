<?php

namespace Ticketpark\HtmlPhpExcel;

use avadim\FastExcelWriter\Excel;
use Ticketpark\HtmlPhpExcel\Elements as HtmlPhpExcelElement;
use Ticketpark\HtmlPhpExcel\Exception\InexistentExcelObjectException;
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
     * The default styles to be applied to all excel cells
     */
    private array $defaultStyles = [];

    /**
     * The default styles additionally to be applied to header cells (<th>)
     */
    private array $defaultHeaderStyles = [];

    /**
     * The document instance which contains the parsed html elements
     */
    private HtmlPhpExcelElement\Document $document;

    /**
     * The instance of the Excel creator used within this library.
     */
    private ?Excel $excel = null;

    public function __construct(
        private string $htmlStringOrFile
    ) {
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

    public function setDefaultStyles(array $defaultStyles): self
    {
        $this->defaultStyles = $defaultStyles;

        return $this;
    }

    public function setDefaultHeaderStyles(array $defaultHeaderStyles): self
    {
        $this->defaultHeaderStyles = $defaultHeaderStyles;

        return $this;
    }

    public function process(?Excel $excel = null): self
    {
        $this->excel = $excel;
        if (null === $this->excel) {
            $this->excel = Excel::create();
        }

        $this->parseHtml();
        $this->createExcel();

        return $this;
    }

    public function download(string $filename): void
    {
        $filename = str_ireplace('.xlsx', '', $filename);
        $this->getExcelObject()->download($filename . '.xlsx');
    }

    public function save(string $filename): bool
    {
        $filename = str_ireplace('.xlsx', '', $filename);

        return $this->getExcelObject()->save($filename . '.xlsx');
    }

    public function getExcelObject(): Excel
    {
        if (null === $this->excel) {
            throw new InexistentExcelObjectException('You must run process() before handling the excel object. ');
        }

        return $this->excel;
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
                $rowStyles = $this->getStyles($row);
                $sheet->setRowStyles(
                    $rowIndex,
                    empty($rowStyles) ? null : $rowStyles
                );

                // Loop over all cells in a row
                $colIndex = 1;
                foreach($row->getCells() as $cell) {
                    $cellStyles = $this->getStyles($cell);
                    $sheet->writeCell(
                        trim($cell->getValue()),
                        empty($cellStyles) ? null : $cellStyles
                    );

                    if (isset($cellStyles['width'])) {
                        $sheet->setColWidth($colIndex, $cellStyles['width']);
                    }

                    if (isset($cellStyles['height'])) {
                        $sheet->setRowHeight($rowIndex, $cellStyles['height']);
                    }

                    $cellComment = $cell->getAttribute('_excel-comment');
                    if ($cellComment) {
                        $sheet->addNote(Excel::cellAddress($rowIndex, $colIndex), $cellComment);
                    }

                    $colIndex++;
                }

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

        return array_merge(
            $this->defaultStyles,
            ($documentElement instanceof HtmlPhpExcelElement\Cell && $documentElement->isHeader()) ? $this->defaultHeaderStyles : [],
            $styles
        );
    }
}
