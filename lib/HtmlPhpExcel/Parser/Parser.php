<?php

namespace Ticketpark\HtmlPhpExcel\Parser;

use Ticketpark\HtmlPhpExcel\Elements\Cell;
use Ticketpark\HtmlPhpExcel\Elements\Document;
use Ticketpark\HtmlPhpExcel\Elements\Row;
use Ticketpark\HtmlPhpExcel\Elements\Table;
use Ticketpark\HtmlPhpExcel\Exception\HtmlPhpExcelException;

class Parser {

    /**
     * The html string to be parsed
     *
     * @var string
     */
    private $html;

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

    public function __construct(string $htmlStringOrFile = null)
    {
        if (null !== $htmlStringOrFile) {
            if (PHP_MAXPATHLEN >= strlen($htmlStringOrFile) && is_file($htmlStringOrFile)) {
                $this->setHtmlFile($htmlStringOrFile);
            } elseif (is_string($htmlStringOrFile)) {
                $this->setHtml($htmlStringOrFile);
            }
        }
    }

    public function setHtml(string $html): self
    {
        $this->html = $html;

        return $this;
    }

    public function setHtmlFile(string $file): self
    {
        $this->html = file_get_contents($file);

        return $this;
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

    public function parse(): Document
    {
        if (null === $this->html) {
            throw new HtmlPhpExcelException('You must provide html content first. Use setHtml() or setHtmlFile().');
        }

        $dom = new \DOMDocument();
        $dom->loadHTML($this->html);

        $xpath = new \DOMXPath($dom);
        $htmlTables = $xpath->query('.//table[contains(concat(" ", normalize-space(@class), " "), "'.$this->tableClass.'")]');

        $document = new Document();

        foreach ($htmlTables as $htmlTable) {

            $table = new Table();

            $htmlRows = $xpath->query('.//tr[contains(concat(" ", normalize-space(@class), " "), "'.$this->rowClass.'")]', $htmlTable);
            foreach($htmlRows as $htmlRow) {

                $row = new Row();
                $htmlCells = $xpath->query(
                     './/td[contains(concat(" ", normalize-space(@class), " "), "'.$this->cellClass.'")]
                    | .//th[contains(concat(" ", normalize-space(@class), " "), "'.$this->cellClass.'")]',
                    $htmlRow
                );

                foreach($htmlCells as $htmlCell) {
                    $cell = new Cell();
                    $cell->setValue($htmlCell->nodeValue);

                    foreach ($htmlCell->attributes as $attribute) {
                        $cell->addAttribute($attribute->name, $attribute->value);
                    }

                    if ('th' == $htmlCell->nodeName) {
                        $cell->setIsHeader(true);
                    }

                    $row->addCell($cell);
                }

                foreach ($htmlRow->attributes as $attribute) {
                    $row->addAttribute($attribute->name, $attribute->value);
                }

                $table->addRow($row);
            }

            foreach ($htmlTable->attributes as $attribute) {
                $table->addAttribute($attribute->name, $attribute->value);
            }

            $document->addTable($table);
        }

        return $document;
    }
}

