<?php

namespace Ticketpark\HtmlPhpExcel\Parser;

use Ticketpark\HtmlPhpExcel\Elements\Cell;
use Ticketpark\HtmlPhpExcel\Elements\Document;
use Ticketpark\HtmlPhpExcel\Elements\Row;
use Ticketpark\HtmlPhpExcel\Elements\Table;

class Parser {

    /**
     * The html string to be parsed
     */
    private string $html;

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

    public function __construct(string $htmlStringOrFile = null)
    {
        if (null !== $htmlStringOrFile) {
            if (PHP_MAXPATHLEN >= strlen($htmlStringOrFile) && is_file($htmlStringOrFile)) {
                $this->html = file_get_contents($htmlStringOrFile);
            } else {
                $this->html = $htmlStringOrFile;
            }
        }
    }

    public function setTableClass(string $class): self
    {
        $this->tableClass = $class;

        return $this;
    }

    public function setRowClass(string $class): self
    {
        $this->rowClass = $class;

        return $this;
    }

    public function setCellClass(string $class): self
    {
        $this->cellClass = $class;

        return $this;
    }

    public function parse(): Document
    {
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

