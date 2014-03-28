<?php

namespace Ticketpark\HtmlPhpExcel\Parser;

use Ticketpark\HtmlPhpExcel\Elements\Cell;
use Ticketpark\HtmlPhpExcel\Elements\Document;
use Ticketpark\HtmlPhpExcel\Elements\Row;
use Ticketpark\HtmlPhpExcel\Elements\Table;
use Ticketpark\HtmlPhpExcel\Exception\HtmlPhpExcelException;

/**
 * HTML Parser
 *
 * @author Manuel Reinhard <manu@sprain.ch>
 */
class Parser {

    /**
     * The html string to be parsed
     *
     * @var string
     */
    protected $html;

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
     * Constructor
     *
     * @param string|null $htmlStringOrFile
     */
    public function __construct($htmlStringOrFile = null)
    {
        if (null !== $htmlStringOrFile) {
            if (is_file($htmlStringOrFile)) {
                $this->setHtmlFile($htmlStringOrFile);
            } elseif (is_string($htmlStringOrFile)) {
                $this->setHtml($htmlStringOrFile);
            }
        }
    }

    /**
     * Set html as string
     *
     * @param string $html
     * @return $this
     */
    public function setHtml($html)
    {
        $this->html = $html;

        return $this;
    }

    /**
     * Set file containing html content
     *
     * @param string $file
     * @return $this
     */
    public function setHtmlFile($file)
    {
        $this->html = file_get_contents($file);

        return $this;
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
     * Parses the html code
     *
     * @return \Ticketpark\HtmlPhpExcel\Elements\Document
     */
    public function parse()
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

