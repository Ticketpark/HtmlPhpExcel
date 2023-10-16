<?php

namespace HtmlPhpExcel\Tests\Parser;

use PHPUnit\Framework\TestCase;
use Ticketpark\HtmlPhpExcel\Parser\Parser;

class ParserTest extends TestCase
{
    private string $pathToTestfiles = __DIR__.'/../../../testfiles';

    public function testSimpleTable()
    {
        $parser = new Parser('<table><tr><td>row1cell1</td><td>row1cell2</td></tr><tr><td>row2cell1</td><td>row2cell2</td></tr></table>');
        $document = $parser->parse();

        $this->assertEquals(1, count($document->getTables()));
        $this->assertEquals(2, count($document->getTables()[0]->getRows()));

        foreach($document->getTables()[0]->getRows() as $row){
            $this->assertEquals(2, count($row->getCells()));
        }
    }

    public function testMultipleTables()
    {
        $parser = new Parser('
            <table><tr><td>row1cell1</td><td>row1cell2</td></tr><tr><td>row2cell1</td><td>row2cell2</td></tr></table>
            <p>someotherstuff</p>
            <table><tr><td>row1cell1</td><td>row1cell2</td></tr><tr><td>row2cell1</td><td>row2cell2</td></tr></table>
        ');
        $document = $parser->parse();

        $this->assertEquals(2, count($document->getTables()));
        foreach($document->getTables() as $table){
            $this->assertEquals(2, count($table->getRows()));

            foreach($table->getRows() as $row){
                $this->assertEquals(2, count($row->getCells()));
            }
        }
    }

    public function testMultipleTablesWithTableClass()
    {
        $parser = new Parser('
            <table><tr><td>row1cell1</td><td>row1cell2</td></tr><tr><td>row2cell1</td><td>row2cell2</td></tr></table>
            <p>someotherstuff</p>
            <table class="pickme"><tr><td>row1cell1</td><td>row1cell2</td></tr><tr><td>row2cell1</td><td>row2cell2</td></tr></table>
        ');
        $document = $parser->setTableClass('pickme')->parse();

        $this->assertEquals(1, count($document->getTables()));
    }

    public function testMultipleTablesWithRowClass()
    {
        $parser = new Parser('
            <table><tr class="pickme"><td>row1cell1</td><td>row1cell2</td></tr><tr><td>row2cell1</td><td>row2cell2</td></tr></table>
            <p>someotherstuff</p>
            <table><tr class="pickme"><td>row1cell1</td><td>row1cell2</td></tr><tr><td>row2cell1</td><td>row2cell2</td></tr></table>
        ');
        $document = $parser->setRowClass('pickme')->parse();

        foreach($document->getTables() as $table){
            $this->assertEquals(1, count($table->getRows()));
        }
    }

    public function testMultipleTablesWithCellClass()
    {
        $parser = new Parser('
            <table><tr><td>row1cell1</td><td class="pickme">row1cell2</td></tr><tr><td class="pickme">row2cell1</td><td>row2cell2</td></tr></table>
            <p>someotherstuff</p>
            <table><tr><td>row1cell1</td><td class="pickme">row1cell2</td></tr><tr><td class="pickme">row2cell1</td><td>row2cell2</td></tr></table>
        ');
        $document = $parser->setCellClass('pickme')->parse();

        foreach($document->getTables() as $table){
            foreach($table->getRows() as $row){
                $this->assertEquals(1, count($row->getCells()));
            }
        }
    }

    public function testMultipleTablesWithMixedClasses()
    {
        $parser = new Parser('
            <table class="pickme"><tr class="pickme"><td>row1cell1</td><td class="pickme">row1cell2</td></tr><tr><td>row2cell1</td><td>row2cell2</td></tr></table>
            <p>someotherstuff</p>
            <table><tr><td>row1cell1</td><td>row1cell2</td></tr><tr><td>row2cell1</td><td>row2cell2</td></tr></table>
        ');
        $document = $parser
            ->setTableClass('pickme')
            ->setRowClass('pickme')
            ->setCellClass('pickme')
            ->parse();

        $this->assertEquals(1, count($document->getTables()));
        foreach($document->getTables() as $table){
            $this->assertEquals(1, count($table->getRows()));

            foreach($table->getRows() as $row){
                $this->assertEquals(1, count($row->getCells()));
            }
        }
    }

    public function testMultipleTablesWithMixedClassesAndOtherClasses()
    {
        $parser = new Parser('
            <table class="foo pickme"><tr class="pickme bar"><td>row1cell1</td><td class="foo pickme bar">row1cell2</td></tr><tr><td>row2cell1</td><td>row2cell2</td></tr></table>
            <p>someotherstuff</p>
            <table><tr><td>row1cell1</td><td>row1cell2</td></tr><tr><td>row2cell1</td><td>row2cell2</td></tr></table>
        ');
        $document = $parser
            ->setTableClass('pickme')
            ->setRowClass('pickme')
            ->setCellClass('pickme')
            ->parse();

        $this->assertEquals(1, count($document->getTables()));
        foreach($document->getTables() as $table){
            $this->assertEquals(1, count($table->getRows()));

            foreach($table->getRows() as $row){
                $this->assertEquals(1, count($row->getCells()));
            }
        }
    }

    public function testFindsHeaders()
    {
        $parser = new Parser('<table><tr><th>row1cell1</th><th>row1cell2</th></tr><tr><td>row2cell1</td><td>row2cell2</td></tr></table>');
        $document = $parser->parse();

        foreach($document->getTables()[0]->getRows() as $key => $row){
            foreach($row->getCells() as $cell){
                if (0 === $key) {
                    $this->assertTrue($cell->isHeader());
                } elseif (1 === $key) {
                    $this->assertFalse($cell->isHeader());
                }
            }
        }
    }

    public function testFindsAttributes()
    {
        $parser = new Parser('<table bar="foo"><tr bar="foo"><td bar="foo">row1cell1</td><td>row1cell2</td></tr><tr><td>row2cell1</td><td>row2cell2</td></tr></table>');
        $document = $parser->parse();

        foreach($document->getTables() as $table){
            $this->assertEquals('foo', $table->getAttribute('bar'));

            foreach($table->getRows() as $rowKey => $row){
                if (0 === $rowKey) {
                    $this->assertEquals('foo', $row->getAttribute('bar'));
                } else {
                    $this->assertEquals(null, $row->getAttribute('bar'));
                }

                foreach($row->getCells() as $cellKey => $cell){
                    if (0 === $rowKey && 0 == $cellKey) {
                        $this->assertEquals('foo', $cell->getAttribute('bar'));
                    } else {
                        $this->assertEquals(null, $cell->getAttribute('bar'));
                    }
                }
            }
        }
    }

    public function testWithFile()
    {
        $parser = new Parser($this->pathToTestfiles . '/test.html');
        $document = $parser->parse();
        $this->assertEquals(1, count($document->getTables()));
    }
}
