<?php

namespace HtmlPhpExcel\Tests\Parser;

use Ticketpark\HtmlPhpExcel\Parser\Parser;

class ParserTest extends \PHPUnit_Framework_TestCase
{
    public function setUp()
    {
        $this->pathToTestfiles = __DIR__.'/../../../testfiles/';
    }

    public function testSimpleTable()
    {
        $parser = new Parser('<table><tr><td>row1cell1</td><td>row1cell2</td></tr><tr><td>row2cell1</td><td>row2cell2</td></tr></table>');
        $document = $parser->parse();

        $this->assertEquals(1, $document->getTables()->count());
        $this->assertEquals(2, $document->getTables()->current()->getRows()->count());

        foreach($document->getTables()->current()->getRows() as $row){
            $this->assertEquals(2, $row->getCells()->count());
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

        $this->assertEquals(2, $document->getTables()->count());
        foreach($document->getTables() as $table){
            $this->assertEquals(2, $table->getRows()->count());

            foreach($table->getRows() as $row){
                $this->assertEquals(2, $row->getCells()->count());
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

        $this->assertEquals(1, $document->getTables()->count());
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
            $this->assertEquals(1, $table->getRows()->count());
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
                $this->assertEquals(1, $row->getCells()->count());
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

        $this->assertEquals(1, $document->getTables()->count());
        foreach($document->getTables() as $table){
            $this->assertEquals(1, $table->getRows()->count());

            foreach($table->getRows() as $row){
                $this->assertEquals(1, $row->getCells()->count());
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

        $this->assertEquals(1, $document->getTables()->count());
        foreach($document->getTables() as $table){
            $this->assertEquals(1, $table->getRows()->count());

            foreach($table->getRows() as $row){
                $this->assertEquals(1, $row->getCells()->count());
            }
        }
    }

    public function testFindsHeaders()
    {
        $parser = new Parser('<table><tr><th>row1cell1</th><th>row1cell2</th></tr><tr><td>row2cell1</td><td>row2cell2</td></tr></table>');
        $document = $parser->parse();

        foreach($document->getTables()->current()->getRows() as $key => $row){
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

    public function testSetHtml()
    {
        $parser = new Parser();
        $parser->setHtml('<table></table>');
        $document = $parser->parse();
        $this->assertEquals(1, $document->getTables()->count());
    }

    public function testSetHtmlFile()
    {
        $parser = new Parser();
        $parser->setHtmlFile($this->pathToTestfiles.'test.html');
        $document = $parser->parse();
        $this->assertEquals(1, $document->getTables()->count());
    }

    /**
     * @expectedException \Ticketpark\HtmlPhpExcel\Exception\HtmlPhpExcelException
     */
    public function testExceptionWithoutHtmlContent()
    {
        $parser = new Parser();
        $parser->parse();
    }
}
