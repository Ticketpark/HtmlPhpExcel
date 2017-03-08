<?php

namespace HtmlPhpExcel\Tests;

use Ticketpark\HtmlPhpExcel\HtmlPhpExcel;

/**
 * HtmlPhpExcelTest
 *
 * This only tests basic behaviour of the class, not the actual content of the excel files.
 */
class HtmlPhpExcelTest extends \PHPUnit_Framework_TestCase
{
    protected $pathToTestfiles;

    public function setUp()
    {
        $this->pathToTestfiles = __DIR__.'/../../testfiles/';
    }

    public function testSave()
    {
        if (!is_writable($this->pathToTestfiles)) {
            $this->markTestSkipped(
                sprintf('The directory %s must be writable for this test to run.', realpath($this->pathToTestfiles))
            );
        }

        $file = $this->pathToTestfiles.'test.xlsx';
        if (file_exists($file)) {
            if (!is_writable($file)) {
                $this->markTestSkipped(
                    sprintf('The file %s must be writable for this test to run.', realpath($file))
                );
            }
            unlink($file);
        }

        $htmlphpexcel = new HtmlPhpExcel('<table></table>');
        $htmlphpexcel->process()->save($file);

        $this->assertTrue(file_exists($file));

        unlink($file);
    }

    public function testGetExcelObject()
    {
        $htmlphpexcel = new HtmlPhpExcel('<table></table>');
        $excelObject = $htmlphpexcel->process()->getExcelObject();

        $this->assertInstanceOf('\PhpExcel', $excelObject);
    }

    public function testGetHtmlFromFile()
    {
        $htmlphpexcel = new HtmlPhpExcel($this->pathToTestfiles.'test.html');
        $htmlphpexcel->process();

        $this->assertEquals(1, $htmlphpexcel->getDocument()->getTables()->count());
    }

    /**
     * @expectedException \Ticketpark\HtmlPhpExcel\Exception\HtmlPhpExcelException
     */
    public function testExceptionOutputWithoutProcess()
    {
        $htmlphpexcel = new HtmlPhpExcel('<table></table>');
        $htmlphpexcel->output();
    }

    /**
     * @expectedException \Ticketpark\HtmlPhpExcel\Exception\HtmlPhpExcelException
     */
    public function testExceptionSaveWithoutProcess()
    {
        $htmlphpexcel = new HtmlPhpExcel('<table></table>');
        $htmlphpexcel->save('foo.xls');
    }

    /**
     * @expectedException \Ticketpark\HtmlPhpExcel\Exception\HtmlPhpExcelException
     */
    public function testExceptionGetObjectWithoutProcess()
    {
        $htmlphpexcel = new HtmlPhpExcel('<table></table>');
        $htmlphpexcel->getExcelObject();
    }

    /**
     * @expectedException \Ticketpark\HtmlPhpExcel\Exception\HtmlPhpExcelException
     */
    public function testExceptionGetDocumentWithoutProcess()
    {
        $htmlphpexcel = new HtmlPhpExcel('<table></table>');
        $htmlphpexcel->getDocument();
    }
}
