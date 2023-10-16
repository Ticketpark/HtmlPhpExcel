<?php

namespace HtmlPhpExcel\Tests;

use PHPUnit\Framework\TestCase;
use Ticketpark\HtmlPhpExcel\HtmlPhpExcel;

/**
 * This only tests basic behaviour of the class, not the actual content of the excel files.
 */
class HtmlPhpExcelTest extends TestCase
{
    private string $pathToTestfiles;

    public function setUp(): void
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
}
