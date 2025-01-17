<?php
/*
The MIT License (MIT)

Copyright (c) 2015 PortPHP

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
 */
namespace Port\Spreadsheet\Tests;

use PhpOffice\PhpSpreadsheet\IOFactory;
use Port\Spreadsheet\SpreadsheetWriter;

/**
 * {@inheritDoc}
 */
class SpreadsheetWriterTest extends \PHPUnit_Framework_TestCase
{
    /**
     * {@inheritDoc}
     */
    public function setUp()
    {
        if (!extension_loaded('zip')) {
            static::markTestSkipped();
        }
    }

    /**
     * Test that column names not prepended to first row if SpreadsheetWriter's 4-th
     * parameter not given
     *
     * @author  Igor Mukhin <igor.mukhin@gmail.com>
     */
    public function testHeaderNotPrependedByDefault()
    {
        $file = tempnam(sys_get_temp_dir(), null);

        $writer = new SpreadsheetWriter(new \SplFileObject($file, 'w'), null, 'Xlsx');
        $writer->prepare();
        $writer->writeItem(['col 1 name' => 'col 1 value', 'col 2 name' => 'col 2 value', 'col 3 name' => 'col 3 value']);
        $writer->finish();

        $excel = IOFactory::load($file);
        $sheet = $excel->getActiveSheet()->toArray();

        // Values should be at first line
        static::assertEquals(['col 1 value', 'col 2 value', 'col 3 value'], $sheet[0]);
    }

    /**
     * Test that column names prepended at first row
     * and values have been written at second line
     * if SpreadsheetWriter's 4-th parameter set to true
     *
     * @author  Igor Mukhin <igor.mukhin@gmail.com>
     */
    public function testHeaderPrependedWhenOptionSetToTrue()
    {
        $file = tempnam(sys_get_temp_dir(), null);

        $writer = new SpreadsheetWriter(new \SplFileObject($file, 'w'), null, 'Xlsx', true);
        $writer->prepare();
        $writer->writeItem(['col 1 name' => 'col 1 value', 'col 2 name' => 'col 2 value', 'col 3 name' => 'col 3 value']);
        $writer->finish();

        $excel = IOFactory::load($file);
        $sheet = $excel->getActiveSheet()->toArray();

        // Check column names at first line
        static::assertEquals(['col 1 name', 'col 2 name', 'col 3 name'], $sheet[0]);

        // Check values at second line
        static::assertEquals(['col 1 value', 'col 2 value', 'col 3 value'], $sheet[1]);
    }

    /**
     *
     */
    public function testWriteItemAppendWithSheetTitle()
    {
        $file = tempnam(sys_get_temp_dir(), null);

        $writer = new SpreadsheetWriter(new \SplFileObject($file, 'w'), 'Sheet 1');

        $writer->prepare();
        $writer->writeItem(['first', 'last']);

        $writer->writeItem(['first' => 'James', 'last'  => 'Bond']);

        $writer->writeItem(['first' => '', 'last'  => 'Dr. No']);

        $writer->finish();

        // Open file with append mode ('a') to add a sheet
        $writer = new SpreadsheetWriter(new \SplFileObject($file, 'a'), 'Sheet 2');

        $writer->prepare();

        $writer->writeItem(['first', 'last']);

        $writer->writeItem(['first' => 'Miss', 'last'  => 'Moneypenny']);

        $writer->finish();

        $excel = IOFactory::load($file);

        static::assertTrue($excel->sheetNameExists('Sheet 1'));
        static::assertEquals(3, $excel->getSheetByName('Sheet 1')->getHighestRow());

        static::assertTrue($excel->sheetNameExists('Sheet 2'));
        static::assertEquals(2, $excel->getSheetByName('Sheet 2')->getHighestRow());
    }

    /**
     *
     */
    public function testWriteItemWithoutSheetTitle()
    {
        $outputFile = new \SplFileObject(tempnam(sys_get_temp_dir(), null));
        $writer     = new SpreadsheetWriter($outputFile);

        $writer->prepare();

        $writer->writeItem(['first', 'last']);

        $writer->finish();
    }
}
