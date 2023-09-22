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

use Port\Spreadsheet\SpreadsheetReader;

/**
 * {@inheritDoc}
 */
class SpreadsheetReaderTest extends \PHPUnit_Framework_TestCase
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
     *
     */
    public function testCountWithHeaders()
    {
        $file   = new \SplFileObject(__DIR__.'/fixtures/data_column_headers.xlsx');
        $reader = new SpreadsheetReader($file, 0);
        static::assertEquals(3, $reader->count());
    }

    /**
     *
     */
    public function testCountWithoutHeaders()
    {
        $file   = new \SplFileObject(__DIR__.'/fixtures/data_no_column_headers.xls');
        $reader = new SpreadsheetReader($file);
        static::assertEquals(3, $reader->count());
    }

    /**
     * @author  Derek Chafin <infomaniac50@gmail.com>
     */
    public function testCustomColumnHeadersWithHeaders()
    {
        $file   = new \SplFileObject(__DIR__.'/fixtures/data_column_headers.xlsx');
        $reader = new SpreadsheetReader($file, 0);

        static::assertEquals(['id', 'number', 'description'], $reader->getColumnHeaders());

        $reader->setColumnHeaders(
            ['id2', 'number2', 'description2']
        );

        static::assertEquals(['id2', 'number2', 'description2'], $reader->getColumnHeaders());

        // TODO: Check if row 0 should return the header row if headers are enabled.
        // Row 0 returns the header row as data and indexes.
        $row = $reader->getRow(0);
        static::assertEquals(['id2'          => 'id', 'number2'      => 'number', 'description2' => 'description'], $row);

        $row = $reader->getRow(3);
        static::assertEquals(['id2'          => 7.0, 'number2'      => 7890.0, 'description2' => 'Some more info'], $row);

        $row = $reader->getRow(1);
        static::assertEquals(['id2'          => 50.0, 'number2'      => 123.0, 'description2' => 'Description'], $row);

        $row = $reader->getRow(2);
        static::assertEquals(['id2'          => 6.0, 'number2'      => 456.0, 'description2' => 'Another description'], $row);
    }

    /**
     * @author  Derek Chafin <infomaniac50@gmail.com>
     */
    public function testCustomColumnHeadersWithoutHeaders()
    {
        $file   = new \SplFileObject(__DIR__.'/fixtures/data_no_column_headers.xls');
        $reader = new SpreadsheetReader($file);

        $reader->setColumnHeaders(
            ['id', 'number', 'description']
        );

        $row = $reader->getRow(2);
        static::assertEquals(['id'          => 7.0, 'number'      => 7890.0, 'description' => 'Some more info'], $row);

        $row = $reader->getRow(0);
        static::assertEquals(['id'          => 50.0, 'number'      => 123.0, 'description' => 'Description'], $row);

        $row = $reader->getRow(1);
        static::assertEquals(['id'          => 6.0, 'number'      => 456.0, 'description' => 'Another description'], $row);
    }

    /**
     * @author  Derek Chafin <infomaniac50@gmail.com>
     */
    public function testIterateWithHeaders()
    {
        $file   = new \SplFileObject(__DIR__.'/fixtures/data_column_headers.xlsx');
        $reader = new SpreadsheetReader($file, 0);

        $actualData   = [];
        $expectedData = [['id'          => 50.0, 'number'      => 123.0, 'description' => 'Description'], ['id'          => 6.0, 'number'      => 456.0, 'description' => 'Another description'], ['id'          => 7.0, 'number'      => 7890.0, 'description' => 'Some more info']];

        foreach ($reader as $row) {
            $actualData[] = $row;
        }

        static::assertEquals($expectedData, $actualData);
    }

    /**
     * @author  Derek Chafin <infomaniac50@gmail.com>
     */
    public function testIterateWithoutHeaders()
    {
        $file   = new \SplFileObject(__DIR__.'/fixtures/data_no_column_headers.xls');
        $reader = new SpreadsheetReader($file);

        $actualData   = [];
        $expectedData = [[50.0, 123.0, "Description"], [6.0, 456.0, 'Another description'], [7.0, 7890.0, 'Some more info']];

        foreach ($reader as $row) {
            $actualData[] = $row;
        }

        static::assertEquals($expectedData, $actualData);
    }

    /**
     *
     */
    public function testMaxRowNumb()
    {
        $file   = new \SplFileObject(__DIR__.'/fixtures/data_no_column_headers.xls');
        $reader = new SpreadsheetReader($file, null, null, null, 1000);
        static::assertEquals(3, $reader->count());

        // Without $maxRows, this faulty file causes OOM because of an extremely
        //high last row number
        $file = new \SplFileObject(__DIR__.'/fixtures/data_extreme_last_row.xlsx');

        $max    = 5;
        $reader = new SpreadsheetReader($file, null, null, null, $max);
        static::assertEquals($max, $reader->count());
    }

    /**
     *
     */
    public function testMultiSheet()
    {
        $file         = new \SplFileObject(__DIR__.'/fixtures/data_multi_sheet.xls');
        $sheet1reader = new SpreadsheetReader($file, null, 0);
        static::assertEquals(3, $sheet1reader->count());

        $sheet2reader = new SpreadsheetReader($file, null, 1);
        static::assertEquals(2, $sheet2reader->count());
    }

    /**
     * @author  Derek Chafin <infomaniac50@gmail.com>
     */
    public function testSeekWithHeaders()
    {
        $file   = new \SplFileObject(__DIR__.'/fixtures/data_column_headers.xlsx');
        $reader = new SpreadsheetReader($file, 0);

        // TODO: Check if row 0 should return the header row if headers are enabled.
        // Row 0 returns the header row as data and indexes.
        $row = $reader->getRow(0);
        static::assertEquals(['id'          => 'id', 'number'      => 'number', 'description' => 'description'], $row);
        static::assertEquals(0, $reader->key());

        $row = $reader->getRow(3);
        static::assertEquals(['id'          => 7.0, 'number'      => 7890.0, 'description' => 'Some more info'], $row);
        static::assertEquals(3, $reader->key());

        $row = $reader->getRow(1);
        static::assertEquals(['id'          => 50.0, 'number'      => 123.0, 'description' => 'Description'], $row);
        static::assertEquals(1, $reader->key());

        $row = $reader->getRow(2);
        static::assertEquals(['id'          => 6.0, 'number'      => 456.0, 'description' => 'Another description'], $row);
        static::assertEquals(2, $reader->key());
    }

    /**
     * @author  Derek Chafin <infomaniac50@gmail.com>
     */
    public function testSeekWithoutHeaders()
    {
        $file   = new \SplFileObject(__DIR__.'/fixtures/data_no_column_headers.xls');
        $reader = new SpreadsheetReader($file);

        $row = $reader->getRow(2);
        static::assertEquals([7.0, 7890.0, 'Some more info'], $row);

        $row = $reader->getRow(0);
        static::assertEquals([50.0, 123.0, 'Description'], $row);

        $row = $reader->getRow(1);
        static::assertEquals([6.0, 456.0, 'Another description'], $row);
    }
}
