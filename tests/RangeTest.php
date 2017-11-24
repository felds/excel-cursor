<?php
declare(strict_types=1);

namespace Felds\ExcelCursor;

use PHPUnit\Framework\TestCase;

class RangeTest extends TestCase
{
    /**
     * @test
     */
    function create_a_range()
    {
        $r = new Range();

        $this->assertInstanceOf(Range::class, $r);
    }

    /**
     * @test
     */
    function default_value()
    {
        $r = new Range();

        $this->assertSame('A1:A1', (string) $r);
    }

    /**
     * @test
     */
    function create_from_cursor()
    {
        $cursor = new Cursor('E18');
        $range = Range::fromCursor($cursor);

        $range->moveCol(-2)->goToRow(20);

        $this->assertSame('E18:C20', (string) $range);
    }

    /**
     * @test
     * @dataProvider col_movements
     */
    function move_columns($initial, $movement, $expected)
    {
        $r = new Range($initial);
        $r->moveCol($movement);

        $this->assertSame($expected, (string) $r);
    }

    function col_movements()
    {
        return [
            ['A1', +1, 'A1:B1'],
            ['A1', -1, 'A1:A1'],
            ['A1', 0, 'A1:A1'],
            ['A1', +26, 'A1:AA1'],
            ['B19', +2, 'B19:D19'],
            ['D1', -1, 'D1:C1'],
            ['D1', -10, 'D1:A1'],
        ];
    }

    /**
     * @test
     * @dataProvider row_movements
     */
    function move_rows($initial, $movement, $expected)
    {
        $r = new Range($initial);
        $r->moveRow($movement);

        $this->assertSame($expected, (string) $r);
    }


    function row_movements()
    {
        return [
            ['A1', +1, 'A1:A2'],
            ['A1', -1, 'A1:A1'],
            ['A1', 0, 'A1:A1'],
            ['A1', +26, 'A1:A27'],
            ['B17', +30, 'B17:B47'],
            ['B17', -20, 'B17:B1'],
        ];
    }
}
