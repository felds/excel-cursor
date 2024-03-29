<?php
declare(strict_types=1);

namespace Felds\ExcelCursor;

use PHPUnit\Framework\TestCase;

class CursorTest extends TestCase
{
    /** @test */
    function create_a_cursor()
    {
        $sut = new Cursor('A1');

        $this->assertInstanceOf(Cursor::class, $sut);
    }

    /**
     * @test
     * @dataProvider bad_positions
     */
    function validate_initial_position(string $pos)
    {
        $this->expectException(\InvalidArgumentException::class);
        new Cursor($pos);
    }

    /**
     * @test
     * @dataProvider good_positions
     */
    function get_the_cursor_position(string $pos)
    {
        $sut = new Cursor($pos);

        $this->assertEquals(strtoupper($pos), (string)$sut);
    }

    /**
     * @test
     * @dataProvider row_movement
     */
    function move_row($from, $to, $mov = null)
    {
        $sut = new Cursor($from);

        $sut->moveRow($mov);

        $this->assertSame($to, (string)$sut);
    }

    /**
     * @test
     * @dataProvider col_movement
     */
    function move_col($from, $to, $mov = null)
    {
        $sut = new Cursor($from);

        $sut->moveCol($mov);

        $this->assertSame($to, (string)$sut);
    }

    /**
     * @test
     * @dataProvider good_columns
     */
    function parse_column_name($name, $index)
    {
        $this->assertSame($index, Cursor::columnIndexFromName($name));
    }

    /**
     * @test
     * @dataProvider good_columns
     */
    function parse_column_index($name, $index)
    {
        $this->assertSame(strtoupper($name), Cursor::columnNameFromIndex($index));
    }

    /**
     * @test
     * @dataProvider good_rows
     */
    function go_to_row($initial, $row, $expected)
    {
        $sut = new Cursor($initial);

        $sut->goToRow($row);
        $this->assertSame($expected, (string) $sut);
    }

    /**
     * @test
     * @dataProvider bad_rows
     */
    function dont_go_to_row($row)
    {
        $sut = new Cursor();

        $this->expectException(\InvalidArgumentException::class);
        $sut->goToRow($row);
    }

    /**
     * @test
     * @dataProvider good_col_indices
     */
    function go_to_col($initial, $col, $expected)
    {
        $sut = new Cursor($initial);

        $sut->goToCol($col);
        $this->assertSame($expected, (string) $sut);
    }

    /**
     * @test
     * @dataProvider bad_col_indices
     */
    function dont_go_to_col($index)
    {
        $sut = new Cursor();

        $this->expectException(\InvalidArgumentException::class);
        $sut->goToCol($index);
    }

    /**
     * @test
     */
    function create_range()
    {
        $cursor = new Cursor("B4");
        $range = $cursor->createRange();

        $this->assertSame('B4:B4', (string) $range);
    }

    function bad_positions()
    {
        return [
            [''], ['A'], ['1'], ['+A2'], ['34A']
        ];
    }

    function good_positions()
    {
        return [
            ['A1'], ['b3'], ['AA4'], ['B87'], ['AA989']
        ];
    }

    function row_movement()
    {
        return [
            ['A1', 'A2'],
            ['A9', 'A10'],
            ['A10', 'A9', -1],
            ['A9', 'A9', 0],
            ['J14', 'J1', -13],
            ['Z1', 'Z1', -999],
        ];
    }

    function col_movement()
    {
        return [
            ['A1', 'B1'],
            ['A1', 'C1', 2],
            ['C1', 'A1', -2],
            ['Z23', 'A23', -999],
        ];
    }

    function good_columns()
    {
        return [
            ["A", 1],
            ["B", 2],
            ["AA", 27],
            ["BFG", 1515],
            ["xfd", 16384], // excel limit
            ["amj", 1024], // open office limit
            ["IV", 256], // google sheets limit
            ["zz", 702],
            ["fdp", 4176],
        ];
    }

    function good_rows()
    {
        return [
            ['A1', 2, 'A2'],
            ['B5', 5, 'B5'],
            ['ZZ129', 999, 'ZZ999'],
        ];
    }

    function bad_rows()
    {
        return [[0], [-1]];
    }

    function good_col_indices()
    {
        return [
            ['A1', 1, 'A1'],
            ['C12', 1, 'A12'],
            ['A9', 27, 'AA9'],
        ];
    }

    function bad_col_indices()
    {
        return [
            [0], [-1], [-999]
        ];
    }

}
