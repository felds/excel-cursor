<?php

namespace Felds\ExcelCursor;

use PHPUnit\Framework\TestCase;

class ExcelCursorTest extends TestCase
{
    /** @test */
    function create_a_cursor() {
        $sut = new ExcelCursor('A1');

        $this->assertInstanceOf(ExcelCursor::class, $sut);
    }

    /**
     * @test
     * @dataProvider badPosition
     * @expectedException InvalidArgumentException
     */
    function validate_initial_position(string $pos) {
        new ExcelCursor($pos);
    }

    /**
     * @test
     * @dataProvider goodPositions
     */
    function get_the_cursor_position(string $pos) {
        $sut = new ExcelCursor($pos);

        $this->assertEquals(strtoupper($pos), (string) $sut);
    }

    /**
     * @test
     * @dataProvider rowMovement
     */
    function move_row($from, $to, $mov = null) {
        $sut = new ExcelCursor($from);

        $sut->moveRow($mov);

        $this->assertSame($to, (string) $sut);
    }

    /**
     * @test
     * @dataProvider colMovement
     */
    function move_col($from, $to, $mov = null) {
        $sut = new ExcelCursor($from);

        $sut->moveCol($mov);

        $this->assertSame($to, (string) $sut);
    }

    function badPosition() {
        return [
            [''], ['A'], ['1'], ['+A2'], ['34A']
        ];
    }

    function goodPositions() {
        return [
            ['A1'], ['b3'], ['AA4'], ['B87'], ['AA989']
        ];
    }

    function rowMovement() {
        return [
            ['A1', 'A2'],
            ['A9', 'A10'],
            ['A10', 'A9', -1],
            ['A9', 'A9', 0],
            ['J14', 'J1', -13],
            ['Z1', 'Z1', -999],
        ];
    }

    function colMovement() {
        return [
            ['A1', 'B1'],
            ['A1', 'C1', 2],
            ['C1', 'A1', -2],
            ['Z23', 'A23', -999],
        ];
    }
}