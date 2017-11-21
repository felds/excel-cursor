<?php
declare(strict_types=1);

namespace Felds\ExcelCursor;

use PHPExcel_Cell;

class ExcelCursor
{
    /** @var string */
    private $col;

    /** @var int */
    private $row;

    public function __construct(string $pos)
    {
        $isValid = preg_match('{^([a-z]+)([0-9]+)$}i', $pos, $results);

        if (! $isValid)
            throw new \InvalidArgumentException("Bad initial position “{$pos}”.");

        $this->col = PHPExcel_Cell::columnIndexFromString($results[1]) - 1;
        $this->row = $results[2];
    }

    public function __toString()
    {
        return PHPExcel_Cell::stringFromColumnIndex($this->col) . $this->row;
    }

    public function moveRow(int $mov = null): self
    {
        $this->row = max($this->row + ($mov ?? 1), 1);

        return $this;
    }

    public function moveCol(int $mov = null): self
    {
        $this->col = max($this->col + ($mov ?? 1), 0);

        return $this;
    }
}