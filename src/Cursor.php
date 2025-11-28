<?php
declare(strict_types=1);

namespace Felds\ExcelCursor;

class Cursor
{
    protected int $col {
        get {
            return $this->col;
        }
    }

    protected int $row {
        get {
            return $this->row;
        }
    }

    public function __construct(string $pos = "A1")
    {
        $isValid = preg_match('{^([a-z]+)([0-9]+)$}i', $pos, $results);

        if (! $isValid)
            throw new \InvalidArgumentException("Bad initial position “{$pos}”.");

        $this->col = self::columnIndexFromName($results[1]);
        $this->row = (int) $results[2];
    }

    public function __toString(): string
    {
        return $this->getName();
    }

    /**
     * Get the cell name as `A1`.
     */
    public function getName(): string
    {
        return self::columnNameFromIndex($this->col) . $this->row;
    }

    /**
     * Move `$n` rows down.
     */
    public function moveRow(int $n = 1): self
    {
        $this->row = max($this->row + $n, 1);
        return $this;
    }

    /**
     * Move `$n` columns to the right.
     */
    public function moveCol(int $n = 1): self
    {
        $this->col = max($this->col + $n, 1);
        return $this;
    }

    /**
     * Go to the row `$row`.
     */
    public function goToRow(int $row): self
    {
        if ($row < 1) {
            throw new \InvalidArgumentException("The row can't be smaller than 1.");
        }
        $this->row = $row;
        return $this;
    }

    /**
     * Go to column `$col`.
     */
    public function goToCol(int $col): self
    {
        if ($col < 1) {
            throw new \InvalidArgumentException("The col can't be smaller than 1.");
        }
        $this->col = $col;
        return $this;
    }

    public function createRange(): Range
    {
        return Range::fromCursor($this);
    }

    public static function columnNameFromIndex(int $index): string
    {
        $name = "";
        while ($index > 0) {
            $charValue = ($index - 1) % 26;
            $char = chr(65 + $charValue);
            $name = $char . $name;
            $index = intdiv($index - $charValue - 1, 26);
        }
        return $name;
    }

    public static function columnIndexFromName(string $name): int
    {
        // normalize the name to uppercase, split it into chars and reverse the order
        $chars = array_reverse(str_split(strtoupper($name)));
        $sum = 0;
        foreach ($chars as $i => $char) {
            $charValue = ord($char) - 65 + 1; // 65 = ord("A")
            $sum += $charValue * (26 ** $i); // add current char to the accumulator
        }
        return $sum;
    }
}
