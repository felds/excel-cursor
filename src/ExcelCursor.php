<?php
declare(strict_types=1);

namespace Felds\ExcelCursor;

class ExcelCursor
{
    /** @var string */
    private $col;

    /** @var int */
    private $row;

    public function __construct(string $pos = "A1")
    {
        $isValid = preg_match('{^([a-z]+)([0-9]+)$}i', $pos, $results);

        if (! $isValid)
            throw new \InvalidArgumentException("Bad initial position “{$pos}”.");

        $this->col = self::columnIndexFromName($results[1]);
        $this->row = (int) $results[2];
    }

    public function __toString()
    {
        return self::columnNameFromIndex($this->col) . $this->row;
    }

    public function moveRow(int $mov = null): self
    {
        $this->row = max($this->row + ($mov ?? 1), 1);

        return $this;
    }

    public function moveCol(int $mov = null): self
    {
        $this->col = max($this->col + ($mov ?? 1), 1);

        return $this;
    }

    public function goToRow(int $row): self
    {
        if ($row < 1) throw new \InvalidArgumentException("The row can't be smaller than 1.");

        $this->row = $row;

        return $this;
    }

    public function goToCol(int $col): self
    {
        if ($col < 1) throw new \InvalidArgumentException("The col can't be smaller than 1.");

        $this->col = $col;

        return $this;
    }

    public static function columnNameFromIndex(int $index): string
    {
        $name = "";

        while ($index > 0) {
            $charValue = ($index - 1) % 26; // current char codepoint
            $char = chr(65 + $charValue); // the actual char
            $name = $char . $name; // prepend the char
            $index = ($index - $charValue - 1) / 26; // take of the char and step down the char
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
