<?php
declare(strict_types=1);

namespace Felds\ExcelCursor;

class Range extends Cursor
{
    protected(set) Cursor $origin;

    public function __construct(string $pos = "A1")
    {
        parent::__construct($pos);
        $this->origin = (clone $this);
    }

    public function getName(): string
    {
        $cols = [$this->col, $this->origin->col];
        $rows = [$this->row, $this->origin->row];

        return sprintf('%s%s:%s%s',
            self::columnNameFromIndex(min($cols)), min($rows),
            self::columnNameFromIndex(max($cols)), max($rows));
    }

    public static function fromCursor(Cursor $origin): static
    {
        $self = new static((string) $origin);
        $self->origin = clone $origin;
        return $self;
    }
}
