<?php
declare(strict_types=1);

namespace Felds\ExcelCursor;

class Range extends Cursor
{
    /**
     * @var Cursor
     */
    private $origin;

    public function __construct($pos = "A1")
    {
        parent::__construct($pos);

        $this->origin = (clone $this);
    }

    public function __toString(): string
    {
        return sprintf('%s:%s', $this->origin->getName(), $this->getName());
    }

    public static function fromCursor(Cursor $origin): self
    {
        $self = new static((string) $origin);
        $self->origin = clone $origin;

        return $self;
    }

    public function getOrigin(): Cursor
    {
        return $this->origin;
    }
}