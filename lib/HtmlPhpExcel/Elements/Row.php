<?php

namespace Ticketpark\HtmlPhpExcel\Elements;

class Row extends BaseElement implements Element
{
    private array $cells = [];

    public function addCell(Cell $cell): void
    {
        $this->cells[] = $cell;
    }

    /**
     * @return array<Cell>
     */
    public function getCells(): array
    {
        return $this->cells;
    }
}
