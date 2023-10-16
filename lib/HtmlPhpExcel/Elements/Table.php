<?php

namespace Ticketpark\HtmlPhpExcel\Elements;

class Table extends BaseElement implements Element
{
    private array $rows = [];

    public function addRow(Row $row): void
    {
        $this->rows[] = $row;
    }

    /**
     * @return array<Row>
     */
    public function getRows(): array
    {
        return $this->rows;
    }
}
