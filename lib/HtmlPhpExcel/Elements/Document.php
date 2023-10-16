<?php

namespace Ticketpark\HtmlPhpExcel\Elements;

class Document
{
    private array $tables = [];

    public function addTable(Table $table): void
    {
        $this->tables[] = $table;
    }

    /**
     * @return array<Table>
     */
    public function getTables(): array
    {
        return $this->tables;
    }
}
