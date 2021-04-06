<?php

namespace Ticketpark\HtmlPhpExcel\Elements;

class Cell extends BaseElement implements Element
{
    /**
     * The value of a table cell
     *
     * @var string
     */
    private $value;

    /**
     * Flag whether the cell is a header cell (<th>)
     *
     * @var bool
     */
    private $isHeader = false;

    public function setValue(string $value): void
    {
        $this->value = $value;
    }

    public function getValue(): string
    {
        return $this->value;
    }

    public function setIsHeader(bool $isHeader): void
    {
        $this->isHeader = $isHeader;
    }

    public function isHeader(): bool
    {
        return $this->isHeader;
    }
}