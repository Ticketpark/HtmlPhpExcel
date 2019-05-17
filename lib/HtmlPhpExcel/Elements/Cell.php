<?php

namespace Ticketpark\HtmlPhpExcel\Elements;

use Ticketpark\HtmlPhpExcel\Elements\Base\BaseElement;

class Cell extends BaseElement
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