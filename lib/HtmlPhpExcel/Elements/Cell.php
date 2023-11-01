<?php

declare(strict_types=1);

namespace Ticketpark\HtmlPhpExcel\Elements;

class Cell extends BaseElement implements Element
{
    private ?string $value = null;
    private bool $isHeader = false;

    public function setValue(string $value): void
    {
        $this->value = $value;
    }

    public function getValue(): ?string
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
