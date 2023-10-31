<?php

declare(strict_types=1);

namespace Ticketpark\HtmlPhpExcel\Elements;

abstract class BaseElement
{
    private array $attributes = [];

    public function addAttribute(string $key, string $value): void
    {
        $this->attributes[$key] = $value;
    }

    public function getAttribute(string $key): ?string
    {
        if (!isset($this->attributes[$key])) {
            return null;
        }

        return $this->attributes[$key];
    }
}
