<?php

namespace Ticketpark\HtmlPhpExcel\Elements\Base;

use Doctrine\Common\Collections\ArrayCollection;

/**
 * Base Element
 *
 * @author Manuel Reinhard <manu@sprain.ch>
 */
class BaseElement
{
    /**
     * A collection of attributes of this cell
     *
     * @var \Doctrine\Common\Collections\ArrayCollection
     */
    protected $attributes;

    public function __construct()
    {
        $this->attributes = new ArrayCollection();
    }

    public function addAttribute(string $key, string $value): void
    {
        $this->attributes->set($key, $value);
    }

    public function getAttribute($key): ?string
    {
        return $this->attributes->get($key);
    }
}