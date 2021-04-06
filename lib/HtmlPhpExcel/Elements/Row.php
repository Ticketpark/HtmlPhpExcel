<?php

namespace Ticketpark\HtmlPhpExcel\Elements;

use Doctrine\Common\Collections\ArrayCollection;

class Row extends BaseElement implements Element
{
    /**
     * A collection of cells within the row
     *
     * @var \Doctrine\Common\Collections\ArrayCollection
     */
    protected $cells;

    public function __construct()
    {
        $this->cells = new ArrayCollection();
        parent::__construct();
    }

    public function addCell(Cell $cell): void
    {
        $this->cells->add($cell);
    }

    public function getCells(): ArrayCollection
    {
        return $this->cells;
    }
}