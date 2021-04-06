<?php

namespace Ticketpark\HtmlPhpExcel\Elements;

use Doctrine\Common\Collections\ArrayCollection;

class Table extends BaseElement implements Element
{
    /**
     * A collection of rows within the table
     *
     * @var \Doctrine\Common\Collections\ArrayCollection
     */
    protected $rows;

    public function __construct()
    {
        $this->rows = new ArrayCollection();
        parent::__construct();
    }

    public function addRow(Row $row): void
    {
        $this->rows->add($row);
    }

    public function getRows(): ArrayCollection
    {
        return $this->rows;
    }
}