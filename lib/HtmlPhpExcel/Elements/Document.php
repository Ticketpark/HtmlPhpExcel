<?php

namespace Ticketpark\HtmlPhpExcel\Elements;

use Doctrine\Common\Collections\ArrayCollection;

class Document
{
    /**
     * A collection of tables within the document
     *
     * @var \Doctrine\Common\Collections\ArrayCollection
     */
    protected $tables;

    /**
     * Constructor
     */
    public function __construct()
    {
        $this->tables = new ArrayCollection();
    }

    public function addTable(Table $table): void
    {
        $this->tables->add($table);
    }

    public function getTables(): ArrayCollection
    {
        return $this->tables;
    }
}