<?php

namespace Ticketpark\HtmlPhpExcel\Elements;

use Doctrine\Common\Collections\ArrayCollection;
use Ticketpark\HtmlPhpExcel\Elements\Base\BaseElement;

/**
 * Row
 *
 * @author Manuel Reinhard <manu@sprain.ch>
 */
class Row extends BaseElement
{
    /**
     * A collection of cells within the row
     *
     * @var \Doctrine\Common\Collections\ArrayCollection
     */
    protected $cells;

    /**
     * Constructor
     */
    public function __construct()
    {
        $this->cells = new ArrayCollection();
        parent::__construct();
    }

    /**
     * Add a cell to the row
     *
     * @param \Ticketpark\HtmlPhpExcel\Elements\Cell $cell
     */
    public function addCell(Cell $cell)
    {
        $this->cells->add($cell);
    }

    /**
     * Get cells/columns of the row
     *
     * @return ArrayCollection
     */
    public function getCells()
    {
        return $this->cells;
    }
}