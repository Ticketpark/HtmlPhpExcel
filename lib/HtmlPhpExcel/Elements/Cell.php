<?php

namespace Ticketpark\HtmlPhpExcel\Elements;

use Ticketpark\HtmlPhpExcel\Elements\Base\BaseElement;

/**
 * Cell
 *
 * @author Manuel Reinhard <manu@sprain.ch>
 */
class Cell extends BaseElement
{
    /**
     * The value of a table cell
     *
     * @var string
     */
    protected $value;

    /**
     * Flag whether the cell is a header cell (<th>)
     *
     * @var bool
     */
    protected $isHeader = false;

    /**
     * Set value
     *
     * @param string $value
     */
    public function setValue($value)
    {
        $this->value = $value;
    }

    /**
     * Get value
     *
     * @return string
     */
    public function getValue()
    {
        return $this->value;
    }

    /**
     * Set if the cell is a header cell (<th>)
     *
     * @param bool $isHeader
     */
    public function setIsHeader($isHeader)
    {
        $this->isHeader = $isHeader;
    }

    /**
     * Returns if the cell is a header cell (<th>)
     *
     * @return bool
     */
    public function isHeader()
    {
        return $this->isHeader;
    }
}