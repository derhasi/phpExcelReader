<?php

require_once 'Spreadsheet/Excel/Reader/Workbook.php';
require_once 'Spreadsheet/Excel/Reader/Cell.php';

class Excel_Worksheet
{
    /**
     * Name of the worksheet as appears on the tab
     */
    public $name;

    private $_workbook;

    public $cells = array();
    public $selected_cells = array();

    public $first_row = 0;
    public $last_row = 0;
    
    public $first_col = 0;
    public $last_col = 0;

    public $numRows = 0;
    public $numCols = 0;
    
    function __construct(Excel_Workbook $workbook)
    {
        $this->workbook = $workbook;
    }

    public function addCell($row, $col, $xf, $value, $type = 'Unknown', $raw = null)
    {
        if (is_null($type)) {
            $raw = $value;
        }

        $cell = new Excel_Cell($this, $row, $col, $xf, $value);
        $this->cells[$row][$col] = $cell;
    }

    public function toArray()
    {
    }

    public function getCell($x, $y)
    {
        if (is_string($x) && !is_numeric($x)) {
            $col = 0;
            foreach (str_split(strtoupper($x)) as $letter) {
                $col += ord($letter) - 65;
            }
            $row = $y - 1;
        } else {
            $row = $x;
            $col = $y;
        }
        return $this->cells[$row][$col];
    }

    public function getSelectedCells()
    {
        
    }
}

?>
