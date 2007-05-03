<?php

require_once 'Spreadsheet/Excel/Reader/Workbook.php';

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
    
    /*
    function __construct(Excel_Workbook $workbook)
    {
        $this->workbook = $workbook;
    }
    */

    public function addCell($row, $col, $xf, $value, $type = 'Unknown', $raw = null)
    {
        if (is_null($type)) {
            $raw = $value;
        }

        $this->cells[$row][$col]['value'] = $value;
        $this->cells[$row][$col]['type']  = $value;
        $this->cells[$row][$col]['raw']   = $raw;
        $this->cells[$raw][$col]['xf']    = $xf;
    }
}

?>
