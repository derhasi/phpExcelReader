<?php

class Excel_Workbook
{
    const USER_DEFINED_FORMATS = 164;

    /**
     * List of worksheets that form part of this workbook
     * 
     * @access public
     */
    public $worksheets = array();

    /**
     * SST - Shared String Table
     */
    public $sst = array();

    /**
     * Datemode - Epoch used to calculate dates
     * 
     * 0 - Use 30 Dec 1899
     * 1 - Use  1 Jan 1904
     */
    public $datemode = 0;

    public $font_records = array();

    /**
     * Holds the format records as defined in Section 6.45
     * 
     * The defaults here are built in and are not written (except for the money ones)
     * 
     * TODO: store default records.
     * 
     */
    public $format_records = array(
        0  => '#',
        1  => '0',
        2  => '0.00',
        3  => '#,##0',
        4  => '#,##0.00',
        5  => '"$"#,##0_);("$"#,##0)',
        6  => '"$"#,##0_);[Red]("$"#,##0)',
        7  => '"$"#,##0.00_);("$"#,##0.00)',
        8  => '"$"#,##0.00_);[Red]("$"#,##0.00)',
        9  => '0%',
        10 => '0.00%',
        11 => '0.00E+00',
        12 => '# ?/?',
        13 => '# ??/??',
        14 => 'D/M/YY',
        15 => 'D-MMM-YY',
        16 => 'D-MMM',
        17 => 'MMM-YY',
        18 => 'h:mm AM/PM',
        19 => 'h:mm:ss AM/PM',
        20 => 'h:mm',
        21 => 'h:mm:ss',
        22 => 'D/M/YY h:mm',
        37 => '_(#,##0_);(#,##0)',
        38 => '_(#,##0_);[Red](#,##0)',
        39 => '_(#,##0.00_);(#,##0.00)',
        40 => '_(#,##0.00_);[Red](#,##0.00)',
        41 => '_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)',
        42 => '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
        43 => '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)',
        44 => '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
        45 => 'mm:ss',
        46 => '[h]:mm:ss',
        47 => 'mm:ss.0',
        48 => '##0.0E+0',
        49 => '@',
    );


    public $xf_records = array();


    public $style_records = array();

    public $active_worksheet;
    public $user;
    public $name;

/*
    public $number_formats = array(
        0  => array('type' => 'General', 'format' => )
    )
*/

    public function is1904()
    {
        return $this->datemode == 1;
    }

    public function set1904($is_1904)
    {
        $this->datemode = (int) $is_1904;
    }

    public function getWorksheet($worksheet_name)
    {
        foreach ($this->worksheets as $worksheet) {
            if ($worksheet->name === $worksheet_name) {
                return $worksheet;
            }
        }

        return new Spreadsheet_Excel_Reader_Exception('Unknown worksheet');
    }

    public function toArray()
    {
    }
}

?>
