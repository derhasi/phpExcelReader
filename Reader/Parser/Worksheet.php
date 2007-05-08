<?php

require_once 'Spreadsheet/Excel/Reader/Parser.php';
require_once 'Spreadsheet/Excel/Reader/Worksheet.php';

class Spreadsheet_Excel_Reader_BIFFParser_Worksheet extends Spreadsheet_Excel_Reader_BIFFParser
{
    private $first_row_index;
    private $last_row_index;

    private $workbook;

    function __construct($stream, $workbook, $version)
    {
        parent::__construct($stream);
        $this->workbook = $workbook;
        $this->version = $version;
    }

    private function isDate($xf_index)
    {
        if ($this->workbook->xf_records[$xf_index]['type'] == 'date') {
            $this->curformat = $this->workbook->xf_records[$xf_index]['format'];
            $this->rectype = 'date';
            return true;
        } else {
            if ($this->workbook->xf_records[$xf_index]['type'] == 'number') {
                $this->curformat = $this->workbook->xf_records[$xf_index]['format'];
                $this->rectype = 'number';
                if (($xf_index == 0x9) || ($xf_index == 0xa)){
                    $this->multiplier = 100;
                }
            }else{
                $this->curformat = Spreadsheet_Excel_Reader::DEF_NUM_FORMAT;
                $this->rectype = 'unknown';
            }
            return false;
        }
    }

    /**
     * Convert the raw Excel date into a human readable format
     *
     * Dates in Excel are stored as number of seconds from an epoch.  On 
     * Windows, the epoch is 30/12/1899 and on Mac it's 01/01/1904
     *
     * @access private
     * @param integer The raw Excel value to convert
     * @return array First element is the converted date, the second element is number a unix timestamp
     */ 
    function createDate($numValue)
    {
        if ($numValue > 1) {
            $utcDays = $numValue - ($this->workbook->datemode === 1 ? Spreadsheet_Excel_Reader::UTCOFFSETDAYS1904 : Spreadsheet_Excel_Reader::UTCOFFSETDAYS);
            $utcValue = round(($utcDays+1) * Spreadsheet_Excel_Reader::MSINADAY);
            $string = date ($this->curformat, $utcValue);
            $raw = $utcValue;
        } else {
            $raw = $numValue;
            $hours = floor($numValue * 24);
            $mins = floor($numValue * 24 * 60) - $hours * 60;
            $secs = floor($numValue * Spreadsheet_Excel_Reader::MSINADAY) - $hours * 60 * 60 - $mins * 60;
            $string = date ($this->curformat, mktime($hours, $mins, $secs));
        }

        return array($string, $raw);
    }


    /**
     * Parse a worksheet
     *
     * @access public
     */
    public function parse()
    {
        $worksheet = new Excel_Worksheet;

        $pos = ftell($this->_stream);

        $code   = $this->_readInt(2);
        $length = $this->_readInt(2);

// TODO change me
assert($code == Spreadsheet_Excel_Reader::TYPE_BOF);

        $row_block_count = 0;

        while($code != Spreadsheet_Excel_Reader::TYPE_EOF) {

            $this->multiplier = 1; // need for format with %

            switch ($code) {


                // Section 6.8
                case Spreadsheet_Excel_Reader::TYPE_BOF:

                    // The version in worksheet streams cannot be trusted
                    fseek($this->_stream, 2, SEEK_CUR);
                    $substream_type = $this->_readInt(2);
                    $build_id       = $this->_readInt(2);
                    $build_year     = $this->_readInt(2);

// TODO change me
assert($substream_type == Spreadsheet_Excel_Reader::WORKSHEET);

                    if ($this->version == 8) {
                        $file_history_flags = $this->_readInt(4);
                        $lowest_version     = $this->_readInt(4);
                    }

                    break;


                // Formulas not recalculated
                // Section 6.104
                // The contents are not used, the presence of this record flags the property
                // TODO store?
                case Spreadsheet_Excel_Reader::TYPE_UNCALCED:

                    fseek($this->_stream, 2, SEEK_CUR);
                    break;


                // Index - contains range of used rows and stream positions to several records
                // Section 6.55
                case Spreadsheet_Excel_Reader::TYPE_INDEX:
//echo "type index\n";

                    fseek($this->_stream, 4, SEEK_CUR);
                    $int_size = $this->version == 7 ? 2 : 4;
                    $this->first_row_index = $this->_readInt($int_size);
                    $this->last_row_index  = $this->_readInt($int_size) - 1;

                    // The remaining part of the record is not relevant

/*
                    fseek($this->_stream, 4, SEEK_CUR);
                    // FIXME
                    // floor or ceil?
                    //$nm = floor(($rl - $rf - 1) / (32 + 1));
                    $nm = ceil(($rl - $rf - 1) / (32 + 1));
                    fseek($this->_stream, $nm * 4, SEEK_CUR);
*/

                    break;


                // Default height for rows that do not have a corresponding
                // ROW record
                // Section 6.28
                case Spreadsheet_Excel_Reader::TYPE_DEFAULTROWHEIGHT:
//echo "DEFAULTROWHEIGHT\n";
                    $options        = $this->_readInt(2);
                    $default_unused = $this->_readInt(2);
                    break;


                //
                // --- BEGIN worksheet protection ---
                //

                case Spreadsheet_Excel_Reader::TYPE_PROTECT:
echo "protect\n";
                    $protected = $this->_readInt(2);
                    break;

                case Spreadsheet_Excel_Reader::TYPE_WINDOWPROTECT:
echo "windowprotect\n";
                    $window_settings_protected = $this->_readInt(2);
                    break;

                case Spreadsheet_Excel_Reader::TYPE_OBJECTPROTECT:
echo "objectprotect\n";
                    $objects_protected = $this->_readInt(2);
                    break;

                case Spreadsheet_Excel_Reader::TYPE_SCENPROTECT:
echo "scenprotect\n";
                    $scenarios_protected = $this->_readInt(2);
                    break;

                case Spreadsheet_Excel_Reader::TYPE_PASSWORD:
echo "password\n";
                    $password = fread($this->_stream, 2);
                    break;

                //
                // --- END worksheet protection ---
                //

                case Spreadsheet_Excel_Reader::TYPE_DEFCOLWIDTH:
echo "defcolwidth\n";
                    $col_width = $this->_readInt(2);
                    break;

                case Spreadsheet_Excel_Reader::TYPE_COLINFO:
echo "colinfo\n";
                    $first_col  = $this->_readInt(2);
                    $last_col   = $this->_readInt(2);
                    $col_width  = $this->_readInt(2);
                    $xf_index   = $this->_readInt(2);
                    $options    = $this->_readInt(2);
                    fseek($this->_stream, 2, SEEK_CUR);
                    break;


                // Section 6.31
                case Spreadsheet_Excel_Reader::TYPE_DIMENSIONS:
echo "dimensions\n";
                    if ($this->version == 7){
                        $first_row = $this->_readInt(2);
                        $last_row  = $this->_readInt(2);
                    } else {
                        $first_row = $this->_readInt(4);
                        $last_row  = $this->_readInt(4);
                    }

                    $first_col = $this->_readInt(2);
                    $last_col  = $this->_readInt(2);

                    $this->sheets['numRows'] = $last_row - $first_row;
                    $this->sheets['numCols'] = $last_col - $first_col;

                    fseek($this->_stream, 2, SEEK_CUR);

                    break;


                case Spreadsheet_Excel_Reader::TYPE_ROW:
echo "row\n";
                    $row_index  = $this->_readInt(2);
                    $first_col  = $this->_readInt(2);
                    $last_col   = $this->_readInt(2);
                    $row_height = $this->_readInt(2);
                    fseek($this->_stream, 4, SEEK_CUR);
                    $options    = $this->_readInt(4);

                    $row_block_count++;

                    break;



                //
                // --- BEGIN cell block ---
                //

                case Spreadsheet_Excel_Reader::TYPE_BLANK:
echo "BLANK\n";
                    $row_index = $this->_readInt(2);
                    $col_index = $this->_readInt(2);
                    $xf_index  = $this->_readInt(2);

                    $worksheet->addCell($row_index, $col_index, $xf_index, null);
                    break;


                case Spreadsheet_Excel_Reader::TYPE_BOOLERR:
echo "type boolerr\n";

                    // TODO
                    // Retrieve error codes as described in Section 3.7

                    $row_index = $this->_readInt(2);
                    $col_index = $this->_readInt(2);
                    $xf_index  = $this->_readInt(2);
                    $type      = $this->_readInt(1);
                    $data      = $this->_readInt(1);
                    
                    $worksheet->addCell($row_index, $col_index, $xf_index, $data);
                    break;


                case Spreadsheet_Excel_Reader::TYPE_LABEL:
echo "Type label\n";
                    $row_index = $this->_readInt(2);
                    $col_index = $this->_readInt(2);
                    $xf_index  = $this->_readInt(2);

                    if ($this->version == 7) {
                        $label = $this->_readString(2);
                    } else {
                        $label = $this->_readUnicodeString(2);
                    }

                    $worksheet->addCell($row_index, $col_index, $xf_index, $label);

                    break;


                // Section 6.61
                case Spreadsheet_Excel_Reader::TYPE_LABELSST:
echo "labelsst"."\n";

                    $row_index = $this->_readInt(2);
                    $col_index = $this->_readInt(2);
                    $xf_index  = $this->_readInt(2);
                    $sst_index = $this->_readInt(4);

                    $worksheet->addCell($row_index, $col_index, $xf_index, $this->workbook->sst[$sst_index]);

                    break;


                // Section 6.64
                // Multiple blank - all cells in same row
                case Spreadsheet_Excel_Reader::TYPE_MULBLANK:
echo "MULBLANK\n";
                    $row_index       = $this->_readInt(2);
                    $first_col_index = $this->_readInt(2);
                    
                    // the last col index appears after the data!
                    $temp_pos = ftell($this->_stream);
                    fseek($this->_stream, $length - 6, SEEK_CUR);
                    $last_col_index  = $this->_readInt(2);
                    fseek($this->_stream, $temp_pos);

                    $num_cols = $last_col_index - $first_col_index + 1;
                    for ($i = 0; $i < $num_cols; $i++) {
                        $xf_index = $this->_readInt(2);
                        $worksheet->addCell($row_index, $first_col_index + $i, $xf_index, null);
                    }

                    fseek($this->_stream, 2, SEEK_CUR);
                    break;


                // Section 6.64
                // Multiple RK - all cells in same row
                case Spreadsheet_Excel_Reader::TYPE_MULRK:
echo "type mulrk"."\n";
                    
                    $row_index       = $this->_readInt(2);
                    $first_col_index = $this->_readInt(2);
                    
                    // the last col index appears after the data!
                    $temp_pos = ftell($this->_stream);
                    fseek($this->_stream, $length - 6, SEEK_CUR);
                    $last_col_index  = $this->_readInt(2);
                    fseek($this->_stream, $temp_pos);

                    $num_cols = $last_col_index - $first_col_index + 1;
                    for ($i = 0; $i < $num_cols; $i++) {

                        $xf_index = $this->_readInt(2);
                        $value    = $this->_readInt(4);

/*
                        if ($this->isDate($xf_index)) {
                            list($string, $raw) = $this->createDate($value);
                        } else {
                            $raw = $value;

                            if (isset($this->_columnsFormat[$first_col_index + $i + 1])) {
                                $this->curformat = $this->_columnsFormat[$first_col_index + $i + 1];
                            }

                            $string = sprintf($this->curformat, $value * $this->multiplier);
                        }
*/

                        $worksheet->addCell($row_index, $first_col_index + $i, $xf_index, $value);
                    }

                    break;


                // Stores floating point values that cannot be stored as an RK value?
                // Section 6.68

                case Spreadsheet_Excel_Reader::TYPE_NUMBER:
echo "type number\n";

                    $row_index = $this->_readInt(2);
                    $col_index = $this->_readInt(2);
                    $xf_index  = $this->_readInt(2);
                    $number    = $this->_readDouble();

/*
                    if ($this->isDate($xf_index)) {
                        list($string, $raw) = $this->createDate($number);
                    } else {
                        if (isset($this->_columnsFormat[$col_index + 1])) {
                            $this->curformat = $this->_columnsFormat[$col_index + 1];
                        }
                        $raw = $number;
                        $string = sprintf($this->curformat, $number);
                    }
*/

                    $worksheet->addCell($row_index, $col_index, $xf_index, $number);

                    break;


                case Spreadsheet_Excel_Reader::TYPE_RK:

echo 'Spreadsheet_Excel_Reader::TYPE_RK'."\n";

                    $row_index = $this->_readInt(2);        
                    $col_index = $this->_readInt(2);        
                    $xf_index  = $this->_readInt(2);        

                    $rk_value  = $this->_readInt(4);
                    $number    = $this->_convertRKValue($rk_value);

/*
                    if ($this->isDate($xf_index)) {
                        list($string, $raw) = $this->createDate($number);
                    } else {
                        $raw = $number;
                        if (isset($this->_columnsFormat[$col_index + 1])){
                                $this->curformat = $this->_columnsFormat[$col_index + 1];
                        }

                        $string = sprintf($this->curformat, $number * $this->multiplier);
                    }
*/

                    $worksheet->addCell($row_index, $col_index, $xf_index, $number);

                    break;


                // BIFF7 Rich text strings
                // BIFF8 Uses this record only for the clipboard
                // Section 6.84

                case Spreadsheet_Excel_Reader::TYPE_RSTRING:

                    // TODO
                    // store

                    $row_index = $this->_readInt(2);
                    $col_index = $this->_readInt(2);
                    $xf_index  = $this->_readInt(2);
                    $string    = $this->_readString(2);
                    $num_runs  = $this->_readInt(1);

                    for ($i = 0; $i < $num_runs; $i++) {
                        $run_list  = $this->_readInt(2);
                    }

                    break;


                case Spreadsheet_Excel_Reader::TYPE_FORMULA:
echo "type formula\n";

                    $row_index = $this->_readInt(2);
                    $col_index = $this->_readInt(2);
                    $xf_index  = $this->_readInt(2);
                    // todo result may be stored in 1 of 5 ways
                    $result    = $this->_readResult();
                    $options   = $this->_readInt(2);
                    fseek($this->_stream, 4, SEEK_CUR);
                    // todo read the formula data
                    $formula   = $this->_readFormula();


                    // todo good enough check?
                    if (is_float($result)) {

/*
                        if ($this->isDate($xf_index)) {
                            list($string, $raw) = $this->createDate($result);
                        } else {
                            if (isset($this->_columnsFormat[$col_index + 1])) {
                                $this->curformat = $this->_columnsFormat[$col_index + 1];
                            }

                            $raw = $result;
                            $string = sprintf($this->curformat, $raw * $this->multiplier);
                        }
*/

                        $worksheet->addCell($row_index, $col_index, $xf_index, $result);
                    }

                    break;


                case Spreadsheet_Excel_Reader::TYPE_ARRAY:

                    // TODO
                    fseek($this->_stream, 12, SEEK_CUR);
                    $this->_readFormula();
                    break;


                // Shared Formula
                // Section 6.94
                case Spreadsheet_Excel_Reader::TYPE_SHRFMLA:
                    
                    // TODO
                    fseek($this->_stream, 8, SEEK_CUR);
                    $this->_readFormula();
                    break;


                case Spreadsheet_Excel_Reader::TYPE_TABLEOP:

                    // TODO
                    fseek($this->_stream, 16, SEEK_CUR);
                    break;


                case Spreadsheet_Excel_Reader::TYPE_DBCELL:
echo 'type dbcell'."\n";

                    //todo?
                    fseek($this->_stream, $row_block_count * 2 + 4, SEEK_CUR);
                    $row_block_count = 0;
                    break;


                //
                // Worksheet View Settings Block
                //

                case Spreadsheet_Excel_Reader::TYPE_WINDOW2:
echo "type window2\n";
                    $options   = $this->_readInt(2);
                    $row_index = $this->_readInt(2);
                    $col_index = $this->_readInt(2);

                    if ($this->version == 7) {

                        $gridline_colour = $this->_readInt(4);

                    } else {

                        $gridline_colour_index    = $this->_readInt(2);

                        fseek($this->_stream, 2, SEEK_CUR);

                        $page_break_magnification = $this->_readInt(2);
                        $normal_magnification     = $this->_readInt(2);

                        fseek($this->_stream, 4, SEEK_CUR);
                    }

                    break;


                case Spreadsheet_Excel_Reader::TYPE_PANE:

                    // todo this might be useful
                    fseek($this->_stream, 9, SEEK_CUR);
                    break;


                case Spreadsheet_Excel_Reader::TYPE_SELECTION:

                    // todo store on a pane basis
                    $pane_id                   = $this->_readInt(1);
                    $row_index                 = $this->_readInt(2);
                    $col_index                 = $this->_readInt(2);
                    $cell_range_index          = $this->_readInt(2);
                    $worksheet->selected_cells = $this->_readCellRangeAddressList(1);
                    break;


                case Spreadsheet_Excel_Reader::TYPE_PHONETIC:
echo "type phonetic\n";
                    // todo store
                    $font_index     = $this->_readInt(2);
                    $settings       = $this->_readInt(2);
                    $phonetic_cells = $this->_readCellRangeAddressList(2);
                    break;


                case Spreadsheet_Excel_Reader::TYPE_MERGEDCELLS:

                    $worksheet->merged_cells = $this->_readCellRangeAddressList(1);        
                    break;



                //
                // Other records not yet stored
                //

                // calculation settings block
                // occurs in every stream and is global for the entire workbook
                case Spreadsheet_Excel_Reader::TYPE_CALCCOUNT:
                case Spreadsheet_Excel_Reader::TYPE_CALCMODE:
                case Spreadsheet_Excel_Reader::TYPE_PRECISION:
                case Spreadsheet_Excel_Reader::TYPE_REFMODE:
                case Spreadsheet_Excel_Reader::TYPE_ITERATION:
                case Spreadsheet_Excel_Reader::TYPE_SAVERECALC:
                case Spreadsheet_Excel_Reader::TYPE_DELTA:

                // print dialog options
                // could be useful if using to determine whether to display
                // gridlines or headers
                case Spreadsheet_Excel_Reader::TYPE_PRINTHEADERS:
                case Spreadsheet_Excel_Reader::TYPE_PRINTGRIDLINES:
                case Spreadsheet_Excel_Reader::TYPE_GRIDSET:         // gridlines ever been set?

                case Spreadsheet_Excel_Reader::TYPE_GUTS:            // outline symbol area display options

                case Spreadsheet_Excel_Reader::TYPE_WSBOOL:          // Worksheet boolean options


                // page settings block
                case Spreadsheet_Excel_Reader::TYPE_HORIZONTALPAGEBREAKS:
                case Spreadsheet_Excel_Reader::TYPE_VERTICALPAGEBREAKS:
                case Spreadsheet_Excel_Reader::TYPE_HEADER:
                case Spreadsheet_Excel_Reader::TYPE_FOOTER:
                case Spreadsheet_Excel_Reader::TYPE_HCENTER:
                case Spreadsheet_Excel_Reader::TYPE_VCENTER:
                case Spreadsheet_Excel_Reader::TYPE_LEFTMARGIN:
                case Spreadsheet_Excel_Reader::TYPE_RIGHTMARGIN:
                case Spreadsheet_Excel_Reader::TYPE_TOPMARGIN:
                case Spreadsheet_Excel_Reader::TYPE_BOTTOMMARGIN:
                // Undocumented
                //case Spreadsheet_Excel_Reader::TYPE_PLS:
                case Spreadsheet_Excel_Reader::TYPE_SETUP:
                case Spreadsheet_Excel_Reader::TYPE_BITMAP:


                // view settings
                case Spreadsheet_Excel_Reader::TYPE_SCL:             // view magnification


                default:
                    break;
            }

            // failsafe
            fseek($this->_stream, $pos + $length + 4);

            $pos    = ftell($this->_stream);
            $code   = $this->_readInt(2);
            $length = $this->_readInt(2);

echo "\n";
echo "*** NEW RECORD ***\n";
echo "File position: 0x". dechex(ftell($this->_stream))."\n";
echo "code:          0x". dechex($code)."\n";
echo "length:        $length which is 0x".dechex($length)."\n";
        }

        if (!isset($this->sheets['numRows'])) {
             $this->sheets['numRows'] = $this->sheets['maxrow'];
        }

        if (!isset($this->sheets['numCols'])) {
             $this->sheets['numCols'] = $this->sheets['maxcol'];
        }

        return $worksheet;
    }
}

?>
