<?php
/* vim: set expandtab tabstop=4 shiftwidth=4 softtabstop=4: */

/**
 * Class for parsing workbook streams within a BIFF file
 * 
 * @category   Spreadsheet
 * @package    Spreadsheet_Excel_Reader
 * @author     David Sanders <shangxiao@php.net>
 * @copyright  Copyright &copy; 2007, David Sanders
 * @license    LGPL <http://www.gnu.org/licenses/lgpl.html>
 * @version    Release: @release@
 * @link       http://pear.php.net/package/Spreadsheet_Excel_Reader
 * @see        OLE
 * @see        Spreadsheet_Excel
 */

require_once 'Spreadsheet/Excel/Reader/BIFFParser.php';
require_once 'Spreadsheet/Excel/Workbook.php';

/**
 * Class for parsing workbook streams within a BIFF file
 * 
 * @category   Spreadsheet
 * @package    Spreadsheet_Excel_Reader
 * @author     David Sanders <shangxiao@php.net>
 * @copyright  Copyright &copy; 2007, David Sanders
 * @license    LGPL <http://www.gnu.org/licenses/lgpl.html>
 * @version    Release: @release@
 * @link       http://pear.php.net/package/Spreadsheet_Excel_Reader
 * @see        OLE
 * @see        Spreadsheet_Excel
 */

class Spreadsheet_Excel_Reader_BIFFParser_Workbook extends Spreadsheet_Excel_Reader_BIFFParser
{
    private $_boundsheets = array();

    private $_date_formats = array(
        0xe  => "d/m/Y",
        0xf  => "d-M-Y",
        0x10 => "d-M",
        0x11 => "M-Y",
        0x12 => "h:i a",
        0x13 => "h:i:s a",
        0x14 => "H:i",
        0x15 => "H:i:s",
        0x16 => "d/m/Y H:i",
        0x2d => "i:s",
        0x2e => "H:i:s",
        0x2f => "i:s.S");

    private $_number_formats = array(
        0x1  => "%1.0f",    // "0"
        0x2  => "%1.2f",    // "0.00",
        0x3  => "%1.0f",    //"#,##0",
        0x4  => "%1.2f",    //"#,##0.00",
        0x5  => "%1.0f",    /*"$#,##0;($#,##0)",*/
        0x6  => '$%1.0f',   /*"$#,##0;($#,##0)",*/
        0x7  => '$%1.2f',   //"$#,##0.00;($#,##0.00)",
        0x8  => '$%1.2f',   //"$#,##0.00;($#,##0.00)",
        0x9  => '%1.0f%%',  // "0%"
        0xa  => '%1.2f%%',  // "0.00%"
        0xb  => '%1.2f',    // 0.00E00",
        0x25 => '%1.0f',    // "#,##0;(#,##0)",
        0x26 => '%1.0f',    //"#,##0;(#,##0)",
        0x27 => '%1.2f',    //"#,##0.00;(#,##0.00)",
        0x28 => '%1.2f',    //"#,##0.00;(#,##0.00)",
        0x29 => '%1.0f',    //"#,##0;(#,##0)",
        0x2a => '$%1.0f',   //"$#,##0;($#,##0)",
        0x2b => '%1.2f',    //"#,##0.00;(#,##0.00)",
        0x2c => '$%1.2f',   //"$#,##0.00;($#,##0.00)",
        0x30 => '%1.0f');   //"##0.0E0";


    /**
     * Open Office Excel file format 
     *
     * Each record (Section 3.1)
     * 
     * Code: Record identifier (2 bytes)
     * Length: Size of the data (2 bytes)
     *
     */
    public function parse()
    {
        $workbook = new Excel_Workbook;

        $pos = ftell($this->_stream);

        $code   = $this->_readInt(2);
        $length = $this->_readInt(2);

        assert($code === Spreadsheet_Excel_Reader_BIFFParser::TYPE_BOF);

        while ($code != Spreadsheet_Excel_Reader_BIFFParser::TYPE_EOF) {

            switch ($code) {

                // Beginning Of File
                // Section 6.8
                case Spreadsheet_Excel_Reader_BIFFParser::TYPE_BOF:

                    $version        = $this->_readInt(2);
                    $substream_type = $this->_readInt(2);

                    if ($version == Spreadsheet_Excel_Reader_BIFFParser::BIFF8) {

                        $this->version = 8;

                    } else if ($version == Spreadsheet_Excel_Reader_BIFFParser::BIFF7) {

                        $this->version = 7;

                    } else {

                        throw new Spreadsheeet_Excel_Reader_Exception('Unsupported Excel Version');
                    }

                    assert($substream_type == Spreadsheet_Excel_Reader_BIFFParser::WORKBOOKGLOBALS);

                    $workbook->build_id   = $this->_readInt(2);
                    $workbook->build_year = $this->_readInt(2);

                    if ($this->version == 8) {
                        $workbook->history_flags  = $this->_readInt(4);
                        $workbook->lowest_version = $this->_readInt(4);
                    }

                    break;


                // SST Record - Shared String Table
                // Section 6.96
                // BIFF8 only
                case Spreadsheet_Excel_Reader_BIFFParser::TYPE_SST:
//echo "Type_SST\n";

                    $total_strings  = $this->_readInt(4);
                    $unique_strings = $this->_readInt(4); 

                    for ($i = 0; $i < $unique_strings; $i++) {
                        $workbook->sst[]= $this->_readUnicodeString(2);
                    }

                    break;


                // File Password
                // Ignore?
                case Spreadsheet_Excel_Reader_BIFFParser::TYPE_FILEPASS:
//echo "Type_filepass\n";
                    //TODO
                    if ($this->version == Spreadsheet_Excel_Reader_BIFFParser::BIFF7) {

                        $encryption_key = $this->_readInt(2);
                        $hash_value     = $this->_readInt(2);

                    } else {

                        $encryption = $this->_readInt(2);

                        if ($encryption == Spreadsheet_Excel_Reader_BIFFParser::ENCRYPTION_WEAK) {

                            $encryption_key = $this->_readInt(2);
                            $hash_value     = $this->_readInt(2);

                        } else {

                            fseek($this->_stream, 2, SEEK_CUR);
                            $encryption2 = $this->_readInt(2);

                            if ($encryption2 == Spreadsheet_Excel_Reader_BIFFParser::ENCRYPTION_STANDARD) {

                                fseek($this->_stream, 48, SEEK_CUR);

                            } else {

                                fseek($this->_stream, 4, SEEK_CUR);
                                $size = $this->_readInt(4);
                                fseek($this->_stream, $size, SEEK_CUR);
                                $size = $this->_readInt(4);
                                fseek($this->_stream, $size * 2, SEEK_CUR);
                                $size = $this->_readInt(4);
                                fseek($this->_stream, $size, SEEK_CUR);
                            }
                        }

                    }

                    break;


                // Not relevant
                case Spreadsheet_Excel_Reader_BIFFParser::TYPE_NAME:
//echo "Type_NAME\n";
                    fseek($this->_stream, $length, SEEK_CUR);
                    break;


                // Section 6.45
                // Note: Currency records are always written
                case Spreadsheet_Excel_Reader_BIFFParser::TYPE_FORMAT:

                    $index = $this->_readInt(2);

                    if ($this->version == 8) {
                        $format_string = $this->_readUnicodeString(2);
                    } else {
                        $format_string = $this->_readString(2);
                    }

                    $workbook->format_records[$index] = $format_string;
                    break;


                // Section 6.115
                case Spreadsheet_Excel_Reader_BIFFParser::TYPE_XF:

                    $font_index   = $this->_readInt(2);
                    $format_index = $this->_readInt(2);

                    if (array_key_exists($format_index, $this->_date_formats)) {

                        $workbook->xf_records[] = array(
                            'type' => 'date',
                            'format_index' => $format_index,
                            'format' => $this->_date_formats[$format_index]
                            );

                    } elseif (array_key_exists($format_index, $this->_number_formats)) {

                        $workbook->xf_records[] = array(
                            'type' => 'number',
                            'format_index' => $format_index,
                            'format' => $this->_number_formats[$format_index]
                            );

                    } else {

                        $isdate = false;

                        if ($format_index > 0 && isset($workbook->format_records[$format_index])) {

                            $formatstr = $workbook->format_records[$format_index];
//echo '.other.';
//echo "\ndate-time=$formatstr=\n";
                            if ($formatstr && preg_match("/[^hmsday\/\-:\s]/i", $formatstr) == 0) { // found day and time format
                                $isdate = true;
                                $formatstr = str_replace('mm', 'i', $formatstr);
                                $formatstr = str_replace('h', 'H', $formatstr);
//echo "\ndate-time $formatstr \n";
                            }
                        }

                        if ($isdate){
                            $workbook->xf_records[] = array(
                                'type' => 'date',
                                'format_index' => $format_index,
                                'format' => $formatstr,
                                );
                        } else {
                            $workbook->xf_records[] = array(
                                'type' => 'other',
                                'format' => '',
                                'format_index' => $format_index,
                                'code' => $format_index
                                );
                        }
                    }

                    break;


                // Section 6.25
                case Spreadsheet_Excel_Reader_BIFFParser::TYPE_DATEMODE:

                    $workbook->datemode = $this->_readInt(2);
                    break;


                // Section 6.12
                case Spreadsheet_Excel_Reader_BIFFParser::TYPE_BOUNDSHEET:

                    $offset     = $this->_readInt(4);
                    $visibility = $this->_readInt(1);
                    $type       = $this->_readInt(1);

                    if ($this->version == 8) {
                        $name = $this->_readUnicodeString(1);
                    } else {
                        $name = $this->_readString(1);
                    }

                    // Type worksheet and visible
                    if ($visibility == 0 && $type == 0) {
                        $this->boundsheets[] = array('name'   => $name,
                                                     'offset' => $offset);
                    }

                    break;

                case Spreadsheet_Excel_Reader_BIFFParser::TYPE_WRITEACCESS:

                    if ($this->version == Spreadsheet_Excel_Reader_BIFFParser::BIFF7) {
                        $workbook->user = $this->_readString(1);
                    } else {
                        $workbook->user = $this->_readUnicodeString(2);
                    }

                    break;


                // text encoding is only relevant for biff7
                case Spreadsheet_Excel_Reader_BIFFParser::TYPE_CODEPAGE:
//echo "type codepage\n";
                    fseek($this->_stream, 2, SEEK_CUR);
                    break;


                // Double Stream File flag
                // not relevant
                case Spreadsheet_Excel_Reader_BIFFParser::TYPE_DSF:
//echo "type dsf\n";
                    fseek($this->_stream, 2, SEEK_CUR);
                    break;


                // Window Settings Protection
                // not relevant
                case Spreadsheet_Excel_Reader_BIFFParser::TYPE_WINDOWPROTECT:
//echo "type windowprotect\n";
                    fseek($this->_stream, 2, SEEK_CUR);
                    break;


                // Workbook Protection
                // Flags whether a worksheet/workbook is protected
                case Spreadsheet_Excel_Reader_BIFFParser::TYPE_PROTECT:
//echo "type protect\n";
                    // todo store
                    $protection = $this->_readInt(2);
                    break;


                // PASSWORD
                // 0 means no password
                case Spreadsheet_Excel_Reader_BIFFParser::TYPE_PASSWORD:
//echo "type password\n";
                    // todo store
                    $password = $this->_readInt(2);
                    break;


                case Spreadsheet_Excel_Reader_BIFFParser::TYPE_WINDOW1:
//echo "type window1\n";
                    // useful information?
                    $horizontal_position        = $this->_readInt(2);
                    $vertical_position          = $this->_readInt(2);
                    $width                      = $this->_readInt(2);
                    $height                     = $this->_readInt(2);
                    $options                    = $this->_readInt(2);
                    $workbook->active_worksheet = $this->_readInt(2);
                    $first_visible_tab          = $this->_readInt(2);
                    $num_selected_sheets        = $this->_readInt(2);
                    $tab_bar_width              = $this->_readInt(2);
                    break;


                // Formula Calculation Precision
                // Flags whether formulas use the real values or displayed values for calculation
                // todo store?
                case Spreadsheet_Excel_Reader_BIFFParser::TYPE_PRECISION:
//echo "type precision\n";
                    fseek($this->_stream, 2, SEEK_CUR);
                    break;


                case Spreadsheet_Excel_Reader_BIFFParser::TYPE_FONT:
//echo "type font\n";
                    $font_record = array();
                    $font_record['height']     = $this->_readInt(2);
                    $font_record['options']    = $this->_readInt(2);
                    $font_record['colour']     = $this->_readInt(2);
                    $font_record['weight']     = $this->_readInt(2);
                    $font_record['escapement'] = $this->_readInt(2);
                    $font_record['underline']  = $this->_readInt(1);
                    $font_record['family']     = $this->_readInt(1);
                    $font_record['charset']    = $this->_readInt(1);

                    fseek($this->_stream, 1, SEEK_CUR);

                    if ($this->version == Spreadsheet_Excel_Reader_BIFFParser::BIFF7) {
                        $font_record['name'] = $this->_readString(1);
                    } else {
                        $font_record['name'] = $this->_readString(2);
                    }

                    $workbook->font_records[] = $font_record;

                    break;


                // Style record
                // Stores the name for user defined, options for built-in
                // Section 6.99
                case Spreadsheet_Excel_Reader_BIFFParser::TYPE_STYLE:
//echo "type style\n";
                    $xf_index = $this->_readInt(2);

                    if (($xf_index & 0x8000) == 0x8000) {

                        // built in style
                        $style = $this->_readInt(1);
                        $level = $this->_readInt(1);

                        $workbook->style_records['built_in'][$xf_index] = array(
                            'style' => $style,
                            'level' => $level);
                    } else {

                        // user defined style
                        if ($this->version == Spreadsheet_Excel_Reader_BIFFParser::BIFF7) {

                            $name = $this->_readString(1);

                        } else {

                            $name = $this->_readUnicodeString(2);
                        }

                        $workbook->style_records['user'][$xf_index] = $name;
                    }

                    break;


                case Spreadsheet_Excel_Reader_BIFFParser::TYPE_PALETTE:
//echo "type palette\n";
                    $num_colours = $this->_readInt(2);

                    $palette = array();

                    for ($i = 0; $i < $num_colours; $i++) {
                        $colour_index = $this->_readInt(4);

                        // convert a built in colour to a standard hexadecimal format

                        // TODO: store system colours?
                        switch ($colour_index) {

                            case 0x1:
                            $colour = 0xffffff;
                            break;

                            case 0x2:
                            $colour = 0xff0000;
                            break;

                            case 0x3:
                            $colour = 0x00ff00;
                            break;

                            case 0x4:
                            $colour = 0x0000ff;
                            break;

                            case 0x5:
                            $colour = 0xffff00;
                            break;

                            case 0x6:
                            $colour = 0xff00ff;
                            break;

                            case 0x7:
                            $colour = 0x00ffff;
                            break;

                            case 0x4f:
                            default:
                            $colour = 0x000000;
                            break;
                        }

                        $palette[] = $colour;
                    }

                    $workbook->palette = $palette;

                    break;


                case Spreadsheet_Excel_Reader_BIFFParser::TYPE_USESELFS:
//echo "type useselfs\n";
                    // todo store
                    $use_natural_language_formulas = $this->_readInt(2);
                    break;


                case Spreadsheet_Excel_Reader_BIFFParser::TYPE_COUNTRY:
//echo "type country\n";
                    // todo store
                    $ui_country       = $this->_readInt(2);
                    $settings_country = $this->_readInt(2);
                    break;


                // Create a backup on saving
                // not relevant
                case Spreadsheet_Excel_Reader_BIFFParser::TYPE_BACKUP:

                // Hide Objects
                // not relevant
                case Spreadsheet_Excel_Reader_BIFFParser::TYPE_HIDEOBJ:

                // save values from external books
                // not relevant
                case Spreadsheet_Excel_Reader_BIFFParser::TYPE_BOOKBOOL:

                // Extended SST
                //
                // hash table to offsets within the sst
                // not relevant
                case Spreadsheet_Excel_Reader_BIFFParser::TYPE_EXTSST:

                default:

                    break;

/*
echo "WARNING: UNKNOWN RECORD TYPE\n";
echo 'File position: '. dechex(ftell($this->_stream))."\n";
echo "Default data:\n";
if ($length > 0) {
    echo fread($this->_stream, $length);
}
echo "\n\n";
break;
*/
            }


            // failsafe
            fseek($this->_stream, $pos + $length + 4);
            $pos = ftell($this->_stream);

            $code          = $this->_readInt(2);
            $length        = $this->_readInt(2);

/*
echo "\n";
echo "*** NEW RECORD ***\n";
echo "File position: 0x". dechex(ftell($this->_stream))."\n";
echo "code:          0x". dechex($code)."\n";
echo "length:        $length which is 0x".dechex($length)."\n";
*/
        }

        foreach ($this->boundsheets as $sheet_index => $boundsheet) {
            fseek($this->_stream, $boundsheet['offset']);
echo '** Parsing sheet at offset: '.dechex($boundsheet['offset'])."\n";

            require_once 'Spreadsheet/Excel/Reader/BIFFParser/Worksheet.php';
            $parser = new Spreadsheet_Excel_Reader_BIFFParser_Worksheet($this->_stream, $workbook, $this->version);
            $worksheet = $parser->parse();
            $worksheet->name = $boundsheet['name'];

            $workbook->worksheets[] = $worksheet;
        }

        return $workbook;
    }
}

?>
