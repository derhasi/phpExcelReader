<?php

require_once 'Spreadsheet/Excel/Reader.php';

abstract class Spreadsheet_Excel_Reader_BIFFParser
{
    const BIFF8               = 0x600;
    const BIFF7               = 0x500;
    const WORKBOOKGLOBALS     = 0x5;
    const WORKSHEET           = 0x10;


    const TYPE_BOF            = 0x809;
    const TYPE_EOF            = 0xa;
    const TYPE_BOUNDSHEET     = 0x85;         // 6.12
    const TYPE_DIMENSIONS     = 0x200;
    const TYPE_ROW            = 0x208;
    const TYPE_DBCELL         = 0xd7;
    const TYPE_NOTE           = 0x1c;
    const TYPE_TXO            = 0x1b6;
    const TYPE_INDEX          = 0x20b;
    const TYPE_SST            = 0xfc;         // 6.96
    const TYPE_EXTSST         = 0xff;         // 6.40
    const TYPE_CONTINUE       = 0x3c;
    const TYPE_NAME           = 0x18;
    const TYPE_STRING         = 0x207;
    const TYPE_FORMAT         = 0x41e;        // 6.45
    const TYPE_XF             = 0xe0;         // 6.115
    const TYPE_UNKNOWN        = 0xffff;
    const TYPE_NINETEENFOUR   = 0x22;         // 6.25
    const TYPE_MERGEDCELLS    = 0xE5;

    const TYPE_UNCALCED       = 0x5e;

    const TYPE_CODEPAGE       = 0x42;

    const TYPE_DSF            = 0x161;

    const TYPE_WINDOW1        = 0x3d;

    const TYPE_BACKUP         = 0x40;

    const TYPE_HIDEOBJ        = 0x8d;

    const TYPE_FONT           = 0x31;

    const TYPE_BOOKBOOL       = 0xda;

    const TYPE_STYLE          = 0x293;

    const TYPE_PALETTE        = 0x92;

    const TYPE_USESELFS       = 0x160;

    const TYPE_COUNTRY        = 0x8c;


// file protection
    const TYPE_FILEPASS       = 0x2f;
    const TYPE_WRITEACCESS    = 0x5c;


// calculation settings
    const TYPE_CALCCOUNT      = 0xc;
    const TYPE_CALCMODE       = 0xd;
    const TYPE_PRECISION      = 0xe;
    const TYPE_REFMODE        = 0xf;
    const TYPE_DELTA          = 0x10;
    const TYPE_ITERATION      = 0x11;
    const TYPE_DATEMODE       = 0x22;
    const TYPE_SAVERECALC     = 0x5F;

    const TYPE_PRINTHEADERS       = 0x2a;
    const TYPE_PRINTGRIDLINES     = 0x2b;
    const TYPE_GRIDSET            = 0x82;
    const TYPE_GUTS               = 0x80;
    const TYPE_DEFAULTROWHEIGHT   = 0x225;
    const TYPE_WSBOOL             = 0x81;

// page settings
    const TYPE_HORIZONTALPAGEBREAKS   = 0x1b;
    const TYPE_VERTICALPAGEBREAKS     = 0x1a;
    const TYPE_HEADER                 = 0x14;
    const TYPE_FOOTER                 = 0x15;
    const TYPE_HCENTER                = 0x83;
    const TYPE_VCENTER                = 0x84;
    const TYPE_LEFTMARGIN             = 0x26;
    const TYPE_RIGHTMARGIN            = 0x27;
    const TYPE_TOPMARGIN              = 0x28;
    const TYPE_BOTTOMMARGIN           = 0x29;
//PLS UNDOCUMENTED
    const TYPE_SETUP                  = 0xa1;
    const TYPE_BITMAP                 = 0xe9;

// worksheet protection block
    const TYPE_PROTECT                = 0x12;
    const TYPE_WINDOWPROTECT          = 0x19;
    const TYPE_OBJECTPROTECT          = 0x63;
    const TYPE_SCENPROTECT            = 0xdd;
    const TYPE_PASSWORD               = 0x13;


    const TYPE_DEFCOLWIDTH            = 0x55;
    const TYPE_COLINFO                = 0x7d;


// cell block

    const TYPE_BLANK                 = 0x201;
    const TYPE_BOOLERR               = 0x205;
    const TYPE_LABEL                 = 0x204;
    const TYPE_LABELSST              = 0xfd;         // 6.61
    const TYPE_MULBLANK              = 0xbe;
    const TYPE_MULRK                 = 0xbd;
    const TYPE_NUMBER                = 0x203;
    const TYPE_RK                    = 0x27e;
    const TYPE_RSTRING               = 0xd6;


// formula cell block

    const TYPE_FORMULA               = 0x6;
    const TYPE_ARRAY                 = 0x221;
    const TYPE_SHRFMLA               = 0x4bc;
    const TYPE_TABLEOP               = 0x236;

    const RESULT_NOTFLOAT            = 0xffff;

    const RESULT_STRING              = 0x00;
    const RESULT_BOOL                = 0x01;
    const RESULT_ERROR               = 0x02;
    const RESULT_EMPTY               = 0x03;


// worksheet view settings block

    const TYPE_WINDOW2               = 0x23e;
    const TYPE_SCL                   = 0xa0;
    const TYPE_PANE                  = 0x41;
    const TYPE_SELECTION             = 0x1d;


    const TYPE_PHONETIC              = 0xef;



    const EXPONENT_BIAS   = 1023;


    /**
     * BIFF version - Either 7 or 8
     */
    public $version;

    /**
     * File Stream from within the OLE container
     * 
     */
    protected $_stream;

    /**
     * Constructor
     */
    function __construct(&$stream_handle)
    {
        $this->_stream = $stream_handle;
    }

    /**
     * Read a Little Endian integer from the data stream's current position.
     * Expects sizes either 1, 2 or 4
     *
     * @access protected
     * @param int $size Size in bytes
     * @return integer
     */
    protected function _readInt($size)
    {
        switch ($size) {
            case 1:
            $format = 'C';
            break;

            case 2:
            $format = 'v';
            break;

            case 4:
            $format = 'V';
            break;
        }

        if (($value = fread($this->_stream, $size)) === false) {
            return new Spreadsheet_Excel_Reader_Exception('Error reading stream');
        }

        list(, $value) = unpack($format, $value);
        return $value;
    }

    /**
     * Read a double precision floating point value in Little Endian order
     *
     * Determine Endian order as described in the php manual online
     * User contributed note:
     * info at dreystone dot com
     * 5-5-2005 4:31
     *
     * @access protected
     * @return float
     */
    protected function _readDouble()
    {
        if (($value = fread($this->_stream, 8)) === false) {
            return new Spreadsheet_Excel_Reader_Exception('Error reading stream');
        }
        list(, $t) = unpack('C', pack('S', 256));

        // always read little endian order
        if ($t == 1) {
            list(, $a) = unpack('d', strrev($value));
        } else {
            list(, $a) = unpack('d', $value);
        }

        return $a;
    }

    /**
     * BIFF7
     *
     * See Section 3.3
     */
    protected function _readString($length_size)
    {
        $length = $this->_readInt($length_size);
        return fread($this->_stream, $length);
    }

    /**
     * BIFF8 only
     *
     * See Section 3.4
     *
     */
    protected function _readUnicodeString($length_size)
    {
        $length = $this->_readInt($length_size);
        $options = ord(fread($this->_stream, 1));

        $ccompr   = ($options & 0x01) == 0x01;
        $phonetic = ($options & 0x04) == 0x04;
        $richtext = ($options & 0x08) == 0x08;

        if ($richtext) {
var_dump('rich text');
            $num_formatting_runs = $this->_readInt(2);
        }

        if ($phonetic) {
var_dump('phonetic');
            $extended_run_length = $this->_readInt(4);
        }

        $size = $ccompr ? $length * 2 : $length;
        $string = fread($this->_stream, $size); 

        if ($richtext) {
var_dump('rich text2');
            for ($i = 0; $i < $num_formatting_runs; $i++) {
                //FIXME split up and parse
                $format = $this->_readInt(4);
            }
        }

        if ($phonetic) {
var_dump('phonetic 2');
            //FIXME split up and parse
            $asian_settings = $this->_readInt($extended_run_length);
        }

        return $string;
    }


    /**
     * Section 6.46
     * Result of a formula
     *
     */
    protected function _readResult()
    {
        $pos = ftell($this->_stream);
        fseek($this->_stream, 6, SEEK_CUR);
        $is_float = ($this->_readInt(2) !== Spreadsheet_Excel_Reader_BIFFParser::RESULT_NOTFLOAT);
        fseek($this->_stream, $pos);
    
        if ($is_float) {

            $result = $this->_readDouble();

        } else {

            $type = $this->_readInt(1);

            switch ($type) {
                
                case Spreadsheet_Excel_Reader_BIFFParser::RESULT_STRING:
                case Spreadsheet_Excel_Reader_BIFFParser::RESULT_EMPTY:
                default:
                    fseek($this->_stream, 7, SEEK_CUR);
                    $result = null;
                    break;

                case Spreadsheet_Excel_Reader_BIFFParser::RESULT_BOOL:
                case Spreadsheet_Excel_Reader_BIFFParser::RESULT_ERROR:
                    fseek($this->_stream, 1, SEEK_CUR);
                    $result = $this->_readInt(1);
                    fseek($this->_stream, 5, SEEK_CUR);
                    break;
            }
        }

        return $result;
    }


    /**
     * TODO
     */
    protected function _readFormula()
    {
        $size = $this->_readInt(2);
        $formula = fread($this->_stream, $size);
        // todo detect additional data
    }



    /**
     * Convert an RK value into its proper format
     *
     * See Section 3.6
     *
     *
     */
    protected function _convertRKValue($value)
    {
echo "_convertRKValue()\n";
echo "input: 0x".dechex($value)."\n";

        $divide = ($value & 0x00000001) == 0x1;
        $is_int = ($value & 0x00000002) == 0x2;
        $number = ($value & 0xFFFFFFFC) >> 2;

        if (!$is_int) {
echo "converting float...\n";

            // todo
            // what about signed infinity and NaN?

            $sign     = ($number & 0x20000000) >> 29;
            $exponent = ($number & 0x1FFC0000) >> 18;
            $mantissa = ($number & 0x0003FFFF);

            // automatic float conversion
            $number = pow(-1, $sign) *
                      pow( 2, $exponent - Spreadsheet_Excel_Reader_BIFFParser::EXPONENT_BIAS) *
                      (1 + $mantissa / pow(2, 18));
        }

        if ($divide) {
echo "dividing...\n";
            $number /= 100;
        }

        return $number;
    }


    /**
     *
     */
    protected function _readCellRangeAddress($col_size)
    {
        $first_row_index = $this->_readInt(2);
        $last_row_index  = $this->_readInt(2);
        $first_col_index = $this->_readInt($col_size);
        $last_col_index  = $this->_readInt($col_size);

        if ($last_row_index == Spreadsheet_Excel_Reader::ROW_LIMIT - 1) {
            $last_row_index = '*';
        }


        if ($last_col_index == Spreadsheet_Excel_Reader::COL_LIMIT - 1) {
            $last_col_index = '*';
        }

        return array($first_row_index,
                     $first_col_index,
                     $last_row_index,
                     $last_col_index);
    }


    /**
     *
     */
    protected function _readCellRangeAddressList($col_size)
    {
        $an_array = array();
        $num_cells = $this->_readInt(2);
        for ($i = 0; $i < $num_cells; $i++) {
            $an_array[] = $this->_readCellRangeAddress($col_size);
        }

        return $an_array;
    }


    /**
     * todo
     */
    protected function _readPhoneticSettings($fh)
    {
    }
}

?>