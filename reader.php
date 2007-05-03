<?php
/* vim: set expandtab tabstop=4 shiftwidth=4 softtabstop=4: */

/**
 * A class for reading Microsoft Excel Spreadsheets.
 *
 * Based on PHPExcelReader by Vadim Tkachenko.
 * (http://sourceforge.net/projects/phpexcelreader)
 * Reads only Biff 7 and Biff 8 formats.
 *
 * PHP version 5
 *
 * @category   Spreadsheet
 * @package    Spreadsheet_Excel_Reader
 * @author     Vadim Tkachenko <vt@apachephp.com>
 * @author     David Sanders <shangxiao@php.net>
 * @license    TODO
 * @version    CVS: $Id: reader.php 19 2007-03-13 12:42:41Z shangxiao $
 * @link       http://pear.php.net/package/Spreadsheet_Excel_Reader
 * @see        OLE
 */

require_once 'PEAR/Exception.php';
class Spreadsheet_Excel_Reader_Exception extends PEAR_Exception{}

require_once 'OLE.php';

define('SPREADSHEET_EXCEL_READER_BIFF8',             0x600);
define('SPREADSHEET_EXCEL_READER_BIFF7',             0x500);
define('SPREADSHEET_EXCEL_READER_WORKBOOKGLOBALS',   0x5);
define('SPREADSHEET_EXCEL_READER_WORKSHEET',         0x10);

define('SPREADSHEET_EXCEL_READER_STREAM_NAME_BIFF7', 'Book');
define('SPREADSHEET_EXCEL_READER_STREAM_NAME_BIFF8', 'Workbook');

define('SPREADSHEET_EXCEL_READER_TYPE_BOF',          0x809);
define('SPREADSHEET_EXCEL_READER_TYPE_EOF',          0xa);
define('SPREADSHEET_EXCEL_READER_TYPE_BOUNDSHEET',   0x85);         // 6.12
define('SPREADSHEET_EXCEL_READER_TYPE_DIMENSIONS',   0x200);
define('SPREADSHEET_EXCEL_READER_TYPE_ROW',          0x208);
define('SPREADSHEET_EXCEL_READER_TYPE_DBCELL',       0xd7);
define('SPREADSHEET_EXCEL_READER_TYPE_NOTE',         0x1c);
define('SPREADSHEET_EXCEL_READER_TYPE_TXO',          0x1b6);
define('SPREADSHEET_EXCEL_READER_TYPE_INDEX',        0x20b);
define('SPREADSHEET_EXCEL_READER_TYPE_SST',          0xfc);         // 6.96
define('SPREADSHEET_EXCEL_READER_TYPE_EXTSST',       0xff);         // 6.40
define('SPREADSHEET_EXCEL_READER_TYPE_CONTINUE',     0x3c);
define('SPREADSHEET_EXCEL_READER_TYPE_NAME',         0x18);
define('SPREADSHEET_EXCEL_READER_TYPE_STRING',       0x207);
define('SPREADSHEET_EXCEL_READER_TYPE_FORMAT',       0x41e);        // 6.45
define('SPREADSHEET_EXCEL_READER_TYPE_XF',           0xe0);         // 6.115
define('SPREADSHEET_EXCEL_READER_TYPE_UNKNOWN',      0xffff);
define('SPREADSHEET_EXCEL_READER_TYPE_NINETEENFOUR', 0x22);         // 6.25
define('SPREADSHEET_EXCEL_READER_TYPE_MERGEDCELLS',  0xE5);

define('SPREADSHEET_EXCEL_READER_TYPE_UNCALCED',     0x5e);

define('SPREADSHEET_EXCEL_READER_TYPE_CODEPAGE',     0x42);

define('SPREADSHEET_EXCEL_READER_TYPE_DSF',          0x161);

define('SPREADSHEET_EXCEL_READER_TYPE_WINDOW1',      0x3d);

define('SPREADSHEET_EXCEL_READER_TYPE_BACKUP',       0x40);

define('SPREADSHEET_EXCEL_READER_TYPE_HIDEOBJ',      0x8d);

define('SPREADSHEET_EXCEL_READER_TYPE_FONT',         0x31);

define('SPREADSHEET_EXCEL_READER_TYPE_BOOKBOOL',     0xda);

define('SPREADSHEET_EXCEL_READER_TYPE_STYLE',        0x293);

define('SPREADSHEET_EXCEL_READER_TYPE_PALETTE',      0x92);

define('SPREADSHEET_EXCEL_READER_TYPE_USESELFS',     0x160);

define('SPREADSHEET_EXCEL_READER_TYPE_COUNTRY',      0x8c);


// file protection
define('SPREADSHEET_EXCEL_READER_TYPE_FILEPASS',     0x2f);
define('SPREADSHEET_EXCEL_READER_TYPE_WRITEACCESS',  0x5c);


// calculation settings
define('SPREADSHEET_EXCEL_READER_TYPE_CALCCOUNT',    0xc);
define('SPREADSHEET_EXCEL_READER_TYPE_CALCMODE',     0xd);
define('SPREADSHEET_EXCEL_READER_TYPE_PRECISION',    0xe);
define('SPREADSHEET_EXCEL_READER_TYPE_REFMODE',      0xf);
define('SPREADSHEET_EXCEL_READER_TYPE_DELTA',        0x10);
define('SPREADSHEET_EXCEL_READER_TYPE_ITERATION',    0x11);
define('SPREADSHEET_EXCEL_READER_TYPE_DATEMODE',     0x22);
define('SPREADSHEET_EXCEL_READER_TYPE_SAVERECALC',   0x5F);

define('SPREADSHEET_EXCEL_READER_TYPE_PRINTHEADERS',     0x2a);
define('SPREADSHEET_EXCEL_READER_TYPE_PRINTGRIDLINES',   0x2b);
define('SPREADSHEET_EXCEL_READER_TYPE_GRIDSET',          0x82);
define('SPREADSHEET_EXCEL_READER_TYPE_GUTS',             0x80);
define('SPREADSHEET_EXCEL_READER_TYPE_DEFAULTROWHEIGHT', 0x225);
define('SPREADSHEET_EXCEL_READER_TYPE_WSBOOL',           0x81);

// page settings
define('SPREADSHEET_EXCEL_READER_TYPE_HORIZONTALPAGEBREAKS', 0x1b);
define('SPREADSHEET_EXCEL_READER_TYPE_VERTICALPAGEBREAKS',   0x1a);
define('SPREADSHEET_EXCEL_READER_TYPE_HEADER',               0x14);
define('SPREADSHEET_EXCEL_READER_TYPE_FOOTER',               0x15);
define('SPREADSHEET_EXCEL_READER_TYPE_HCENTER',              0x83);
define('SPREADSHEET_EXCEL_READER_TYPE_VCENTER',              0x84);
define('SPREADSHEET_EXCEL_READER_TYPE_LEFTMARGIN',           0x26);
define('SPREADSHEET_EXCEL_READER_TYPE_RIGHTMARGIN',          0x27);
define('SPREADSHEET_EXCEL_READER_TYPE_TOPMARGIN',            0x28);
define('SPREADSHEET_EXCEL_READER_TYPE_BOTTOMMARGIN',         0x29);
//PLS UNDOCUMENTED
define('SPREADSHEET_EXCEL_READER_TYPE_SETUP',                0xa1);
define('SPREADSHEET_EXCEL_READER_TYPE_BITMAP',               0xe9);

// worksheet protection block
define('SPREADSHEET_EXCEL_READER_TYPE_PROTECT',              0x12);
define('SPREADSHEET_EXCEL_READER_TYPE_WINDOWPROTECT',        0x19);
define('SPREADSHEET_EXCEL_READER_TYPE_OBJECTPROTECT',        0x63);
define('SPREADSHEET_EXCEL_READER_TYPE_SCENPROTECT',          0xdd);
define('SPREADSHEET_EXCEL_READER_TYPE_PASSWORD',             0x13);


define('SPREADSHEET_EXCEL_READER_TYPE_DEFCOLWIDTH',          0x55);
define('SPREADSHEET_EXCEL_READER_TYPE_COLINFO',              0x7d);


// cell block

define('SPREADSHEET_EXCEL_READER_TYPE_BLANK',               0x201);
define('SPREADSHEET_EXCEL_READER_TYPE_BOOLERR',             0x205);
define('SPREADSHEET_EXCEL_READER_TYPE_LABEL',               0x204);
define('SPREADSHEET_EXCEL_READER_TYPE_LABELSST',            0xfd);         // 6.61
define('SPREADSHEET_EXCEL_READER_TYPE_MULBLANK',            0xbe);
define('SPREADSHEET_EXCEL_READER_TYPE_MULRK',               0xbd);
define('SPREADSHEET_EXCEL_READER_TYPE_NUMBER',              0x203);
define('SPREADSHEET_EXCEL_READER_TYPE_RK',                  0x27e);
define('SPREADSHEET_EXCEL_READER_TYPE_RSTRING',             0xd6);


// formula cell block

define('SPREADSHEET_EXCEL_READER_TYPE_FORMULA',             0x6);
define('SPREADSHEET_EXCEL_READER_TYPE_ARRAY',               0x221);
define('SPREADSHEET_EXCEL_READER_TYPE_SHRFMLA',             0x4bc);
define('SPREADSHEET_EXCEL_READER_TYPE_TABLEOP',             0x236);

define('SPREADSHEET_EXCEL_READER_RESULT_NOTFLOAT',          0xffff);

define('SPREADSHEET_EXCEL_READER_RESULT_STRING',            0x00);
define('SPREADSHEET_EXCEL_READER_RESULT_BOOL',              0x01);
define('SPREADSHEET_EXCEL_READER_RESULT_ERROR',             0x02);
define('SPREADSHEET_EXCEL_READER_RESULT_EMPTY',             0x03);


// worksheet view settings blick

define('SPREADSHEET_EXCEL_READER_TYPE_WINDOW2',             0x23e);
define('SPREADSHEET_EXCEL_READER_TYPE_SCL',                 0xa0);
define('SPREADSHEET_EXCEL_READER_TYPE_PANE',                0x41);
define('SPREADSHEET_EXCEL_READER_TYPE_SELECTION',           0x1d);


define('SPREADSHEET_EXCEL_READER_TYPE_PHONETIC',            0xef);


define('SPREADSHEET_EXCEL_READER_UTCOFFSETDAYS' ,    25569);
define('SPREADSHEET_EXCEL_READER_UTCOFFSETDAYS1904', 24107);
define('SPREADSHEET_EXCEL_READER_MSINADAY',          86400);
//define('SPREADSHEET_EXCEL_READER_MSINADAY', 24 * 60 * 60);

//define('SPREADSHEET_EXCEL_READER_DEF_NUM_FORMAT', "%.2f");
define('SPREADSHEET_EXCEL_READER_DEF_NUM_FORMAT',    "%s");


define('SPREADSHEET_EXCEL_READER_EXPONENT_BIAS', 1023);


// stuff from schmitty
define('NUM_BIG_BLOCK_DEPOT_BLOCKS_POS', 0x2c);
define('SMALL_BLOCK_DEPOT_BLOCK_POS', 0x3c);
define('ROOT_START_BLOCK_POS', 0x30);
define('BIG_BLOCK_SIZE', 0x200);
define('SMALL_BLOCK_SIZE', 0x40);
define('EXTENSION_BLOCK_POS', 0x44);
define('NUM_EXTENSION_BLOCK_POS', 0x48);
define('PROPERTY_STORAGE_BLOCK_SIZE', 0x80);
define('BIG_BLOCK_DEPOT_BLOCKS_POS', 0x4c);
define('SMALL_BLOCK_THRESHOLD', 0x1000);


/*
* Place includes, constant defines and $_GLOBAL settings here.
* Make sure they have appropriate docblocks to avoid phpDocumentor
* construing they are documented by the page-level docblock.
*/

/**
 * A class for reading Microsoft Excel Spreadsheets.
 *
 * Originally developed by Vadim Tkachenko under the name PHPExcelReader.
 * (http://sourceforge.net/projects/phpexcelreader)
 * Based on the Java version by Andy Khan (http://www.andykhan.com).  Now
 * maintained by David Sanders.  Reads only Biff 7 and Biff 8 formats.
 *
 * @category   Spreadsheet
 * @package    Spreadsheet_Excel_Reader
 * @author     Vadim Tkachenko <vt@phpapache.com>
 * @copyright  1997-2005 The PHP Group
 * @license    http://www.php.net/license/3_0.txt  PHP License 3.0
 * @version    Release: @package_version@
 * @link       http://pear.php.net/package/PackageName
 * @see        OLE, Spreadsheet_Excel_Writer
 */
class Spreadsheet_Excel_Reader
{
    var $version;

    /**
     * Array of worksheets found
     *
     * @var array
     * @access public
     */
    var $boundsheets = array();

    /**
     * Array of format records found
     * 
     * @var array
     * @access public
     */
    var $formatRecords = array();

    /**
     * todo
     *
     * @var array
     * @access public
     */
    var $sst = array();

    /**
     * Array of worksheets
     *
     * The data is stored in 'cells' and the meta-data is stored in an array
     * called 'cellsInfo'
     *
     * Example:
     *
     * $sheets  -->  'cells'  -->  row --> column --> Interpreted value
     *          -->  'cellsInfo' --> row --> column --> 'type' - Can be 'date', 'number', or 'unknown'
     *                                            --> 'raw' - The raw data that Excel stores for that data cell
     *
     * @var array
     * @access public
     */
    var $sheets = array();

    /**
     * The data returned by OLE
     *
     * @var string
     * @access public
     */
    var $data;

    /**
     * OLE object for reading the file
     *
     * @var OLE object
     * @access private
     */
    var $_ole;

    /**
     * Default encoding
     *
     * @var string
     * @access private
     */
    var $_defaultEncoding;

    /**
     * Default number format
     *
     * @var integer
     * @access private
     */
    var $_defaultFormat = SPREADSHEET_EXCEL_READER_DEF_NUM_FORMAT;

    /**
     * todo
     * List of formats to use for each column
     *
     * @var array
     * @access private
     */
    var $_columnsFormat = array();

    /**
     * todo
     *
     * @var integer
     * @access private
     */
    var $_rowoffset = 1;

    /**
     * todo
     *
     * @var integer
     * @access private
     */
    var $_coloffset = 1;

    /**
     * List of default date formats used by Excel
     *
     * @see Section 6.45
     * @var array
     * @access public
     */
    var $dateFormats = array (
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

    /**
     * Default number formats used by Excel
     *
     * @see Section 6.45
     * @var array
     * @access public
     */
    var $numberFormats = array(
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




    var $calcmode;

    // }}}
    // {{{ Spreadsheet_Excel_Reader()

    /**
     * Constructor
     *
     * Some basic initialisation
     */ 
    function Spreadsheet_Excel_Reader()
    {
        //$this->_ole =& new OLERead();
        $this->setUTFEncoder('iconv');
    }

    // }}}
    // {{{ setOutputEncoding()

    /**
     * Set the encoding method
     *
     * @param string Encoding to use
     * @access public
     */
    function setOutputEncoding($encoding)
    {
        $this->_defaultEncoding = $encoding;
    }

    // }}}
    // {{{ setUTFEncoder()

    /**
     *  $encoder = 'iconv' or 'mb'
     *  set iconv if you would like use 'iconv' for encode UTF-16LE to your encoding
     *  set mb if you would like use 'mb_convert_encoding' for encode UTF-16LE to your encoding
     *
     * @access public
     * @param string Encoding type to use.  Either 'iconv' or 'mb'
     */
    function setUTFEncoder($encoder = 'iconv')
    {
        $this->_encoderFunction = '';

        if ($encoder == 'iconv') {
            $this->_encoderFunction = function_exists('iconv') ? 'iconv' : '';
        } elseif ($encoder == 'mb') {
            $this->_encoderFunction = function_exists('mb_convert_encoding') ?
                                      'mb_convert_encoding' :
                                      '';
        }
    }

    // }}}
    // {{{ setRowColOffset()

    /**
     * todo
     *
     * @access public
     * @param offset
     */
    function setRowColOffset($iOffset)
    {
        $this->_rowoffset = $iOffset;
        $this->_coloffset = $iOffset;
    }

    // }}}
    // {{{ setDefaultFormat()

    /**
     * Set the default number format
     *
     * @access public
     * @param Default format
     */
    function setDefaultFormat($sFormat)
    {
        $this->_defaultFormat = $sFormat;
    }

    // }}}
    // {{{ setColumnFormat()

    /**
     * Force a column to use a certain format
     *
     * @access public
     * @param integer Column number
     * @param string Format
     */
    function setColumnFormat($column, $sFormat)
    {
        $this->_columnsFormat[$column] = $sFormat;
    }


    private $data_stream;


    // }}}
    // {{{ read()

    /**
     * Read the spreadsheet file using OLE, then parse
     *
     * @access public
     * @param filename
     * @todo return a valid value
     */
    function read($sFileName)
    {
        $ole = new OLE;
        $ole->read($sFileName);

        foreach ($ole->_list as $pps) {
            if (//$pps->Size >= SMALL_BLOCK_THRESHOLD &&
                ($pps->Name == 'Workbook' || $pps->Name == 'Book')) {
                $this->data = $ole->getStream($pps);
                break;
            }
        }

        require_once 'Spreadsheet/Excel/Reader/Parser/Workbook.php';
        $parser = new Spreadsheet_Excel_Reader_Parser_Workbook($this->data);
        return $parser->parse();


/*
        foreach ($ole->_list as $i => $pps) {
            if (($pps->Name == 'Workbook' || $pps->Name == 'Book') &&
                $pps->Size >= SMALL_BLOCK_THRESHOLD) {

                $this->data = $ole->getData($i, 0, $ole->getDataLength($i));
            } elseif ($pps->Name == 'Root Entry') {
                $this->data = $ole->getData($i, 0, $ole->getDataLength($i));
            }
            //var_dump(strlen($ole->getData($i, 0, $ole->getDataLength($i))), $pps->Name, md5($this->data), $ole->getDataLength($i));
        }
//exit;
*/
     //   $this->_parse2();

       // return sizeof($this->sheets) > 0;

/*
        $res = $this->_ole->read($sFileName);

        // oops, something goes wrong (Darko Miljanovic)
        if($res === false) {
            // check error code
            if($this->_ole->error == 1) {
            // bad file
                die('The filename ' . $sFileName . ' is not readable');
            }
            // check other error codes here (eg bad fileformat, etc...)
        }

        $this->data = $this->_ole->getWorkBook();
*/

        /*
        $res = $this->_ole->read($sFileName);

        if ($this->isError($res)) {
//        var_dump($res);
            return $this->raiseError($res);
        }

        $total = $this->_ole->ppsTotal();
        for ($i = 0; $i < $total; $i++) {
            if ($this->_ole->isFile($i)) {
                $type = unpack("v", $this->_ole->getData($i, 0, 2));
                if ($type[''] == 0x0809)  { // check if it's a BIFF stream
                    $this->_index = $i;
                    $this->data = $this->_ole->getData($i, 0, $this->_ole->getDataLength($i));
                    break;
                }
            }
        }

        if ($this->_index === null) {
            return $this->raiseError("$file doesn't seem to be an Excel file");
        }

        */

    //echo "data =".$this->data;
        //$this->readRecords();
        $this->_parse();
    }


    /**
     * Read a Little Endian integer from the data stream's current position.
     * Expects sizes either 1, 2 or 4
     *
     * @access private
     * @param int $size Size in bytes
     */
    function _readInt($size = 2)
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

        if (($value = fread($this->data, $size)) === false) {
            // TODO
            // throw exception?
            die('Error reading stream');
        }
        list(, $value) = unpack($format, $value);
        return $value;
    }


    /**
     * Read a double precision floating point value in Little Endian order
     *
     * As described in the php manual online
     * User contributed note:
     * info at dreystone dot com
     * 5-5-2005 4:31
     *
     */
    function _readDouble()
    {
        if (($value = fread($this->data, 8)) === false) {
            // TODO
            // throw exception?
            die('Error reading stream');
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
    function _readString($length_size)
    {
        $length = $this->_readInt($length_size);
        return fread($this->data, $length);
    }

    /*
     * BIFF8 only
     *
     * See Section 3.4
     *
     */
    function _readUnicodeString($length_size)
    {
echo "Reading Unicode String\n";

        $length = $this->_readInt($length_size);
        $options = ord(fread($this->data, 1));

        $ccompr   = ($options & 0x01) == 0x01;
        $phonetic = ($options & 0x04) == 0x04;
        $richtext = ($options & 0x08) == 0x08;

echo "length:   $length\n";
echo "ccompr:   $ccompr\n";
echo "phonetic: $phonetic\n";
echo "richtext: $richtext\n";

        if ($richtext) {
        var_dump('rich text');
            $num_formatting_runs = $this->_readInt(2);
        }

        if ($phonetic) {
        var_dump('phonetic');
            $extended_run_length = $this->_readInt(4);
        }

        $size = $ccompr ? $length * 2 : $length;
        $string = fread($this->data, $size); 

        if ($richtext) {
        var_dump('rich text2');
            for ($i = 0; $i < $num_formatting_rums; $i++) {
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
    function _readResult()
    {
        $pos = ftell($this->data);
        fseek($this->data, 6, SEEK_CUR);
        $is_float = ($this->_readInt(2) !== SPREADSHEET_EXCEL_READER_RESULT_NOTFLOAT);
        fseek($this->data, $pos);

        if ($is_float) {

            $result = $this->_readDouble();

        } else {

            $type = $this->_readInt(1);

            switch ($type) {
                
                case SPREADSHEET_EXCEL_READER_RESULT_STRING:
                case SPREADSHEET_EXCEL_READER_RESULT_EMPTY:
                default:
                    fseek($this->data, 7, SEEK_CUR);
                    $result = null;
                    break;

                case SPREADSHEET_EXCEL_READER_RESULT_BOOL:
                case SPREADSHEET_EXCEL_READER_RESULT_ERROR:
                    fseek($this->data, 1, SEEK_CUR);
                    $result = $this->_readInt(1);
                    fseek($this->data, 5, SEEK_CUR);
                    break;
            }
        }

        return $result;
    }


    /**
     * TODO
     */
    function _readFormula()
    {
        $size = $this->_readInt(2);
        $formula = fread($this->data, $size);
        // todo detect additional data
    }



    /**
     * Convert an RK value into its proper format
     *
     * See Section 3.6
     *
     *
     */
    function _convertRKValue($value)
    {
echo "_convertRKValue()\n";
echo "input: 0x".dechex($value)."\n";

        $divide = ($value & 0x00000001) == 0x1;
        $is_int = ($value & 0x00000002) == 0x2;
        $number = ($value & 0XFFFFFFFC) >> 2;

        if (!$is_int) {
echo "converting float...\n";

            // todo
            // what about signed infinity and NaN?

            $sign     = ($number & 0x20000000) >> 29;
            $exponent = ($number & 0x1FFC0000) >> 18;
            $mantissa = ($number & 0x0003FFFF);

            // automatic float conversion
            $number = pow(-1, $sign) *
                      pow( 2, $exponent - SPREADSHEET_EXCEL_READER_EXPONENT_BIAS) *
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
    function _readCellRangeAddress($col_size)
    {
        $first_row_index = $this->_readInt(2);
        $last_row_index  = $this->_readInt(2);
        $first_col_index = $this->_readInt($col_size);
        $last_col_index  = $this->_readInt($col_size);

        return array($first_row_index,
                     $last_row_index,
                     $first_col_index,
                     $last_col_index);
    }


    /**
     *
     */
    function _readCellRangeAddressList($col_size)
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
    function _readPhoneticSettings($fh)
    {
    }


    /**
     *
     * Open Office Excel file format 
     *
     * Each record (Section 3.1)
     * 
     * Code: Record identifier (2 bytes)
     * Length: Size of the data (2 bytes)
     *
     */ 
     var $type;
    function _parse2()
    {
        $pos = ftell($this->data);
        echo "Pos: ".$pos."\n";

        $code          = $this->_readInt(2);
        $length        = $this->_readInt(2);

        echo "Code: 0x".dechex($code)."\n";
        echo "Length: ".$length."\n";

        assert($code === SPREADSHEET_EXCEL_READER_TYPE_BOF);

        while ($code != SPREADSHEET_EXCEL_READER_TYPE_EOF) {

            switch ($code) {


                // Beginning Of File
                // Section 6.8
                case SPREADSHEET_EXCEL_READER_TYPE_BOF:

                    $this->version = $this->_readInt(2);
                    $substreamType = $this->_readInt(2);

                    echo "Version: 0x".dechex($this->version)."\n";
                    echo "Substream Type: 0x".dechex($substreamType)."\n";


                    if ($this->version != SPREADSHEET_EXCEL_READER_BIFF8 &&
                        $this->version != SPREADSHEET_EXCEL_READER_BIFF7) {
                        // TODO exception("Unsupported Excel Version");
                        return false;
                    }

assert($substreamType == SPREADSHEET_EXCEL_READER_WORKBOOKGLOBALS);

                    $build_id   = $this->_readInt(2);
                    $build_year = $this->_readInt(2);

                    if ($this->version == SPREADSHEET_EXCEL_READER_BIFF8) {
                        $history_flags   = $this->_readInt(4);
                        $loweset_version = $this->_readInt(4);
                    }

                    break;


                // SST Record - Shared String Table
                // Section 6.96
                // BIFF8 only
                case SPREADSHEET_EXCEL_READER_TYPE_SST:
echo "Type_SST\n";
                    
                    $total_strings  = $this->_readInt(4);
                    $unique_strings = $this->_readInt(4); 

echo "Total strings: $total_strings\n";
echo "Unique strings: $unique_strings\n";

                    for ($i = 0; $i < $unique_strings; $i++) {
                        $this->sst[]= $this->_readUnicodeString(2);
                    }
echo "Strings:\n";
var_dump($this->sst);

                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_FILEPASS:
echo "Type_filepass\n";
                    if ($this->version == SPREADSHEET_EXCEL_READER_BIFF7) {

                        $encryption_key = $this->_readInt(2);
                        $hash_value     = $this->_readInt(2);

                    } else {

                        $encryption = $this->_readInt(2);

                        if ($encryption == SPREADSHEET_EXCEL_READER_ENCRYPTION_WEAK) {

                            $encryption_key = $this->_readInt(2);
                            $hash_value     = $this->_readInt(2);

                        } else {

                            fseek($this->data, 2, SEEK_CUR);
                            $encryption2 = $this->_readInt(2);

                            if ($encryption2 == SPREADSHEET_EXCEL_READER_ENCRYPTION_STANDARD) {

                                // TODO
                                fseek($this->data, 48, SEEK_CUR);

                            } else {

                                // TODO
                                fseek($this->data, 4, SEEK_CUR);
                                $size = $this->_readInt(4);
                                fseek($this->data, $size, SEEK_CUR);
                                $size = $this->_readInt(4);
                                fseek($this->data, $size * 2, SEEK_CUR);
                                $size = $this->_readInt(4);
                                fseek($this->data, $size, SEEK_CUR);
                            }
                        }

                    }

                    break;


                case SPREADSHEET_EXCEL_READER_TYPE_NAME:
echo "Type_NAME\n";
                    fseek($this->data, $length, SEEK_CUR);
                    break;


                case SPREADSHEET_EXCEL_READER_TYPE_FORMAT:
echo "Type_format\n";

                    $indexCode = $this->_readInt(2);

                    if ($this->version == SPREADSHEET_EXCEL_READER_BIFF8) {
                        $formatString = $this->_readUnicodeString(2);
echo "Format string: $formatString\n";
                    } else {
                        // TODO CHECK THIS
                        $formatString = $this->_readString(2);
                        // why multiply 2?
                        //$numchars = ord($this->data[$pos+6]);
                        //$formatString = substr($this->data, $pos+7, $numchars*2);
                    }

                    $this->formatRecords[$indexCode] = $formatString;
                    break;

    
                // Section 6.115
                case SPREADSHEET_EXCEL_READER_TYPE_XF:
echo 'Type_xf'."\n";

                    $fontIndex   = $this->_readInt(2);
                    $formatIndex = $this->_readInt(2);

echo "XF font index: $fontIndex\n";
echo "XF format index: $formatIndex\n";

                    if (array_key_exists($formatIndex, $this->dateFormats)) {
echo "isdate ".$this->dateFormats[$formatIndex]."\n";
                        $this->formatRecords['xfrecords'][] = array(
                            'type' => 'date',
                            'format' => $this->dateFormats[$formatIndex]
                            );
                    } elseif (array_key_exists($formatIndex, $this->numberFormats)) {
echo "isnumber ".$this->numberFormats[$formatIndex]."\n";
                            $this->formatRecords['xfrecords'][] = array(
                            'type' => 'number',
                            'format' => $this->numberFormats[$formatIndex]
                            );
                    } else {
                        $isdate = FALSE;
                        if ($formatIndex > 0 && isset($this->formatRecords[$formatIndex])) {
                            $formatstr = $this->formatRecords[$formatIndex];
echo '.other.';
echo "\ndate-time=$formatstr=\n";
                            if ($formatstr && preg_match("/[^hmsday\/\-:\s]/i", $formatstr) == 0) { // found day and time format
                                $isdate = TRUE;
                                $formatstr = str_replace('mm', 'i', $formatstr);
                                $formatstr = str_replace('h', 'H', $formatstr);
echo "\ndate-time $formatstr \n";
                            }
                        }

                        if ($isdate){
                            $this->formatRecords['xfrecords'][] = array(
                                'type' => 'date',
                                'format' => $formatstr,
                                );
                        } else {
                            $this->formatRecords['xfrecords'][] = array(
                                'type' => 'other',
                                'format' => '',
                                'code' => $formatIndex
                                );
                        }
                    }

                    break;


                // Section 6.25
                case SPREADSHEET_EXCEL_READER_TYPE_DATEMODE:
echo "Type.DATEMODE\n";
                    $this->nineteenFour = $this->_readInt(2) == 1;
                    break;

                // Section 6.12
                case SPREADSHEET_EXCEL_READER_TYPE_BOUNDSHEET:
echo "Type.BOUNDSHEET\n";

                    // ORDER CHANGED FROM ORIGINAL
                    $offset     = $this->_readInt(4);
                    $visibility = $this->_readInt(1);
                    $type       = $this->_readInt(1);

echo "offset:     $offset\n";
echo "visibility: $visibility\n";
echo "type:       $type\n";

assert($visibility >= 0 && $visibility <= 2);
assert($type === 0 || $type === 2 || $type === 6);

$this->type = 'boundsheet';

                    if ($this->version == SPREADSHEET_EXCEL_READER_BIFF8) {
                        $rec_name = $this->_readUnicodeString(1);
                    } else {
                        $rec_name = $this->_readString(1);
                    }
echo "sheet name: $rec_name\n";
            
                    $this->boundsheets[] = array('name'   => $rec_name,
                                                 'offset' => $offset);

                    break;

            }


            fseek($this->data, $pos + $length + 4);
            $pos = ftell($this->data);

            $code          = $this->_readInt(2);
            $length        = $this->_readInt(2);

echo "\n";
echo "File position:   $pos\n";
echo "Record code  : 0x".dechex($code)."\n";
echo "Record length:   $length\n";
        }

        foreach ($this->boundsheets as $sheet_index => $boundsheet) {
            fseek($this->data, $boundsheet['offset']);
echo '** Parsing sheet at offset: '.dechex($boundsheet['offset'])."\n";
            $this->_parsesheet2($sheet_index);
        }

        return true;
    }


    // }}}
    // {{{ _parse()

    /**
     * Parse a workbook
     *
     * @access private
     * @return bool
     */
    function _parse()
    {
//~
        $fh = $this->data;
        $this->data = stream_get_contents($fh);

        $pos = 0;

        $code = ord($this->data[$pos]) | ord($this->data[$pos+1])<<8;
        $length = ord($this->data[$pos+2]) | ord($this->data[$pos+3])<<8;

        $version = ord($this->data[$pos + 4]) | ord($this->data[$pos + 5])<<8;
        $substreamType = ord($this->data[$pos + 6]) | ord($this->data[$pos + 7])<<8;
        //echo "Start parse code=".base_convert($code,10,16)." version=".base_convert($version,10,16)." substreamType=".base_convert($substreamType,10,16).""."\n";

        if (($version != SPREADSHEET_EXCEL_READER_BIFF8) &&
            ($version != SPREADSHEET_EXCEL_READER_BIFF7)) {
            return false;
        }

        if ($substreamType != SPREADSHEET_EXCEL_READER_WORKBOOKGLOBALS){
            return false;
        }

        //print_r($rec);
        $pos += $length + 4;

        $code = ord($this->data[$pos]) | ord($this->data[$pos+1])<<8;
        $length = ord($this->data[$pos+2]) | ord($this->data[$pos+3])<<8;

        while ($code != SPREADSHEET_EXCEL_READER_TYPE_EOF) {
            switch ($code) {
                case SPREADSHEET_EXCEL_READER_TYPE_SST:
                    //echo "Type_SST\n";
                     $spos = $pos + 4;
                     $limitpos = $spos + $length;
                     $uniqueStrings = $this->_GetInt4d($this->data, $spos+4);
                                                $spos += 8;
                                       for ($i = 0; $i < $uniqueStrings; $i++) {
        // Read in the number of characters
                                                if ($spos == $limitpos) {
                                                $opcode = ord($this->data[$spos]) | ord($this->data[$spos+1])<<8;
                                                $conlength = ord($this->data[$spos+2]) | ord($this->data[$spos+3])<<8;
                                                        if ($opcode != 0x3c) {
                                                                return -1;
                                                        }
                                                $spos += 4;
                                                $limitpos = $spos + $conlength;
                                                }
                                                $numChars = ord($this->data[$spos]) | (ord($this->data[$spos+1]) << 8);
                                                //echo "i = $i pos = $pos numChars = $numChars ";
                                                $spos += 2;
                                                $optionFlags = ord($this->data[$spos]);
                                                $spos++;
                                        $asciiEncoding = (($optionFlags & 0x01) == 0) ;
                                                $extendedString = ( ($optionFlags & 0x04) != 0);

                                                // See if string contains formatting information
                                                $richString = ( ($optionFlags & 0x08) != 0);

                                                if ($richString) {
                                        // Read in the crun
                                                        $formattingRuns = ord($this->data[$spos]) | (ord($this->data[$spos+1]) << 8);
                                                        $spos += 2;
                                                }

                                                if ($extendedString) {
                                                  // Read in cchExtRst
                                                  $extendedRunLength = $this->_GetInt4d($this->data, $spos);
                                                  $spos += 4;
                                                }

                                                $len = ($asciiEncoding)? $numChars : $numChars*2;
                                                if ($spos + $len < $limitpos) {
                                                                $retstr = substr($this->data, $spos, $len);
                                                                $spos += $len;
                                                }else{
                                                        // found countinue
                                                        $retstr = substr($this->data, $spos, $limitpos - $spos);
                                                        $bytesRead = $limitpos - $spos;
                                                        $charsLeft = $numChars - (($asciiEncoding) ? $bytesRead : ($bytesRead / 2));
                                                        $spos = $limitpos;

                                                         while ($charsLeft > 0){
                                                                $opcode = ord($this->data[$spos]) | ord($this->data[$spos+1])<<8;
                                                                $conlength = ord($this->data[$spos+2]) | ord($this->data[$spos+3])<<8;
                                                                        if ($opcode != 0x3c) {
                                                                                return -1;
                                                                        }
                                                                $spos += 4;
                                                                $limitpos = $spos + $conlength;
                                                                $option = ord($this->data[$spos]);
                                                                $spos += 1;
                                                                  if ($asciiEncoding && ($option == 0)) {
                                                                                $len = min($charsLeft, $limitpos - $spos); // min($charsLeft, $conlength);
                                                                    $retstr .= substr($this->data, $spos, $len);
                                                                    $charsLeft -= $len;
                                                                    $asciiEncoding = true;
                                                                  }elseif (!$asciiEncoding && ($option != 0)){
                                                                                $len = min($charsLeft * 2, $limitpos - $spos); // min($charsLeft, $conlength);
                                                                    $retstr .= substr($this->data, $spos, $len);
                                                                    $charsLeft -= $len/2;
                                                                    $asciiEncoding = false;
                                                                  }elseif (!$asciiEncoding && ($option == 0)) {
                                                                // Bummer - the string starts off as Unicode, but after the
                                                                // continuation it is in straightforward ASCII encoding
                                                                                $len = min($charsLeft, $limitpos - $spos); // min($charsLeft, $conlength);
                                                                        for ($j = 0; $j < $len; $j++) {
                                                                 $retstr .= $this->data[$spos + $j].chr(0);
                                                                }
                                                            $charsLeft -= $len;
                                                                $asciiEncoding = false;
                                                                  }else{
                                                            $newstr = '';
                                                                    for ($j = 0; $j < strlen($retstr); $j++) {
                                                                      $newstr = $retstr[$j].chr(0);
                                                                    }
                                                                    $retstr = $newstr;
                                                                                $len = min($charsLeft * 2, $limitpos - $spos); // min($charsLeft, $conlength);
                                                                    $retstr .= substr($this->data, $spos, $len);
                                                                    $charsLeft -= $len/2;
                                                                    $asciiEncoding = false;
                                                                        //echo "Izavrat\n";
                                                                  }
                                                          $spos += $len;

                                                         }
                                                }
                                                $retstr = ($asciiEncoding) ? $retstr : $this->_encodeUTF16($retstr);
//                                              echo "Str $i = $retstr\n";
                                        if ($richString){
                                                  $spos += 4 * $formattingRuns;
                                                }

                                                // For extended strings, skip over the extended string data
                                                if ($extendedString) {
                                                  $spos += $extendedRunLength;
                                                }
                                                        //if ($retstr == 'Derby'){
                                                        //      echo "bb\n";
                                                        //}
                                                $this->sst[]=$retstr;
                                       }
                    /*$continueRecords = array();
                    while ($this->getNextCode() == Type_CONTINUE) {
                        $continueRecords[] = &$this->nextRecord();
                    }
                    //echo " 1 Type_SST\n";
                    $this->shareStrings = new SSTRecord($r, $continueRecords);
                    //print_r($this->shareStrings->strings);
                     */
                     // echo 'SST read: '.($time_end-$time_start)."\n";
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_FILEPASS:
                    return false;
                    break;
                case SPREADSHEET_EXCEL_READER_TYPE_NAME:
                    //echo "Type_NAME\n";
                    break;
                case SPREADSHEET_EXCEL_READER_TYPE_FORMAT:
                        $indexCode = ord($this->data[$pos+4]) | ord($this->data[$pos+5]) << 8;

                        if ($version == SPREADSHEET_EXCEL_READER_BIFF8) {
                            $numchars = ord($this->data[$pos+6]) | ord($this->data[$pos+7]) << 8;
                            if (ord($this->data[$pos+8]) == 0){
                                $formatString = substr($this->data, $pos+9, $numchars);
                            } else {
                                $formatString = substr($this->data, $pos+9, $numchars*2);
                            }
                        } else {
                            $numchars = ord($this->data[$pos+6]);
                            $formatString = substr($this->data, $pos+7, $numchars*2);
                        }

                    $this->formatRecords[$indexCode] = $formatString;
                   // echo "Type.FORMAT\n";
                    break;
                case SPREADSHEET_EXCEL_READER_TYPE_XF:
                        //global $dateFormats, $numberFormats;
                        $indexCode = ord($this->data[$pos+6]) | ord($this->data[$pos+7]) << 8;
                        //echo "\nType.XF ".count($this->formatRecords['xfrecords'])." $indexCode ";
                        if (array_key_exists($indexCode, $this->dateFormats)) {
                            //echo "isdate ".$dateFormats[$indexCode];
                            $this->formatRecords['xfrecords'][] = array(
                                    'type' => 'date',
                                    'format' => $this->dateFormats[$indexCode]
                                    );
                        }elseif (array_key_exists($indexCode, $this->numberFormats)) {
                        //echo "isnumber ".$this->numberFormats[$indexCode];
                            $this->formatRecords['xfrecords'][] = array(
                                    'type' => 'number',
                                    'format' => $this->numberFormats[$indexCode]
                                    );
                        }else{
                            $isdate = FALSE;
                            if ($indexCode > 0){
                                if (isset($this->formatRecords[$indexCode]))
                                    $formatstr = $this->formatRecords[$indexCode];
                                //echo '.other.';
                                //echo "\ndate-time=$formatstr=\n";
                                if ($formatstr)
                                if (preg_match("/[^hmsday\/\-:\s]/i", $formatstr) == 0) { // found day and time format
                                    $isdate = TRUE;
                                    $formatstr = str_replace('mm', 'i', $formatstr);
                                    $formatstr = str_replace('h', 'H', $formatstr);
                                    //echo "\ndate-time $formatstr \n";
                                }
                            }

                            if ($isdate){
                                $this->formatRecords['xfrecords'][] = array(
                                        'type' => 'date',
                                        'format' => $formatstr,
                                        );
                            }else{
                                $this->formatRecords['xfrecords'][] = array(
                                        'type' => 'other',
                                        'format' => '',
                                        'code' => $indexCode
                                        );
                            }
                        }
                        //echo "\n";
                    break;
                case SPREADSHEET_EXCEL_READER_TYPE_NINETEENFOUR:
                    //echo "Type.NINETEENFOUR\n";
                    $this->nineteenFour = (ord($this->data[$pos+4]) == 1);
                    break;
                case SPREADSHEET_EXCEL_READER_TYPE_BOUNDSHEET:
                    //echo "Type.BOUNDSHEET\n";
                        $rec_offset = $this->_GetInt4d($this->data, $pos+4);
                        $rec_typeFlag = ord($this->data[$pos+8]);
                        $rec_visibilityFlag = ord($this->data[$pos+9]);
                        $rec_length = ord($this->data[$pos+10]);
//~
echo "rec_offset: $rec_offset\n";
echo "visibility: $rec_visibilityFlag\n";
echo "type: $rec_typeFlag\n";
echo "string length: $rec_length\n";

                        if ($version == SPREADSHEET_EXCEL_READER_BIFF8){
                            $chartype =  ord($this->data[$pos+11]);
                            if ($chartype == 0){
                                $rec_name    = substr($this->data, $pos+12, $rec_length);
                            } else {
                                $rec_name    = $this->_encodeUTF16(substr($this->data, $pos+12, $rec_length*2));
                            }
                        }elseif ($version == SPREADSHEET_EXCEL_READER_BIFF7){
                                $rec_name    = substr($this->data, $pos+11, $rec_length);
                        }
//~
echo "rec_name: $rec_name\n";
return;
                    $this->boundsheets[] = array('name'=>$rec_name,
                                                 'offset'=>$rec_offset);

                    break;

            }

            //echo "Code = ".base_convert($r['code'],10,16)."\n";
            $pos += $length + 4;
            $code = ord($this->data[$pos]) | ord($this->data[$pos+1])<<8;
            $length = ord($this->data[$pos+2]) | ord($this->data[$pos+3])<<8;

            //$r = &$this->nextRecord();
            //echo "1 Code = ".base_convert($r['code'],10,16)."\n";
        }

        foreach ($this->boundsheets as $key=>$val){
            $this->sn = $key;
            $this->_parsesheet($val['offset']);
        }
        return true;

    }

    /**
     * Parse a worksheet
     *
     * @access private
     * @param todo
     * @todo fix return codes
     */
    function _parsesheet2($sheet_index)
    {

        $code   = $this->_readInt(2);
        $length = $this->_readInt(2);

echo "code:          0x". dechex($code)."\n";
echo "length:        $length which is 0x".dechex($length)."\n";

        assert($code == SPREADSHEET_EXCEL_READER_TYPE_BOF);

        $row_block_count = 0;

        while($code != SPREADSHEET_EXCEL_READER_TYPE_EOF) {

            $this->sheets[$sheet_index]['maxrow'] = $this->_rowoffset - 1;
            $this->sheets[$sheet_index]['maxcol'] = $this->_coloffset - 1;

            unset($this->rectype);
            $this->multiplier = 1; // need for format with %

            switch ($code) {


                // Section 6.8
                case SPREADSHEET_EXCEL_READER_TYPE_BOF:

                    // The version in worksheet streams cannot be trusted
                    //$version       = $this->_readInt(2);
                    fseek($this->data, 2, SEEK_CUR);
                    $substreamType = $this->_readInt(2);
                    $build_id      = $this->_readInt(2);
                    $build_year    = $this->_readInt(2);

echo "substreamType: 0x". dechex($substreamType)."\n";
echo "build id:      ".   $build_id."\n";
echo "build year:    ".   $build_year."\n";

                    if ($this->version == SPREADSHEET_EXCEL_READER_BIFF8) {
                        $file_history_flags = $this->_readInt(4);
                        $lowest_version     = $this->_readInt(4);

echo "file history flags: ".   $file_history_flags."\n";
echo "lowest version:     ".   $lowest_version."\n";
                    }


                    if ($substreamType != SPREADSHEET_EXCEL_READER_WORKSHEET) {
                        return -2;
                    }

                    break;


                // Section 6.104 - not used
                case SPREADSHEET_EXCEL_READER_TYPE_UNCALCED:
        
                    fseek($this->data, 2, SEEK_CUR);
                    break;

                // Section 6.55
                case SPREADSHEET_EXCEL_READER_TYPE_INDEX:
echo "type index\n";

                    //TODO - store
                    fseek($this->data, 4, SEEK_CUR);
                    $int_size = $this->version == SPREADSHEET_EXCEL_READER_BIFF7 ? 2 : 4;
                    $rf = $this->_readInt($int_size);
                    $rl = $this->_readInt($int_size);
                    fseek($this->data, 4, SEEK_CUR);
                    // FIXME
                    // floor or ceil?
                    //$nm = floor(($rl - $rf - 1) / (32 + 1));
                    $nm = ceil(($rl - $rf - 1) / (32 + 1));
                    fseek($this->data, $nm * 4, SEEK_CUR);

                    break;


                // 
                // --- BEGIN calculation settings ---
                // 

                // occurs in every stream and is global for the entire workbook

                case SPREADSHEET_EXCEL_READER_TYPE_CALCCOUNT:
                case SPREADSHEET_EXCEL_READER_TYPE_CALCMODE:
                case SPREADSHEET_EXCEL_READER_TYPE_PRECISION:
                case SPREADSHEET_EXCEL_READER_TYPE_REFMODE:
                case SPREADSHEET_EXCEL_READER_TYPE_ITERATION:
                // datemode should appear in the globals stream only
                //case SPREADSHEET_EXCEL_READER_TYPE_DATEMODE:
                case SPREADSHEET_EXCEL_READER_TYPE_SAVERECALC:
echo "calc settings\n";
                    fseek($this->data, 2, SEEK_CUR);
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_DELTA:

echo "calc settings\n";
                    fseek($this->data, 8, SEEK_CUR);
                    break;

                //
                // --- END calculation settings ---
                //

                case SPREADSHEET_EXCEL_READER_TYPE_PRINTHEADERS:
echo "print headers\n";
                    $print_headers = $this->_readInt(2);
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_PRINTGRIDLINES:
echo "print gridlines\n";
                    $print_gridlines = $this->_readInt(2);
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_GRIDSET:
echo "gridset\n";
                    $gridset = $this->_readInt(2);
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_GUTS:
echo "guts\n";
                    $width  = $this->_readInt(2);
                    $height = $this->_readInt(2);
                    $num_visible_row_outline_levels = $this->_readInt(2);
                    $num_visible_col_outline_levels = $this->_readInt(2);
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_DEFAULTROWHEIGHT:
echo "DEFAULTROWHEIGHT\n";
                    $options        = $this->_readInt(2);
                    $default_unused = $this->_readInt(2);
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_WSBOOL:
echo "wsbool\n";
                    $wsbool = $this->_readInt(2);
                    break;


                //
                // --- BEGIN page settings ---
                //

                case SPREADSHEET_EXCEL_READER_TYPE_HORIZONTALPAGEBREAKS:
                case SPREADSHEET_EXCEL_READER_TYPE_VERTICALPAGEBREAKS:
echo "page breaks\n";

                    $num_indexes = $this->_readInt(2);

                    if ($this->version == SPREADSHEET_EXCEL_READER_BIFF8) {
                        fseek($this->data, $num_indexes * 6, SEEK_CUR);
                    } else {
                        fseek($this->data, $num_indexes * 2, SEEK_CUR);
                    }
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_HEADER:
echo "header\n";

                    if ($length == 0) {
                        break;
                    }

                    if ($this->version == SPREADSHEET_EXCEL_READER_BIFF8) {
                        $header = $this->_readUnicodeString(2);
                    } else {
                        $header = $this->_readString(1);
                    }

                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_FOOTER:
echo "footer\n";

                    if ($length == 0) {
                        break;
                    }

                    if ($this->version == SPREADSHEET_EXCEL_READER_BIFF8) {
                        $footer = $this->_readUnicodeString(2);
                    } else {
                        $footer = $this->_readString(1);
                    }

                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_HCENTER:
echo "hcenter\n";
                    $hcenter = $this->_readInt(2);
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_VCENTER:
echo "vcenter\n";
                    $vcenter = $this->_readInt(2);
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_LEFTMARGIN:
echo "leftmargin\n";
                    // TODO
                    // store
                    fseek($this->data, 8, SEEK_CUR);
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_RIGHTMARGIN:
echo "rightmargin\n";
                    // TODO
                    // store
                    fseek($this->data, 8, SEEK_CUR);
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_TOPMARGIN:
echo "topmargin\n";
                    // TODO
                    // store
                    fseek($this->data, 8, SEEK_CUR);
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_BOTTOMMARGIN:
echo "bottommargin\n";
                    // TODO
                    // store
                    fseek($this->data, 8, SEEK_CUR);
                    break;

                // Undocumented
                //case SPREADSHEET_EXCEL_READER_TYPE_PLS:
                    //break;

                case SPREADSHEET_EXCEL_READER_TYPE_SETUP:
echo "setup\n";
                    $paper_size                = $this->_readInt(2);
                    $scaling_factor            = $this->_readInt(2);
                    $start_pageno              = $this->_readInt(2);
                    $width_restriction         = $this->_readInt(2);
                    $height_restriction        = $this->_readInt(2);
                    $options                   = $this->_readInt(2);
                    $print_resolution          = $this->_readInt(2);
                    $vertical_print_resolution = $this->_readInt(2);
                    // todo store header and footer margins
                    fseek($this->data, 16, SEEK_CUR);
                    $num_copies_to_print       = $this->_readInt(2);
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_BITMAP:
echo "bitmap\n";
                    // unknown values
                    fseek($this->data, 4, SEEK_CUR);

                    $size          = $this->_readInt(4);

                    fseek($this->data, 4, SEEK_CUR);

                    $bitmap_width  = $this->_readInt(2);
                    $bitmap_height = $this->_readInt(2);
                    $num_planes    = $this->_readInt(2);
                    $colour_depth  = $this->_readInt(2);
                    $colour_depth  = $this->_readInt(2);

                    $line_size = floor($width * 3 / 4) + 4;
                    $bitmap    = fread($this->data, $height * $line_size); 

                    break;


                //
                // --- END page settings ---
                //


                //
                // --- BEGIN worksheet protection ---
                //

                case SPREADSHEET_EXCEL_READER_TYPE_PROTECT:
echo "protect\n";
                    $protected = $this->_readInt(2);
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_WINDOWPROTECT:
echo "windowprotect\n";
                    $window_settings_protected = $this->_readInt(2);
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_OBJECTPROTECT:
echo "objectprotect\n";
                    $objects_protected = $this->_readInt(2);
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_SCENPROTECT:
echo "scenprotect\n";
                    $scenarios_protected = $this->_readInt(2);
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_PASSWORD:
echo "password\n";
                    $password = fread($this->data, 2);
                    break;

                //
                // --- END worksheet protection ---
                //

                case SPREADSHEET_EXCEL_READER_TYPE_DEFCOLWIDTH:
echo "defcolwidth\n";
                    $col_width = $this->_readInt(2);
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_COLINFO:
echo "colinfo\n";
                    $first_col  = $this->_readInt(2);
                    $last_col   = $this->_readInt(2);
                    $col_width  = $this->_readInt(2);
                    $xf_index   = $this->_readInt(2);
                    $options    = $this->_readInt(2);
                    fseek($this->data, 2, SEEK_CUR);
                    break;


                // Section 6.31
                case SPREADSHEET_EXCEL_READER_TYPE_DIMENSIONS:
echo "dimensions\n";
                    if ($this->version == SPREADSHEET_EXCEL_READER_BIFF7){
                        $first_row = $this->_readInt(2);
                        $last_row  = $this->_readInt(2);
                    } else {
                        $first_row = $this->_readInt(4);
                        $last_row  = $this->_readInt(4);
                    }

                    $first_col = $this->_readInt(2);
                    $last_col  = $this->_readInt(2);

                    $this->sheets[$sheet_index]['numRows'] = $last_row - $first_row;
                    $this->sheets[$sheet_index]['numCols'] = $last_col - $first_col;

                    fseek($this->data, 2, SEEK_CUR);

                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_ROW:
echo "row\n";
                    $row_index  = $this->_readInt(2);
                    $first_col  = $this->_readInt(2);
                    $last_col   = $this->_readInt(2);
                    $row_height = $this->_readInt(2);
                    fseek($this->data, 4, SEEK_CUR);
                    $options    = $this->_readInt(4);

                    $row_block_count++;

                    break;



                //
                // --- BEGIN cell block ---
                //

                case SPREADSHEET_EXCEL_READER_TYPE_BLANK:
echo "BLANK\n";
                    // TODO
                    // store information

                    $row_index = $this->_readInt(2);
                    $col_index = $this->_readInt(2);
                    $xf_index  = $this->_readInt(2);

                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_BOOLERR:
echo "type boolerr\n";

                    // TODO
                    // Retrieve error codes as described in Section 3.7

                    $row_index = $this->_readInt(2);
                    $col_index = $this->_readInt(2);
                    $xf_index  = $this->_readInt(2);
                    $type      = $this->_readInt(1);
                    $data      = $this->_readInt(1);
                    
                    $this->addcell($sheet_index, $row_index, $col_index, $data);
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_LABEL:
echo "Type label\n";
                    $row_index = $this->_readInt(2);
                    $col_index = $this->_readInt(2);
                    $xf_index  = $this->_readInt(2);

                    if ($this->version == SPREADSHEET_EXCEL_READER_BIFF7) {
                        $label = $this->_readString(2);
                    } else {
                        $label = $this->_readUnicodeString(2);
                    }

                    $this->addcell($sheet_index, $row_index, $col_index, $label);

                    break;

                // Section 6.61
                case SPREADSHEET_EXCEL_READER_TYPE_LABELSST:
echo "labelsst"."\n";

                    $row_index = $this->_readInt(2);
                    $col_index = $this->_readInt(2);
                    $xf_index  = $this->_readInt(2);
                    $sst_index = $this->_readInt(4);

                    $this->addcell($sheet_index, $row_index, $col_index, $this->sst[$sst_index]);

                    break;

                // Section 6.64
                // Multiple blank
                case SPREADSHEET_EXCEL_READER_TYPE_MULBLANK:
echo "MULBLANK\n";

                    // TODO
                    // Store information

                    $row_index       = $this->_readInt(2);
                    $first_col_index = $this->_readInt(2);
                    
                    // the last col index appears after the data!

                    $temp_pos = ftell($this->data);
                    fseek($this->data, $length - 6, SEEK_CUR);
                    $last_col_index  = $this->_readInt(2);
                    fseek($this->data, $temp_pos);

                    $xf_indexes = array();
                    $num_cols = $last_col_index - $first_col_index + 1;
                    for ($i = 0; $i < $num_cols; $i++) {
                        $xf_indexes[] = $this->_readInt(2);
                    }

                    fseek($this->data, 2, SEEK_CUR);

                    break;

                // Section 6.64
                // Multiple RK
                case SPREADSHEET_EXCEL_READER_TYPE_MULRK:
echo "type mulrk"."\n";
                    
                    $row_index       = $this->_readInt(2);
                    $first_col_index = $this->_readInt(2);
                    
                    // the last col index appears after the data!

                    $temp_pos = ftell($this->data);
                    fseek($this->data, $length - 6, SEEK_CUR);
                    $last_col_index  = $this->_readInt(2);
                    fseek($this->data, $temp_pos);

                    $num_cols = $last_col_index - $first_col_index + 1;
                    for ($i = 0; $i < $num_cols; $i++) {

                        $xf_index = $this->_readInt(2);
                        $value    = $this->_readInt(4);

                        if ($this->isDate2($xf_index)) {
                            list($string, $raw) = $this->createDate($value);
                        } else {
                            $raw = $numValue;

                            if (isset($this->_columnsFormat[$colFirst + $i + 1])) {
                                $this->curformat = $this->_columnsFormat[$colFirst + $i + 1];
                            }

                            $string = sprintf($this->curformat, $numValue * $this->multiplier);
                        }

                        $this->addcell($sheet_index, $row_index, $first_col_index + $i, $string, $raw);
                    }

                    break;


                // Stores floating point values that cannot be stored as an RK value?
                // Section 6.68

                case SPREADSHEET_EXCEL_READER_TYPE_NUMBER:
echo "type number\n";

                    $row_index = $this->_readInt(2);
                    $col_index = $this->_readInt(2);
                    $xf_index  = $this->_readInt(2);
                    $number    = $this->_readDouble();

                    if ($this->isDate2($xf_index)) {
                        list($string, $raw) = $this->createDate($number);
                    } else {
                        if (isset($this->_columnsFormat[$col_index + 1])) {
                            $this->curformat = $this->_columnsFormat[$col_index + 1];
                        }
                        $raw = $number;
                        $string = sprintf($this->curformat, $number);
                    }

                    $this->addcell($sheet_index, $row_index, $col_index, $string, $raw);

                    break;


                case SPREADSHEET_EXCEL_READER_TYPE_RK:

echo 'SPREADSHEET_EXCEL_READER_TYPE_RK'."\n";

                    $row_index = $this->_readInt(2);        
                    $col_index = $this->_readInt(2);        
                    $xf_index  = $this->_readInt(2);        

                    $rk_value  = $this->_readInt(4);
                    $number    = $this->_convertRKValue($rk_value);
echo "RK number: $number\n";

                    if ($this->isDate2($xf_index)) {
echo "is date\n";
                        list($string, $raw) = $this->createDate($number);
echo "date: $string \n";
                    } else {
                        $raw = $number;
                        if (isset($this->_columnsFormat[$col_index + 1])){
                                $this->curformat = $this->_columnsFormat[$col_index + 1];
                        }

                        // todo
                        // multipler??
                        $string = sprintf($this->curformat, $number * $this->multiplier);
                    }

                    $this->addcell($sheet_index, $row_index, $col_index, $string, $raw);

                    break;


                // BIFF7 Rich text strings
                // BIFF8 Uses this record only for the clipboard
                // Section 6.84

                case SPREADSHEET_EXCEL_READER_TYPE_RSTRING:

                    $row_index = $this->_readInt(2);
                    $col_index = $this->_readInt(2);
                    $xf_index  = $this->_readInt(2);
                    $string    = $this->_readString(2);
                    $num_runs  = $this->_readInt(1);

                    for ($i = 0; $i < $num_runs; $i++) {
                        $run_list  = $this->_readInt(2);
                    }

                    break;


                case SPREADSHEET_EXCEL_READER_TYPE_FORMULA:
echo "type formula\n";

                    $row_index = $this->_readInt(2);
                    $col_index = $this->_readInt(2);
                    $xf_index  = $this->_readInt(2);
                    // todo result may be stored in 1 of 5 ways
                    $result    = $this->_readResult();
                    $options   = $this->_readInt(2);
                    fseek($this->data, 4, SEEK_CUR);
                    // todo read the formula data
                    $formula   = $this->_readFormula();


                    // todo good enough check?
                    if (is_float($result)) {
                        if ($this->isDate2($xf_index)) {
                            list($string, $raw) = $this->createDate($result);
                        } else {
                            if (isset($this->_columnsFormat[$col_index + 1])) {
                                $this->curformat = $this->_columnsFormat[$col_index + 1];
                            }

                            $raw = $result;
                            $string = sprintf($this->curformat, $raw * $this->multiplier);
                        }

                        $this->addcell($sheet_index, $row_index, $col_index, $string, $raw);
                    }

                    break;


                case SPREADSHEET_EXCEL_READER_TYPE_ARRAY:

                    // TODO
                    fseek($this->data, 12, SEEK_CUR);
                    $this->_readFormula();
                    break;


                // Shared Formula
                // Section 6.94
                case SPREADSHEET_EXCEL_READER_TYPE_SHRFMLA:
                    
                    // TODO
                    fseek($this->data, 8, SEEK_CUR);
                    $this->_readFormula();
                    break;


                case SPREADSHEET_EXCEL_READER_TYPE_TABLEOP:

                    // TODO
                    fseek($this->data, 16, SEEK_CUR);
                    break;


                case SPREADSHEET_EXCEL_READER_TYPE_DBCELL:
echo 'type dbcell'."\n";

                    //todo?
                    fseek($this->data, $row_block_count * 2 + 4, SEEK_CUR);
                    $row_block_count = 0;
                    break;


                //
                // Worksheet View Settings Block
                //

                case SPREADSHEET_EXCEL_READER_TYPE_WINDOW2:
echo "type window2\n";
                    $options   = $this->_readInt(2);
                    $row_index = $this->_readInt(2);
                    $col_index = $this->_readInt(2);

                    if ($this->version == SPREADSHEET_EXCEL_READER_BIFF7) {

                        $gridline_colour = $this->_readInt(4);

                    } else {

                        $gridline_colour_index    = $this->_readInt(2);

                        fseek($this->data, 2, SEEK_CUR);

                        $page_break_magnification = $this->_readInt(2);
                        $normal_magnification     = $this->_readInt(2);

                        fseek($this->data, 4, SEEK_CUR);
                    }

                    break;


                case SPREADSHEET_EXCEL_READER_TYPE_SCL:
echo "type scl\n";
                    // todo needed?
                    fseek($this->data, 2, SEEK_CUR);
                    break;


                case SPREADSHEET_EXCEL_READER_TYPE_PANE:
echo "type pane\n";
                    // todo needed?
                    fseek($this->data, 9, SEEK_CUR);
                    break;


                case SPREADSHEET_EXCEL_READER_TYPE_SELECTION:
echo "type selection\n";
                    // todo store
                    $pane_id          = $this->_readInt(1);
                    $row_index        = $this->_readInt(2);
                    $col_index        = $this->_readInt(2);
                    $cell_range_index = $this->_readInt(2);
                    $selected_cells   = $this->_readCellRangeAddressList(1);
                    break;


                case SPREADSHEET_EXCEL_READER_TYPE_PHONETIC:
echo "type phonetic\n";
                    // todo store
                    $font_index     = $this->_readInt(2);
                    $settings       = $this->_readInt(2);
                    $phonetic_cells = $this->_readCellRangeAddressList(2);
                    break;


                case SPREADSHEET_EXCEL_READER_TYPE_MERGEDCELLS:
echo 'type merged cell'."\n";
exit;
                    $cellRanges = ord($this->data[$spos]) | ord($this->data[$spos+1])<<8;
                    for ($i = 0; $i < $cellRanges; $i++) {
                        $fr =  ord($this->data[$spos + 8*$i + 2]) | ord($this->data[$spos + 8*$i + 3])<<8;
                        $lr =  ord($this->data[$spos + 8*$i + 4]) | ord($this->data[$spos + 8*$i + 5])<<8;
                        $fc =  ord($this->data[$spos + 8*$i + 6]) | ord($this->data[$spos + 8*$i + 7])<<8;
                        $lc =  ord($this->data[$spos + 8*$i + 8]) | ord($this->data[$spos + 8*$i + 9])<<8;
                        //$this->sheets[$this->sn]['mergedCells'][] = array($fr + 1, $fc + 1, $lr + 1, $lc + 1);
                        if ($lr - $fr > 0) {
                            $this->sheets[$this->sn]['cellsInfo'][$fr+1][$fc+1]['rowspan'] = $lr - $fr + 1;
                        }
                        if ($lc - $fc > 0) {
                            $this->sheets[$this->sn]['cellsInfo'][$fr+1][$fc+1]['colspan'] = $lc - $fc + 1;
                        }
                    }
                    //echo "Merged Cells $cellRanges $lr $fr $lc $fc\n";
                    break;


                case SPREADSHEET_EXCEL_READER_TYPE_CALCMODE:

                    $this->calcmode = $this->_readInt(2);
echo "Calcmode: ".dechex($this->calcmode)."\n";
return;

                    break;

                default:
                    // 0x8c8 ?

                    echo "WARNING: UNKNOWN RECORD TYPE\n";
                    echo 'File position: '. dechex(ftell($this->data))."\n";
                    echo "Default data:\n";
                    echo fread($this->data, $length);
                    echo "\n\n";
                    break;
            }

            $code   = $this->_readInt(2);
            $length = $this->_readInt(2);
echo "\n";
echo "*** NEW RECORD ***\n";
echo "File position: 0x". dechex(ftell($this->data))."\n";
echo "code:          0x". dechex($code)."\n";
echo "length:        $length which is 0x".dechex($length)."\n";
        }

        if (!isset($this->sheets[$sheet_index]['numRows'])) {
             $this->sheets[$sheet_index]['numRows'] = $this->sheets[$sheet_index]['maxrow'];
        }

        if (!isset($this->sheets[$sheet_index]['numCols'])) {
             $this->sheets[$sheet_index]['numCols'] = $this->sheets[$sheet_index]['maxcol'];
        }
    }

    /**
     * Parse a worksheet
     *
     * @access private
     * @param todo
     * @todo fix return codes
     */
    function _parsesheet($spos)
    {
        $cont = true;
        // read BOF
        $code = ord($this->data[$spos]) | ord($this->data[$spos+1])<<8;
        $length = ord($this->data[$spos+2]) | ord($this->data[$spos+3])<<8;

        $version = ord($this->data[$spos + 4]) | ord($this->data[$spos + 5])<<8;
        $substreamType = ord($this->data[$spos + 6]) | ord($this->data[$spos + 7])<<8;

        if (($version != SPREADSHEET_EXCEL_READER_BIFF8) && ($version != SPREADSHEET_EXCEL_READER_BIFF7)) {
            return -1;
        }

        if ($substreamType != SPREADSHEET_EXCEL_READER_WORKSHEET){
            return -2;
        }
        //echo "Start parse code=".base_convert($code,10,16)." version=".base_convert($version,10,16)." substreamType=".base_convert($substreamType,10,16).""."\n";
        $spos += $length + 4;
        //var_dump($this->formatRecords);
    //echo "code $code $length";
        while($cont) {
            //echo "mem= ".memory_get_usage()."\n";
//            $r = &$this->file->nextRecord();
            $lowcode = ord($this->data[$spos]);
            if ($lowcode == SPREADSHEET_EXCEL_READER_TYPE_EOF) break;
            $code = $lowcode | ord($this->data[$spos+1])<<8;
            $length = ord($this->data[$spos+2]) | ord($this->data[$spos+3])<<8;
            $spos += 4;
            $this->sheets[$this->sn]['maxrow'] = $this->_rowoffset - 1;
            $this->sheets[$this->sn]['maxcol'] = $this->_coloffset - 1;
            //echo "Code=".base_convert($code,10,16)." $code\n";
            unset($this->rectype);
            $this->multiplier = 1; // need for format with %
            switch ($code) {
                case SPREADSHEET_EXCEL_READER_TYPE_DIMENSIONS:
                    //echo 'Type_DIMENSION ';
                    if (!isset($this->numRows)) {
                        if (($length == 10) ||  ($version == SPREADSHEET_EXCEL_READER_BIFF7)){
                            $this->sheets[$this->sn]['numRows'] = ord($this->data[$spos+2]) | ord($this->data[$spos+3]) << 8;
                            $this->sheets[$this->sn]['numCols'] = ord($this->data[$spos+6]) | ord($this->data[$spos+7]) << 8;
                        } else {
                            $this->sheets[$this->sn]['numRows'] = ord($this->data[$spos+4]) | ord($this->data[$spos+5]) << 8;
                            $this->sheets[$this->sn]['numCols'] = ord($this->data[$spos+10]) | ord($this->data[$spos+11]) << 8;
                        }
                    }
                    //echo 'numRows '.$this->numRows.' '.$this->numCols."\n";
                    break;
                case SPREADSHEET_EXCEL_READER_TYPE_MERGEDCELLS:
                    $cellRanges = ord($this->data[$spos]) | ord($this->data[$spos+1])<<8;
                    for ($i = 0; $i < $cellRanges; $i++) {
                        $fr =  ord($this->data[$spos + 8*$i + 2]) | ord($this->data[$spos + 8*$i + 3])<<8;
                        $lr =  ord($this->data[$spos + 8*$i + 4]) | ord($this->data[$spos + 8*$i + 5])<<8;
                        $fc =  ord($this->data[$spos + 8*$i + 6]) | ord($this->data[$spos + 8*$i + 7])<<8;
                        $lc =  ord($this->data[$spos + 8*$i + 8]) | ord($this->data[$spos + 8*$i + 9])<<8;
                        //$this->sheets[$this->sn]['mergedCells'][] = array($fr + 1, $fc + 1, $lr + 1, $lc + 1);
                        if ($lr - $fr > 0) {
                            $this->sheets[$this->sn]['cellsInfo'][$fr+1][$fc+1]['rowspan'] = $lr - $fr + 1;
                        }
                        if ($lc - $fc > 0) {
                            $this->sheets[$this->sn]['cellsInfo'][$fr+1][$fc+1]['colspan'] = $lc - $fc + 1;
                        }
                    }
                    //echo "Merged Cells $cellRanges $lr $fr $lc $fc\n";
                    break;
                case SPREADSHEET_EXCEL_READER_TYPE_RK:
                case SPREADSHEET_EXCEL_READER_TYPE_RK2:
                    //echo 'SPREADSHEET_EXCEL_READER_TYPE_RK'."\n";
                    $row = ord($this->data[$spos]) | ord($this->data[$spos+1])<<8;
                    $column = ord($this->data[$spos+2]) | ord($this->data[$spos+3])<<8;
                    $rknum = $this->_GetInt4d($this->data, $spos + 6);
                    $numValue = $this->_GetIEEE754($rknum);
                    //echo $numValue." ";
                    if ($this->isDate($spos)) {
                        list($string, $raw) = $this->createDate($numValue);
                    }else{
                        $raw = $numValue;
                        if (isset($this->_columnsFormat[$column + 1])){
                                $this->curformat = $this->_columnsFormat[$column + 1];
                        }
                        $string = sprintf($this->curformat, $numValue * $this->multiplier);
                        //$this->addcell(RKRecord($r));
                    }
                    $this->addcell($row, $column, $string, $raw);
                    //echo "Type_RK $row $column $string $raw {$this->curformat}\n";
                    break;
                case SPREADSHEET_EXCEL_READER_TYPE_LABELSST:
                        $row        = ord($this->data[$spos]) | ord($this->data[$spos+1])<<8;
                        $column     = ord($this->data[$spos+2]) | ord($this->data[$spos+3])<<8;
                        $xfindex    = ord($this->data[$spos+4]) | ord($this->data[$spos+5])<<8;
                        $index  = $this->_GetInt4d($this->data, $spos + 6);
            //var_dump($this->sst);
                        $this->addcell($row, $column, $this->sst[$index]);
                        //echo "LabelSST $row $column $string\n";
                    break;
                case SPREADSHEET_EXCEL_READER_TYPE_MULRK:
                    $row        = ord($this->data[$spos]) | ord($this->data[$spos+1])<<8;
                    $colFirst   = ord($this->data[$spos+2]) | ord($this->data[$spos+3])<<8;
                    $colLast    = ord($this->data[$spos + $length - 2]) | ord($this->data[$spos + $length - 1])<<8;
                    $columns    = $colLast - $colFirst + 1;
                    $tmppos = $spos+4;
                    for ($i = 0; $i < $columns; $i++) {
                        $numValue = $this->_GetIEEE754($this->_GetInt4d($this->data, $tmppos + 2));
                        if ($this->isDate($tmppos-4)) {
                            list($string, $raw) = $this->createDate($numValue);
                        }else{
                            $raw = $numValue;
                            if (isset($this->_columnsFormat[$colFirst + $i + 1])){
                                        $this->curformat = $this->_columnsFormat[$colFirst + $i + 1];
                                }
                            $string = sprintf($this->curformat, $numValue * $this->multiplier);
                        }
                      //$rec['rknumbers'][$i]['xfindex'] = ord($rec['data'][$pos]) | ord($rec['data'][$pos+1]) << 8;
                      $tmppos += 6;
                      $this->addcell($row, $colFirst + $i, $string, $raw);
                      //echo "MULRK $row ".($colFirst + $i)." $string\n";
                    }
                     //MulRKRecord($r);
                    // Get the individual cell records from the multiple record
                     //$num = ;

                    break;
                case SPREADSHEET_EXCEL_READER_TYPE_NUMBER:
                    $row    = ord($this->data[$spos]) | ord($this->data[$spos+1])<<8;
                    $column = ord($this->data[$spos+2]) | ord($this->data[$spos+3])<<8;
                    $tmp = unpack("ddouble", substr($this->data, $spos + 6, 8)); // It machine machine dependent
                    if ($this->isDate($spos)) {
                        list($string, $raw) = $this->createDate($tmp['double']);
                     //   $this->addcell(DateRecord($r, 1));
                    }else{
                        //$raw = $tmp[''];
                        if (isset($this->_columnsFormat[$column + 1])){
                                $this->curformat = $this->_columnsFormat[$column + 1];
                        }
                        $raw = $this->createNumber($spos);
                        $string = sprintf($this->curformat, $raw * $this->multiplier);

                     //   $this->addcell(NumberRecord($r));
                    }
                    $this->addcell($row, $column, $string, $raw);
                    //echo "Number $row $column $string\n";
                    break;
                case SPREADSHEET_EXCEL_READER_TYPE_FORMULA:
                case SPREADSHEET_EXCEL_READER_TYPE_FORMULA2:
                    $row    = ord($this->data[$spos]) | ord($this->data[$spos+1])<<8;
                    $column = ord($this->data[$spos+2]) | ord($this->data[$spos+3])<<8;
                    if ((ord($this->data[$spos+6])==0) && (ord($this->data[$spos+12])==255) && (ord($this->data[$spos+13])==255)) {
                        //String formula. Result follows in a STRING record
                        //echo "FORMULA $row $column Formula with a string<br>\n";
                    } elseif ((ord($this->data[$spos+6])==1) && (ord($this->data[$spos+12])==255) && (ord($this->data[$spos+13])==255)) {
                        //Boolean formula. Result is in +2; 0=false,1=true
                    } elseif ((ord($this->data[$spos+6])==2) && (ord($this->data[$spos+12])==255) && (ord($this->data[$spos+13])==255)) {
                        //Error formula. Error code is in +2;
                    } elseif ((ord($this->data[$spos+6])==3) && (ord($this->data[$spos+12])==255) && (ord($this->data[$spos+13])==255)) {
                        //Formula result is a null string.
                    } else {
                        // result is a number, so first 14 bytes are just like a _NUMBER record
                        $tmp = unpack("ddouble", substr($this->data, $spos + 6, 8)); // It machine machine dependent
                        if ($this->isDate($spos)) {
                            list($string, $raw) = $this->createDate($tmp['double']);
                         //   $this->addcell(DateRecord($r, 1));
                        }else{
                            //$raw = $tmp[''];
                            if (isset($this->_columnsFormat[$column + 1])){
                                    $this->curformat = $this->_columnsFormat[$column + 1];
                            }
                            $raw = $this->createNumber($spos);
                            $string = sprintf($this->curformat, $raw * $this->multiplier);

                         //   $this->addcell(NumberRecord($r));
                        }
                        $this->addcell($row, $column, $string, $raw);
                        //echo "Number $row $column $string\n";
                    }
                    break;
                case SPREADSHEET_EXCEL_READER_TYPE_BOOLERR:
                    $row    = ord($this->data[$spos]) | ord($this->data[$spos+1])<<8;
                    $column = ord($this->data[$spos+2]) | ord($this->data[$spos+3])<<8;
                    $string = ord($this->data[$spos+6]);
                    $this->addcell($row, $column, $string);
                    //echo 'Type_BOOLERR '."\n";
                    break;
                case SPREADSHEET_EXCEL_READER_TYPE_ROW:
                case SPREADSHEET_EXCEL_READER_TYPE_DBCELL:
                case SPREADSHEET_EXCEL_READER_TYPE_MULBLANK:
                    break;
                case SPREADSHEET_EXCEL_READER_TYPE_LABEL:
                    $row    = ord($this->data[$spos]) | ord($this->data[$spos+1])<<8;
                    $column = ord($this->data[$spos+2]) | ord($this->data[$spos+3])<<8;
                    $this->addcell($row, $column, substr($this->data, $spos + 8, ord($this->data[$spos + 6]) | ord($this->data[$spos + 7])<<8));

                   // $this->addcell(LabelRecord($r));
                    break;

                case SPREADSHEET_EXCEL_READER_TYPE_EOF:
                    $cont = false;
                    break;
                default:
                    //echo ' unknown :'.base_convert($r['code'],10,16)."\n";
                    break;

            }
            $spos += $length;
        }

        if (!isset($this->sheets[$this->sn]['numRows']))
             $this->sheets[$this->sn]['numRows'] = $this->sheets[$this->sn]['maxrow'];
        if (!isset($this->sheets[$this->sn]['numCols']))
             $this->sheets[$this->sn]['numCols'] = $this->sheets[$this->sn]['maxcol'];

    }

    /**
     * Check whether the current record read is a date
     *
     * @param todo
     * @return boolean True if date, false otherwise
     */
    function isDate2($xfindex)
    {
        if ($this->formatRecords['xfrecords'][$xfindex]['type'] == 'date') {
            $this->curformat = $this->formatRecords['xfrecords'][$xfindex]['format'];
            $this->rectype = 'date';
            return true;
        } else {
            if ($this->formatRecords['xfrecords'][$xfindex]['type'] == 'number') {
                $this->curformat = $this->formatRecords['xfrecords'][$xfindex]['format'];
                $this->rectype = 'number';
                if (($xfindex == 0x9) || ($xfindex == 0xa)){
                    $this->multiplier = 100;
                }
            }else{
                $this->curformat = $this->_defaultFormat;
                $this->rectype = 'unknown';
            }
            return false;
        }
    }


    function isDate($spos)
    {
        //$xfindex = GetInt2d(, 4);
        $xfindex = ord($this->data[$spos+4]) | ord($this->data[$spos+5]) << 8;
        //echo 'check is date '.$xfindex.' '.$this->formatRecords['xfrecords'][$xfindex]['type']."\n";
        //var_dump($this->formatRecords['xfrecords'][$xfindex]);
        if ($this->formatRecords['xfrecords'][$xfindex]['type'] == 'date') {
            $this->curformat = $this->formatRecords['xfrecords'][$xfindex]['format'];
            $this->rectype = 'date';
            return true;
        } else {
            if ($this->formatRecords['xfrecords'][$xfindex]['type'] == 'number') {
                $this->curformat = $this->formatRecords['xfrecords'][$xfindex]['format'];
                $this->rectype = 'number';
                if (($xfindex == 0x9) || ($xfindex == 0xa)){
                    $this->multiplier = 100;
                }
            }else{
                $this->curformat = $this->_defaultFormat;
                $this->rectype = 'unknown';
            }
            return false;
        }
    }

    //}}}
    //{{{ createDate()

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
            $utcDays = $numValue - ($this->nineteenFour ? SPREADSHEET_EXCEL_READER_UTCOFFSETDAYS1904 : SPREADSHEET_EXCEL_READER_UTCOFFSETDAYS);
            $utcValue = round(($utcDays+1) * SPREADSHEET_EXCEL_READER_MSINADAY);
            $string = date ($this->curformat, $utcValue);
            $raw = $utcValue;
        } else {
            $raw = $numValue;
            $hours = floor($numValue * 24);
            $mins = floor($numValue * 24 * 60) - $hours * 60;
            $secs = floor($numValue * SPREADSHEET_EXCEL_READER_MSINADAY) - $hours * 60 * 60 - $mins * 60;
            $string = date ($this->curformat, mktime($hours, $mins, $secs));
        }

        return array($string, $raw);
    }



    /**
     * Create a number using IEEE 754 64-bit double precision
     *
     * @access private
     * @
     */
    function createNumber2($value_high, $value_low)
    {
        $sign         = $value_high >> 31;
        $exp          = ($value_high & 0x7ff00000) >> 20;
        $mantissa     = (0x100000 | ($value_high & 0x000fffff));
        $mantissalow1 = ($rknumlow & 0x80000000) >> 31;
        $mantissalow2 = ($rknumlow & 0x7fffffff);
        $value        = $mantissa / pow( 2 , (20- ($exp - 1023)));
        if ($mantissalow1 != 0) $value += 1 / pow (2 , (21 - ($exp - 1023)));
        $value += $mantissalow2 / pow (2 , (52 - ($exp - 1023)));
        if ($sign) {$value = -1 * $value;}

        return  $value;
    }

    function createNumber($spos)
    {
        $rknumhigh = $this->_GetInt4d($this->data, $spos + 10);
        $rknumlow = $this->_GetInt4d($this->data, $spos + 6);
        //for ($i=0; $i<8; $i++) { echo ord($this->data[$i+$spos+6]) . " "; } echo "<br>";
        $sign = ($rknumhigh & 0x80000000) >> 31;
        $exp =  ($rknumhigh & 0x7ff00000) >> 20;
        $mantissa = (0x100000 | ($rknumhigh & 0x000fffff));
        $mantissalow1 = ($rknumlow & 0x80000000) >> 31;
        $mantissalow2 = ($rknumlow & 0x7fffffff);
        $value = $mantissa / pow( 2 , (20- ($exp - 1023)));
        if ($mantissalow1 != 0) $value += 1 / pow (2 , (21 - ($exp - 1023)));
        $value += $mantissalow2 / pow (2 , (52 - ($exp - 1023)));
        //echo "Sign = $sign, Exp = $exp, mantissahighx = $mantissa, mantissalow1 = $mantissalow1, mantissalow2 = $mantissalow2<br>\n";
        if ($sign) {$value = -1 * $value;}
        return  $value;
    }

    function addcell($sheet_index, $row, $col, $string, $raw = '')
    {
        //echo "ADD cel $row-$col $string\n";
        $this->sheets[$sheet_index]['maxrow'] = max($this->sheets[$sheet_index]['maxrow'], $row + $this->_rowoffset);
        $this->sheets[$sheet_index]['maxcol'] = max($this->sheets[$sheet_index]['maxcol'], $col + $this->_coloffset);
        $this->sheets[$sheet_index]['cells'][$row + $this->_rowoffset][$col + $this->_coloffset] = $string;
        if ($raw)
            $this->sheets[$sheet_index]['cellsInfo'][$row + $this->_rowoffset][$col + $this->_coloffset]['raw'] = $raw;
        if (isset($this->rectype))
            $this->sheets[$sheet_index]['cellsInfo'][$row + $this->_rowoffset][$col + $this->_coloffset]['type'] = $this->rectype;

    }


    function _GetIEEE754($rknum)
    {
        if (($rknum & 0x02) != 0) {
                $value = $rknum >> 2;
        } else {
echo "old way float\n";
//mmp
// first comment out the previously existing 7 lines of code here
//                $tmp = unpack("d", pack("VV", 0, ($rknum & 0xfffffffc)));
//                //$value = $tmp[''];
//                if (array_key_exists(1, $tmp)) {
//                    $value = $tmp[1];
//                } else {
//                    $value = $tmp[''];
//                }
// I got my info on IEEE754 encoding from
// http://research.microsoft.com/~hollasch/cgindex/coding/ieeefloat.html
// The RK format calls for using only the most significant 30 bits of the
// 64 bit floating point value. The other 34 bits are assumed to be 0
// So, we use the upper 30 bits of $rknum as follows...
         $sign = ($rknum & 0x80000000) >> 31;
        $exp = ($rknum & 0x7ff00000) >> 20;
        $mantissa = (0x100000 | ($rknum & 0x000ffffc));
        $value = $mantissa / pow( 2 , (20- ($exp - 1023)));
        if ($sign) {$value = -1 * $value;}
//end of changes by mmp

        }

        if (($rknum & 0x01) != 0) {
            $value /= 100;
        }
        return $value;
    }

    function _encodeUTF16($string)
    {
        $result = $string;
        if ($this->_defaultEncoding){
            switch ($this->_encoderFunction){
                case 'iconv' :     $result = iconv('UTF-16LE', $this->_defaultEncoding, $string);
                                break;
                case 'mb_convert_encoding' :     $result = mb_convert_encoding($string, $this->_defaultEncoding, 'UTF-16LE' );
                                break;
            }
        }
        return $result;
    }

    function _GetInt4d($data, $pos)
    {
        $value = ord($data[$pos]) | (ord($data[$pos+1]) << 8) | (ord($data[$pos+2]) << 16) | (ord($data[$pos+3]) << 24);
        if ($value>=4294967294)
        {
            $value=-2;
        }
        return $value;
    }

}

/*
 * Local variables:
 * tab-width: 4
 * c-basic-offset: 4
 * c-hanging-comment-ender-p: nil
 * End:
 */

?>
