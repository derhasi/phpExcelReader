<?php

require_once 'Spreadsheet/Excel/Reader.php';

abstract class Spreadsheet_Excel_Reader_BIFFParser
{
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
        $is_float = ($this->_readInt(2) !== Spreadsheet_Excel_Reader::RESULT_NOTFLOAT);
        fseek($this->_stream, $pos);

        if ($is_float) {

            $result = $this->_readDouble();

        } else {

            $type = $this->_readInt(1);

            switch ($type) {
                
                case Spreadsheet_Excel_Reader::RESULT_STRING:
                case Spreadsheet_Excel_Reader::RESULT_EMPTY:
                default:
                    fseek($this->_stream, 7, SEEK_CUR);
                    $result = null;
                    break;

                case Spreadsheet_Excel_Reader::RESULT_BOOL:
                case Spreadsheet_Excel_Reader::RESULT_ERROR:
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
                      pow( 2, $exponent - Spreadsheet_Excel_Reader::EXPONENT_BIAS) *
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