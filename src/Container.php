<?php
/**
 * Created by PhpStorm.
 * User: buphmin
 * Date: 12/6/15
 * Time: 4:05 PM
 */

namespace PHPExcelDataFiles;


class Container
{

    /**
     * @var \PHPExcel
     */
    private $workbook;

    /**
     * @var \PHPExcel_Worksheet
     */
    private $worksheet;

    /**
     * @var array
     */
    private $fileInfo = array();

    /**
     * @var string
     */
    private $primaryKey;

    /**
     * @var string = the file extension
     */
    private $fileExtension;

    /**
     * @var string = the full file path
     */
    private $filePath;

    /**
     * Container constructor.
     * Takes everything required in order to make a container
     * @param $filePath
     * @param $primaryKey
     * @param array $fileInfo
     */
    public function __construct($filePath, $primaryKey, $fileInfo = array())
    {
        if (empty($fileInfo)) // if no explicit file info set try to determine
        {
            $this->analyzeFile($filePath);
        }
        else
        {
            $this->setFileInfo($fileInfo);
        }

        $this->primaryKey = $primaryKey;
        $this->filePath = $filePath;
        $this->createWorkBook($filePath);
    }

    /**
     * Creates an instance of PHPExcel based of a file path.
     * @param $filePath
     * @return \PHPExcel
     */
    public function createWorkBook($filePath)
    {
        $info = pathinfo($filePath);
        $this->fileExtension = $info['extension'];

        switch($this->fileExtension)
        {
            // if text or csv use PHPExcel_Reader_CSV else use loader
            //due to inability of the PHPExcel IOfactory being
            //unable to handle csv/txt
            case 'txt':

                $objReader = new \PHPExcel_Reader_CSV();
                $objReader->setInputEncoding($this->fileInfo['encoding']);
                $objReader->setDelimiter($this->fileInfo['delimiter']);
                $objReader->setEnclosure($this->fileInfo['enclosure']);
                $objReader->setSheetIndex(0);
                $this->workbook = $objReader->load($filePath);
                break;
            case 'csv':
                $objReader = new \PHPExcel_Reader_CSV();
                $objReader->setInputEncoding($this->fileInfo['encoding']);
                $objReader->setDelimiter($this->fileInfo['delimiter']);
                $objReader->setEnclosure($this->fileInfo['enclosure']);
                $objReader->setSheetIndex(0);
                $this->workbook = $objReader->load($filePath);
                break;
            default:
                $this->workbook = \PHPExcel_IOFactory::load($filePath);
                break;

        }

        return $this->workbook; // return created workbook
    }

    /**
     * @return string
     */
    public function getFilePath()
    {
        return $this->filePath;
    }

    /**
     * Expects an array with the delimiter, enclosure, and encoding of
     * the file.
     * EX:
     * array(
     *  'delimiter' => ',',
     *  'enclosure' => '"',
     *  'encoding' => UTF8,
     *  'lineEnding' => "\r\n"
     *       );
     * @param array $fileInfo
     * @return $this
     */
    public function setFileInfo(array $fileInfo)
    {
        $this->fileInfo = $fileInfo;

        return $this;
    }


    /**
     * @return \PHPExcel
     */
    public function getWorkBook()
    {
        return $this->workbook;
    }

    /**
     * Gets a sheet by a loose name match unless
     * exactMatch is true. Returns a PHPExcel_Worksheet object
     * if a match is found, false if no match.
     * @param $name
     * @param bool|false $exactMatch
     * @return bool|\PHPExcel_Worksheet
     */
    private function getSheetByName($name, $exactMatch = false)
    {
        if($exactMatch === false)
        {
            $name = strtolower($name);

            // determine a case insensitive match
            foreach ($this->workbook->getSheetNames() as $sheet)
            {
                if (strtolower($sheet) === $name)
                {
                    $use = $sheet;
                    break;
                }
            }

            if (!isset($use))
            {
                return false;
            }

            $this->worksheet = $this->workbook->getSheetByName($use);
        }
        else
        {
            if($this->workbook->sheetNameExists($name))
            {
                $this->worksheet = $this->workbook->getSheetByName($name);
            }
            else
            {
                return false;
            }
        }

        return $this->worksheet;
    }

    /**
     * Returns a PHPExcel_Worksheet object if it the sheet index
     * exists, false otherwise
     * @param $index
     * @return bool|\PHPExcel_Worksheet
     */
    private function getSheetByIndex($index)
    {
        try
        {
            $this->worksheet = $this->workbook->getSheet($index);
        }
        catch (\PHPExcel_Exception $e)
        {
            //sheet out of bounds
            return false;
        }

        return $this->worksheet;
    }

    /**
     * Gets a PHPExcel_Worksheet object either by name
     * or by sheet index. Returns false if not found.
     * @param $sheet
     * @return bool|\PHPExcel_Worksheet
     */
    public function getSheet($sheet)
    {
        $this->worksheetName = $sheet;

        if(is_numeric($sheet))
        {
            return $this->getSheetByIndex($sheet);
        }
        else
        {
            return $this->getSheetByName($sheet);
        }
    }


    /**
     * Returns the last column of the worksheet or
     * returns false if unable to determine a max column
     * @param $sheet = name or index of the worksheet
     * @param int $maxColumn
     * @param int $row = the row to check where the max column is
     * @return bool|int
     */
    public function getLastColumn($sheet, $maxColumn = 50,  $row = 1)
    {
        $step = false;

        for($i = 0; $i < $maxColumn; $i++)
        {
            $value = $this->getSheet($sheet)->getCellByColumnAndRow($i, $row)->getValue();


            if($value == "")
            {
                $step = true;
            }

            if($step == true)
            {
                if($i == 0)
                    return false; //no values populated

                return $i - 1; //the last column is the last not blank
            }
        }

        return false; //last column is past maxColumn
    }


    /**
     * Gets the last row with data at a certain column index.
     * I will refine this method at some point, but for now its better than
     * brute force finding the last row.
     * @param $sheet
     * @param int $start
     * @param int $times
     * @param int $col
     * @return int
     */
    public function getLastRow($sheet, $start = 100000, $times = 0, $col = 0)
    {
        $times += 1;
        $minus = false;
        $value = $this->getSheet($sheet)->getCellByColumnAndRow($col, $start)->getValue();

        if($times < 100)
        {
            if(empty($value))
            {
                $fin = $this->getLastRow($sheet, round($start/2), $times, $col);
            }
            else
            {
                $fin = $this->getLastRow($sheet, round($start * 1.5), $times, $col);
            }
        }
        else
        {
            if(!isset($fin))
            {
                while(true)
                {
                    if($minus === TRUE)
                    {
                        $value = $this->getSheet($sheet)->getCellByColumnAndRow($col, $start)->getValue();

                        if(!empty($value))
                        {
                            $fin = $start;
                            $this->lastrow = $fin;
                            return $fin;
                        }

                        $start -= 1;
                    }
                    else
                    {
                        $value = $this->getSheet($sheet)->getCellByColumnAndRow($col, $start)->getValue();
                        if(empty($value))
                        {
                            $minus = true;
                            continue;
                        }

                        $start += 1;
                    }
                }
            }
            else
            {

                return $fin;
            }
        }


        return (int) $fin;

    }

    /**
     * Converts a column index to a column letter
     * @param $val
     * @return string
     */
    public function colNumberToLetter($val)
    {
        return $charCol = chr(64 + $val);
    }

    /**
     * Attempts to save the workbook or file.
     * It will attempt to save in the original file format
     * but if not it will default to xlsx
     * @throws \Exception
     * @throws \PHPExcel_Reader_Exception
     */
    public function saveBook()
    {
        switch($this->fileExtension)
        {
            case 'xlsx':
                $save = \PHPExcel_IOFactory::createWriter($this->workbook, 'Excel2007');
                break;
            case 'xlsm':
                $save = \PHPExcel_IOFactory::createWriter($this->workbook, 'Excel2007');
                break;
            case 'xltx':
                $save = \PHPExcel_IOFactory::createWriter($this->workbook, 'Excel2007');
                break;
            case 'xltm':
                $save = \PHPExcel_IOFactory::createWriter($this->workbook, 'Excel2007');
                break;
            case 'xls':
                $save = \PHPExcel_IOFactory::createWriter($this->workbook, 'Excel5');
                break;
            case 'xlt':
                $save = \PHPExcel_IOFactory::createWriter($this->workbook, 'Excel5');
                break;
            case 'ods':
                $save = \PHPExcel_IOFactory::createWriter($this->workbook, 'Excel2007');
                break;
            case 'ots':
                $save = \PHPExcel_IOFactory::createWriter($this->workbook, 'Excel2007');
                break;
            case 'xml':
                $save = \PHPExcel_IOFactory::createWriter($this->workbook, 'Excel2007');
                break;
            case 'txt':
            case 'csv':
                $save = new \PHPExcel_Writer_CSV($this->workbook);
                $save->setDelimiter($this->fileInfo['delimiter']);
                $save->setEnclosure($this->fileInfo['enclosure']);
                $save->setLineEnding($this->fileInfo['lineEnding']);
                $save->setSheetIndex(0);
                break;
        }


        if(isset($save))
        {
            if($this->fileExtension == 'ods' || $this->fileExtension == 'odt')
            {
                $this->filePath = str_replace($this->fileExtension, 'xlsx', $this->filePath);
            }

            $save->save($this->getFilePath());

            if($this->fileExtension == 'txt' || $this->fileExtension == 'csv')
            {
                // clean up extra blank rows made by PHPExcel
                $data = file_get_contents($this->filePath);
                $data = rtrim($data, "\t\r\n");
                $data .= $this->fileInfo['lineEnding'];
                file_put_contents($this->filePath, $data);
            }
        }
        else
        {
            throw new \Exception("Unable to create PHPExcel Writer");
        }
    }

    /**
     * Scans the file in the path to determine the file info.
     * Sets the delimiter, encoding, enclosure, and line ending
     * if it properly detects.
     * @param $filePath
     */
    public function analyzeFile($filePath)
    {
        $file = new \SplFileObject($filePath);
        $checkLines = 5;
        $delimiters = array(
            ',',
            "\t",
            ';',
            '|',
            ':'
        );
        $results = array();
        $i = 0;
        $single = 0;
        $double = 0;
        $lines = "";
        $line = '';

        while(!$file->eof() || $i <= $checkLines)
        {
            $line = $file->fgets();
            $lines .= $line;
            foreach ($delimiters as $delimiter)
            {
                $regExp = '/['.$delimiter.']/';
                $fields = preg_split($regExp, $line);

                if(count($fields) > 1)
                {
                    if(!empty($results[$delimiter]))
                    {
                        $results[$delimiter]++;
                    }
                    else
                    {
                        $results[$delimiter] = 1;
                    }
                }
            }

            $single += substr_count($line, "'");
            $double += substr_count($line, '"');
            $i++;
        }

        $this->fileInfo['lineEnding'] = $this->detectEol($lines);

        $encoding = mb_detect_encoding($lines);
        $results = array_keys($results, max($results));
        $this->fileInfo['delimiter'] = $results[0];
        $this->fileInfo['encoding'] = $encoding;
        $cellCount = count(explode($this->fileInfo['delimiter'], $line));

        if($cellCount * 2 > $single || $cellCount * 2 > $double)
        {

            if ($single > $double)
            {
                $this->fileInfo['enclosure'] = $single;
            }
            else
            {
                $this->fileInfo['enclosure'] = $double;
            }
        }
        else
        {
            $this->fileInfo['enclosure'] = '"';
        }

    }

    /**
     * Builds an associative array (aka dictionary) of the
     * header to column index. EX:
     * array("id" => 0, "title" => 1)
     * @param $sheet
     * @return array|bool
     */
    public function getHeaderBySheet($sheet)
    {
        $i = 0;
        $header = array();
        $lastColumn = $this->getLastColumn($sheet);

        while(true)
        {
            $value = $this->getSheet($sheet)->getCellByColumnAndRow($i, 1)->getValue();

            if(empty($value) || $i > $lastColumn)
            {
                if($i = 0)
                {
                    return false;
                }
                else
                {
                    return $header;
                }
            }
            else
            {
                $header[$value] = $i;
                $i++;
            }
        }

        return false;
    }

    /**
     * Builds a associative array (aka dictionary) of the primary keys
     * of the data file to the row number. EX:
     * array("28374932" => 34, "28394757" => 35)
     * @param $sheet
     * @return array|bool
     */
    public function getPrimaryKeyList($sheet)
    {

        $head = $this->getHeaderBySheet($sheet);
        $lastRow = $this->getLastRow($sheet);
        $idList = array();
        // add 1 because column index starts at 0
        $col = $this->colNumberToLetter($head['id']+1);
        $temp = $this->getSheet($sheet)->rangeToArray("$col"."2:$col$lastRow");
        $idRow = 2;

        foreach($temp as $arr)
        {
            $idList[$arr[0]] = $idRow;
            $idRow++;
        }

        if(!empty($idList))
        {
            return $idList;
        }
        else
        {
            return false;
        }

    }

    /**
     * Detects the end-of-line character of a string.
     * @param string $str The string to check.
     * @param string $default Default EOL (if not detected).
     * @return string The detected EOL, or default one.
     */
    function detectEol($str, $default=''){
        static $eols = array(
            "\r\n", // [UNICODE] CR+LF: CR (U+000D) followed by LF (U+000A)
            "\n",     // [UNICODE] LF: Line Feed, U+000A
            "\r",     // [UNICODE] VT: Vertical Tab, U+000B
        );
        $cur_cnt = 0;
        $cur_eol = $default;
        foreach($eols as $eol)
        {
            if(($count = substr_count($str, $eol)) > $cur_cnt)
            {
                $cur_cnt = $count;
                $cur_eol = $eol;
            }
        }
        return $cur_eol;
    }


}