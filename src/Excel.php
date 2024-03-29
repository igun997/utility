<?php

namespace Igun997\Utility;
require_once "Bootstrap.php";

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Exception;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;


/**
 * Excel Class
 */
class Excel
{
    const UP = 0;
    const DOWN = 1;
    public int $row = -1;
    public int $col = -1;
    public int $sheet = 0;
    public array $options;
    public $content = null;
    public $label = null;
    public $instance = NULL;
    public $property = null;
    public $is_download;

    public ?string $type = null;

    function __construct(string $data = NULL, array $options = [])
    {
        if ($data) {

            $spreadsheet = IOFactory::load($data);
            $this->options = $options;
            $this->content = $spreadsheet;
        } else {

            $this->instance = new Spreadsheet();
        }

        return $this;
    }

    //Export Array to Excel
    public function properties(array $property)
    {
        if ($this->instance == NULL) {
            throw new \Exception('Instance NULL');
        }
        $this->property = $property;
        $this->instance->getProperties()
            ->setCreator($property["creator"])
            ->setTitle($property["title"])
            ->setSubject($property["subject"])
            ->setDescription($property["description"]);
        return $this;
    }

    /**
     * @param $func
     * @param array $cb
     * @return $this
     */
    public function operation($func, array &$cb)
    {

        $cb = $func($this->content);

        return $this;
    }

    /**
     * @param int $sheet
     * @return $this
     */
    public function setSheet(int $sheet = 0)
    {
        $this->sheet = $sheet;
        return $this;
    }

    /**
     * @param array $data
     * @param String $title
     * @param bool $is_download
     * @param array $hidden_col
     * @return $this
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function write(array $data, string $title, bool $is_download = FALSE, array $hidden_col = [])
    {

        if ($this->instance == NULL) {
            throw new \Exception('Instance NULL');
        }

        $this->is_download = $is_download;
        //Check if want writer existing
        if ($this->content === null){
            $this->instance->createSheet();
            $init = $this->instance->setActiveSheetIndex($this->sheet);
        }else{
            $this->instance = $this->content;
        }

        $this->instance->getActiveSheet()->setTitle($title);

        if (count($hidden_col) > 0) {
            foreach ($hidden_col as $index => $item) {
                $this->instance->getActiveSheet()->getColumnDimension($item)->setVisible(false);
            }
        }

        foreach ($data as $key => $value) {

            $init->setCellValue($key, $value);
        }

        return $this;
    }

    //Import Excel to Array

    /**
     * @param int $row
     * @param int $col
     * @return $this
     */
    public function setLabel(int $row, int $col = -1)
    {
        $this->row = $row;
        $this->col = $col;
        return $this;
    }


    /**
     * @param String $type
     * @return $this
     */
    public function type(string $type)
    {

        $this->type = $type;
        if ($this->type == "raw") {


        } elseif ($this->type == "json") {


        } elseif ($this->type == "xml") {


        } elseif ($this->type == "array") {

            $this->content = $this->arrayFormat();
        } else {
            $this->content = false;
        }
        return $this;
    }

    /**
     * @return array
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    private function arrayFormat()
    {
        $init = $this->content->setActiveSheetIndex($this->sheet);
        $sheetData = $this->content->getActiveSheet()->toArray(false, true, true, true);
//    Reformat To Numeric KEY
        $this->_replacer($sheetData);
//    Remove Fucking Empty Array
        $this->_replacer($sheetData);
        return $sheetData;

    }

    /**
     * @param $data
     */
    private function _replacer(&$data)
    {
        $i = 0;
        if (is_array($data)) {
            foreach ($data as $key => &$value) {

                if (is_array($value)) {
                    if (empty($value)) {
                        unset($data[$key]);
                        continue;
                    }
                }

                if ($value === NULL) {
                    unset($data[$key]);
                    continue;
                }

                if ($value === FALSE) {
                    unset($data[$key]);
                    continue;
                }

                if (!is_numeric($key)) {

                    unset($data[$key]);
                    $data[$i] = $value;
                } else {
                    $this->_replacer($value);
                }

                $i++;
            }
        }

    }

    /**
     * @param array $options
     * @return $this
     * @throws \Exception
     */
    public function reformat(array $options = [])
    {

        $use_col = true;

        if (empty($options)) {
            $options = $this->options;
        }

        if (empty($options)) {

            throw new \Exception('Reformat Need Options Params on `reformat()` or on initial construct');
            exit();
        }


        if ($this->row === -1) {

            throw new \Exception('setLabel must be contruct first');
            exit();
        }


        if ($this->col === -1) {
            $use_col = false;
        }

        if (!$use_col) {
            if (!isset($this->content[$this->row])) {

                throw new \Exception('label not found on index ' . $this->row);
                exit();

            }

            $this->label = $this->content[$this->row];
            $startFrom = [$this->row + 1];

        } else {

            if (!isset($this->content[$this->col][$this->row])) {
                throw new \Exception('label not found on index ' . $this->row . ' - ' . $this->col);
                exit();

            }

            $this->label = $this->content[$this->col][$this->row];
            $startFrom = [($this->col), $this->row + 1];

        }

        $in_array = [];

        foreach ($this->label as $key => $value) {
            $in_array[] = strtolower($value);
        }


        if (count($startFrom) == 1) {

            // $startFormat = $this->content[$startFrom[0]];
            unset($this->content[$this->row]);
        } else {

            // $startFormat = $this->content[$startFrom[0]][$startFrom[1]];
            unset($this->content[$startFrom[0]][$startFrom[1]]);

            $change = $this->content;
            $this->content = $change[$startFrom[0]];

            unset($change);


        }

        $temp = [];

        foreach ($this->content as $ks => $vs) {

            if ($ks == $this->row) {
                continue;
            }

            $build = [];

            $startFormat = $vs;

            foreach ($options as $k => $v) {
                if (!is_array($v)) {

                    if (in_array(strtolower($v), $in_array)) {
                        if (isset($startFormat[$this->_index($in_array, strtolower($v))])) {
                            $build[$k] = $startFormat[$this->_index($in_array, strtolower($v))];
                        } else {
                            $build[$k] = NULL;
                        }
                    } else {
                        $build[$k] = NULL;
                    }
                } else {

                    foreach ($v as $key => &$value) {
                        if (in_array(strtolower($value), $in_array)) {
                            $build[$k][$key] = $startFormat[$this->_index($in_array, strtolower($value))];
                        } else {
                            $build[$k][$key] = NULL;
                        }
                    }
                }

            }

            $temp[] = $build;
        }

        foreach ($temp as $index => $item) {
            if (count($item) == 0) {
                unset($temp[$index]);
            }
        }
        $this->content = $temp;

        unset($temp);

        return $this;

    }

    /**
     * @param $data
     * @param $val
     * @return int|string
     * @throws \Exception
     */
    private function _index($data, $val)
    {
        foreach ($data as $key => $value) {
            if ($val == $value) {
                return $key;
            }
        }
        throw new \Exception('index data notfound');
    }

    /**
     * @return int
     */
    public function num_rows()
    {

        return count($this->content);
    }


    /**
     * @param null $costum_path
     * @return Spreadsheet|void
     * @throws Exception
     */
    public function output($costum_path = NULL)
    {

        if ($this->instance != NULL) {

            $writer = new Xlsx($this->instance);
            $name = strtolower(trim(str_replace(" ", "_", $this->property["title"]))) . '.xlsx';
            if ($costum_path) {
                $name = $costum_path . strtolower(trim(str_replace(" ", "_", $this->property["title"]))) . '.xlsx';
            }
            if ($this->is_download) {

                header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
                header('Content-Disposition: attachment;filename="' . $name . '"');
                header('Cache-Control: max-age=0');

                $writer->save('php://output');

            } else {

                return $writer->save($name);
            }


        }

        return $this->content;

    }

}
