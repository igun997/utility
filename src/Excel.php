<?php
namespace Indie\Utility;
require_once "Bootstrap.php";
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;


/**
 * Excel Class
 */
class Excel
{
  CONST UP = 0;
  CONST DOWN = 1;
  public int $row = -1;
  public int $col = -1;
  public int $sheet = 0;
  public array $options;
  public $content;
  public $label;
  public $instance = NULL;
  public $property;
  public $is_download;

  public ?string $type = null;

  function __construct(String $data=NULL,Array $options = [])
  {
    if ($data) {

      $spreadsheet = IOFactory::load($data);
      $this->options = $options;
      $this->content = $spreadsheet;
    }else {

        $this->instance = new Spreadsheet();
    }

    return $this;
  }

  //Export Array to Excel
  public function properties(Array $property)
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

  public function operation($func,array &$cb)
  {

        $cb = $func($this->content);

        return $this;
  }

  public function setSheet(int $sheet = 0){
      $this->sheet = $sheet;
      return $this;
  }

  public function write(Array $data,String $title,bool $is_download = FALSE)
  {

    if ($this->instance == NULL) {
      throw new \Exception('Instance NULL');
    }

    $this->is_download = $is_download;
    $this->instance->createSheet();
    $init = $this->instance->setActiveSheetIndex($this->sheet);

    $this->instance->getActiveSheet()->setTitle($title);


    foreach ($data as $key => $value) {

      $init->setCellValue($key,$value);
    }

    return $this;
  }

  //Import Excel to Array
  public function setLabel(int $row,int $col = -1)
  {
    $this->row = $row;
    $this->col = $col;
    return $this;
  }



  public function type(String $type)
  {

    $this->type = $type;
    if ($this->type == "raw") {


    }elseif ($this->type == "json") {


    }elseif ($this->type == "xml") {


    }elseif ($this->type == "array") {

      $this->content =  $this->arrayFormat();
    }else {
      $this->content =  false;
    }
    return $this;
  }

  private function arrayFormat()
  {

    $sheetData = $this->content->getActiveSheet()->toArray(false, true, true, true);
//    Reformat To Numeric KEY
    $this->_replacer($sheetData);
//    Remove Fucking Empty Array
    $this->_replacer($sheetData);
    return $sheetData;

  }

  private function _replacer(&$data)
  {
    $i = 0;
    if (is_array($data)) {
      foreach ($data as $key => &$value) {

          if(empty($value)){
              unset($data[$key]);
              continue;
          }

          if(!$value){
              unset($data[$key]);
              continue;
          }

          if (!is_numeric($key)) {
              unset($data[$key]);
              $data[$i] = $value;
          }else {
              $this->_replacer($value);
          }

          $i++;
      }
    }

  }

  private function _index($data,$val)
  {
    foreach ($data as $key => $value) {
      if ($val == $value) {
        return $key;
      }
    }
    throw new \Exception('index data notfound');
  }

  public function reformat(Array $options = [])
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

        throw new \Exception('label not found on index '.$this->row);
        exit();

      }

      $this->label = $this->content[$this->row];
      $startFrom = [$this->row+1];

    }else {

      if (!isset($this->content[$this->col][$this->row])) {
        throw new \Exception('label not found on index '.$this->row.' - '.$this->col);
        exit();

      }

      $this->label = $this->content[$this->col][$this->row];
      $startFrom = [($this->col),$this->row+1];

    }

    $in_array = [];

    foreach ($this->label as $key => $value) {
      $in_array[] = strtolower($value);
    }





    if (count($startFrom) == 1) {

       // $startFormat = $this->content[$startFrom[0]];
       unset($this->content[$this->row]);
    }else {

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

          if (in_array(strtolower($v),$in_array)) {
              if(isset($startFormat[$this->_index($in_array,strtolower($v))])){
                  $build[$k] = $startFormat[$this->_index($in_array,strtolower($v))];
              }else{
                  $build[$k] = NULL;
              }
          }else {
            $build[$k] = NULL;
          }
        }else {

          foreach ($v as $key => &$value) {
            if (in_array(strtolower($value),$in_array)) {
              $build[$k][$key] = $startFormat[$this->_index($in_array,strtolower($value))];
            }else {
              $build[$k][$key] = NULL;
            }
          }
        }

      }

      $temp[] = $build;
    }

    foreach ($temp as $index => $item) {
          if (count($item) == 0){
              unset($temp[$index]);
          }
    }
    $this->content = $temp;

    unset($temp);

    return $this;

  }

  public function num_rows()
  {

    return count($this->content);
  }


  public function output()
  {

    if ($this->instance != NULL) {

      $writer = new Xlsx($this->instance);
      $name = strtolower(trim(str_replace(" ","_",$this->property["title"]))).'.xlsx';
      if ($this->is_download) {

         header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
         header('Content-Disposition: attachment;filename="'.$name.'"');
         header('Cache-Control: max-age=0');

         $writer->save('php://output');

      }else {

        return $writer->save($name);
      }


    }

    return $this->content;

  }

}
