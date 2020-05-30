<?php

namespace Indie\Utility;

/**
 * Debug Class
 */
class Debug
{

  public static function log($data){

    if (is_array($data)) {
      var_dump($data).PHP_EOL;
      return;
    }
    echo $data.PHP_EOL;

  }
}
