<?php

namespace MyProject\Samples\CompleteFramework;

use ms\Typer;

use ms2\System\netString;
use ms2\System\netInt32;
use ms2\System\netDateTime;

/**
 * .Net Date Time example
 */
class Example1 {

  public static function Example1() {
    // Make sure the Runtime is Initialized.
    RuntimeManager::Instance()->InitializeRuntime();

    $today = \ms2\System\netDateTime::Today();
    $day_of_week = $today->DayOfWeek()->ToString()->Val();
    $month = $today->Month()->Val();

    echo "<br>Today is: $day_of_week</br>";
    echo "<br>Month: $month</br>";

    // Open a directory with process explorer
    \ms2\System\Diagnostics\netProcess::_Start("explorer.exe", "c:\\");
  }

 

}
