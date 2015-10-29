<?php

namespace MyProject\Samples\Spreadsheetlight;

use ms\Typer;

use ms\SpreadsheetLight\netSLConvert;
use ms\SpreadsheetLight\netSLStyle;
use ms\System\netString;
use ms\System\netInt32;
use ms\System\netDateTime;

/**
 * This example is a port of the official
 * 
 * http://spreadsheetlight.com/downloads/samplecode/HelloWorld.cs
 * 
 * Hello World example
 */
class Example1 {

  public static function Example1() {

    // Make sure the Runtime is Initialized.
    RuntimeManager::Instance()->InitializeRuntime();

    $sl = \ms\SpreadsheetLight\netSLDocument::SLDocument_Constructor();

    $sl->SetCellValue("A1", TRUE);

    for ($i = 1; $i <= 20; $i++) {
      $sl->SetCellValue(2, $i, $i);
    }

    $sl->SetCellValue("B3", Typer::cDouble(3.14159));

    // set the value of PI at row 4, column 2 (or "B4") in string form.
    // use this when you already have numeric data in string form and don't
    // want to parse it to a double or float variable type
    // and then set it as a value.
    // Note that "3,14159" is invalid. Excel (or Open XML) stores numerals in
    // invariant culture mode. Frankly, even "1,234,567.89" is invalid because
    // of the comma. If you can assign it in code, then it's fine, like so:
    // double fTemp = 1234567.89;
    $sl->SetCellValueNumeric(4, 2, "3.14159");

    // normal string data
    $sl->SetCellValue("C6", "This is at C6!");

    // typical XML-invalid characters are taken care of,
    // in particular the & and < and >
    $sl->SetCellValue("I6", "Dinner & Dance costs < $10");

    // this sets a cell formula
    // Note that if you want to set a string that starts with the equal sign,
    // but is not a formula, prepend a single quote.
    // For example, "'==" will display 2 equal signs
    $sl->SetCellValue(7, 3, "=SUM(A2:T2)");

    // if you need cell references and cell ranges *really* badly, consider the SLConvert class.
    $sl->SetCellValue(netSLConvert::_ToCellReference(7, 4), netString::_Format("=SUM({0})", netSLConvert::_ToCellRange(2, 1, 2, 20)));

    // dates need the format code to be displayed as the typical date.
    // Otherwise it just looks like a floating point number.
    $sl->SetCellValue("C8",  netDateTime::DateTime_Constructor(3141, 5, 9));

    $style = $sl->CreateStyle();
    $style->FormatCode("d-mmm-yyyy");
    $sl->SetCellStyle("C8", $style);

    $sl->SetCellValue(8, 6, "I predict this to be a significant date. Why, I do not know...");

    $sl->SetCellValue(9, 4, 456.123789);
    //// we don't have to create a new SLStyle because
    //// we only used the FormatCode property
    $style->FormatCode = "0.000%";
    $sl->SetCellStyle(9, 4, $style);

    $sl->SetCellValue(9, 6, "Perhaps a phenomenal growth in something?");

    $sl->SaveAs("d:\\HelloWorld.xlsx");

    echo '<br>Check out your first Excel file at d:\HelloWorld.xlsx</br>';
  }

  public static function Example2() {

    // Make sure the Runtime is Initialized.
    RuntimeManager::Instance()->InitializeRuntime();

    $css = <<<CSS

body {
    color: purple;
    background-color: #d8da3d 
}

CSS;

    $minifier = \ms\Microsoft\Ajax\Utilities\netMinifier::Minifier_Constructor();
    $settings = \ms\Microsoft\Ajax\Utilities\netCodeSettings::CodeSettings_Constructor();
    $csssettings = \ms\Microsoft\Ajax\Utilities\netCssSettings::CssSettings_Constructor();
    $settings->OutputMode(\ms\Microsoft\Ajax\Utilities\netOutputMode::SingleLine());
    $settings->PreserveFunctionNames(FALSE);
    $settings->QuoteObjectLiteralProperties(TRUE);

    $result = $minifier->MinifyStyleSheet($css, $csssettings, $settings)->Val();
  
    echo $result;

  }

  public static function ProceduralExample() {
  
    // Instantiate and initialize the Model
    $runtime = new \NetPhp\Core\NetPhpRuntime('COM', 'netutilities.NetPhpRuntime');
    $runtime->Initialize();

    // Register the assemblies that we will be using. The runtime has bundles with some
    // of the most used assemblies of the .Net framework.
    $runtime->RegisterNetFramework2();

    // Add both SpreadsheetLight and the OpenXML it depends on.
    $runtime->RegisterAssemblyFromFile(APPLICATION_ROOT . '/binaries/SpreadsheetLight.dll', 'SpreadsheetLight');
    $runtime->RegisterAssemblyFromFile(APPLICATION_ROOT . '/binaries/DocumentFormat.OpenXml.dll', 'DocumentFormat.OpenXml');
    $runtime->RegisterAssemblyFromFile(APPLICATION_ROOT . '/binaries/AjaxMin.dll', 'AjaxMin');

    // Using the FullName of a type that belongs to an assembly that has already been registered.
    $datetime = $runtime->TypeFromName("System.DateTime");

    // Using the FullName of a type that has not been registered yet (from a file)
    $minifier = $runtime->TypeFromFile("Microsoft.Ajax.Utilities.Minifier", APPLICATION_ROOT . '/binaries/AjaxMin.dll');

    // Using the FullName of a type that has not been registered yet (autodiscoverable)
    $datetime2 = $runtime->TypeFromAssembly("System.DateTime", "mscorlib, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089");
  
    $datetime->Instantiate();
    echo $datetime->ToShortDateString()->Val(); // Outputs 01/01/0001

    // We can only use Int32 from native PHP, so parse
    // an Int64 that is equivalent to (long) in the DateTime constructor.
    $ticks = $runtime->TypeFromName("System.Int64")->Parse('98566569856565656');

    $ticks = \ms\System\netInt64::_Parse('98566569856565656');

    $datetime->Instantiate($ticks);
    echo $datetime->ToShortDateString()->Val(); // Outputs 07/05/0313

    $now = $runtime->TypeFromName("System.DateTime")->Now;
    echo $now->ToShortDateString()->Val(); // Outputs "23/10/2015"

    $now = $runtime->TypeFromName("System.DateTime")->Now();
    echo $now->ToShortDateString()->Val(); // Outputs "23/10/2015"

    $timer = $runtime->TypeFromName("System.Timers.Timer")->Instantiate();
    $timer->AutoReset(TRUE);
    $timer->AutoReset = TRUE;

    $datetime->AddYears(25);

    $data = $timer->GetPhpFromJson();

    var_dump($data);

    // Compare Enum
    $IsMonday = $runtime->TypeFromName("System.DateTime")->Now->DayOfWeek->Equals($runtime->TypeFromName("System.DayOfWeek")->Enum('Monday'));

    // Working with collections.
    $list = $runtime->TypeFromName("System.Collections.ArrayList")->Instantiate()->AsIterator();

    $list->Add('My first thing');
    $list->Add(2562);

    echo count($list); // Outputs 2

    foreach ($list as $item) {
      echo "item: {$item->Val()} </br>";
    }

  }
}
