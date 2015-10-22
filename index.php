<?php

use \ms\Typer;

include_once('vendor/autoload.php');

define('APPLICATION_ROOT', dirname(__FILE__));

// On the first RUN the static models will be created, this might take A WHILE!
\MyProject\Samples\CompleteFramework\RuntimeManager::Instance()->GenerateStaticClassModel();
\MyProject\Samples\Spreadsheetlight\RuntimeManager::Instance()->GenerateStaticClassModel();

// Now call the samples.
\MyProject\Samples\Spreadsheetlight\Example1::Example1();
\MyProject\Samples\CompleteFramework\Example1::Example1();

?>