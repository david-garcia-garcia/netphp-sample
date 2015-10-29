<?php

use \ms\Typer;

include_once('vendor/autoload.php');

define('APPLICATION_ROOT', dirname(__FILE__));

header("Content-Type: text/html; charset=utf-8");

print '<html><title>Flush Test</title><head></head><body>';

function flush_buffers(){
  flush();
}

print "Generating static class model... this might take a while on the first run.";
flush_buffers();

// On the first RUN the static models will be created, this might take A WHILE!
\MyProject\Samples\CompleteFramework\RuntimeManager::Instance()->GenerateStaticClassModel();
\MyProject\Samples\Spreadsheetlight\RuntimeManager::Instance()->GenerateStaticClassModel();
\MyProject\Samples\WordInterop\RuntimeManager::Instance()->GenerateStaticClassModel();

print "Running examples...";
flush_buffers();

//\MyProject\Samples\WordInterop\Example1::Example1();

// Now call the samples.
\MyProject\Samples\Spreadsheetlight\Example1::Example1();
\MyProject\Samples\Spreadsheetlight\Example1::Example2();
\MyProject\Samples\CompleteFramework\Example1::Example1();

print "</body></html>";

?>