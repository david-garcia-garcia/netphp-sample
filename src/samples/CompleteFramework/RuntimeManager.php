<?php

namespace MyProject\Samples\CompleteFramework;



/**
 * Runtime to use for the SpreadsheetLight exammples
 */
class RuntimeManager {

  private static $manager = NULL;

  public static function Instance() {
    if (empty(static::$manager)) {
      static::$manager = new RuntimeManager();
    }
    return static::$manager;
  }

  /**
   *  @var \NetPhp\Core\NetPhpRuntime
   */
  private $runtime = NULL;

  /**
   * Make sure that the runtime is initialized and the
   * static type maps are binded to this runtime.
   */
  public function InitializeRuntime() {

    if (empty($this->runtime)) {

      // Instantiate and initialize the Model
      $this->runtime = new \NetPhp\Core\NetPhpRuntime('COM', 'netutilities.NetPhpRuntime');
      $this->runtime->Initialize();

      // Register the assemblies that we will be using. The runtime has bundles with some
      // of the most used assemblies of the .Net framework.
      $this->runtime->RegisterNetFramework4();

      // Once we have a Static Typed PHP model
      // we need to explictly tell it to use
      // a runtime.

      // We check if it exists because we might not have yet dumped
      // any class model.
      if (class_exists(\ms2\TypeMap::class)) {
        \ms2\TypeMap::SetRuntime($this->runtime);
      }
    }

  }

  /**
   * Generate a PhpRuntime that will be used both at runtime
   * and to generate the NetPhp class model.
   *
   * @return \NetPhp\Core\NetPhpRuntime
   */
  public function GetRuntime() {

    $this->InitializeRuntime();

    return $this->runtime;
  }

  /**
   * Use this code to generate a static PHP class model
   * that you can use work with strong typed .Net objects
   * from PHP.
   */
  public function GenerateStaticClassModel() {

    // Initialize the dumper.
    $dumper = new \NetPhp\Core\TypeDumper();
    $dumper->Initialize();

    // If you don't call this, and the destination directory is not empty
    // the static class model will not be generated. Just a safe guard.
    //$dumper->AllowDestinationDirectoryClear();

    // Set the destination path and base namespace.
    // Al the .Net namespaces will be nested inside this
    // namespace.
    $dumper->SetDestination(APPLICATION_ROOT . '/ms2');
    $dumper->SetBaseNamespace('ms2');

    // Get a copy of the runtime.
    $runtime = $this->GetRuntime();

    // Tell the runtime to register the assemblies in the Dumper.
    $runtime->RegisterAssembliesInDumper($dumper);

    // Dump the complete SpreadshetLight namespace.
    // You need to explicitly add one or more regular expressions
    // that will thell the dumper what Types to consider
    // when dumping the model.

    // Dump EVERYTHING!
    $dumper->AddDumpFilter('^System\\.DateTime$');
    $dumper->AddDumpFilter('^System\\.Diagnostics\\.Process$');

    // Dump EVERYTHING!
    $dumper->SetDumpDepth(2);

    // Generate the static class model.
    try {
      $dumper->GenerateModel();
      // Just make sure that PHP can start using this new classes
      // during this same request.
      clearstatcache();
    }
    catch (\Exception $e){}

    // Just make sure that we will be missing no dependencies
    // on the deployment environment!
    $report = $runtime->GetAssemblyReport();
  }

}
