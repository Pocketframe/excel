<?php

declare(strict_types=1);

namespace Pocketframe\Excel\Console\Commands;

use Pocketframe\Contracts\CommandInterface;

class CreateImporterCommand implements CommandInterface
{
  protected array $args;
  protected string $stubPath;
  protected string $importerName;
  protected ?string $entityName = null;

  public function __construct(array $args)
  {
    $this->args = $args;
    $this->stubPath = dirname(__DIR__, 3) . '/stubs';
  }

  public function handle(): void
  {
    if (empty($this->args)) {
      $this->showUsage();
      exit(1);
    }

    $this->parseOptions();

    if (!$this->importerName) {
      $this->showUsage();
      exit(1);
    }

    $this->createImporter();
  }

  protected function createImporter(): void
  {
    $targetDir = base_path("app/Excel/Imports");
    $targetPath = $targetDir . "/" . ucfirst($this->importerName) . "Importer.php";

    if (!is_dir($targetDir)) mkdir($targetDir, 0777, true);

    if (file_exists($targetPath)) {
      echo "Importer " . ucfirst($this->importerName) . "Importer already exists.\n";
      exit(1);
    }

    $stub = file_get_contents("{$this->stubPath}/importer.stub");

    // Set namespace and className defaults.
    $namespace = 'App\\Excel\\Imports';
    $className = ucfirst($this->importerName) . "Importer";
    // If an entity name has been provided via --entity=, use it; otherwise provide a default placeholder.
    $entity = $this->entityName ?? 'YourEntity';

    $content = str_replace(
      ['{{namespace}}', '{{className}}', '{{entity}}'],
      [$namespace, $className, $entity],
      $stub
    );

    file_put_contents($targetPath, $content);
    echo "💪 Importer created: {$targetPath}\n";
  }

  protected function parseOptions(): void
  {
    // Process arguments looking for the entity flag.
    $this->entityName = null;
    foreach ($this->args as $key => $arg) {
      if (str_starts_with($arg, '--entity=')) {
        $parts = explode('=', $arg, 2);
        $this->entityName = trim($parts[1]);
        // Remove this option from the args.
        unset($this->args[$key]);
      }
    }
    // Remove any remaining flags (arguments starting with '-') and reindex args.
    $this->args = array_values(array_filter($this->args, fn($arg) => !str_starts_with($arg, '-')));
    $this->importerName = $this->args[0] ?? null;
  }

  protected function showUsage(): void
  {
    echo "Usage: php pocket excel:create:importer ImporterName [--entity=EntityName]\n";
  }
}
