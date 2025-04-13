<?php

declare(strict_types=1);

namespace Pocketframe\Excel\Console\Commands;

use Pocketframe\Contracts\CommandInterface;

class CreateExporterCommand implements CommandInterface
{
  protected array $args;
  protected string $stubPath;
  protected string $exporterName;
  protected ?string $entityName = null;
  protected bool $isMultiSheet = false;

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

    if (!$this->exporterName) {
      $this->showUsage();
      exit(1);
    }

    $this->createExporter();
  }

  protected function createExporter(): void
  {
    $targetDir = base_path("app/Excel/Exports");
    $targetPath = $targetDir . "/" . ucfirst($this->exporterName) . "Exporter.php";

    if (!is_dir($targetDir)) {
      mkdir($targetDir, 0777, true);
    }

    if (file_exists($targetPath)) {
      echo "Exporter " . ucfirst($this->exporterName) . "Exporter already exists.\n";
      exit(1);
    }

    // Choose stub file based on multi-sheet flag.
    $stubFile = $this->isMultiSheet
      ? "{$this->stubPath}/exporter.multisheet.stub"
      : "{$this->stubPath}/exporter.stub";

    $stub = file_get_contents($stubFile);
    $namespace = 'App\\Excel\\Exports';
    $className = ucfirst($this->exporterName) . "Exporter";
    $entity = $this->entityName ?? 'YourEntity';

    $content = str_replace(
      ['{{namespace}}', '{{className}}', '{{entity}}'],
      [$namespace, $className, $entity],
      $stub
    );

    file_put_contents($targetPath, $content);
    echo "💪 Exporter created: {$targetPath}\n";
  }

  protected function parseOptions(): void
  {
    // Process flags:
    // --entity=, --multi or -m are supported.
    $this->entityName = null;
    $this->isMultiSheet = false;

    foreach ($this->args as $key => $arg) {
      if (str_starts_with($arg, '--entity=')) {
        $parts = explode('=', $arg, 2);
        $this->entityName = trim($parts[1]);
        unset($this->args[$key]);
      }
      if ($arg === '--multi' || $arg === '-m') {
        $this->isMultiSheet = true;
        unset($this->args[$key]);
      }
    }
    // Remove any remaining flags and reindex.
    $this->args = array_values(array_filter($this->args, fn($arg) => !str_starts_with($arg, '-')));
    $this->exporterName = $this->args[0] ?? null;
  }

  protected function showUsage(): void
  {
    echo "Usage: php pocket excel:create:exporter ExporterName [--entity=EntityName] [--multi|-m]\n";
  }
}
