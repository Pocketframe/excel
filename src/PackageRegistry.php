<?php

declare(strict_types=1);

namespace Pocketframe\Excel;

use Pocketframe\Container\Container;
use Pocketframe\Excel\Mask\Excel;
use Pocketframe\Package\Contract\PackageInterface;
use Pocketframe\Console\Kernel;

class PackageRegistry implements PackageInterface
{
  /**
   * Register Excel package bindings and commands.
   *
   * @param Container $container
   * @return void
   */
  public function register(Container $container): void
  {
    // Bind the Excel mask.
    $container->bind('excel', fn() => new Excel(null));

    // Register Excel console commands.
    // These are optional: commands may be resolved via the container.
    $container->bind('command.excel.create.importer', fn() => new Console\Commands\CreateImporterCommand([]));
    $container->bind('command.excel.create.exporter', fn() => new Console\Commands\CreateExporterCommand([]));

    /** @var Kernel $kernel */
    $kernel = $container->get(Kernel::class);

    $kernel->addCommand(
      'excel:create:importer',
      Console\Commands\CreateImporterCommand::class,
      'Generate a new Excel importer class.'
    );

    $kernel->addCommand(
      'excel:create:exporter',
      Console\Commands\CreateExporterCommand::class,
      'Generate a new Excel exporter class.'
    );
  }
}
