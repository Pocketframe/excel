<?php

declare(strict_types=1);

namespace Pocketframe\Excel\Contracts;

interface ImporterInterface
{
  /**
   * Map a raw row (an array of cell values) into an associative array.
   *
   * @param array $row
   * @return array Mapped row data.
   */
  public function map(array $row): array;

  /**
   * Process the complete set of mapped data.
   *
   * @param array $data An array of mapped rows.
   * @return void
   */
  public function handle(array $data): void;
}
