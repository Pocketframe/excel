<?php

declare(strict_types=1);

namespace Pocketframe\Excel\Contracts;

interface ExporterInterface
{
  /**
   * Provide the column headings for the export.
   *
   * @return array
   */
  public function headings(): array;

  /**
   * Provide the data rows for the export.
   *
   * @return array
   */
  public function data(): array;

  /**
   * (Optional) Provide style settings for the export.
   *
   * The method should return an associative array where keys are cell ranges
   * (e.g., 'A1:C1') and values are style arrays as defined by PhpSpreadsheet.
   *
   * @return array|null
   */
  public function styles(): ?array;

  /**
   * (Optional) Provide multiple sheets for export.
   *
   * Each sheet should be defined as an array with keys:
   *  - 'name': (string) Sheet name.
   *  - 'headings': (array) Column headings.
   *  - 'data': (array) Data rows.
   *  - 'styles': (array, optional) Styling rules (same format as styles() above).
   *
   * If this method is implemented and returns a non-empty array, the engine will iterate
   * over each sheet definition.
   *
   * @return array|null
   */
  public function sheets(): ?array;
}
