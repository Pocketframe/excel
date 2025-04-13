<?php

declare(strict_types=1);

namespace Pocketframe\Excel\Engine;

use PhpOffice\PhpSpreadsheet\Reader\IReadFilter;

class ChunkReadFilter implements IReadFilter
{
  private int $startRow = 0;
  private int $endRow = 0;

  /**
   * Set the rows to be read.
   *
   * @param int $startRow The starting row.
   * @param int $chunkSize The number of rows to read.
   */
  public function setRows(int $startRow, int $chunkSize): void
  {
    $this->startRow = $startRow;
    $this->endRow   = $startRow + $chunkSize - 1;
  }

  /**
   * Decide whether the current cell should be read.
   *
   * @param string $column Column letter.
   * @param int $row Row number.
   * @param string $worksheetName Worksheet name.
   * @return bool True if the row is within the desired range.
   */
  public function readCell($column, $row, $worksheetName = ''): bool
  {
    return ($row >= $this->startRow && $row <= $this->endRow);
  }
}
