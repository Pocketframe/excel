<?php

declare(strict_types=1);

namespace Pocketframe\Excel\Engine;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use Pocketframe\PocketORM\Essentials\DataSet;
use Pocketframe\Excel\Contracts\ExporterInterface;

class PhpSpreadsheetEngine
{
  /**
   * Read an Excel or CSV file into a DataSet.
   *
   * @param string $filePath Full file path.
   * @param int|null $chunkSize Optional number of rows per chunk.
   * @param string|null $sheetName Optional sheet name to load.
   * @return DataSet
   */
  public function read(string $filePath, ?int $chunkSize = null, ?string $sheetName = null): DataSet
  {
    // Auto-detect CSV extension
    $ext = strtolower(pathinfo($filePath, PATHINFO_EXTENSION));
    if ($ext === 'csv') {
      $rows = [];
      $skipHeader = true;
      if (($handle = fopen($filePath, 'r')) !== false) {
        while (($data = fgetcsv($handle)) !== false) {
          if ($skipHeader) {
            $skipHeader = false;
            continue;
          }
          $rows[] = $data;
        }
        fclose($handle);
      }
      return new DataSet($rows);
    }

    // Create a reader for Excel files.
    $reader = IOFactory::createReaderForFile($filePath);
    if ($sheetName !== null) {
      $reader->setLoadSheetsOnly($sheetName);
    }

    // Use chunking if specified.
    if ($chunkSize !== null && $chunkSize > 0) {
      $chunkFilter = new ChunkReadFilter();
      $reader->setReadFilter($chunkFilter);

      $rows = [];
      $startRow = 1;
      $skipHeader = true;

      do {
        $chunkFilter->setRows($startRow, $chunkSize);
        $spreadsheet = $reader->load($filePath);
        $sheet = $spreadsheet->getActiveSheet();
        $chunkRows = [];

        foreach ($sheet->getRowIterator($startRow, $startRow + $chunkSize - 1) as $row) {
          $cellIterator = $row->getCellIterator();
          $cellIterator->setIterateOnlyExistingCells(true);
          $cells = [];
          foreach ($cellIterator as $cell) {
            $cells[] = $cell->getValue();
          }
          // Only include non-empty rows
          if (array_filter($cells)) {
            $chunkRows[] = $cells;
          }
        }
        if (empty($chunkRows)) {
          break;
        }
        $rows = array_merge($rows, $chunkRows);
        $startRow += $chunkSize;
        $spreadsheet->disconnectWorksheets();
        unset($spreadsheet);
      } while (true);

      return new DataSet($rows);
    }

    // Otherwise, load all rows at once.
    $spreadsheet = $reader->load($filePath);
    $sheet = $spreadsheet->getActiveSheet();
    $rows = [];
    $skipHeader = true;
    foreach ($sheet->getRowIterator() as $row) {
      if ($skipHeader) {
        $skipHeader = false;
        continue;
      }
      $cellIterator = $row->getCellIterator();
      $cellIterator->setIterateOnlyExistingCells(true);
      $cells = [];
      foreach ($cellIterator as $cell) {
        $cells[] = $cell->getValue();
      }
      $rows[] = $cells;
    }
    return new DataSet($rows);
  }

  /**
   * Create a Spreadsheet from an exporter.
   *
   * Supports single-sheet and multi-sheet exports.
   *
   * @param Exporter $exporter
   * @return Spreadsheet
   */
  public function createSpreadsheet(ExporterInterface $exporter): Spreadsheet
  {
    $spreadsheet = new Spreadsheet();

    // If the exporter has a sheets() method and returns an array, handle multi-sheet export.
    if (method_exists($exporter, 'sheets') && !empty($exporter->sheets())) {
      $sheetDefs = $exporter->sheets();
      foreach ($sheetDefs as $index => $sheetDef) {
        if ($index > 0) {
          $spreadsheet->createSheet();
        }
        $sheet = $spreadsheet->setActiveSheetIndex($index);
        $sheet->setTitle($sheetDef['name'] ?? "Sheet " . ($index + 1));
        $sheet->fromArray($sheetDef['headings'] ?? [], null, 'A1');
        $sheet->fromArray($sheetDef['data'] ?? [], null, 'A2');
        // Apply styles if provided
        if (isset($sheetDef['styles']) && is_array($sheetDef['styles'])) {
          foreach ($sheetDef['styles'] as $range => $styleArray) {
            $sheet->getStyle($range)->applyFromArray($styleArray);
          }
        }
      }
      $spreadsheet->setActiveSheetIndex(0);
    } else {
      // Single-sheet export.
      $sheet = $spreadsheet->getActiveSheet();
      $sheet->fromArray($exporter->headings(), null, 'A1');
      $sheet->fromArray($exporter->data(), null, 'A2');
      // Apply styles if exporter provides a styles() method.
      if (method_exists($exporter, 'styles') && ($styles = $exporter->styles())) {
        foreach ($styles as $range => $styleArray) {
          $sheet->getStyle($range)->applyFromArray($styleArray);
        }
      }
    }

    return $spreadsheet;
  }
}
