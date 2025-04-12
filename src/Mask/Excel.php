<?php

declare(strict_types=1);

namespace Pocketframe\Excel\Mask;

use Pocketframe\Excel\Engine\PhpSpreadsheetEngine;
use Pocketframe\Excel\Contracts\ImporterInterface;
use Pocketframe\Excel\Contracts\ExporterInterface;
use Pocketframe\Storage\Facade\Storage;
use Pocketframe\Http\Response\Response;

class Excel
{
  protected mixed $spreadsheet;

  /**
   * Retrieve an instance of the PhpSpreadsheetEngine.
   *
   * @return PhpSpreadsheetEngine
   */
  protected static function engine(): PhpSpreadsheetEngine
  {
    return new PhpSpreadsheetEngine();
  }

  /**
   * Import data from a file using the given Importer.
   *
   * @param string $importerClass Class name implementing Importer.
   * @param string $fileName File name relative to the storage disk.
   * @param int|null $chunkSize Optional number of rows per chunk.
   * @param string|null $sheetName Optional sheet name to load.
   * @return void
   */
  public static function import(string $importerClass, string $fileName, ?int $chunkSize = null, ?string $sheetName = null): void
  {
    /** @var Importer $importer */
    $importer = new $importerClass;

    // Resolve the full file path using the Storage mask.
    $filePath = Storage::getInstance()->path($fileName);

    // Read the file (with chunking if set) into a DataSet.
    $dataSet = static::engine()->read($filePath, $chunkSize, $sheetName);

    // Map each raw row through the importer.
    $mappedData = $dataSet->map(fn($row) => $importer->map($row))->toArray();

    // Let the importer process the full mapped data.
    $importer->handle($mappedData);
  }

  /**
   * Export data using the given Exporter.
   *
   * @param string $exporterClass Class name implementing Exporter.
   * @return self
   */
  public static function export(string $exporterClass): self
  {
    /** @var Exporter $exporter */
    $exporter = new $exporterClass;
    $spreadsheet = static::engine()->createSpreadsheet($exporter);
    return new static($spreadsheet);
  }

  /**
   * Constructor.
   *
   * @param mixed $spreadsheet A Spreadsheet object.
   */
  protected function __construct(mixed $spreadsheet)
  {
    $this->spreadsheet = $spreadsheet;
  }

  /**
   * Initiate a download of the generated spreadsheet.
   *
   * @param string $fileName The file name for the downloaded file.
   * @return void
   */
  public function download(string $fileName): void
  {
    // Auto-detect file extension for writer.
    $ext = strtolower(pathinfo($fileName, PATHINFO_EXTENSION));
    if ($ext === 'csv') {
      $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($this->spreadsheet, 'Csv');
    } else {
      $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($this->spreadsheet, 'Xlsx');
    }

    // Capture output in a temporary file.
    ob_start();
    $writer->save('php://output');
    $content = ob_get_clean();

    $tempFile = tempnam(sys_get_temp_dir(), 'excel_');
    file_put_contents($tempFile, $content);

    // Use Pocketframe's Response to send the file as a download.
    Response::file($tempFile, $fileName)->send();

    // Clean up the temporary file.
    unlink($tempFile);
  }
}
