<?php

declare(strict_types=1);

namespace Pocketframe\Excel\Mask;

use Pocketframe\Excel\Engine\PhpSpreadsheetEngine;
use Pocketframe\Excel\Contracts\ImporterInterface;
use Pocketframe\Excel\Contracts\ExporterInterface;
use Pocketframe\Storage\Mask\Storage;
use Pocketframe\Http\Response\Response;

class Excel
{
  protected mixed $spreadsheet;

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
  public static function import(
    string $importerClass,
    string|\Pocketframe\Http\Request\UploadedFile $fileName,
    ?int $chunkSize = null,
    ?string $sheetName = null
  ): void {
    // If $fileName is an instance of UploadedFile, convert it to the real path.
    if ($fileName instanceof \Pocketframe\Http\Request\UploadedFile) {
      $fullPath = $fileName->getRealPath();
    } else {
      // For regular file paths, use concrete Storage class
      $fullPath = (new \Pocketframe\Storage\Storage())->path($fileName);
    }
    // Proceed as normal.
    /** @var ImporterInterface $importer */
    $importer = new $importerClass;

    // Read the file into a DataSet.
    $dataSet = self::engine()->read($fullPath, $chunkSize, $sheetName);

    // Map and handle the data.
    $mappedData = $dataSet->map(fn($row) => $importer->map($row))->toArray();
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
    /** @var ExporterInterface $exporter */
    $exporter = new $exporterClass;
    $spreadsheet = static::engine()->createSpreadsheet($exporter);
    return new static($spreadsheet);
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
