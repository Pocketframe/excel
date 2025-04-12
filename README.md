# Pocketframe Excel

A robust Pocketframe package for Excel and CSV import/export built on PhpSpreadsheet. It supports advanced features such as cell styling, multiple sheets, chunked processing for very large files, and auto-detection of CSV files.

## Features

- **Auto-detection of file type:**
  Process Excel (.xlsx, .xls) and CSV files seamlessly.

- **Chunked Reading:**
  Use a custom ChunkReadFilter to limit memory usage by processing files in defined row chunks.

- **Advanced Export Options:**
  Apply cell formatting and styling, create multi-sheet exports, and handle complex excel features like merged cells and formulas.

- **Easy to Use API:**
  The package exposes a clean, static-style API through the Excel Mask that integrates with Pocketframe’s Storage, DataSet, and Response systems.

- **Extensible:**
  Use standard contracts for custom importer and exporter classes.

## Installation

Install via Composer:

```bash
composer require pocketframe/excel
```

## Usage

### Importing Data
Generating an importer by running the following command:

```bash
php pocket excel:create:importer UsersImporter
```
This will generate a file named UsersImporter.php in the app/Excel/Imports directory. You can then customize the file accordingly.

> [!NOTE]
> You can pass an entity name to the command to generate an importer for that entity.
> For example, `php pocket excel:create:importer --entity=User` will generate an importer for the User entity.

**Create an importer by implementing**

Pocketframe\Excel\Contracts\ImporterInterface. For example, create a file at app/Excel/Imports/UsersImporter.php:

```php
<?php
declare(strict_types=1);

namespace App\Excel\Imports;

use Pocketframe\Excel\Contracts\ImporterInterface;
use Pocketframe\PocketORM\Essentials\DataSet;
use App\Entities\User;

class UsersImporter implements ImporterInterface
{
    public function map(array $row): array
    {
        // Map raw row data to your data structure.
        return [
            'name'  => $row[0],
            'email' => $row[1],
            'age'   => (int)$row[2],
        ];
    }

    public function handle(array $data): void
    {
        // Process each mapped row,
        // for example by saving a new User entity.
        foreach ($data as $row) {
            (new User())->fill($row)->save();
        }
    }
}
```

Import your Excel file (with optional chunking and sheet selection) in your controller:

```php

use \Pocketframe\Excel\Excel;
use \App\Excel\Imports\UsersImporter;

Excel::import(UsersImporter::class, 'uploads/users.xlsx', 1000, 'DataSheet');
```

### Exporting Data
Generating an exporter by running the following command:

```bash
php pocket excel:create:exporter UsersExporter
```
This will generate a file named UsersExporter.php in the app/Excel/Exports directory. You can then customize the file accordingly.

> [!NOTE]
> You can pass an entity name to the command to generate an exporter for that entity.
> For example, `php pocket excel:create:exporter --entity=User` will generate an exporter for the User entity.

**Create an exporter by implementing**

Pocketframe\Excel\Contracts\ExporterInterface. For example, at app/Excel/Exports/UsersExporter.php:

```php
<?php
declare(strict_types=1);

namespace App\Excel\Exports;

use Pocketframe\Excel\Contracts\ExporterInterface;
use PhpOffice\PhpSpreadsheet\Style\Border;
use App\Entities\User;

class UsersExporter implements ExporterInterface
{
    public function headings(): array
    {
        return ['Name', 'Email', 'Age'];
    }

    public function data(): array
    {
        // Retrieve data from your entities.
        return array_map(function($user) {
            return [
              $user->name,
              $user->email,
              $user->age
            ];
        }, User::all()->toArray());
    }

    public function styles(): ?array
    {
        // Optional: Apply styling to the header row.
        return [
            'A1:C1' => [
                'font'    => ['bold' => true, 'size' => 12, 'color' => ['argb' => 'FF0000FF']],
                'borders' => [
                    'allBorders' => [
                        'borderStyle' => Border::BORDER_THIN,
                        'color'       => ['argb' => 'FF000000']
                    ]
                ],
            ]
        ];
    }

    public function sheets(): ?array
    {
        // Return null for single-sheet export.
        return null;
    }
}
```

> [!NOTE]
> For styling you can use the PhpSpreadsheet documentation for styling options. [PhpSpreadsheet Documentation](https://phpspreadsheet.readthedocs.io/en/latest/)

Export your Excel file (with optional chunking and sheet selection) in your controller:

```php

use \Pocketframe\Excel\Excel;
use \App\Excel\Exports\UsersExporter;

Excel::export(UsersExporter::class)->download('users.xlsx');
```
For a CSV export, simply use a CSV file name:

```php
Excel::export(UsersExporter::class)->download('users.csv');
```

### Multi-Sheet Export with Styling
If your exporter supports multiple sheets, implement the sheets() method:

To generate a multi-sheet exporter, run the following command:

```bash
php pocket excel:create:exporter UserMultiExporter --entity=User --multi
```

This will generate a file in the app/Excel/Exports directory named UserMultiExporter.php that includes the multi‑sheet boilerplate (with a sheets() method) and correctly references the User entity. You can then modify the file to customize the sheet names, columns, data, and styles.

```php
<?php
declare(strict_types=1);

namespace App\Exports;

use Pocketframe\Excel\Contracts\ExporterInterface;
use PhpOffice\PhpSpreadsheet\Style\Border;
use App\Entities\User;

class MultiSheetUsersExporter implements ExporterInterface
{
    public function headings(): array { return []; }
    public function data(): array { return []; }

    public function styles(): ?array { return null; }

    public function sheets(): ?array
    {
        return [
            [
                'name'     => 'Active Users',
                'headings' => ['Name', 'Email', 'Age'],
                'data'     => array_map(function($user) {
                    return [$user->name, $user->email, $user->age];
                }, User::active()->toArray()),
                'styles'   => [
                    'A1:C1' => [
                        'font'    => ['bold' => true, 'size' => 12, 'color' => ['argb' => 'FF0000FF']],
                        'borders' => [
                            'allBorders' => [
                                'borderStyle' => Border::BORDER_THIN,
                                'color'       => ['argb' => 'FF000000']
                            ]
                        ],
                    ]
                ],
            ],
            [
                'name'     => 'Inactive Users',
                'headings' => ['Name', 'Email', 'Age'],
                'data'     => array_map(function($user) {
                    return [$user->name, $user->email, $user->age];
                }, User::inactive()->toArray()),
            ],
        ];
    }
}
```

Then export with:

```php
Excel::export(MultiSheetUsersExporter::class)->download('users.xlsx');
```

### API Reference

#### Excel Mask Methods

```php
Excel::import($importerClass, $fileName, $chunkSize, $sheetName)
```
**Imports data from a file.**

Parameters:

- **`$importerClass (string)`**: The importer class name that implements Importer.

- **`$fileName (string)`**: File name relative to the storage disk.

- **`$chunkSize (int|null)`**: Optional chunk size for processing large files.

- **`$sheetName (string|null)`**: Optional sheet name (defaults to active sheet).

```php
Excel::export($exporterClass)
```
Exports data using the provided exporter. Returns an instance for chaining (e.g., `->download()`).

Parameters:

- **`$exporterClass (string)`**: The exporter class name that implements Exporter.

```php
->download($fileName)
```

Downloads the generated spreadsheet. The writer auto-detects the file format based on extension.
