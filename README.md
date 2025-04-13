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
php pocket excel:create:importer Users
```
>[!important]
> You should not add Importer or Import to the file name. This will be added automatically and the file name will become `UsersImporter.php`. If you add Importer or Import to the file name will now become `UsersImporterImporter.php` which may not look nice. This also applies to the exporters.

This will generate a file named UsersImporter.php in the `app/Excel/Imports` directory. You can then customize the file accordingly.

> [!NOTE]
> You can pass an entity name to the command to generate an importer for that entity.
> For example, `php pocket excel:create:importer Users --entity=User` will generate an importer for the User entity.

**Create an importer by implementing**

`Pocketframe\Excel\Contracts\ImporterInterface`. For example, create a file at `app/Excel/Imports/UsersImporter.php` with the following content:

```php
<?php
declare(strict_types=1);

namespace App\Excel\Imports;

use Pocketframe\Excel\Contracts\ImporterInterface;
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
            $users = new User([
              'name'  => $row['name'],
              'email' => $row['email'],
              'age'   => $row['age'],
            ]);

            $users->save();
        }
    }
}
```

Import your Excel file (with optional chunking and sheet selection) in your controller:

```php

use \Pocketframe\Excel\Excel;
use \App\Excel\Imports\UsersImporter;

$path = $request->file('file')->store('uploads');

Excel::import(UsersImporter::class, $path, 1000, 'user_sheet');
```

> [!tip]
> You can also check for duplicates in the `handle` method. For example, you can check if a user with the same email already exists before saving it to the database.
> ```php
>  public function handle(array $data): void
>  {
>    foreach ($data as $row) {
>      $existing_user = (new QueryEngine(User::class))
>        ->where('email', '=', $row['email'])
>        ->first();
>
>     if ($existing_user) {
>        continue;
>      }
>
>      $category = new Category([
>        'category_name' => $row['category_name'],
>       'slug'          => $row['slug'],
>        'description'   => $row['description'],
>        'status'        => $row['status'],
>      ]);
>      $category->save();
>    }
>  }

### Exporting Data
Generating an exporter by running the following command:

```bash
php pocket excel:create:exporter Users
```
This will generate a file named UsersExporter.php in the `app/Excel/Exports` directory. You can then customize the file accordingly.

> [!tip]
> You can pass an entity name to the command to generate an exporter for that entity.
> For example, `php pocket excel:create:exporter Users --entity=User` will generate an exporter for the User entity.

**Create an exporter by implementing**

`Pocketframe\Excel\Contracts\ExporterInterface`. For example, at `app/Excel/Exports/UsersExporter.php`:

```php
<?php
declare(strict_types=1);

namespace App\Excel\Exports;

use Pocketframe\Excel\Contracts\ExporterInterface;
use PhpOffice\PhpSpreadsheet\Style\Border;
use App\Entities\User;
use Pocketframe\PocketORM\Database\QueryEngine;

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
        }, QueryEngine::for(User::class)->get()->all());
    }

    public function styles(): ?array
    {
        // Optional: Apply styling to the header row.
        return [
            'A1:C1' => [
                'font'    => [
                  'bold' => true,
                  'size' => 12,
                  'color' => [
                    'argb' => 'FF0000FF'
                  ]
                ],
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
On top of fetching data, you can also add filters by applying where conditions that are part of the QueryEngine.

```php
return array_map(function($user) {
    return [
      $user->name,
      $user->email,
      $user->age
    ];
}, QueryEngine::for(User::class)
  ->where('age', '>', 18)
  ->get()
  ->all());
```

> [!NOTE]
> For styling you can use the PhpSpreadsheet documentation for styling options. [PhpSpreadsheet Documentation](https://phpspreadsheet.readthedocs.io/en/latest/). Native support for styling will be added in the future.

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

#### Additional Styling
You can also add additional styling to the exported file through the `configureSheet` method. For example you can autosize the columns, set the default font, set conditional columns and set the default row height.

```php
public function configureSheet(Worksheet $sheet): void
{
    // Auto-size all columns from A through E.
    foreach (range('A', 'E') as $columnID) {
      $sheet->getColumnDimension($columnID)->setAutoSize(true);
    }

    // Determine the highest row with data.
    $highestRow = $sheet->getHighestRow();

    // Loop through each data row (assuming headers are in row 1).
    for ($row = 2; $row <= $highestRow; $row++) {
      // Get the value of the "Status" cell (column D).
      $statusCell = $sheet->getCell("D{$row}");
      $status = strtolower(trim($statusCell->getValue()));

      if ($status === 'active') {
        // Apply a light green fill for "active" status.
        $sheet->getStyle("D{$row}")
          ->getFill()
          ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
          ->getStartColor()->setARGB('FFB6D7A8');
      } elseif ($status === 'inactive') {
        // Apply a light red/pink fill for "inactive" status.
        $sheet->getStyle("D{$row}")
          ->getFill()
          ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
          ->getStartColor()->setARGB('FFF4B7B2');
      }
    }

    // Optionally, apply a border style to all data cells.
    $sheet->getStyle("A2:E{$highestRow}")->applyFromArray([
      'borders' => [
        'allBorders' => [
          'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
          'color'       => ['argb' => 'FFCCCCCC']
        ],
      ],
    ]);
  }
  ```

### Multi-Sheet Export with Styling

To generate a multi-sheet exporter, run the following command:

```bash
php pocket excel:create:exporter UserMulti --entity=User --multi
```

This will generate a file in the `app/Excel/Exports` directory named `UserMultiExporter.php` that includes the multi‑sheet boilerplate (with a sheets() method) and correctly references the User entity. You can then modify the file to customize the sheet names, columns, data, and styles.

```php
<?php
declare(strict_types=1);

namespace App\Exports;

use Pocketframe\Excel\Contracts\ExporterInterface;
use PhpOffice\PhpSpreadsheet\Style\Border;
use App\Entities\User;
use Pocketframe\PocketORM\Database\QueryEngine;

class MultiSheetUsersExporter implements ExporterInterface
{
    public function sheets(): ?array
    {
        return [
            [
                'name' => 'Active Users',
                'headings' => [
                  'Name',
                  'Email',
                  'Age'
                ],
                'data' => array_map(function($user) {
                    return [
                      $user->name,
                      $user->email,
                      $user->age
                    ];
                }, QueryEngine::for(User::class)
                  ->where('status', 'active')
                  ->get()
                  ->all()),
                'styles'   => [
                    'A1:C1' => [
                        'font' => [
                          'bold' => true,
                          'size' => 12,
                          'color' => ['argb' => 'FF0000FF']
                        ],
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
                'name'=> 'Inactive Users',
                'headings' => [
                  'Name',
                  'Email',
                  'Age'
                ],
                'data' => array_map(function($user) {
                    return [
                      $user->name,
                      $user->email,
                      $user->age
                    ];
                }, QueryEngine::for(User::class)
                  ->where('status', 'inactive')
                  ->get()
                  ->all()),
            ],
        ];
    }

    // do not remove these methods
    public function headings(): array { return []; }
    public function data(): array { return []; }
    public function styles(): ?array { return null; }
}
```

Then export with:

```php
use \App\Excel\Exports\MultiSheetUsersExporter;
use \Pocketframe\Excel\Excel;

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
