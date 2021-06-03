# HtmlPhpExcel

[![Build Status](https://github.com/Ticketpark/HtmlPhpExcel/actions/workflows/tests.yml/badge.svg)](https://github.com/Ticketpark/HtmlPhpExcel/actions)

This is a php library based on [PhpSpreadsheet](https://github.com/PHPOffice/PhpSpreadsheet) which simplifies converting html tables to excel files. It allows styling right within the html template with specific attributes.

## Installation

Add HtmlPhpExcel to your composer.json:

```
composer require ticketpark/htmlphpexcel
```

## Simple example
```php
<?php

require_once('../vendor/autoload.php');

$html = '<table><tr><th>Column A</th><th>Column B</th></tr><tr><td>Value A</td><td>Value B</td></tr></table>';
$htmlPhpExcel = new \Ticketpark\HtmlPhpExcel\HtmlPhpExcel($html);

// Create and output the excel file to the browser
$htmlPhpExcel->process()->output();

// Alternatively create the excel and save to a file
$htmlPhpExcel->process()->save('myFile.xlsx');

// or get the \PhpOffice\PhpSpreadsheet\Spreadsheet object to do further work with it
$phpExcelObject = $htmlPhpExcel->process()->getExcelObject();

```

For a more complex example with styling options see [example directory](example).

## Styling
There is support for specific html attributes to allow styling of the excel output. The attributes expect the content to be json_encoded.

* `_excel-styles`<br>Supports everything which is possible with PhpSpreadsheet's `applyFromArray()` method ([also see here](https://phpspreadsheet.readthedocs.io/en/latest/topics/recipes/#valid-array-keys-for-style-applyfromarray)).

Example:
```html
<table>
    <tr>
        <td _excel-styles='{"font":{"size":16,"color":{"rgb":"FF0000"}}}'>Foo</td>
    </tr>
</table>
```

* `_excel-dimensions`<br>Supports changing dimensions of rows (when applied to a `<tr>` or `<td>`) or column (when applied to a `<td>`),

Example:
```html
<table>
    <tr _excel-dimensions='{"row":{"rowHeight":50}}'>
        <td _excel-dimensions='{"column":{"width":20}}'>Foo</td>
    </tr>
</table>
```

* `_excel-explicit`<br>Supports applying an explicit cell value type.

Example:
```html
<table>
    <tr>
        <td _excel-explicit='PhpSpreadsheet_Cell_DataType::TYPE_STRING'>0022</td>
    </tr>
</table>
```