#HtmlPhpExcel
This is a php library based on [PHPExcel](https://phpexcel.codeplex.com/) which simplifies converting html tables to excel files. It allows styling right within the html template with specific attributes.

## Todos
* Write documentation for usage of table class, row class and cell class filtering
* Write documentation for cell merging functionalities
* Write documentation for usage of `utf8EncodeValues()` and `utf8DecodeValues()`

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
$htmlPhpExcel->process()->save('myFile.xls');

// or get the PHPExcel object to do further work with it
$phpExcelObject = $htmlPhpExcel->process()->getExcelObject();

```

For a more complex example with styling options see [example directory](example).

## Styling
There is support for specific html attributes to allow styling of the excel output. The attributes expect the content to be json_encoded.

* `_excel-styles`<br>Supports everything which is possible with PHPExcel's `applyFromArray()` method ([also see here](http://phpexcel.codeplex.com/discussions/206914)).

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
        <td _excel-explicit='PHPExcel_Cell_DataType::TYPE_STRING'>0022</td>
    </tr>
</table>
```

## License
This bundle is under the MIT license. See the complete license in the bundle:

    Resources/meta/LICENSE