# HtmlPhpExcel

[![Build Status](https://github.com/Ticketpark/HtmlPhpExcel/actions/workflows/tests.yml/badge.svg)](https://github.com/Ticketpark/HtmlPhpExcel/actions)

This is a php library based on [FastExcelWriter](https://github.com/aVadim483/fast-excel-writer), simplifying converting html tables to excel files. It allows styling right within the html template with specific attributes.

## Installation

Add HtmlPhpExcel to your composer.json:

```
composer require ticketpark/htmlphpexcel
```

## Simple example
```php
<?php

require_once('vendor/autoload.php');

$html = '<table><tr><th>Column A</th><th>Column B</th></tr><tr><td>Value A</td><td>Value B</td></tr></table>';
$htmlPhpExcel = new \Ticketpark\HtmlPhpExcel\HtmlPhpExcel($html);

$htmlPhpExcel->process()->save('myFile.xlsx');

```

For a more complex example see [example directory](example).

## Styling
Styles are set with an html element `_excel-styles`. The attribute expects the content to be in json format.

Example:
```html
<table>
    <tr _excel-styles='{"height": 50}'>
        <td _excel-styles='{"font-size": 16, "font-color": "#FF0000", "width": 200}'>
            Cell value
        </td>
    </tr>
</table>
```

You can use any style supported by `fast-excel-writer`, of which the most common are:

* border-color
* border-style
* fill-color
* fill-pattern
* font-color
* font-size
* format
* format-text-wrap
* height
* text-align
* text-color
* text-rotation
* text-wrap
* vertical-align
* width

More information (though unfortunately limited) is available in the [docs of FastExcelWriter](https://github.com/aVadim483/fast-excel-writer/blob/master/docs/04-styles.md).

## Adding comments to cells

Example:
```html
<table>
    <tr >
        <td _excel-comment="This is a comment.">
            Cell value
        </td>
    </tr>
</table>
```

## History

* v2.x of this library is based on `FastExcelWriter`
* v1.x of this library was based on `PhpSpreadsheet`
* v0.x of this library was based on `PhpExcel`