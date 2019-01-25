<?php

require_once(__DIR__ . '/../vendor/autoload.php');

$htmlPhpExcel = new \Ticketpark\HtmlPhpExcel\HtmlPhpExcel(__DIR__  . '/example.html');
$htmlPhpExcel->process()->save(__DIR__ . '/test.xlsx');
