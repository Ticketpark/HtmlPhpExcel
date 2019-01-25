<?php

require_once(__DIR__ . '/../vendor/autoload.php');

$htmlPhpExcel = new \Ticketpark\HtmlPhpExcel\HtmlPhpExcel('example.html');
$htmlPhpExcel->process()->output();
