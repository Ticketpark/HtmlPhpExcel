<?php

require_once('../vendor/autoload.php');

$htmlPhpExcel = new \Ticketpark\HtmlPhpExcel\HtmlPhpExcel('example.html');
$htmlPhpExcel->process()->output();
