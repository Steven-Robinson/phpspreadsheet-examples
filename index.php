<?php

require __DIR__ . '/vendor/autoload.php';

$app = new \App\App();

$app['spreadsheet.controller']->index();
