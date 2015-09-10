<?php

require 'HarvestSheet.php';
date_default_timezone_set('Europe/Amsterdam');

$path = $argv[1];
$sheet = new HarvestSheet($path);
$sheet->removeColumns(array(
	'Billable?', 'Invoiced?', 'Approved?', 'Employee?', 'Billable Rate',
	'Billable Amount', 'Cost Rate', 'Cost Amount', 'Currency'
));
$sheet->splitColumn('B', 'Client Code');
$sheet->output();
