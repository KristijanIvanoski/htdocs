<?php

$vars = $_GET;

if (empty($vars['fileName'])) {
	echo "No fileName specified";
	exit;
}

$x = explode("\\", dirname(__FILE__));
$safeRoot = '';
for ($i = 0; $i <= count($x) - 3; $i++) {
    $safeRoot .= $x[$i] . '\\';
}

$r = explode('.',  $vars["fileName"]);
$safeOutputName = $r[0] . '.xls';

$output = shell_exec( '..\..\helpers\ConvertExcelTo2003.exe "' . $safeRoot . 'attachment\\' . $vars["fileName"] . '" "' . $safeRoot . 'attachment\\' . $safeOutputName . '"');

if (empty($output)) {
	echo "Converted";
} else {
	echo "<pre>$output</pre>";
}