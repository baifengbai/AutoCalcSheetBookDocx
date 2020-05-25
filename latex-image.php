<?php
$ch = curl_init();
$location = curl_escape($ch, $_GET['math']);
$remoteImage = 'http://chart.googleapis.com/chart?cht=tx&chl='.$location;
header("Content-type: image/png");
readfile($remoteImage);
curl_close($ch);
?>
