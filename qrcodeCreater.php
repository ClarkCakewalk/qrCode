<?php
include('phpqrcode/qrlib.php');
QRcode::png($_REQUEST['data'],false,'L',4);
?>