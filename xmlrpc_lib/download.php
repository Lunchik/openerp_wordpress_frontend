<?php
   $file_path = "http://tgtoil.com/wp-content/themes/MyCuisine/xmlrpc_lib/xls_files/".$_GET['path'];
   //echo $file_path;
   header("Content-Type: application/octet-stream");
   header("Content-Disposition: attachment; filename=".$_GET['path']);
   readfile($file_path);
?>