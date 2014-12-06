<?php

require_once "/home/webmaster/tgtoil.com/wp-includes/utilites/ExClasses/PHPExcel.php";

$report_id = @intval($_GET['report_id']);

$Reports = array("Revenue_Client_Cost-Centre", "Revenue_Client_Country", "Revenue_Cost-Centre_Country", "Revenue_Operator_Country",
                 "Revenue_Service_Category-1", "Revenue_Service_Category_Country", "Revenue_Service_Category-2", "Revenue_Service_Client", "Revenue_Service_Country",
                 "Revenue_Service_Operator", "Revenue_Technology_Country", "Revenue_Actual_vs_Target", "Revenue_Client_Country_Target", "Avg_Revenue_by_Country", "Avg_Revenue_by_Client",
                 "Avg_Revenue_by_Operator", "Avg_Revenue_by_Category", "assets1", "assets_summary", "assets_quant_summary",
                 "Asset_Util_by_Jobs_Count", "Asset_Util_by_Revenue", "Cons_Asset_Util", "hr_vacation_balance_v2", "hr_loading_balance_v2", "hr_vacation_balance_v2");

$file_url = "/home/webmaster/tgtoil.com/wp-content/themes/MyCuisine/xmlrpc_lib/xls_files/".$Reports[$report_id].".xlsx";
//echo $file_url;
$input_file_type = PHPExcel_IOFactory::identify($file_url);
$objReader = new PHPExcel_Reader_Excel2007($inputFileType);

$objPHPExcel = $objReader->load($file_url);

//$objWorksheet = $objPHPExcel->getActiveSheet();


$objWriter = new PHPExcel_Writer_HTML($objPHPExcel);

$result1 = '';
$result2 = '';

if ($objPHPExcel->getSheetCount() > 1) {
    foreach ($objPHPExcel->getWorksheetIterator() as $worksheet) {
        $result1 .= "<a id='title_".$objPHPExcel->getIndex($worksheet)."' href='#' class='button_bar_sm'>".$worksheet->getTitle()."</a>";

        // rest of code above

        $objWriter->setSheetIndex($objPHPExcel->getIndex($worksheet));

        $result2 .= $objWriter->generateHTMLHeader();


        $result2 .= "<style><!--";
        $result2 .= "html { font-family: Times New Roman; font-size: 9pt; background-color: white; }";

        $result2 .= $objWriter->generateStyles(false); // do not write <style> and </style>
        $result2 .= "--></style>";


        $result2 .= $objWriter->generateSheetData();
        $result2 .= $objWriter->generateHTMLFooter();
    } //

    print "<div id='sheet_links'>".$result1."</div><div id='sheet_tables'>".$result2."</div>";
}
else {
    echo $objWriter->generateHTMLHeader();

    ?>

    <style>
    <!--
    html {
      font-family: Times New Roman;
      font-size: 9pt;
      background-color: white;
    }

    <?php
    echo $objWriter->generateStyles(false); // do not write <style> and </style>
    ?>

    -->
    </style>

    <?php


    echo $objWriter->generateSheetData();
    echo $objWriter->generateHTMLFooter();
} //end else

?>