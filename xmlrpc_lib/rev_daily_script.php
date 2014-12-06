<?php

include 'xmlrpc.inc';
/** Include PHPExcel **/
require_once "/home/webmaster/tgtoil.com/wp-includes/utilites/ExClasses/PHPExcel.php";
require_once 'report_class.php';
require_once 'xml_styles.php';

set_time_limit(0);
ignore_user_abort(1);


/**********
Target Revenue Report (single file) - connect, get data
generate file call -- below, above the definition of functions
**********/
echo "Hello!";

$file_names1 = array(
                      "Actual_vs_Target"
                );
$Tar_Report = new Report;
$relation_tar = array(
                      "location.rev"
                );
$required_keys_tar = array(
            array(
                       new xmlrpcval('country_id', "string"),
                       new xmlrpcval('tar_mo', "string"),
                       new xmlrpcval('mon', "string")
            )
    );
$search_keys_tar = array(
        array(new xmlrpcval(array(new xmlrpcval("year" , "string"),
                         new xmlrpcval("=","string"),
                         new xmlrpcval("2014","string")),"array"),
        )
    );
$Tar_Report->constructor('rpc_user', 'rpc_password', 'db_name', 'xmlrpc_interface_url', $relation_tar, $required_keys_tar, $search_keys_tar);
$Tar_Report->connect_xmlrpc();
$Tar_Report->get_xml_data(0);



/**********
Simple Revenue Reports - connect, get data, generate file
**********/

$Country_Cache = array(); //to be used in Actual_vs_Target report

$file_names = array(
                      "Client_Cost-Centre",
                      "Cost-Centre_Country",
                      "Service_Client",
                      "Service_Country",
                      "Service_Category-1",
                      "Service_Category_Country", //5
                      "Service_Category-2",
                      "Client_Country", //7  -- client here
                      "Service_Operator",
                      "Operator_Country",
                      "Technology_Country" //10
                );
$Rev_Report = new Report;
$relation_rev = array(
                      "sale.by.cc.client.report",
                      "sale.by.cc.report",
                      "sale.by.service.client.report",
                      "sale.by.service.c.report",
                      "sale.by.service.report",
                      "sale.by.service.cat.cou.report",
                      "sale.by.service.cat2.report",
                      "sale.by.client.report", //7
                      "sale.by.service.op.report",
                      "sale.by.operator.report",
                      "sale.by.service.cat3.report"
                );
$required_keys_rev = array(
            array(
                       new xmlrpcval('defcol_one_id', "string"),
                       new xmlrpcval('defcol_two_id', "string"),
                       new xmlrpcval('short_name', "string"),
                       new xmlrpcval('jan', "string"),
                       new xmlrpcval('feb', "string"),
                       new xmlrpcval('mar', "string"),
                       new xmlrpcval('q1', "string"),
                       new xmlrpcval('apr', "string"),
                       new xmlrpcval('may', "string"),
                       new xmlrpcval('jun', "string"),
                       new xmlrpcval('q2', "string"),
                       new xmlrpcval('jul', "string"),
                       new xmlrpcval('aug', "string"),
                       new xmlrpcval('sep', "string"),
                       new xmlrpcval('q3', "string"),
                       new xmlrpcval('oct', "string"),
                       new xmlrpcval('nov', "string"),
                       new xmlrpcval('des', "string"),
                       new xmlrpcval('q4', "string"),
                       new xmlrpcval('total', "string")
            )
    );
$search_keys_rev = array(
            array( //for hr.vacation.report
                        new xmlrpcval(array(new xmlrpcval("id" , "string"),
                        new xmlrpcval("!=","string"),
                        new xmlrpcval("-1","string")),"array"),
            )
    );
$Rev_Report->constructor('rpc_user', 'rpc_password', 'db_name', 'xmlrpc_interface_url', $relation_rev, $required_keys_rev, $search_keys_rev);
$Rev_Report->connect_xmlrpc();
$res = array();
for ($i=0; $i < count($Rev_Report->relation_names); $i++){
    $Rev_Report->get_xml_data_reg($i);
    if (($i==0)) {
        $temp = array();
        foreach ($Rev_Report->data_arr[$i] as $key => $val) {
            $temp = $val["defcol_one_id"];
            $Rev_Report->data_arr[$i][$key]["defcol_one_id"] = $val["defcol_two_id"];
            $Rev_Report->data_arr[$i][$key]["defcol_two_id"] = $temp;
        }
        save_rev_report_excel($Rev_Report->data_arr[$i], $file_names[$i], $style_header, $style_borders, $style_normal, $style_ctitle, $style_stotal, $style_total);
    } else if ($i==1) {
        save_rev_report_excel3($Rev_Report->data_arr[$i], $file_names[$i], $Country_Cache, $style_header, $style_borders, $style_normal, $style_ctitle, $style_stotal, $style_total);
    }else if (($i==6)) {
        //category 2
        save_rev_report_excel2($Rev_Report->data_arr[$i], $file_names[$i], $style_header, $style_borders, $style_normal, $style_ctitle, $style_stotal, $style_total);
    } else if ($i == 10) {
        $temp = array();
        foreach ($Rev_Report->data_arr[$i] as $key => $val) {
            $temp[0] = 0;
            $temp[1] = strtoupper($val["defcol_one_id"]);
            $Rev_Report->data_arr[$i][$key]["defcol_one_id"] = $temp;
        }
        save_rev_report_excel5($Rev_Report->data_arr[$i], $file_names[$i], $style_header, $style_borders, $style_normal, $style_ctitle, $style_stotal, $style_total);
    }else {
    //create file each time after getting the data
        if ($i == 5) {
            var_dump($Rev_Report->data_arr[$i]);
        }

        if ($i == 7) {

            foreach ($Rev_Report->data_arr[$i] as $key => $val) {
                if ($val["short_name"]) {
                    $temp[0] = 0;
                    $temp[1] = $val["short_name"];
                    $Rev_Report->data_arr[$i][$key]["defcol_one_id"] = $temp;
                }
            }

            save_rev_report_excel4($Rev_Report->data_arr[$i], $Tar_Report->data_arr[0], 'Client_Country_Target', $style_header, $style_borders, $style_normal, $style_ctitle, $style_stotal, $style_total);
        }
        if ($i == 9) {

            foreach ($Rev_Report->data_arr[$i] as $key => $val) {
                if ($val["short_name"]) {
                    $temp[0] = 0;
                    $temp[1] = $val["short_name"];
                    $Rev_Report->data_arr[$i][$key]["defcol_one_id"] = $temp;
                }
            }
        }
        save_rev_report_excel($Rev_Report->data_arr[$i], $file_names[$i], $style_header, $style_borders, $style_normal, $style_ctitle, $style_stotal, $style_total);
    }
}
//$res = array_replace_recursive($Rev_Report->data_arr[0], $CA_Report->data_arr[1], $CA_Report->data_arr[2], $CA_Report->data_arr[3]);


/***** OLD TARGET REPORT. FUNCTION INITIATED HERE, DATA EXTRACTED AT THE TOP *****/

save_tar_report_excel($Tar_Report->data_arr[0], $file_names1[0], $Country_Cache, $style_header, $style_borders, $style_normal, $style_ctitle, $style_stotal, $style_total);



/***** FUNCTIONS DEFINITIONS *****/

function save_tar_report_excel($xarr, $report_name, $Country_Cache, $style_header, $style_borders, $style_normal, $style_ctitle, $style_stotal, $style_total) {

    echo "ENTERED THE TARGET SCRIPT!";

    usort($xarr, function ($a, $b) {
        if ($a["country_id"][1] == $b["country_id"][1]) return 0;
        return ($a["country_id"][1] < $b["country_id"][1]) ? -1 : 1;
    });
        // Create new PHPExcel object
        $objPHPExcel = new PHPExcel();

        // Set document properties
        $objPHPExcel->getProperties()->setCreator("OpenERP")
        							 ->setLastModifiedBy("OpenERP")
        							 ->setTitle("AVG Revenue Report")
        							 ->setSubject("Office 2007 XLSX Test Document")
        							 ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
        							 ->setKeywords("office 2007 openxml php")
        							 ->setCategory("Test result file");

        // Set default font
        //$objPHPExcel->getDefaultStyle()->getFont()->setName('Arial')
                                                  //->setSize(10);

    $objPHPExcel->getActiveSheet()->setCellValue('A1', "Country")
                                  ->setCellValue('B1', 'Jan')
                                  ->setCellValue('C1', 'Feb')
                                  ->setCellValue('D1', 'Mar')
                                  ->setCellValue('E1', 'Apr')
                                  ->setCellValue('F1', 'May')
                                  ->setCellValue('G1', 'Jun')
                                  ->setCellValue('H1', 'Jul')
                                  ->setCellValue('I1', 'Aug')
                                  ->setCellValue('J1', 'Sep')
                                  ->setCellValue('K1', 'Nov')
                                  ->setCellValue('L1', 'Dec')
                                  ->setCellValue('M1', 'YTD Actual')
                                  ->setCellValue('N1', 'YTD Target')
                                  ->setCellValue('P1', 'Year Target')
                                  ->setCellValue('O1', 'Variance');

    $style_variance = array(
                        'font' => array(
                            'name' => 'Arial',
                            'size' => 10,
                            'bold' => true,
                        ),
                        'fill' => array(
                            'type' => PHPExcel_Style_Fill::FILL_SOLID,
                            'color' => array('argb' => 'FFDDDDDD')
                        )
                      );

    $tar_total = 0;
    $tar_cou = 0;
    $tar_yr = 0;

    $i=2;
    $j = 0;
    $mon_format = array('jan' => 1, 'feb' => 2, 'mar' => 3, 'apr' => 4, 'may' => 5, 'jun' => 6, 'jul' => 7, 'aug' => 8, 'sep' => 9,
                        'oct' => 10, 'nov' => 11, 'des' => 12);
    $today = getdate();
    foreach($Country_Cache as $k => $cval) {

        $cou = $cval['name'];
        if (isset($xarr[$j]['country_id'][1]) && ($xarr[$j]['country_id'][1] == $cou)) {
            while (isset($xarr[$j]['country_id'][1]) && ($xarr[$j]['country_id'][1] == $cou)) { //count the targets for the country $cou
                $mo = $xarr[$j]['mon'];
                $tar_cou += ($mon_format[$mo] <= $today['mon']) ? $xarr[$j]['tar_mo'] : 0;
                $tar_yr += $xarr[$j]['tar_mo'];
                $j++;
            }
            $objPHPExcel->getActiveSheet()->setCellValue('A'.$i, $cval['name'])
                                  ->setCellValue('B'.$i, num_form($cval['jan']))
                                  ->setCellValue('C'.$i, num_form($cval['feb']))
                                  ->setCellValue('D'.$i, num_form($cval['mar']))
                                  ->setCellValue('E'.$i, num_form($cval['apr']))
                                  ->setCellValue('F'.$i, num_form($cval['may']))
                                  ->setCellValue('G'.$i, num_form($cval['jun']))
                                  ->setCellValue('H'.$i, num_form($cval['jul']))
                                  ->setCellValue('I'.$i, num_form($cval['aug']))
                                  ->setCellValue('J'.$i, num_form($cval['sep']))
                                  ->setCellValue('K'.$i, num_form($cval['nov']))
                                  ->setCellValue('L'.$i, num_form($cval['des']))
                                  ->setCellValue('M'.$i, num_form($cval['total']))
                                  ->setCellValue('N'.$i, num_form($tar_cou,0,'.',','))
                                  ->setCellValue('O'.$i, ($tar_cou > 0)? round((float)($cval['total']/$tar_cou), 2) : 1)
                                  ->setCellValue('P'.$i, num_form($tar_yr));
            $objPHPExcel->getActiveSheet()->getStyle('A'.$i.':'.$objPHPExcel->getActiveSheet()->getHighestColumn().$i)
                        ->applyFromArray($style_normal);
            $i++;
            $tar_total += $tar_cou;
            $tar_cou = 0;
            $tar_yr = 0;
        } else { //display the $Country_Cache data with 0 target
            $objPHPExcel->getActiveSheet()->setCellValue('A'.$i, $cval['name'])
                                   ->setCellValue('B'.$i, num_form($cval['jan']))
                                   ->setCellValue('C'.$i, num_form($cval['feb']))
                                   ->setCellValue('D'.$i, num_form($cval['mar']))
                                   ->setCellValue('E'.$i, num_form($cval['apr']))
                                   ->setCellValue('F'.$i, num_form($cval['may']))
                                   ->setCellValue('G'.$i, num_form($cval['jun']))
                                   ->setCellValue('H'.$i, num_form($cval['jul']))
                                   ->setCellValue('I'.$i, num_form($cval['aug']))
                                   ->setCellValue('J'.$i, num_form($cval['sep']))
                                   ->setCellValue('K'.$i, num_form($cval['nov']))
                                   ->setCellValue('L'.$i, num_form($cval['des']))
                                   ->setCellValue('M'.$i, num_form($cval['total']))
                                  ->setCellValue('N'.$i, 0)
                                  ->setCellValue('O'.$i, 1)
                                  ->setCellValue('P'.$i, 0);
            $i++;

        }
        $objPHPExcel->getActiveSheet()->getStyle('A'.($i-1).':'.$objPHPExcel->getActiveSheet()->getHighestColumn().($i-1))
                        ->applyFromArray($style_normal);
    } //foreach $Country_Cache

/*    $objPHPExcel->getActiveSheet()->getStyle('O2:O'.$objPHPExcel->getActiveSheet()->getHighestRow())
                ->getNumberFormat()->setFormatCode('[Green][>=1]$#,##0.00"%";[Red][<1]$#,##0.00"%"'); */
    $objPHPExcel->getActiveSheet()->getStyle('A1:A'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_ctitle);

    $objPHPExcel->getActiveSheet()->getStyle('M1:N'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_ctitle);
    $objPHPExcel->getActiveSheet()->getStyle('P1:P'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_ctitle);

    $objPHPExcel->getActiveSheet()->getStyle('A1:'.$objPHPExcel->getActiveSheet()->getHighestColumn().'1')->applyFromArray($style_total);
    $objPHPExcel->getActiveSheet()->getStyle('A'.
            $objPHPExcel->getActiveSheet()->getHighestRow().':'.
            $objPHPExcel->getActiveSheet()->getHighestColumn().
            $objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_total);

    $objPHPExcel->getActiveSheet()->getStyle('O2:O'.$objPHPExcel->getActiveSheet()->getHighestRow())
        ->getNumberFormat()->applyFromArray(
            array(
                'code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE
            )
        );

    /*$col = 14;
    for ($row=0; $row <= $objPHPExcel->getActiveSheet()->getHighestRow(); $row++) {
        if ($objPHPExcel->getActiveSheet()->getCellByColumnAndRow($col, $row)->getValue() >= 1) {
            $objPHPExcel->getActiveSheet()->getCellByColumnAndRow($col, $row)->getFont()->getColor()->setRGB('009900');
        } else {
            $objPHPExcel->getActiveSheet()->getCellByColumnAndRow($col, $row)->getFont()->getColor()->setRGB('FF0000');
        }
    }*/

    $objPHPExcel->getActiveSheet()->getStyle(
            'A1:' .
            $objPHPExcel->getActiveSheet()->getHighestColumn().
            $objPHPExcel->getActiveSheet()->getHighestRow()
    )->applyFromArray($style_borders);

    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, "Excel2007");
    $objWriter->save("/home/webmaster/tgtoil.com/wp-content/themes/MyCuisine/xmlrpc_lib/xls_files/Revenue_".$report_name.".xlsx");
} //save xls for target report

function save_rev_report_excel($xarr, $report_name, $style_header, $style_borders, $style_normal, $style_ctitle, $style_stotal, $style_total) {

        foreach ($xarr as $key => $val) {
            $loc_arr[$key] = $val["defcol_one_id"][1];
            $cat_arr[$key] = $val["defcol_two_id"][1];
        }

        array_multisort($cat_arr, SORT_ASC, $loc_arr, SORT_ASC, $xarr);

        $objPHPExcel = new PHPExcel();

        $objPHPExcel->getProperties()->setCreator("OpenERP")
        							 ->setLastModifiedBy("OpenERP")
        							 ->setTitle("AVG Revenue Report")
        							 ->setSubject("Office 2007 XLSX Test Document")
        							 ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
        							 ->setKeywords("office 2007 openxml php")
        							 ->setCategory("Test result file");


        $objPHPExcel->getDefaultStyle()->getFont()->setName('Arial')
                                                  ->setSize(10);

    $sub_rows = array();

    $objPHPExcel->getActiveSheet()->setCellValue('A1', str_replace("_", "/", $report_name))
                                  ->setCellValue('B1', 'Jan')
                                  ->setCellValue('C1', 'Feb')
                                  ->setCellValue('D1', 'Mar')
                                  ->setCellValue('E1', 'Q1')
                                  ->setCellValue('F1', 'Apr')
                                  ->setCellValue('G1', 'May')
                                  ->setCellValue('H1', 'Jun')
                                  ->setCellValue('I1', 'Q2')
                                  ->setCellValue('J1', 'Jul')
                                  ->setCellValue('K1', 'Aug')
                                  ->setCellValue('L1', 'Sep')
                                  ->setCellValue('M1', 'Q3')
                                  ->setCellValue('N1', 'Oct')
                                  ->setCellValue('O1', 'Nov')
                                  ->setCellValue('P1', 'Dec')
                                  ->setCellValue('Q1', 'Q4')
                                  ->setCellValue('R1', 'Total')
                                  ->setCellValue('S1', '% of Country Revenue')
                                  ->setCellValue('T1', '% of Corporate Revenue');

    $cat = '';
    $cat_rev = array('jan'=>0, 'feb'=>0, 'mar'=>0, 'q1'=>0, 'apr'=>0, 'may'=>0, 'jun'=>0, 'q2'=>0,
                     'jul'=>0, 'aug'=>0, 'sep'=>0, 'q3'=>0, 'oct'=>0, 'nov'=>0, 'des'=>0, 'q4'=>0);
    $tot_rev = array('jan'=>0, 'feb'=>0, 'mar'=>0, 'q1'=>0, 'apr'=>0, 'may'=>0, 'jun'=>0, 'q2'=>0,
                     'jul'=>0, 'aug'=>0, 'sep'=>0, 'q3'=>0, 'oct'=>0, 'nov'=>0, 'des'=>0, 'q4'=>0);

    $i=2;
    $start = 2;
    foreach($xarr as $k => $xval) {

        if ($cat=='') {
            $cat = $xval['defcol_two_id'][1];
            //fill cat_rev
            foreach ($cat_rev as $id=>$sval){
                $cat_rev[$id] = $xval[$id];
            }
            $total_col = $xval['q1']+$xval['q2']+$xval['q3']+$xval['q4'];
            fill_row($objPHPExcel, $xval, $total_col, $xval['defcol_one_id'][1], $i, $style_normal);
            $i++;
        }
        else if ($cat == $xval['defcol_two_id'][1]) {
            //sum cat_rev
            foreach ($cat_rev as $id=>$sval){
                $cat_rev[$id] += $xval[$id];
            }
            //xls the line
            $total_col = $xval['q1']+$xval['q2']+$xval['q3']+$xval['q4'];
            fill_row($objPHPExcel, $xval, $total_col, $xval['defcol_one_id'][1], $i, $style_normal);
            $i++;
        }
        else {
            //sum tot_rev
            foreach ($tot_rev as $id=>$sval){
                $tot_rev[$id] += $cat_rev[$id];
            }
            //xls the cat_rev
            $total_col = $cat_rev['q1']+$cat_rev['q2']+$cat_rev['q3']+$cat_rev['q4'];
            fill_row($objPHPExcel, $cat_rev, $total_col, strtoupper($cat), $i, $style_stotal);
            //here fill the "special" percentage column
            fill_persentage($objPHPExcel, 'R', 'S', $start, $i, $style_stotal);
            $i++;
            $start = $i;
            $total_col = $xval['q1']+$xval['q2']+$xval['q3']+$xval['q4'];
            fill_row($objPHPExcel, $xval, $total_col, $xval['defcol_one_id'][1], $i, $style_normal);
            $i++;
            $cat = $xval['defcol_two_id'][1];
            //new cat_rev
            foreach ($cat_rev as $id=>$sval){
                $cat_rev[$id] = $xval[$id];
            }
        }
    } //foreach $xarr
    $last_elem = end($xarr);
    foreach ($tot_rev as $id=>$sval){
        $tot_rev[$id] += $cat_rev[$id];
    }
    //xls cat and grand total
    $total_col = $cat_rev['q1']+$cat_rev['q2']+$cat_rev['q3']+$cat_rev['q4'];
    fill_row($objPHPExcel, $cat_rev, $total_col, strtoupper($cat), $i, $style_stotal);
    fill_persentage($objPHPExcel, 'R', 'S', $start, $i, $style_stotal);
    $i++;
    $total_col = $tot_rev['q1']+$tot_rev['q2']+$tot_rev['q3']+$tot_rev['q4'];
    fill_row($objPHPExcel, $tot_rev, $total_col, 'TGT CORPORATE', $i, $style_total);

    fill_persentage($objPHPExcel, 'R', 'T', 2, $i, $style_total);

    $objPHPExcel->getActiveSheet()->getStyle('A1:A'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_ctitle);

    $objPHPExcel->getActiveSheet()->getStyle('E1:E'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_stotal);
    $objPHPExcel->getActiveSheet()->getStyle('I1:I'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_stotal);
    $objPHPExcel->getActiveSheet()->getStyle('M1:M'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_stotal);
    $objPHPExcel->getActiveSheet()->getStyle('Q1:Q'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_stotal);
    $objPHPExcel->getActiveSheet()->getStyle('R1:R'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_total);

    $objPHPExcel->getActiveSheet()->getStyle('A1:'.$objPHPExcel->getActiveSheet()->getHighestColumn().'1')->applyFromArray($style_total);

    $objPHPExcel->getActiveSheet()->getStyle(
        'B2:' .
        $objPHPExcel->getActiveSheet()->getHighestColumn().
        $objPHPExcel->getActiveSheet()->getHighestRow()
    )->getNumberFormat()->setFormatCode('#,##0_-');

    /*$objPHPExcel->getActiveSheet()->getStyle('S2:T'.$objPHPExcel->getActiveSheet()->getHighestRow())
        ->getNumberFormat()->applyFromArray(
            array(
                'code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE
            )
        );*/

    $objPHPExcel->getActiveSheet()->getStyle(
        'A1:' .
        $objPHPExcel->getActiveSheet()->getHighestColumn().
        $objPHPExcel->getActiveSheet()->getHighestRow()
    )->applyFromArray($style_borders);

    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, "Excel2007");
    $objWriter->save("/home/webmaster/tgtoil.com/wp-content/themes/MyCuisine/xmlrpc_lib/xls_files/Revenue_".$report_name.".xlsx");
} //standard save xls

function save_rev_report_excel2($xarr, $report_name, $style_header, $style_borders, $style_normal, $style_ctitle, $style_stotal, $style_total) {

        foreach ($xarr as $key => $val) {
            $loc_arr[$key] = $val["defcol_one_id"][1];
            $cat_arr[$key] = $val["defcol_two_id"];
        }

        array_multisort($cat_arr, SORT_ASC, $loc_arr, SORT_ASC, $xarr);

        // Create new PHPExcel object
        $objPHPExcel = new PHPExcel();

        // Set document properties
        $objPHPExcel->getProperties()->setCreator("OpenERP")
        							 ->setLastModifiedBy("OpenERP")
        							 ->setTitle("AVG Revenue Report")
        							 ->setSubject("Office 2007 XLSX Test Document")
        							 ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
        							 ->setKeywords("office 2007 openxml php")
        							 ->setCategory("Test result file");

        // Set default font
        $objPHPExcel->getDefaultStyle()->getFont()->setName('Arial')
                                                  ->setSize(10);

    $objPHPExcel->getActiveSheet()->setCellValue('A1', str_replace("_", "/", $report_name))
                                  ->setCellValue('B1', 'Jan')
                                  ->setCellValue('C1', 'Feb')
                                  ->setCellValue('D1', 'Mar')
                                  ->setCellValue('E1', 'Q1')
                                  ->setCellValue('F1', 'Apr')
                                  ->setCellValue('G1', 'May')
                                  ->setCellValue('H1', 'Jun')
                                  ->setCellValue('I1', 'Q2')
                                  ->setCellValue('J1', 'Jul')
                                  ->setCellValue('K1', 'Aug')
                                  ->setCellValue('L1', 'Sep')
                                  ->setCellValue('M1', 'Q3')
                                  ->setCellValue('N1', 'Oct')
                                  ->setCellValue('O1', 'Nov')
                                  ->setCellValue('P1', 'Dec')
                                  ->setCellValue('Q1', 'Q4')
                                  ->setCellValue('R1', 'Total')
                                  ->setCellValue('S1', '% of Country Revenue')
                                  ->setCellValue('T1', '% of Corporate Revenue');

    $cat = '';
    $cat_rev = array('jan'=>0, 'feb'=>0, 'mar'=>0, 'q1'=>0, 'apr'=>0, 'may'=>0, 'jun'=>0, 'q2'=>0,
                     'jul'=>0, 'aug'=>0, 'sep'=>0, 'q3'=>0, 'oct'=>0, 'nov'=>0, 'des'=>0, 'q4'=>0);
    $tot_rev = array('jan'=>0, 'feb'=>0, 'mar'=>0, 'q1'=>0, 'apr'=>0, 'may'=>0, 'jun'=>0, 'q2'=>0,
                     'jul'=>0, 'aug'=>0, 'sep'=>0, 'q3'=>0, 'oct'=>0, 'nov'=>0, 'des'=>0, 'q4'=>0);

    $i=2;
    $start = 2;
    foreach($xarr as $k => $xval) {

        if ($cat=='') {
            $cat = $xval['defcol_two_id'];
            //fill cat_rev
            foreach ($cat_rev as $id=>$sval){
                $cat_rev[$id] = $xval[$id];
            }
            $total_col = $xval['q1']+$xval['q2']+$xval['q3']+$xval['q4'];
            fill_row($objPHPExcel, $xval, $total_col, $xval['defcol_one_id'][1], $i, $style_normal);
            $i++;
        }
        else if ($cat == $xval['defcol_two_id']) {
            //sum cat_rev
            foreach ($cat_rev as $id=>$sval){
                $cat_rev[$id] += $xval[$id];
            }
            //xls the line
            $total_col = $xval['q1']+$xval['q2']+$xval['q3']+$xval['q4'];
            fill_row($objPHPExcel, $xval, $total_col, $xval['defcol_one_id'][1], $i, $style_normal);
            $i++;
        }
        else {
            //sum tot_rev
            foreach ($tot_rev as $id=>$sval){
                //echo "id = ".$id.", sval = ".$sval.", tot_rev[id] = ".$tot_rev[$id];
                $tot_rev[$id] += $cat_rev[$id];
            }
            //xls the cat_rev
            $total_col = $cat_rev['q1']+$cat_rev['q2']+$cat_rev['q3']+$cat_rev['q4'];
            fill_row($objPHPExcel, $cat_rev, $total_col, strtoupper($cat), $i, $style_stotal);
            fill_persentage($objPHPExcel, 'R', 'S', $start, $i, $style_stotal);
            $i++;
            $start = $i;
            $total_col = $xval['q1']+$xval['q2']+$xval['q3']+$xval['q4'];
            fill_row($objPHPExcel, $xval, $total_col, $xval['defcol_one_id'][1], $i, $style_normal);
            $i++;
            $cat = $xval['defcol_two_id'];
            //new cat_rev
            foreach ($cat_rev as $id=>$sval){
                $cat_rev[$id] = $xval[$id];
            }
        }
    } //foreach $xarr
    $last_elem = end($xarr);
    foreach ($tot_rev as $id=>$sval){
        $tot_rev[$id] += $cat_rev[$id];
    }
    //xls cat and grand total
    $total_col = $cat_rev['q1']+$cat_rev['q2']+$cat_rev['q3']+$cat_rev['q4'];
    fill_row($objPHPExcel, $cat_rev, $total_col, strtoupper($cat), $i, $style_stotal);
    fill_persentage($objPHPExcel, 'R', 'S', $start, $i, $style_stotal);
    $i++;
    $total_col = $tot_rev['q1']+$tot_rev['q2']+$tot_rev['q3']+$tot_rev['q4'];
    fill_row($objPHPExcel, $tot_rev, $total_col, 'TOTAL', $i, $style_total);

    fill_persentage($objPHPExcel, 'R', 'T', 2, $i, $style_total);

    $objPHPExcel->getActiveSheet()->getStyle('A1:A'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_ctitle);

    $objPHPExcel->getActiveSheet()->getStyle('E1:E'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_stotal);
    $objPHPExcel->getActiveSheet()->getStyle('I1:I'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_stotal);
    $objPHPExcel->getActiveSheet()->getStyle('M1:M'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_stotal);
    $objPHPExcel->getActiveSheet()->getStyle('Q1:Q'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_stotal);
    $objPHPExcel->getActiveSheet()->getStyle('R1:R'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_total);

    $objPHPExcel->getActiveSheet()->getStyle('A1:'.$objPHPExcel->getActiveSheet()->getHighestColumn().'1')->applyFromArray($style_total);

    $objPHPExcel->getActiveSheet()->getStyle(
        'B2:' .
        $objPHPExcel->getActiveSheet()->getHighestColumn().
        $objPHPExcel->getActiveSheet()->getHighestRow()
    )->getNumberFormat()->setFormatCode('#,##0_-');

    /*$objPHPExcel->getActiveSheet()->getStyle('S2:T'.$objPHPExcel->getActiveSheet()->getHighestRow())
        ->getNumberFormat()->applyFromArray(
            array(
                'code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00
            )
        );*/

    $objPHPExcel->getActiveSheet()->getStyle(
        'A1:' .
        $objPHPExcel->getActiveSheet()->getHighestColumn().
        $objPHPExcel->getActiveSheet()->getHighestRow()
    )->applyFromArray($style_borders);

    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, "Excel2007");
    $objWriter->save("/home/webmaster/tgtoil.com/wp-content/themes/MyCuisine/xmlrpc_lib/xls_files/Revenue_".$report_name.".xlsx");
} // save_xls for Category 2

function save_rev_report_excel3($xarr, $report_name, &$Country_Cache, $style_header, $style_borders, $style_normal, $style_ctitle, $style_stotal, $style_total) {

        foreach ($xarr as $key => $val) {
            $loc_arr[$key] = $val["defcol_one_id"][1];
            $cat_arr[$key] = $val["defcol_two_id"][1];
        }

        array_multisort($cat_arr, SORT_ASC, $loc_arr, SORT_ASC, $xarr);

        // Create new PHPExcel object
        $objPHPExcel = new PHPExcel();

        // Set document properties
        $objPHPExcel->getProperties()->setCreator("OpenERP")
        							 ->setLastModifiedBy("OpenERP")
        							 ->setTitle("AVG Revenue Report")
        							 ->setSubject("Office 2007 XLSX Test Document")
        							 ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
        							 ->setKeywords("office 2007 openxml php")
        							 ->setCategory("Test result file");

        // Set default font
        $objPHPExcel->getDefaultStyle()->getFont()->setName('Arial')
                                                  ->setSize(10);

    $objPHPExcel->getActiveSheet()->setCellValue('A1', str_replace("_", "/", $report_name))
                                  ->setCellValue('B1', 'Jan')
                                  ->setCellValue('C1', 'Feb')
                                  ->setCellValue('D1', 'Mar')
                                  ->setCellValue('E1', 'Q1')
                                  ->setCellValue('F1', 'Apr')
                                  ->setCellValue('G1', 'May')
                                  ->setCellValue('H1', 'Jun')
                                  ->setCellValue('I1', 'Q2')
                                  ->setCellValue('J1', 'Jul')
                                  ->setCellValue('K1', 'Aug')
                                  ->setCellValue('L1', 'Sep')
                                  ->setCellValue('M1', 'Q3')
                                  ->setCellValue('N1', 'Oct')
                                  ->setCellValue('O1', 'Nov')
                                  ->setCellValue('P1', 'Dec')
                                  ->setCellValue('Q1', 'Q4')
                                  ->setCellValue('R1', 'Total')
                                  ->setCellValue('S1', '% of Country Revenue')
                                  ->setCellValue('T1', '% of Corporate Revenue');

    $cat = '';
    $cat_rev = array('jan'=>0, 'feb'=>0, 'mar'=>0, 'q1'=>0, 'apr'=>0, 'may'=>0, 'jun'=>0, 'q2'=>0,
                     'jul'=>0, 'aug'=>0, 'sep'=>0, 'q3'=>0, 'oct'=>0, 'nov'=>0, 'des'=>0, 'q4'=>0);
    $tot_rev = array('jan'=>0, 'feb'=>0, 'mar'=>0, 'q1'=>0, 'apr'=>0, 'may'=>0, 'jun'=>0, 'q2'=>0,
                     'jul'=>0, 'aug'=>0, 'sep'=>0, 'q3'=>0, 'oct'=>0, 'nov'=>0, 'des'=>0, 'q4'=>0);

    $i=2;
    $e=0;
    $start = 2;
    foreach($xarr as $k => $xval) {
        if ($cat=='') {
            $cat = $xval['defcol_two_id'][1];
            //fill cat_rev
            foreach ($cat_rev as $id=>$sval){
                $cat_rev[$id] = $xval[$id];
            }
            $total_col = $xval['q1']+$xval['q2']+$xval['q3']+$xval['q4'];
            fill_row($objPHPExcel, $xval, $total_col, $xval['defcol_one_id'][1], $i, $style_normal);
            $i++;
        }
        else if ($cat == $xval['defcol_two_id'][1]) {
            //sum cat_rev
            foreach ($cat_rev as $id=>$sval){
                $cat_rev[$id] += $xval[$id];
            }
            //xls the line
            $total_col = $xval['q1']+$xval['q2']+$xval['q3']+$xval['q4'];
            fill_row($objPHPExcel, $xval, $total_col, $xval['defcol_one_id'][1], $i, $style_normal);
            $i++;
        }
        else {
            //sum tot_rev
            $Country_Cache[$e] = array();
            $Country_Cache[$e] = array('jan'=>0, 'feb'=>0, 'mar'=>0, 'apr'=>0, 'may'=>0, 'jun'=>0, 'jul'=>0, 'aug'=>0, 'sep'=>0, 'oct'=>0, 'nov'=>0, 'des'=>0);
            foreach ($Country_Cache[$e] as $id=>$sv) {
                $Country_Cache[$e][$id] = $cat_rev[$id];
            }
            $Country_Cache[$e]['name'] = $cat;
            $Country_Cache[$e]['total'] = $cat_rev['q1'] + $cat_rev['q2'] + $cat_rev['q3'] + $cat_rev['q4'];
            $e++;
            foreach ($tot_rev as $id=>$sval){
                //echo "id = ".$id.", sval = ".$sval.", tot_rev[id] = ".$tot_rev[$id];
                $tot_rev[$id] += $cat_rev[$id];
            }
            //xls the cat_rev
            $total_col = $cat_rev['q1']+$cat_rev['q2']+$cat_rev['q3']+$cat_rev['q4'];
            fill_row($objPHPExcel, $cat_rev, $total_col, strtoupper($cat), $i, $style_stotal);
            fill_persentage($objPHPExcel, 'R', 'S', $start, $i, $style_stotal);
            $i++;
            $start = $i;
            $total_col = $xval['q1']+$xval['q2']+$xval['q3']+$xval['q4'];
            fill_row($objPHPExcel, $xval, $total_col, $xval['defcol_one_id'][1], $i, $style_normal);
            $i++;
            $cat = $xval['defcol_two_id'][1];
            //new cat_rev
            foreach ($cat_rev as $id=>$sval){
                $cat_rev[$id] = $xval[$id];
            }
        }
    } //foreach $xarr
    $last_elem = end($xarr);
    foreach ($tot_rev as $id=>$sval){
        $tot_rev[$id] += $cat_rev[$id];
    }
    //xls cat and grand total
    $Country_Cache[$e] = array();
    $Country_Cache[$e] = array('jan'=>0, 'feb'=>0, 'mar'=>0, 'apr'=>0, 'may'=>0, 'jun'=>0, 'jul'=>0, 'aug'=>0, 'sep'=>0, 'oct'=>0, 'nov'=>0, 'des'=>0);
    foreach ($Country_Cache[$e] as $id=>$sv) {
        $Country_Cache[$e][$id] = $cat_rev[$id];
    }
    $Country_Cache[$e]['name'] = $cat;
    $Country_Cache[$e]['total'] = $cat_rev['q1'] + $cat_rev['q2'] + $cat_rev['q3'] + $cat_rev['q4'];
    $e++;
    $Country_Cache[$e] = array();
    $Country_Cache[$e] = array('jan'=>0, 'feb'=>0, 'mar'=>0, 'apr'=>0, 'may'=>0, 'jun'=>0, 'jul'=>0, 'aug'=>0, 'sep'=>0, 'oct'=>0, 'nov'=>0, 'des'=>0);
    foreach ($Country_Cache[$e] as $id=>$sv) {
        $Country_Cache[$e][$id] = $tot_rev[$id];
    }
    $Country_Cache[$e]['name'] = "TOTAL";
    $Country_Cache[$e]['total'] = $tot_rev['q1'] + $tot_rev['q2'] + $tot_rev['q3'] + $tot_rev['q4'];
    $e++;
    $total_col = $cat_rev['q1']+$cat_rev['q2']+$cat_rev['q3']+$cat_rev['q4'];
    fill_row($objPHPExcel, $cat_rev, $total_col, strtoupper($cat), $i, $style_stotal);
    fill_persentage($objPHPExcel, 'R', 'S', $start, $i, $style_stotal);
    $i++;
    $total_col = $tot_rev['q1']+$tot_rev['q2']+$tot_rev['q3']+$tot_rev['q4'];

    $objPHPExcel->getActiveSheet()->getStyle('A1:A'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_ctitle);
    fill_row($objPHPExcel, $tot_rev, $total_col, 'TGT CORPORATE', $i, $style_total);

    fill_persentage($objPHPExcel, 'R', 'T', 2, $i, $style_total);

    $objPHPExcel->getActiveSheet()->getStyle('E1:E'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_stotal);
    $objPHPExcel->getActiveSheet()->getStyle('I1:I'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_stotal);
    $objPHPExcel->getActiveSheet()->getStyle('M1:M'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_stotal);
    $objPHPExcel->getActiveSheet()->getStyle('Q1:Q'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_stotal);
    $objPHPExcel->getActiveSheet()->getStyle('R1:R'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_total);

    $objPHPExcel->getActiveSheet()->getStyle('A1:'.$objPHPExcel->getActiveSheet()->getHighestColumn().'1')->applyFromArray($style_total);

    $objPHPExcel->getActiveSheet()->getStyle(
        'B2:' .
        $objPHPExcel->getActiveSheet()->getHighestColumn().
        $objPHPExcel->getActiveSheet()->getHighestRow()
    )->getNumberFormat()->setFormatCode('#,##0_-');

    /*$objPHPExcel->getActiveSheet()->getStyle('S2:T'.$objPHPExcel->getActiveSheet()->getHighestRow())
        ->getNumberFormat()->applyFromArray(
            array(
                'code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00
            )
        );*/

    $objPHPExcel->getActiveSheet()->getStyle(
        'A1:' .
        $objPHPExcel->getActiveSheet()->getHighestColumn().
        $objPHPExcel->getActiveSheet()->getHighestRow()
    )->applyFromArray($style_borders);

    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, "Excel2007");
    $objWriter->save("/home/webmaster/tgtoil.com/wp-content/themes/MyCuisine/xmlrpc_lib/xls_files/Revenue_".$report_name.".xlsx");
    //var_dump($Country_Cache);
} //CC/Country + fill the Country_Cache



function save_rev_report_excel4($rev_arr, $tar_arr, $report_name, $style_header, $style_borders, $style_normal, $style_ctitle, $style_stotal, $style_total)
{
        foreach ($rev_arr as $key => $val) {
            $loc_arr[$key] = $val["defcol_one_id"][1];
            $cat_arr[$key] = $val["defcol_two_id"][1];
        }

        array_multisort($cat_arr, SORT_ASC, $loc_arr, SORT_ASC, $rev_arr);

    usort($tar_arr, function ($a, $b) {
        if ($a["country_id"][1] == $b["country_id"][1]) return 0;
        return ($a["country_id"][1] < $b["country_id"][1]) ? -1 : 1;
    });

        $objPHPExcel = new PHPExcel();

        $objPHPExcel->getProperties()->setCreator("OpenERP")
        							 ->setLastModifiedBy("OpenERP")
        							 ->setTitle("AVG Revenue Report")
        							 ->setSubject("Office 2007 XLSX Test Document")
        							 ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
        							 ->setKeywords("office 2007 openxml php")
        							 ->setCategory("Test result file");


        $objPHPExcel->getDefaultStyle()->getFont()->setName('Arial')
                                                  ->setSize(10);

    $sub_rows = array();

    $objPHPExcel->getActiveSheet()->setCellValue('A1', str_replace("_", "/", $report_name))
                                  ->setCellValue('B1', 'Jan')
                                  ->setCellValue('C1', 'Feb')
                                  ->setCellValue('D1', 'Mar')
                                  ->setCellValue('E1', 'Q1')
                                  ->setCellValue('F1', 'Apr')
                                  ->setCellValue('G1', 'May')
                                  ->setCellValue('H1', 'Jun')
                                  ->setCellValue('I1', 'Q2')
                                  ->setCellValue('J1', 'Jul')
                                  ->setCellValue('K1', 'Aug')
                                  ->setCellValue('L1', 'Sep')
                                  ->setCellValue('M1', 'Q3')
                                  ->setCellValue('N1', 'Oct')
                                  ->setCellValue('O1', 'Nov')
                                  ->setCellValue('P1', 'Dec')
                                  ->setCellValue('Q1', 'Q4')
                                  ->setCellValue('R1', 'YTD Actual')
                                  ->setCellValue('S1', 'YTD Target')
                                  ->setCellValue('U1', 'Variance')
                                  ->setCellValue('T1', 'Annual Target');

    $cat = '';
    $cat_rev = array('jan'=>0, 'feb'=>0, 'mar'=>0, 'q1'=>0, 'apr'=>0, 'may'=>0, 'jun'=>0, 'q2'=>0,
                     'jul'=>0, 'aug'=>0, 'sep'=>0, 'q3'=>0, 'oct'=>0, 'nov'=>0, 'des'=>0, 'q4'=>0);
    $tot_rev = array('jan'=>0, 'feb'=>0, 'mar'=>0, 'q1'=>0, 'apr'=>0, 'may'=>0, 'jun'=>0, 'q2'=>0,
                     'jul'=>0, 'aug'=>0, 'sep'=>0, 'q3'=>0, 'oct'=>0, 'nov'=>0, 'des'=>0, 'q4'=>0);

    $tar_total = 0; //total ytd
    $tar_atotal = 0; //annual total
    $tar_cou = 0;
    $tar_yr = 0;
    $j = 0;
    $mon_format = array('jan' => 1, 'feb' => 2, 'mar' => 3, 'apr' => 4, 'may' => 5, 'jun' => 6, 'jul' => 7, 'aug' => 8, 'sep' => 9,
                        'oct' => 10, 'nov' => 11, 'des' => 12);
    $today = getdate();

    $t = 0;
    $i=2;
    $start = 2;
    $tot_cou = 0;
    foreach($rev_arr as $k => $xval) {
        if ($cat=='') { //cat - country
            $cat = $xval['defcol_two_id'][1];
            //check if the rev_countryid == tar_countryid
            foreach ($cat_rev as $id=>$sval){
                $cat_rev[$id] = $xval[$id];
            }
            $total_col = $xval['q1']+$xval['q2']+$xval['q3']+$xval['q4'];
            fill_row($objPHPExcel, $xval, $total_col, $xval['defcol_one_id'][1], $i, $style_normal);
            $tot_cou = $xval['total'];
            $i++;
        }
        else if ($cat == $xval['defcol_two_id'][1]) {
            foreach ($cat_rev as $id=>$sval){
                $cat_rev[$id] += $xval[$id];
            }
            $total_col = $xval['q1']+$xval['q2']+$xval['q3']+$xval['q4'];
            fill_row($objPHPExcel, $xval, $total_col, $xval['defcol_one_id'][1], $i, $style_normal);
            $tot_cou = $xval['total'];
            $i++;
        }
        else {
            foreach ($tot_rev as $id=>$sval){
                $tot_rev[$id] += $cat_rev[$id];
            }
            $total_col = $cat_rev['q1']+$cat_rev['q2']+$cat_rev['q3']+$cat_rev['q4'];
            fill_row($objPHPExcel, $cat_rev, $total_col, strtoupper($cat), $i, $style_stotal);
            //add target + variance cells
            if (isset($tar_arr[$j]['country_id'][1]) && ($cat == $tar_arr[$j]['country_id'][1])) {
                while (isset($tar_arr[$j]['country_id'][1]) && ($cat == $tar_arr[$j]['country_id'][1])) { //count the targets for the country $cou
                    $mo = $tar_arr[$j]['mon'];
                    $tar_cou += (($mo != 'year_total') && ($mon_format[$mo] <= $today['mon'])) ? $tar_arr[$j]['tar_mo'] : 0;
                    if ($mo == 'year_total') $tar_yr = $tar_arr[$j]['tar_mo'];
                    $j++;
                }
                $tar_total += $tar_cou;
                $tar_atotal += $tar_yr;
                //add cells to the xls
                $objPHPExcel->getActiveSheet()->setCellValue('S'.$i, $tar_cou)
                                                  ->setCellValue('U'.$i, ($tar_yr != 0)? round((float)($tot_cou/$tar_yr), 2) : 1)
                                                  ->setCellValue('T'.$i, $tar_yr);
                //apply style
                $objPHPExcel->getActiveSheet()->getStyle('S'.$i.':U'.$i)->applyFromArray($style_stotal);
                $tar_yr=0;
                $tar_cou=0;
            } else {
                //add 0 target to the xls
                $objPHPExcel->getActiveSheet()->setCellValue('S'.$i, $tar_cou)
                                                  ->setCellValue('U'.$i, ($tar_yr != 0)? round((float)($tot_cou/$tar_yr), 2) : 1)
                                                  ->setCellValue('T'.$i, $tar_yr);
                $objPHPExcel->getActiveSheet()->getStyle('S'.$i.':U'.$i)->applyFromArray($style_stotal);
            }
            $i++;
            $total_col = $xval['q1']+$xval['q2']+$xval['q3']+$xval['q4'];
            fill_row($objPHPExcel, $xval, $total_col, $xval['defcol_one_id'][1], $i, $style_normal);
            $i++;
            $cat = $xval['defcol_two_id'][1];
            foreach ($cat_rev as $id=>$sval){
                $cat_rev[$id] = $xval[$id];
            }
            $tot_cou = $xval['total'];
        }

    } //foreach $xarr
    $last_elem = end($rev_arr);
    foreach ($tot_rev as $id=>$sval){
        $tot_rev[$id] += $cat_rev[$id];
    }
    //xls cat and grand total
    $total_col = $cat_rev['q1']+$cat_rev['q2']+$cat_rev['q3']+$cat_rev['q4'];
    fill_row($objPHPExcel, $cat_rev, $total_col, strtoupper($cat), $i, $style_stotal);
    if ($cat == $tar_arr[$j]['name']) {
                while ($cat == $tar_arr[$j]['name']) { //count the targets for the country $cou
                    $mo = $tar_arr[$j]['mon'];
                    $tar_cou += ($mon_format[$mo] <= $today['mon']) ? $tar_arr[$j]['tar_mo'] : 0;
                    if ($mo == 'year_total') $tar_yr = $tar_arr[$j]['tar_mo'];
                    $j++;
                }
                $tar_total += $tar_cou;
                $tar_atotal += $tar_yr;
                //add cells to the xls
                $objPHPExcel->getActiveSheet()->setCellValue('S'.$i, $tar_cou)
                                                  ->setCellValue('U'.$i, ($tar_yr != 0)? round((float)($tot_cou/$tar_yr), 2) : 1)
                                                  ->setCellValue('T'.$i, $tar_yr);
                $objPHPExcel->getActiveSheet()->getStyle('S'.$i.':U'.$i)->applyFromArray($style_stotal);
                $tar_yr=0;
                $tar_cou=0;
    } else {
                //add 0 target to the xls
                $objPHPExcel->getActiveSheet()->setCellValue('S'.$i, $tar_cou)
                                                  ->setCellValue('U'.$i, ($tar_yr != 0)? round((float)($tot_cou/$tar_yr), 2) : 1)
                                                  ->setCellValue('T'.$i, $tar_yr);
                $objPHPExcel->getActiveSheet()->getStyle('S'.$i.':U'.$i)->applyFromArray($style_stotal);
    }
    $i++;
    $total_col = $tot_rev['q1']+$tot_rev['q2']+$tot_rev['q3']+$tot_rev['q4'];
    fill_row($objPHPExcel, $tot_rev, $total_col, 'TGT CORPORATE', $i, $style_total);
    $objPHPExcel->getActiveSheet()->getStyle('S'.$i.':U'.$i)->applyFromArray($style_stotal);
    echo "THIS IS TOTAL = ".$xval['total'].", THIS IS TARGET COUNTRY TOTAL = ".$tar_cou;
    $objPHPExcel->getActiveSheet()->setCellValue('S'.$i, $tar_total)
                                                  ->setCellValue('U'.$i, ($tar_atotal != 0)? round((float)($total_col/$tar_atotal), 2) : 1)
                                                  ->setCellValue('T'.$i, $tar_atotal);

    $objPHPExcel->getActiveSheet()->getStyle('A1:A'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_ctitle);

    $objPHPExcel->getActiveSheet()->getStyle('E1:E'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_stotal);
    $objPHPExcel->getActiveSheet()->getStyle('I1:I'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_stotal);
    $objPHPExcel->getActiveSheet()->getStyle('M1:M'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_stotal);
    $objPHPExcel->getActiveSheet()->getStyle('Q1:Q'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_stotal);
    $objPHPExcel->getActiveSheet()->getStyle('R1:R'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_total);

    $objPHPExcel->getActiveSheet()->getStyle('A1:'.$objPHPExcel->getActiveSheet()->getHighestColumn().'1')->applyFromArray($style_total);

    $objPHPExcel->getActiveSheet()->getStyle(
        'B2:' .
        $objPHPExcel->getActiveSheet()->getHighestColumn().
        $objPHPExcel->getActiveSheet()->getHighestRow()
    )->getNumberFormat()->setFormatCode('#,##0_-');

    $objPHPExcel->getActiveSheet()->getStyle('U2:U'.$objPHPExcel->getActiveSheet()->getHighestRow())
        ->getNumberFormat()->applyFromArray(
            array(
                'code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE
            )
        );

    $objPHPExcel->getActiveSheet()->getStyle(
        'A1:' .
        $objPHPExcel->getActiveSheet()->getHighestColumn().
        $objPHPExcel->getActiveSheet()->getHighestRow()
    )->applyFromArray($style_borders);

    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, "Excel2007");
    $objWriter->save("/home/webmaster/tgtoil.com/wp-content/themes/MyCuisine/xmlrpc_lib/xls_files/Revenue_".$report_name.".xlsx");
}


function save_rev_report_excel5($xarr, $report_name, $style_header, $style_borders, $style_normal, $style_ctitle, $style_stotal, $style_total) {

        foreach ($xarr as $key => $val) {
            $loc_arr[$key] = $val["defcol_one_id"][1];
            $cat_arr[$key] = $val["defcol_two_id"][1];
        }

        array_multisort($cat_arr, SORT_ASC, $loc_arr, SORT_ASC, $xarr);

        $objPHPExcel = new PHPExcel();

        $objPHPExcel->getProperties()->setCreator("OpenERP")
        							 ->setLastModifiedBy("OpenERP")
        							 ->setTitle("AVG Revenue Report")
        							 ->setSubject("Office 2007 XLSX Test Document")
        							 ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
        							 ->setKeywords("office 2007 openxml php")
        							 ->setCategory("Test result file");


        $objPHPExcel->getDefaultStyle()->getFont()->setName('Arial')
                                                  ->setSize(10);

    $sub_rows = array();

    $objPHPExcel->getActiveSheet()->setCellValue('A1', str_replace("_", "/", $report_name))
                                  ->setCellValue('B1', 'Jan')
                                  ->setCellValue('C1', 'Feb')
                                  ->setCellValue('D1', 'Mar')
                                  ->setCellValue('E1', 'Q1')
                                  ->setCellValue('F1', 'Apr')
                                  ->setCellValue('G1', 'May')
                                  ->setCellValue('H1', 'Jun')
                                  ->setCellValue('I1', 'Q2')
                                  ->setCellValue('J1', 'Jul')
                                  ->setCellValue('K1', 'Aug')
                                  ->setCellValue('L1', 'Sep')
                                  ->setCellValue('M1', 'Q3')
                                  ->setCellValue('N1', 'Oct')
                                  ->setCellValue('O1', 'Nov')
                                  ->setCellValue('P1', 'Dec')
                                  ->setCellValue('Q1', 'Q4')
                                  ->setCellValue('R1', 'Total')
                                  ->setCellValue('S1', '% of Country Revenue')
                                  ->setCellValue('T1', '% of Corporate Revenue');

    $cat = '';
    $cat_rev = array('jan'=>0, 'feb'=>0, 'mar'=>0, 'q1'=>0, 'apr'=>0, 'may'=>0, 'jun'=>0, 'q2'=>0,
                     'jul'=>0, 'aug'=>0, 'sep'=>0, 'q3'=>0, 'oct'=>0, 'nov'=>0, 'des'=>0, 'q4'=>0);
    $tot_rev = array('jan'=>0, 'feb'=>0, 'mar'=>0, 'q1'=>0, 'apr'=>0, 'may'=>0, 'jun'=>0, 'q2'=>0,
                     'jul'=>0, 'aug'=>0, 'sep'=>0, 'q3'=>0, 'oct'=>0, 'nov'=>0, 'des'=>0, 'q4'=>0);
    $corp = array('hpt' => array('jan'=>0, 'feb'=>0, 'mar'=>0, 'q1'=>0, 'apr'=>0, 'may'=>0, 'jun'=>0, 'q2'=>0,
                                 'jul'=>0, 'aug'=>0, 'sep'=>0, 'q3'=>0, 'oct'=>0, 'nov'=>0, 'des'=>0, 'q4'=>0),
                  'mid' => array('jan'=>0, 'feb'=>0, 'mar'=>0, 'q1'=>0, 'apr'=>0, 'may'=>0, 'jun'=>0, 'q2'=>0,
                                 'jul'=>0, 'aug'=>0, 'sep'=>0, 'q3'=>0, 'oct'=>0, 'nov'=>0, 'des'=>0, 'q4'=>0),
                  'other' => array('jan'=>0, 'feb'=>0, 'mar'=>0, 'q1'=>0, 'apr'=>0, 'may'=>0, 'jun'=>0, 'q2'=>0,
                                   'jul'=>0, 'aug'=>0, 'sep'=>0, 'q3'=>0, 'oct'=>0, 'nov'=>0, 'des'=>0, 'q4'=>0),
                  'snl' => array('jan'=>0, 'feb'=>0, 'mar'=>0, 'q1'=>0, 'apr'=>0, 'may'=>0, 'jun'=>0, 'q2'=>0,
                                 'jul'=>0, 'aug'=>0, 'sep'=>0, 'q3'=>0, 'oct'=>0, 'nov'=>0, 'des'=>0, 'q4'=>0),
                  );
    $i=2;
    $start = 2;
    foreach($xarr as $k => $xval) {
        if ($cat=='') {
            $cat = $xval['defcol_two_id'][1];
            //fill cat_rev
            foreach ($cat_rev as $id=>$sval){
                $cat_rev[$id] = $xval[$id];
            }
            foreach($corp[strtolower($xval['defcol_one_id'][1])] as $id=>$sval) {
                $corp[strtolower($xval['defcol_one_id'][1])][$id] = $xval[$id];
            }
            echo "FIRST CORP VAR_DUMP!";
            var_dump($corp);
            var_dump(strtolower($xval['defcol_one_id'][1]));
            var_dump($corp[strtolower($xval['defcol_one_id'][1])]);
            var_dump($xarr);
            $total_col = $xval['q1']+$xval['q2']+$xval['q3']+$xval['q4'];

            fill_row($objPHPExcel, $xval, $total_col, $xval['defcol_one_id'][1], $i, $style_normal);
            $i++;
        }
        else if ($cat == $xval['defcol_two_id'][1]) {
            //sum cat_rev
            foreach ($cat_rev as $id=>$sval){
                $cat_rev[$id] += $xval[$id];
            }
            foreach($corp[strtolower($xval['defcol_one_id'][1])] as $id=>$sval) {
                $corp[strtolower($xval['defcol_one_id'][1])][$id] += $xval[$id];
            }
            //xls the line
            $total_col = $xval['q1']+$xval['q2']+$xval['q3']+$xval['q4'];
            fill_row($objPHPExcel, $xval, $total_col, $xval['defcol_one_id'][1], $i, $style_normal);
            $i++;
        }
        else {
            //sum tot_rev
            foreach ($tot_rev as $id=>$sval){
                //echo "id = ".$id.", sval = ".$sval.", tot_rev[id] = ".$tot_rev[$id];
                $tot_rev[$id] += $cat_rev[$id];
            }
            //xls the cat_rev
            $total_col = $cat_rev['q1']+$cat_rev['q2']+$cat_rev['q3']+$cat_rev['q4'];
            fill_row($objPHPExcel, $cat_rev, $total_col, strtoupper($cat), $i, $style_stotal);
            //here fill the "special" percentage column
            fill_persentage($objPHPExcel, 'R', 'S', $start, $i, $style_stotal);
            $i++;
            $start = $i;
            $total_col = $xval['q1']+$xval['q2']+$xval['q3']+$xval['q4'];
            fill_row($objPHPExcel, $xval, $total_col, $xval['defcol_one_id'][1], $i, $style_normal);
            $i++;
            $cat = $xval['defcol_two_id'][1];
            //new cat_rev
            foreach ($cat_rev as $id=>$sval){
                $cat_rev[$id] = $xval[$id];
            }
            foreach($corp[strtolower($xval['defcol_one_id'][1])] as $id=>$sval) {
                $corp[strtolower($xval['defcol_one_id'][1])][$id] += $xval[$id];
            }
        }
    } //foreach $xarr
    foreach ($tot_rev as $id=>$sval){
        $tot_rev[$id] += $cat_rev[$id];
    }
    //xls cat and grand total
    $total_col = $cat_rev['q1']+$cat_rev['q2']+$cat_rev['q3']+$cat_rev['q4'];
    fill_row($objPHPExcel, $cat_rev, $total_col, strtoupper($cat), $i, $style_stotal);
    fill_persentage($objPHPExcel, 'R', 'S', $start, $i, $style_stotal);
    $i++;
    $objPHPExcel->getActiveSheet()->mergeCells('A'.$i.':'.$objPHPExcel->getActiveSheet()->getHighestColumn().$i);
    $i++;
    foreach($corp as $tech => $varr) {
        $total_col = $varr['q1'] + $varr['q2'] + $varr['q3'] + $varr['q4'];
        fill_row($objPHPExcel, $varr, $total_col, 'TGT CORPORATE: '.strtoupper($tech), $i, $style_total);
        fill_persentage($objPHPExcel, 'R', 'T', 2, $i, $style_total);
        $i++;
    }
    $objPHPExcel->getActiveSheet()->mergeCells('A'.$i.':'.$objPHPExcel->getActiveSheet()->getHighestColumn().$i);
    $i++;
    $total_col = $tot_rev['q1']+$tot_rev['q2']+$tot_rev['q3']+$tot_rev['q4'];
    fill_row($objPHPExcel, $tot_rev, $total_col, 'TGT CORPORATE', $i, $style_total);
    fill_persentage($objPHPExcel, 'R', 'T', 2, $i, $style_total);

    $objPHPExcel->getActiveSheet()->getStyle('A1:A'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_ctitle);

    $objPHPExcel->getActiveSheet()->getStyle('E1:E'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_stotal);
    $objPHPExcel->getActiveSheet()->getStyle('I1:I'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_stotal);
    $objPHPExcel->getActiveSheet()->getStyle('M1:M'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_stotal);
    $objPHPExcel->getActiveSheet()->getStyle('Q1:Q'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_stotal);
    $objPHPExcel->getActiveSheet()->getStyle('R1:R'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_total);

    $objPHPExcel->getActiveSheet()->getStyle('A1:'.$objPHPExcel->getActiveSheet()->getHighestColumn().'1')->applyFromArray($style_total);

    $objPHPExcel->getActiveSheet()->getStyle(
        'B2:' .
        $objPHPExcel->getActiveSheet()->getHighestColumn().
        $objPHPExcel->getActiveSheet()->getHighestRow()
    )->getNumberFormat()->setFormatCode('#,##0_-');

    /*$objPHPExcel->getActiveSheet()->getStyle('S2:T'.$objPHPExcel->getActiveSheet()->getHighestRow())
        ->getNumberFormat()->applyFromArray(
            array(
                'code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE
            )
        );*/

    $objPHPExcel->getActiveSheet()->getStyle(
        'A1:' .
        $objPHPExcel->getActiveSheet()->getHighestColumn().
        $objPHPExcel->getActiveSheet()->getHighestRow()
    )->applyFromArray($style_borders);

    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, "Excel2007");
    $objWriter->save("/home/webmaster/tgtoil.com/wp-content/themes/MyCuisine/xmlrpc_lib/xls_files/Revenue_".$report_name.".xlsx");
} //standard save xls


/***** ADDITIONAL SUPPORT FUNCTIONS - USED IN ALL REPORTS *****/

function fill_row($objPHPExcel, $tot_rev, $total_col, $name, $i, $style) {

    $objPHPExcel->getActiveSheet()->setCellValue('A'.$i, $name)
                                          ->setCellValue('B'.$i, num_form($tot_rev['jan']))
                                          ->setCellValue('C'.$i, num_form($tot_rev['feb']))
                                          ->setCellValue('D'.$i, num_form($tot_rev['mar']))
                                          ->setCellValue('E'.$i, num_form($tot_rev['q1']))
                                          ->setCellValue('F'.$i, num_form($tot_rev['apr']))
                                          ->setCellValue('G'.$i, num_form($tot_rev['may']))
                                          ->setCellValue('H'.$i, num_form($tot_rev['jun']))
                                          ->setCellValue('I'.$i, num_form($tot_rev['q2']))
                                          ->setCellValue('J'.$i, num_form($tot_rev['jul']))
                                          ->setCellValue('K'.$i, num_form($tot_rev['aug']))
                                          ->setCellValue('L'.$i, num_form($tot_rev['sep']))
                                          ->setCellValue('M'.$i, num_form($tot_rev['q3']))
                                          ->setCellValue('N'.$i, num_form($tot_rev['oct']))
                                          ->setCellValue('O'.$i, num_form($tot_rev['nov']))
                                          ->setCellValue('P'.$i, num_form($tot_rev['des']))
                                          ->setCellValue('Q'.$i, num_form($tot_rev['q4']))
                                          ->setCellValue('R'.$i, num_form($total_col));
    if ($style != 0) $objPHPExcel->getActiveSheet()->getStyle('A'.$i.':R'.$i)->applyFromArray($style);
}

function num_form($number) {
    if ($number != 0)  {
        return number_format($number,0,'.',',');
    }
    else {
        return '';
    }
}

function fill_persentage($objPHPExcel, $column_get, $column_set, $row_start, $row_end, $style) {
    $val_tot = floatval(preg_replace("/[^-0-9\.]/","",$objPHPExcel->getActiveSheet()->getCell($column_get.$row_end)->getValue()));
    for ($i = $row_start; $i <= $row_end; $i++) {
        $val = floatval(preg_replace("/[^-0-9\.]/","",$objPHPExcel->getActiveSheet()->getCell($column_get.$i)->getValue()));
        $the_val = ($val*100/$val_tot);
        if (!is_float($the_val)) var_dump($the_val);
        $objPHPExcel->getActiveSheet()->setCellValue($column_set.$i, round($the_val, 0, PHP_ROUND_HALF_UP).'%');
    }
    $objPHPExcel->getActiveSheet()->getStyle('S'.$row_start.':T'.($row_end-1))->applyFromArray(
        array(
                            'font' => array(
                                'name' => 'Open Sans',
                                'size' => 8,
                                'color' => array('argb' => 'FF036B7E')
                            )
                )
    );
    if ($style != 0) $objPHPExcel->getActiveSheet()->getStyle('S'.$row_end.':T'.$row_end)->applyFromArray($style);
}