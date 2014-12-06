<?php

include 'xmlrpc.inc';
/** Include PHPExcel **/
require_once "/home/webmaster/tgtoil.com/wp-includes/utilites/ExClasses/PHPExcel.php";
require_once 'report_class.php';
require_once 'xml_styles.php';

set_time_limit(0);
ignore_user_abort(1);

/**********
AVG Revenue Reports - connect, get data, generate file
**********/
$CA_Report = new Report;
$relation_ca = array("sale.web.avg.de.reportq1", "sale.web.avg.de.reportq2", "sale.web.avg.de.reportq3", "sale.web.avg.de.reportq4");
$required_keys_ca = array(
            array(
                       new xmlrpcval('defcol_one_id', "string"),
                       new xmlrpcval('defcol_two_id', "string"),
                       new xmlrpcval('jan', "string"),
                       new xmlrpcval('jjan', "string"),
                       new xmlrpcval('feb', "string"),
                       new xmlrpcval('jfeb', "string"),
                       new xmlrpcval('mar', "string"),
                       new xmlrpcval('jmar', "string"),
                       new xmlrpcval('q1', "string"),
                       new xmlrpcval('jq1', "string")
            ),
            array(
                       new xmlrpcval('defcol_one_id', "string"),
                       new xmlrpcval('defcol_two_id', "string"),
                       new xmlrpcval('apr', "string"),
                       new xmlrpcval('japr', "string"),
                       new xmlrpcval('may', "string"),
                       new xmlrpcval('jmay', "string"),
                       new xmlrpcval('jun', "string"),
                       new xmlrpcval('jjun', "string"),
                       new xmlrpcval('q2', "string"),
                       new xmlrpcval('jq2', "string")
            ),
            array(
                       new xmlrpcval('defcol_one_id', "string"),
                       new xmlrpcval('defcol_two_id', "string"),
                       new xmlrpcval('jul', "string"),
                       new xmlrpcval('jjul', "string"),
                       new xmlrpcval('aug', "string"),
                       new xmlrpcval('jaug', "string"),
                       new xmlrpcval('sep', "string"),
                       new xmlrpcval('jsep', "string"),
                       new xmlrpcval('q3', "string"),
                       new xmlrpcval('jq3', "string")
            ),
            array(
                       new xmlrpcval('defcol_one_id', "string"),
                       new xmlrpcval('defcol_two_id', "string"),
                       new xmlrpcval('oct', "string"),
                       new xmlrpcval('joct', "string"),
                       new xmlrpcval('nov', "string"),
                       new xmlrpcval('jnov', "string"),
                       new xmlrpcval('des', "string"),
                       new xmlrpcval('jdes', "string"),
                       new xmlrpcval('q4', "string"),
                       new xmlrpcval('jq4', "string")
            )
    );
$search_keys_ca = array(
            array( //for hr.vacation.report
                        new xmlrpcval(array(new xmlrpcval("id" , "string"),
                        new xmlrpcval("!=","string"),
                        new xmlrpcval("-1","string")),"array"),
            ),
            array( //for hr.vacation.report
                        new xmlrpcval(array(new xmlrpcval("id" , "string"),
                        new xmlrpcval("!=","string"),
                        new xmlrpcval("-1","string")),"array"),
            ),
            array( //for hr.vacation.report
                        new xmlrpcval(array(new xmlrpcval("id" , "string"),
                        new xmlrpcval("!=","string"),
                        new xmlrpcval("-1","string")),"array"),
            ),
            array( //for hr.vacation.report
                        new xmlrpcval(array(new xmlrpcval("id" , "string"),
                        new xmlrpcval("!=","string"),
                        new xmlrpcval("-1","string")),"array"),
            )
    );
$CA_Report->constructor('rpc_user', 'rpc_password', 'db_name', 'xmlrpc_interface_url', $relation_ca, $required_keys_ca, $search_keys_ca);
$CA_Report->connect_xmlrpc();
$res = array();
for ($i=0; $i < count($CA_Report->relation_names); $i++){
    $CA_Report->get_xml_data($i);
    /*usort($CA_Report->data_arr[$i], function ($a, $b) {
            if ($a["defcol_one_id"][1] == $b["defcol_one_id"][1]) return 0;
            return ($a["defcol_one_id"][1] < $b["defcol_one_id"][1]) ? -1 : 1;
        });*/
}
$res = array_replace_recursive($CA_Report->data_arr[0], $CA_Report->data_arr[1], $CA_Report->data_arr[2], $CA_Report->data_arr[3]);

save_avg_report_excel_de($res, $style_header, $style_borders, $style_normal, $style_ctitle, $style_stotal, $style_total);

$CL_Report = new Report;
$relation_cl = array("sale.web.avg.pa.reportq1", "sale.web.avg.pa.reportq2", "sale.web.avg.pa.reportq3", "sale.web.avg.pa.reportq4");
$required_keys_cl = array(
            array(
                       new xmlrpcval('defcol_one_id', "array"),
                       new xmlrpcval('defcol_two_id', "array"),
                       new xmlrpcval('jan', "string"),
                       new xmlrpcval('jjan', "string"),
                       new xmlrpcval('feb', "string"),
                       new xmlrpcval('jfeb', "string"),
                       new xmlrpcval('mar', "string"),
                       new xmlrpcval('jmar', "string"),
                       new xmlrpcval('q1', "string"),
                       new xmlrpcval('jq1', "string")
            ),
            array(
                       new xmlrpcval('defcol_one_id', "array"),
                       new xmlrpcval('defcol_two_id', "array"),
                       new xmlrpcval('apr', "string"),
                       new xmlrpcval('japr', "string"),
                       new xmlrpcval('may', "string"),
                       new xmlrpcval('jmay', "string"),
                       new xmlrpcval('jun', "string"),
                       new xmlrpcval('jjun', "string"),
                       new xmlrpcval('q2', "string"),
                       new xmlrpcval('jq2', "string")
            ),
            array(
                       new xmlrpcval('defcol_one_id', "array"),
                       new xmlrpcval('defcol_two_id', "array"),
                       new xmlrpcval('jul', "string"),
                       new xmlrpcval('jjul', "string"),
                       new xmlrpcval('aug', "string"),
                       new xmlrpcval('jaug', "string"),
                       new xmlrpcval('sep', "string"),
                       new xmlrpcval('jsep', "string"),
                       new xmlrpcval('q3', "string"),
                       new xmlrpcval('jq3', "string")
            ),
            array(
                       new xmlrpcval('defcol_one_id', "array"),
                       new xmlrpcval('defcol_two_id', "array"),
                       new xmlrpcval('oct', "string"),
                       new xmlrpcval('joct', "string"),
                       new xmlrpcval('nov', "string"),
                       new xmlrpcval('jnov', "string"),
                       new xmlrpcval('des', "string"),
                       new xmlrpcval('jdes', "string"),
                       new xmlrpcval('q4', "string"),
                       new xmlrpcval('jq4', "string")
            )
    );
$search_keys_cl = array(
            array( //for hr.vacation.report
                        new xmlrpcval(array(new xmlrpcval("id" , "string"),
                        new xmlrpcval("!=","string"),
                        new xmlrpcval("-1","string")),"array"),
            ),
            array( //for hr.vacation.report
                        new xmlrpcval(array(new xmlrpcval("id" , "string"),
                        new xmlrpcval("!=","string"),
                        new xmlrpcval("-1","string")),"array"),
            ),
            array( //for hr.vacation.report
                        new xmlrpcval(array(new xmlrpcval("id" , "string"),
                        new xmlrpcval("!=","string"),
                        new xmlrpcval("-1","string")),"array"),
            ),
            array( //for hr.vacation.report
                        new xmlrpcval(array(new xmlrpcval("id" , "string"),
                        new xmlrpcval("!=","string"),
                        new xmlrpcval("-1","string")),"array"),
            )
    );
$CL_Report->constructor('rpc_user', 'rpc_password', 'db_name', 'xmlrpc_interface_url', $relation_cl, $required_keys_cl, $search_keys_cl);
$CL_Report->connect_xmlrpc();
$res = array();
for ($i=0; $i < count($CL_Report->relation_names); $i++){
    $CL_Report->get_xml_data($i);
    /*usort($CA_Report->data_arr[$i], function ($a, $b) {
            if ($a["defcol_one_id"][1] == $b["defcol_one_id"][1]) return 0;
            return ($a["defcol_one_id"][1] < $b["defcol_one_id"][1]) ? -1 : 1;
        });*/
}
$res = array_replace_recursive($CL_Report->data_arr[0], $CL_Report->data_arr[1], $CL_Report->data_arr[2], $CL_Report->data_arr[3]);


save_avg_report_excel($res, "Avg_Revenue_by_Client", $style_header, $style_borders, $style_normal, $style_ctitle, $style_stotal, $style_total);

$OP_Report = new Report;
$relation_op = array("sale.web.avg.opr.reportq1", "sale.web.avg.opr.reportq2", "sale.web.avg.opr.reportq3", "sale.web.avg.opr.reportq4");
$required_keys_op = array(
            array(
                       new xmlrpcval('defcol_one_id', "array"),
                       new xmlrpcval('defcol_two_id', "array"),
                       new xmlrpcval('jan', "string"),
                       new xmlrpcval('jjan', "string"),
                       new xmlrpcval('feb', "string"),
                       new xmlrpcval('jfeb', "string"),
                       new xmlrpcval('mar', "string"),
                       new xmlrpcval('jmar', "string"),
                       new xmlrpcval('q1', "string"),
                       new xmlrpcval('jq1', "string")
            ),
            array(
                       new xmlrpcval('defcol_one_id', "array"),
                       new xmlrpcval('defcol_two_id', "array"),
                       new xmlrpcval('apr', "string"),
                       new xmlrpcval('japr', "string"),
                       new xmlrpcval('may', "string"),
                       new xmlrpcval('jmay', "string"),
                       new xmlrpcval('jun', "string"),
                       new xmlrpcval('jjun', "string"),
                       new xmlrpcval('q2', "string"),
                       new xmlrpcval('jq2', "string")
            ),
            array(
                       new xmlrpcval('defcol_one_id', "array"),
                       new xmlrpcval('defcol_two_id', "array"),
                       new xmlrpcval('jul', "string"),
                       new xmlrpcval('jjul', "string"),
                       new xmlrpcval('aug', "string"),
                       new xmlrpcval('jaug', "string"),
                       new xmlrpcval('sep', "string"),
                       new xmlrpcval('jsep', "string"),
                       new xmlrpcval('q3', "string"),
                       new xmlrpcval('jq3', "string")
            ),
            array(
                       new xmlrpcval('defcol_one_id', "array"),
                       new xmlrpcval('defcol_two_id', "array"),
                       new xmlrpcval('oct', "string"),
                       new xmlrpcval('joct', "string"),
                       new xmlrpcval('nov', "string"),
                       new xmlrpcval('jnov', "string"),
                       new xmlrpcval('des', "string"),
                       new xmlrpcval('jdes', "string"),
                       new xmlrpcval('q4', "string"),
                       new xmlrpcval('jq4', "string")
            )
    );
$search_keys_op = array(
            array( //for hr.vacation.report
                        new xmlrpcval(array(new xmlrpcval("id" , "string"),
                        new xmlrpcval("!=","string"),
                        new xmlrpcval("-1","string")),"array"),
            ),
            array( //for hr.vacation.report
                        new xmlrpcval(array(new xmlrpcval("id" , "string"),
                        new xmlrpcval("!=","string"),
                        new xmlrpcval("-1","string")),"array"),
            ),
            array( //for hr.vacation.report
                        new xmlrpcval(array(new xmlrpcval("id" , "string"),
                        new xmlrpcval("!=","string"),
                        new xmlrpcval("-1","string")),"array"),
            ),
            array( //for hr.vacation.report
                        new xmlrpcval(array(new xmlrpcval("id" , "string"),
                        new xmlrpcval("!=","string"),
                        new xmlrpcval("-1","string")),"array"),
            )
    );
$OP_Report->constructor('rpc_user', 'rpc_password', 'db_name', 'xmlrpc_interface_url', $relation_op, $required_keys_op, $search_keys_op);
$OP_Report->connect_xmlrpc();
$res = array();
for ($i=0; $i < count($OP_Report->relation_names); $i++){
    $OP_Report->get_xml_data($i);
    /*usort($CA_Report->data_arr[$i], function ($a, $b) {
            if ($a["defcol_one_id"][1] == $b["defcol_one_id"][1]) return 0;
            return ($a["defcol_one_id"][1] < $b["defcol_one_id"][1]) ? -1 : 1;
        });*/
}
$res = array_replace_recursive($OP_Report->data_arr[0], $OP_Report->data_arr[1], $OP_Report->data_arr[2], $OP_Report->data_arr[3]);

save_avg_report_excel($res, "Avg_Revenue_by_Operator", $style_header, $style_borders, $style_normal, $style_ctitle, $style_stotal, $style_total);

$CO_Report = new Report;
$relation_co = array("sale.web.avg.cou.reportq1", "sale.web.avg.cou.reportq2", "sale.web.avg.cou.reportq3", "sale.web.avg.cou.reportq4");
$required_keys_co = array(
            array(
                       new xmlrpcval('defcol_one_id', "array"),
                       new xmlrpcval('defcol_two_id', "array"),
                       new xmlrpcval('jan', "string"),
                       new xmlrpcval('jjan', "string"),
                       new xmlrpcval('feb', "string"),
                       new xmlrpcval('jfeb', "string"),
                       new xmlrpcval('mar', "string"),
                       new xmlrpcval('jmar', "string"),
                       new xmlrpcval('q1', "string"),
                       new xmlrpcval('jq1', "string")
            ),
            array(
                       new xmlrpcval('defcol_one_id', "array"),
                       new xmlrpcval('defcol_two_id', "array"),
                       new xmlrpcval('apr', "string"),
                       new xmlrpcval('japr', "string"),
                       new xmlrpcval('may', "string"),
                       new xmlrpcval('jmay', "string"),
                       new xmlrpcval('jun', "string"),
                       new xmlrpcval('jjun', "string"),
                       new xmlrpcval('q2', "string"),
                       new xmlrpcval('jq2', "string")
            ),
            array(
                       new xmlrpcval('defcol_one_id', "array"),
                       new xmlrpcval('defcol_two_id', "array"),
                       new xmlrpcval('jul', "string"),
                       new xmlrpcval('jjul', "string"),
                       new xmlrpcval('aug', "string"),
                       new xmlrpcval('jaug', "string"),
                       new xmlrpcval('sep', "string"),
                       new xmlrpcval('jsep', "string"),
                       new xmlrpcval('q3', "string"),
                       new xmlrpcval('jq3', "string")
            ),
            array(
                       new xmlrpcval('defcol_one_id', "array"),
                       new xmlrpcval('defcol_two_id', "array"),
                       new xmlrpcval('oct', "string"),
                       new xmlrpcval('joct', "string"),
                       new xmlrpcval('nov', "string"),
                       new xmlrpcval('jnov', "string"),
                       new xmlrpcval('des', "string"),
                       new xmlrpcval('jdes', "string"),
                       new xmlrpcval('q4', "string"),
                       new xmlrpcval('jq4', "string")
            )
    );
$search_keys_co = array(
            array( //for hr.vacation.report
                        new xmlrpcval(array(new xmlrpcval("id" , "string"),
                        new xmlrpcval("!=","string"),
                        new xmlrpcval("-1","string")),"array"),
            ),
            array( //for hr.vacation.report
                        new xmlrpcval(array(new xmlrpcval("id" , "string"),
                        new xmlrpcval("!=","string"),
                        new xmlrpcval("-1","string")),"array"),
            ),
            array( //for hr.vacation.report
                        new xmlrpcval(array(new xmlrpcval("id" , "string"),
                        new xmlrpcval("!=","string"),
                        new xmlrpcval("-1","string")),"array"),
            ),
            array( //for hr.vacation.report
                        new xmlrpcval(array(new xmlrpcval("id" , "string"),
                        new xmlrpcval("!=","string"),
                        new xmlrpcval("-1","string")),"array"),
            )
    );
$CO_Report->constructor('rpc_user', 'rpc_password', 'db_name', 'xmlrpc_interface_url', $relation_co, $required_keys_co, $search_keys_co);
$CO_Report->connect_xmlrpc();
$res = array();
for ($i=0; $i < count($CO_Report->relation_names); $i++){
    $CO_Report->get_xml_data($i);
    /*usort($CA_Report->data_arr[$i], function ($a, $b) {
            if ($a["defcol_one_id"][1] == $b["defcol_one_id"][1]) return 0;
            return ($a["defcol_one_id"][1] < $b["defcol_one_id"][1]) ? -1 : 1;
        });*/
}
$res = array_replace_recursive($CO_Report->data_arr[0], $CO_Report->data_arr[1], $CO_Report->data_arr[2], $CO_Report->data_arr[3]);

save_avg_report_excel2($res, "Avg_Revenue_by_Country", $style_header, $style_borders, $style_normal, $style_ctitle, $style_stotal, $style_total);

function save_avg_report_excel_de($xarr,$style_header, $style_borders, $style_normal, $style_ctitle, $style_stotal, $style_total) {


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

    $objPHPExcel->getActiveSheet()->setCellValue('A1', 'Service')
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
                                  ->setCellValue('R1', 'Total YTD')
                                  ->setCellValue('S1', 'Revenue / Job');

    $cat = '';
    $cat_rev = array('jan'=>0, 'jjan'=>0, 'feb'=>0, 'jfeb'=>0, 'mar'=>0, 'jmar'=>0, 'q1'=>0, 'jq1'=>0,
                     'apr'=>0, 'japr'=>0, 'may'=>0, 'jmay'=>0, 'jun'=>0, 'jjun'=>0, 'q2'=>0, 'jq2'=>0,
                     'jul'=>0, 'jjul'=>0, 'aug'=>0, 'jaug'=>0, 'sep'=>0, 'jsep'=>0, 'q3'=>0, 'jq3'=>0,
                     'oct'=>0, 'joct'=>0, 'nov'=>0, 'jnov'=>0, 'des'=>0, 'jdes'=>0, 'q4'=>0, 'jq4'=>0);
    $tot_rev = array('jan'=>0, 'jjan'=>0, 'feb'=>0, 'jfeb'=>0, 'mar'=>0, 'jmar'=>0, 'q1'=>0, 'jq1'=>0,
                     'apr'=>0, 'japr'=>0, 'may'=>0, 'jmay'=>0, 'jun'=>0, 'jjun'=>0, 'q2'=>0, 'jq2'=>0,
                     'jul'=>0, 'jjul'=>0, 'aug'=>0, 'jaug'=>0, 'sep'=>0, 'jsep'=>0, 'q3'=>0, 'jq3'=>0,
                     'oct'=>0, 'joct'=>0, 'nov'=>0, 'jnov'=>0, 'des'=>0, 'jdes'=>0, 'q4'=>0, 'jq4'=>0);

    $i=2;
    foreach($xarr as $k => $xval) {


        if ($cat=='') {
            echo "cat == '', service = ".$xval['defcol_one_id'][1];
            $cat = $xval['defcol_two_id'];
            //fill cat_rev
            foreach ($cat_rev as $id=>$sval){
                $cat_rev[$id] = $xval[$id];
            }
            //xls the line
            $tot_div = $xval['jq1']+$xval['jq2']+$xval['jq3']+$xval['jq4'];
            $total_col = $xval['q1']+$xval['q2']+$xval['q3']+$xval['q4'];
            fill_row($objPHPExcel, $xval, $total_col, $tot_div, $xval['defcol_one_id'][1], $i, $style_normal);
            $i += 2;
        }
        else if ($cat == $xval['defcol_two_id']) {
            echo "Same cat == ".$cat." service = ".$xval['defcol_one_id'][1];
            //sum cat_rev
            foreach ($cat_rev as $id=>$sval){
                $cat_rev[$id] += $xval[$id];
            }
            //xls the line
            $tot_div = $xval['jq1']+$xval['jq2']+$xval['jq3']+$xval['jq4'];
            $total_col = $xval['q1']+$xval['q2']+$xval['q3']+$xval['q4'];
            fill_row($objPHPExcel, $xval, $total_col, $tot_div, $xval['defcol_one_id'][1], $i, $style_normal);
            $i += 2;
        }
        else {
            //sum tot_rev
            echo "Old cat == ".$cat." New cat == ".$xval['defcol_two_id']." New service = ".$xval['defcol_one_id'][1];
            foreach ($tot_rev as $id=>$sval){
                $tot_rev[$id] += $cat_rev[$id];
            }
            //xls the cat_rev
            $tot_div = $cat_rev['jq1']+$cat_rev['jq2']+$cat_rev['jq3']+$cat_rev['jq4'];
            $total_col = $cat_rev['q1']+$cat_rev['q2']+$cat_rev['q3']+$cat_rev['q4'];
            fill_row($objPHPExcel, $cat_rev, $total_col, $tot_div, strtoupper($cat), $i, $style_stotal);
            $i += 2;
            $objPHPExcel->getActiveSheet()->mergeCells('A'.$i.':S'.$i);
            $i++;
            $tot_div = $xval['jq1']+$xval['jq2']+$xval['jq3']+$xval['jq4'];
            $total_col = $xval['q1']+$xval['q2']+$xval['q3']+$xval['q4'];
            fill_row($objPHPExcel, $xval, $total_col, $tot_div, $xval['defcol_one_id'][1], $i, $style_normal);
            $i +=2;
            $cat = $xval['defcol_two_id'];
            //new cat_rev
            foreach ($cat_rev as $id=>$sval){
                $cat_rev[$id] = $xval[$id];
            }
        }
    } //foreach $xarr
    $last_elem = end($xarr);
    foreach ($tot_rev as $id=>$sval){
        $tot_rev[$id] += $last_elem[$id];
    }
    //xls cat and grand total
                $tot_div = $cat_rev['jq1']+$cat_rev['jq2']+$cat_rev['jq3']+$cat_rev['jq4'];
                $total_col = $cat_rev['q1']+$cat_rev['q2']+$cat_rev['q3']+$cat_rev['q4'];
                fill_row($objPHPExcel, $cat_rev, $total_col, $tot_div, strtoupper($cat), $i, $style_stotal);
                $i += 2;
            $objPHPExcel->getActiveSheet()->mergeCells('A'.$i.':S'.$i);
            $i++;
    $tot_div = $tot_rev['jq1']+$tot_rev['jq2']+$tot_rev['jq3']+$tot_rev['jq4'];
    $total_col = $tot_rev['q1']+$tot_rev['q2']+$tot_rev['q3']+$tot_rev['q4'];

    $objPHPExcel->getActiveSheet()->getStyle('A1:A'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_ctitle);
    fill_row($objPHPExcel, $tot_rev, $total_col, $tot_div, 'TGT Corporate', $i, $style_total);

    $objPHPExcel->getActiveSheet()->getStyle('E1:E'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_stotal);
    $objPHPExcel->getActiveSheet()->getStyle('I1:I'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_stotal);
    $objPHPExcel->getActiveSheet()->getStyle('M1:M'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_stotal);
    $objPHPExcel->getActiveSheet()->getStyle('Q1:Q'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_stotal);
    $objPHPExcel->getActiveSheet()->getStyle('R1:R'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_total);

    $objPHPExcel->getActiveSheet()->getStyle('A1:'.$objPHPExcel->getActiveSheet()->getHighestColumn().'1')->applyFromArray($style_total);


    $objPHPExcel->getActiveSheet()->getStyle(
        'A1:' .
        $objPHPExcel->getActiveSheet()->getHighestColumn().
        $objPHPExcel->getActiveSheet()->getHighestRow()
    )->applyFromArray($style_borders);

    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, "Excel2007");
    $objWriter->save("/home/webmaster/tgtoil.com/wp-content/themes/MyCuisine/xmlrpc_lib/xls_files/Avg_Revenue_by_Category.xlsx");
}

function save_avg_report_excel($xarr, $report_name,$style_header, $style_borders, $style_normal, $style_ctitle, $style_stotal, $style_total) {

        foreach ($xarr as $key => $val) {
            echo "key = ".$key.", ";
            $loc_arr[$key] = $val["defcol_one_id"][1];
            $cat_arr[$key] = $val["defcol_two_id"][1];
        }

        array_multisort($cat_arr, SORT_ASC, $loc_arr, SORT_ASC, $xarr);

        // Create new PHPExcel object
        echo " Create new PHPExcel object" ;
        $objPHPExcel = new PHPExcel();

        // Set document properties
        echo date('H:i:s') , " Set document properties" , EOL;
        $objPHPExcel->getProperties()->setCreator("OpenERP")
        							 ->setLastModifiedBy("OpenERP")
        							 ->setTitle("AVG Revenue Report")
        							 ->setSubject("Office 2007 XLSX Test Document")
        							 ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
        							 ->setKeywords("office 2007 openxml php")
        							 ->setCategory("Test result file");

        // Set default font
        echo date('H:i:s') , " Set default font" , EOL;
        $objPHPExcel->getDefaultStyle()->getFont()->setName('Arial')
                                                  ->setSize(10);

    echo date('H:i:s') , " Add some data" , EOL;
    $objPHPExcel->getActiveSheet()->setCellValue('A1', 'Service')
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
                                  ->setCellValue('R1', 'Total YTD')
                                  ->setCellValue('S1', 'Revenue / Job');

    $cat = '';
    $cat_rev = array('jan'=>0, 'jjan'=>0, 'feb'=>0, 'jfeb'=>0, 'mar'=>0, 'jmar'=>0, 'q1'=>0, 'jq1'=>0,
                     'apr'=>0, 'japr'=>0, 'may'=>0, 'jmay'=>0, 'jun'=>0, 'jjun'=>0, 'q2'=>0, 'jq2'=>0,
                     'jul'=>0, 'jjul'=>0, 'aug'=>0, 'jaug'=>0, 'sep'=>0, 'jsep'=>0, 'q3'=>0, 'jq3'=>0,
                     'oct'=>0, 'joct'=>0, 'nov'=>0, 'jnov'=>0, 'des'=>0, 'jdes'=>0, 'q4'=>0, 'jq4'=>0);
    $tot_rev = array('jan'=>0, 'jjan'=>0, 'feb'=>0, 'jfeb'=>0, 'mar'=>0, 'jmar'=>0, 'q1'=>0, 'jq1'=>0,
                     'apr'=>0, 'japr'=>0, 'may'=>0, 'jmay'=>0, 'jun'=>0, 'jjun'=>0, 'q2'=>0, 'jq2'=>0,
                     'jul'=>0, 'jjul'=>0, 'aug'=>0, 'jaug'=>0, 'sep'=>0, 'jsep'=>0, 'q3'=>0, 'jq3'=>0,
                     'oct'=>0, 'joct'=>0, 'nov'=>0, 'jnov'=>0, 'des'=>0, 'jdes'=>0, 'q4'=>0, 'jq4'=>0);

    $i=2;
    foreach($xarr as $k => $xval) {

        if ($cat=='') {
            echo "cat == '', service = ".$xval['defcol_one_id'][1];
            $cat = $xval['defcol_two_id'][1];
            //fill cat_rev
            foreach ($cat_rev as $id=>$sval){
                $cat_rev[$id] = $xval[$id];
            }
            //xls the line
            $tot_div = $xval['jq1']+$xval['jq2']+$xval['jq3']+$xval['jq4'];
            $total_col = $xval['q1']+$xval['q2']+$xval['q3']+$xval['q4'];
            fill_row($objPHPExcel, $xval, $total_col, $tot_div, $xval['defcol_one_id'][1], $i, $style_normal);
            $i += 2;
        }
        else if ($cat == $xval['defcol_two_id'][1]) {
            echo "Same cat == ".$cat." service = ".$xval['defcol_one_id'][1];
            //sum cat_rev
            foreach ($cat_rev as $id=>$sval){
                $cat_rev[$id] += $xval[$id];
            }
            //xls the line
            $tot_div = $xval['jq1']+$xval['jq2']+$xval['jq3']+$xval['jq4'];
            $total_col = $xval['q1']+$xval['q2']+$xval['q3']+$xval['q4'];
            fill_row($objPHPExcel, $xval, $total_col, $tot_div, $xval['defcol_one_id'][1], $i, $style_normal);
            $i += 2;
        }
        else {
            //sum tot_rev
            echo "Old cat == ".$cat." New cat == ".$xval['defcol_two_id'][1]." New service = ".$xval['defcol_one_id'][1];
            foreach ($tot_rev as $id=>$sval){
                $tot_rev[$id] += $cat_rev[$id];
            }
            //xls the cat_rev
            $tot_div = $cat_rev['jq1']+$cat_rev['jq2']+$cat_rev['jq3']+$cat_rev['jq4'];
            $total_col = $cat_rev['q1']+$cat_rev['q2']+$cat_rev['q3']+$cat_rev['q4'];
            fill_row($objPHPExcel, $cat_rev, $total_col, $tot_div, strtoupper($cat), $i, $style_stotal);
            $i += 2;
            $objPHPExcel->getActiveSheet()->mergeCells('A'.$i.':S'.$i);
            $i++;
            $tot_div = $xval['jq1']+$xval['jq2']+$xval['jq3']+$xval['jq4'];
            $total_col = $xval['q1']+$xval['q2']+$xval['q3']+$xval['q4'];
            fill_row($objPHPExcel, $xval, $total_col, $tot_div, $xval['defcol_one_id'][1], $i, $style_normal);
            $i += 2;
            $cat = $xval['defcol_two_id'][1];
            //new cat_rev
            foreach ($cat_rev as $id=>$sval){
                $cat_rev[$id] = $xval[$id];
            }
        }
    } //foreach $xarr
    $last_elem = end($xarr);
    foreach ($tot_rev as $id=>$sval){
        $tot_rev[$id] += $last_elem[$id];
    }
    //xls cat and grand total
    $tot_div = $cat_rev['jq1']+$cat_rev['jq2']+$cat_rev['jq3']+$cat_rev['jq4'];
    $total_col = $cat_rev['q1']+$cat_rev['q2']+$cat_rev['q3']+$cat_rev['q4'];
    fill_row($objPHPExcel, $cat_rev, $total_col, $tot_div, strtoupper($cat), $i, $style_stotal);
    $i += 2;
            $objPHPExcel->getActiveSheet()->mergeCells('A'.$i.':S'.$i);
            $i++;
    $tot_div = $tot_rev['jq1']+$tot_rev['jq2']+$tot_rev['jq3']+$tot_rev['jq4'];
    $total_col = $tot_rev['q1']+$tot_rev['q2']+$tot_rev['q3']+$tot_rev['q4'];

    $objPHPExcel->getActiveSheet()->getStyle('A1:A'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_ctitle);
    fill_row($objPHPExcel, $tot_rev, $total_col, $tot_div, 'TGT Corporate', $i, $style_total);

    $objPHPExcel->getActiveSheet()->getStyle('E1:E'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_stotal);
    $objPHPExcel->getActiveSheet()->getStyle('I1:I'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_stotal);
    $objPHPExcel->getActiveSheet()->getStyle('M1:M'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_stotal);
    $objPHPExcel->getActiveSheet()->getStyle('Q1:Q'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_stotal);
    $objPHPExcel->getActiveSheet()->getStyle('R1:S'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_total);

    $objPHPExcel->getActiveSheet()->getStyle('A1:'.$objPHPExcel->getActiveSheet()->getHighestColumn().'1')->applyFromArray($style_total);



    $objPHPExcel->getActiveSheet()->getStyle(
        'A1:' .
        $objPHPExcel->getActiveSheet()->getHighestColumn().
        $objPHPExcel->getActiveSheet()->getHighestRow()
    )->applyFromArray($style_borders);

    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, "Excel2007");
    $objWriter->save("/home/webmaster/tgtoil.com/wp-content/themes/MyCuisine/xmlrpc_lib/xls_files/".$report_name.".xlsx");
    echo "Saved the file";
}


function save_avg_report_excel2($xarr, $report_name,$style_header, $style_borders, $style_normal, $style_ctitle, $style_stotal, $style_total) {

        foreach ($xarr as $key => $val) {
            echo "key = ".$key.", ";
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

    $objPHPExcel->getActiveSheet()->setCellValue('A1', 'Service')
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
                                  ->setCellValue('R1', 'Total YTD')
                                  ->setCellValue('S1', 'Revenue / Job');

    $cat = '';
    $cat_rev = array('jan'=>0, 'jjan'=>0, 'feb'=>0, 'jfeb'=>0, 'mar'=>0, 'jmar'=>0, 'q1'=>0, 'jq1'=>0,
                     'apr'=>0, 'japr'=>0, 'may'=>0, 'jmay'=>0, 'jun'=>0, 'jjun'=>0, 'q2'=>0, 'jq2'=>0,
                     'jul'=>0, 'jjul'=>0, 'aug'=>0, 'jaug'=>0, 'sep'=>0, 'jsep'=>0, 'q3'=>0, 'jq3'=>0,
                     'oct'=>0, 'joct'=>0, 'nov'=>0, 'jnov'=>0, 'des'=>0, 'jdes'=>0, 'q4'=>0, 'jq4'=>0);
    $tot_rev = array('jan'=>0, 'jjan'=>0, 'feb'=>0, 'jfeb'=>0, 'mar'=>0, 'jmar'=>0, 'q1'=>0, 'jq1'=>0,
                     'apr'=>0, 'japr'=>0, 'may'=>0, 'jmay'=>0, 'jun'=>0, 'jjun'=>0, 'q2'=>0, 'jq2'=>0,
                     'jul'=>0, 'jjul'=>0, 'aug'=>0, 'jaug'=>0, 'sep'=>0, 'jsep'=>0, 'q3'=>0, 'jq3'=>0,
                     'oct'=>0, 'joct'=>0, 'nov'=>0, 'jnov'=>0, 'des'=>0, 'jdes'=>0, 'q4'=>0, 'jq4'=>0);

    $corp_arr = array();

    $i=2;
    foreach($xarr as $k => $xval) {

        if ($cat=='') {
            echo "cat == '', service = ".$xval['defcol_one_id'][1];
            $cat = $xval['defcol_two_id'][1];
            //fill cat_rev
            foreach ($cat_rev as $id=>$sval){
                $cat_rev[$id] = $xval[$id];
            }
            if (!isset($corp_arr[$xval['defcol_one_id'][1]])) {
                //create
                $corp_arr[$xval['defcol_one_id'][1]] = array('jan'=>0, 'jjan'=>0, 'feb'=>0, 'jfeb'=>0, 'mar'=>0, 'jmar'=>0, 'q1'=>0, 'jq1'=>0,
                                                         'apr'=>0, 'japr'=>0, 'may'=>0, 'jmay'=>0, 'jun'=>0, 'jjun'=>0, 'q2'=>0, 'jq2'=>0,
                                                         'jul'=>0, 'jjul'=>0, 'aug'=>0, 'jaug'=>0, 'sep'=>0, 'jsep'=>0, 'q3'=>0, 'jq3'=>0,
                                                         'oct'=>0, 'joct'=>0, 'nov'=>0, 'jnov'=>0, 'des'=>0, 'jdes'=>0, 'q4'=>0, 'jq4'=>0);
            }
            //add
            var_dump($corp_arr[$xval['defcol_one_id'][1]]);
            foreach ($corp_arr[$xval['defcol_one_id'][1]] as $id=>$sval){
                $corp_arr[$xval['defcol_one_id'][1]][$id] += $xval[$id];
            }
            //xls the line
            $tot_div = $xval['jq1']+$xval['jq2']+$xval['jq3']+$xval['jq4'];
            $total_col = $xval['q1']+$xval['q2']+$xval['q3']+$xval['q4'];
            fill_row($objPHPExcel, $xval, $total_col, $tot_div, $xval['defcol_one_id'][1], $i, $style_normal);
            $i += 2;
        }
        else if ($cat == $xval['defcol_two_id'][1]) {
            echo "Same cat == ".$cat." service = ".$xval['defcol_one_id'][1];
            //sum cat_rev
            foreach ($cat_rev as $id=>$sval){
                $cat_rev[$id] += $xval[$id];
            }
            if (!isset($corp_arr[$xval['defcol_one_id'][1]])) {
                //create
                $corp_arr[$xval['defcol_one_id'][1]] = array('jan'=>0, 'jjan'=>0, 'feb'=>0, 'jfeb'=>0, 'mar'=>0, 'jmar'=>0, 'q1'=>0, 'jq1'=>0,
                                                         'apr'=>0, 'japr'=>0, 'may'=>0, 'jmay'=>0, 'jun'=>0, 'jjun'=>0, 'q2'=>0, 'jq2'=>0,
                                                         'jul'=>0, 'jjul'=>0, 'aug'=>0, 'jaug'=>0, 'sep'=>0, 'jsep'=>0, 'q3'=>0, 'jq3'=>0,
                                                         'oct'=>0, 'joct'=>0, 'nov'=>0, 'jnov'=>0, 'des'=>0, 'jdes'=>0, 'q4'=>0, 'jq4'=>0);
            }
            echo "MORE CORP_ARR!";
            var_dump($corp_arr[$xval['defcol_one_id'][1]]);
            //add
            foreach ($corp_arr[$xval['defcol_one_id'][1]] as $id=>$sval){
                $corp_arr[$xval['defcol_one_id'][1]][$id] += $xval[$id];
            }
            //xls the line
            $tot_div = $xval['jq1']+$xval['jq2']+$xval['jq3']+$xval['jq4'];
            $total_col = $xval['q1']+$xval['q2']+$xval['q3']+$xval['q4'];
            fill_row($objPHPExcel, $xval, $total_col, $tot_div, $xval['defcol_one_id'][1], $i, $style_normal);
            $i += 2;
        }
        else {
            //sum tot_rev
            echo "Old cat == ".$cat." New cat == ".$xval['defcol_two_id'][1]." New service = ".$xval['defcol_one_id'][1];
            foreach ($tot_rev as $id=>$sval){
                $tot_rev[$id] += $cat_rev[$id];
            }
            //xls the cat_rev
            $tot_div = $cat_rev['jq1']+$cat_rev['jq2']+$cat_rev['jq3']+$cat_rev['jq4'];
            $total_col = $cat_rev['q1']+$cat_rev['q2']+$cat_rev['q3']+$cat_rev['q4'];
            fill_row($objPHPExcel, $cat_rev, $total_col, $tot_div, strtoupper($cat), $i, $style_stotal);
            $i += 2;
            $objPHPExcel->getActiveSheet()->mergeCells('A'.$i.':S'.$i);
            $i++;
            if (!isset($corp_arr[$xval['defcol_one_id'][1]])) {
                //create
                $corp_arr[$xval['defcol_one_id'][1]] = array('jan'=>0, 'jjan'=>0, 'feb'=>0, 'jfeb'=>0, 'mar'=>0, 'jmar'=>0, 'q1'=>0, 'jq1'=>0,
                                                         'apr'=>0, 'japr'=>0, 'may'=>0, 'jmay'=>0, 'jun'=>0, 'jjun'=>0, 'q2'=>0, 'jq2'=>0,
                                                         'jul'=>0, 'jjul'=>0, 'aug'=>0, 'jaug'=>0, 'sep'=>0, 'jsep'=>0, 'q3'=>0, 'jq3'=>0,
                                                         'oct'=>0, 'joct'=>0, 'nov'=>0, 'jnov'=>0, 'des'=>0, 'jdes'=>0, 'q4'=>0, 'jq4'=>0);
            }
            //add
            foreach ($corp_arr[$xval['defcol_one_id'][1]] as $id=>$sval){
                $corp_arr[$xval['defcol_one_id'][1]][$id] += $xval[$id];
            }
            $tot_div = $xval['jq1']+$xval['jq2']+$xval['jq3']+$xval['jq4'];
            $total_col = $xval['q1']+$xval['q2']+$xval['q3']+$xval['q4'];
            fill_row($objPHPExcel, $xval, $total_col, $tot_div, $xval['defcol_one_id'][1], $i, $style_normal);
            $i += 2;
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
    $tot_div = $cat_rev['jq1']+$cat_rev['jq2']+$cat_rev['jq3']+$cat_rev['jq4'];
    $total_col = $cat_rev['q1']+$cat_rev['q2']+$cat_rev['q3']+$cat_rev['q4'];
    fill_row($objPHPExcel, $cat_rev, $total_col, $tot_div, strtoupper($cat), $i, $style_stotal);
    $i += 2;
            $objPHPExcel->getActiveSheet()->mergeCells('A'.$i.':S'.$i);
            $i++;
    foreach($corp_arr as $tech => $varr) {
        $tot_div = $varr['jq1'] + $varr['jq2'] + $varr['jq3'] + $varr['jq4'];
        $total_col = $varr['q1'] + $varr['q2'] + $varr['q3'] + $varr['q4'];
        fill_row($objPHPExcel, $varr, $total_col, $tot_div, $tech, $i, $style_normal);
        $i += 2;
    }
    $tot_div = $tot_rev['jq1']+$tot_rev['jq2']+$tot_rev['jq3']+$tot_rev['jq4'];
    $total_col = $tot_rev['q1']+$tot_rev['q2']+$tot_rev['q3']+$tot_rev['q4'];

    $objPHPExcel->getActiveSheet()->getStyle('A1:A'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_ctitle);
    fill_row($objPHPExcel, $tot_rev, $total_col, $tot_div, 'TGT CORPORATE', $i, $style_total);

    $objPHPExcel->getActiveSheet()->getStyle('E1:E'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_stotal);
    $objPHPExcel->getActiveSheet()->getStyle('I1:I'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_stotal);
    $objPHPExcel->getActiveSheet()->getStyle('M1:M'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_stotal);
    $objPHPExcel->getActiveSheet()->getStyle('Q1:Q'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_stotal);
    $objPHPExcel->getActiveSheet()->getStyle('R1:S'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_total);

    $objPHPExcel->getActiveSheet()->getStyle('A1:'.$objPHPExcel->getActiveSheet()->getHighestColumn().'1')->applyFromArray($style_total);

    $objPHPExcel->getActiveSheet()->getStyle(
        'A1:' .
        $objPHPExcel->getActiveSheet()->getHighestColumn().
        $objPHPExcel->getActiveSheet()->getHighestRow()
    )->applyFromArray($style_borders);

    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, "Excel2007");
    $objWriter->save("/home/webmaster/tgtoil.com/wp-content/themes/MyCuisine/xmlrpc_lib/xls_files/".$report_name.".xlsx");
    echo "Saved the file";
}


function fill_row($objPHPExcel, $tot_rev, $total_col, $tot_div, $name, $i, $style) {
    $objPHPExcel->getActiveSheet()->setCellValue('A'.$i, $name.' Revenue')
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
                                          ->setCellValue('R'.$i, num_form($total_col))
                                          ->setCellValue('S'.$i, num_form(($total_col/(($tot_div)? $tot_div : 1))));
    $j = $i+1;
    $objPHPExcel->getActiveSheet()->setCellValue('A'.$j, $name.' Jobs')
                                          ->setCellValue('B'.$j, "x".num_form($tot_rev['jjan']))
                                          ->setCellValue('C'.$j, "x".num_form($tot_rev['jfeb']))
                                          ->setCellValue('D'.$j, "x".num_form($tot_rev['jmar']))
                                          ->setCellValue('E'.$j, "x".num_form($tot_rev['jq1']))
                                          ->setCellValue('F'.$j, "x".num_form($tot_rev['japr']))
                                          ->setCellValue('G'.$j, "x".num_form($tot_rev['jmay']))
                                          ->setCellValue('H'.$j, "x".num_form($tot_rev['jjun']))
                                          ->setCellValue('I'.$j, "x".num_form($tot_rev['jq2']))
                                          ->setCellValue('J'.$j, "x".num_form($tot_rev['jjul']))
                                          ->setCellValue('K'.$j, "x".num_form($tot_rev['jaug']))
                                          ->setCellValue('L'.$j, "x".num_form($tot_rev['jsep']))
                                          ->setCellValue('M'.$j, "x".num_form($tot_rev['jq3']))
                                          ->setCellValue('N'.$j, "x".num_form($tot_rev['joct']))
                                          ->setCellValue('O'.$j, "x".num_form($tot_rev['jnov']))
                                          ->setCellValue('P'.$j, "x".num_form($tot_rev['jdes']))
                                          ->setCellValue('Q'.$j, "x".num_form($tot_rev['jq4']))
                                          ->setCellValue('R'.$j, "x".num_form($tot_div));
    if ($style != 0) $objPHPExcel->getActiveSheet()->getStyle('A'.$i.':S'.($j))->applyFromArray($style);

}

function num_form($number) {
    if ($number != 0)  {
        return number_format($number,0,'.',',');
    }
    else {
        return '';
    }
}