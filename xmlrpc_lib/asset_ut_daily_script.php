<?php

include 'xmlrpc.inc';
/** Include PHPExcel **/
require_once "/home/webmaster/tgtoil.com/wp-includes/utilites/ExClasses/PHPExcel.php";
require_once 'report_class.php';
require_once 'xml_styles.php';


echo "Hello!";
$AUtil = new Report;
$relation_au = array("sale.by.asst.cou.report", "job.by.asst.cou.report");
$required_keys_au = array(

    );
$search_keys_au = array(

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
$fnames = array("Revenue", "Jobs_Count");
$AUtil->constructor('rpc_user', 'rpc_password', 'db_name', 'xmlrpc_interface_url', $relation_au, $required_keys_au, $search_keys_au);
$AUtil->connect_xmlrpc();
for ($i=0; $i < count($AUtil->relation_names); $i++){
    $AUtil->get_xml_data($i);
    //save_au_report_excel($AUtil->data_arr[$i], $fnames[$i],$style_header, $style_borders, $style_normal, $style_ctitle, $style_stotal, $style_total);
}
save_cau_report_excel($AUtil->data_arr, "Cons_Asset_Util",$style_header, $style_borders, $style_normal, $style_ctitle, $style_stotal, $style_total);


function save_cau_report_excel($ut_arr, $fname,$style_header, $style_borders, $style_normal, $style_ctitle, $style_stotal, $style_total) {

        $objPHPExcel = new PHPExcel();

        $objPHPExcel->getProperties()->setCreator("OpenERP")
        							 ->setLastModifiedBy("OpenERP")
        							 ->setTitle("AVG Revenue Report")
        							 ->setSubject("Office 2007 XLSX Test Document")
        							 ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
        							 ->setKeywords("office 2007 openxml php")
        							 ->setCategory("Test result file");

    $xarr1 = $ut_arr[0];
    $yarr1 = $ut_arr[1];

        foreach ($xarr1 as $key => $val) {
            $asset_arr[$key] = $val["defcol_one_id"];
            $cou_arr[$key] = $val["defcol_two_id"][1];
        }
        array_multisort($asset_arr, SORT_ASC, $cou_arr, SORT_ASC, $xarr1);

        foreach ($yarr1 as $key => $val) {
            $asset_arr[$key] = $val["defcol_one_id"];
            $cou_arr[$key] = $val["defcol_two_id"][1];
        }
        array_multisort($asset_arr, SORT_ASC, $cou_arr, SORT_ASC, $yarr1);



    $size = (count($xarr1) >= count($yarr1)) ? count($xarr1) : count($yarr1);

    $x = 0; $y = 0;
    $yarr = array();
    $xarr = array();

    for ($j = 0; $j < $size; $j++) {
        if (strcmp($yarr1[$y]["defcol_two_id"][1]+" "+$yarr1[$y]["defcol_one_id"],$xarr1[$x]["defcol_two_id"][1]+" "+$xarr1[$x]["defcol_one_id"]) == 0) {
            $yarr[$j] = $yarr1[$y];
            $xarr[$j] = $xarr1[$x];
            $x++;
            $y++;
        }
        else if (strcmp($yarr1[$y]["defcol_two_id"][1]+" "+$yarr1[$y]["defcol_one_id"],$xarr1[$x]["defcol_two_id"][1]+" "+$xarr1[$x]["defcol_one_id"]) < 0) {
            $yarr[$j] = $yarr1[$y];
            $xarr[$j] = array("defcol_one_id"=>$yarr1[$y]["defcol_one_id"], "defcol_two_id"=>$yarr1[$y]["defcol_two_id"],
                     'jan'=>0, 'janc'=>0, 'feb'=>0, 'febc'=>0, 'mar'=>0, 'marc'=>0, 'q1'=>0, 'q1c'=>0,
                     'apr'=>0, 'aprc'=>0, 'may'=>0, 'mayc'=>0, 'jun'=>0, 'junc'=>0, 'q2'=>0, 'q2c'=>0,
                     'jul'=>0, 'julc'=>0, 'aug'=>0, 'augc'=>0, 'sep'=>0, 'sepc'=>0, 'q3'=>0, 'q3c'=>0,
                     'oct'=>0, 'octc'=>0, 'nov'=>0, 'novc'=>0, 'des'=>0, 'desc'=>0, 'q4'=>0, 'q4c'=>0
            );
            $y++;
        } else {
            $xarr[$j] = $xarr1[$x];
            $yarr[$j] = array("defcol_one_id"=>$xarr1[$x]["defcol_one_id"], "defcol_two_id"=>$xarr1[$x]["defcol_two_id"],
                     'jan'=>0, 'janc'=>0, 'feb'=>0, 'febc'=>0, 'mar'=>0, 'marc'=>0, 'q1'=>0, 'q1c'=>0,
                     'apr'=>0, 'aprc'=>0, 'may'=>0, 'mayc'=>0, 'jun'=>0, 'junc'=>0, 'q2'=>0, 'q2c'=>0,
                     'jul'=>0, 'julc'=>0, 'aug'=>0, 'augc'=>0, 'sep'=>0, 'sepc'=>0, 'q3'=>0, 'q3c'=>0,
                     'oct'=>0, 'octc'=>0, 'nov'=>0, 'novc'=>0, 'des'=>0, 'desc'=>0, 'q4'=>0, 'q4c'=>0
            );
            $x++;
        }
    }


    if (count($xarr) != count($yarr)) {
        echo "ARRAY SIZES DON'T MATCH!!!";
        return 0;
    } else {
        echo "ARRAY SIZES OK!!!";
    }

    $i = 1;
    $objPHPExcel->getActiveSheet()->setCellValue('A1', strtoupper($xarr[0]["defcol_one_id"]));
    $objPHPExcel->getActiveSheet()->mergeCells('A'.$i.':R'.$i);
    $objPHPExcel->getActiveSheet()->getStyle('A1:R1')->applyFromArray($style_total);
    $i++;
    $objPHPExcel->getActiveSheet()->mergeCells('A'.$i.':R'.$i);
    $i++;

    $objPHPExcel->getActiveSheet()->setCellValue('A'.$i, 'Country/Technology')
                                  ->setCellValue('B'.$i, 'Jan')
                                  ->setCellValue('C'.$i, 'Feb')
                                  ->setCellValue('D'.$i, 'Mar')
                                  ->setCellValue('E'.$i, 'Q1')
                                  ->setCellValue('F'.$i, 'Apr')
                                  ->setCellValue('G'.$i, 'May')
                                  ->setCellValue('H'.$i, 'Jun')
                                  ->setCellValue('I'.$i, 'Q2')
                                  ->setCellValue('J'.$i, 'Jul')
                                  ->setCellValue('K'.$i, 'Aug')
                                  ->setCellValue('L'.$i, 'Sep')
                                  ->setCellValue('M'.$i, 'Q3')
                                  ->setCellValue('N'.$i, 'Oct')
                                  ->setCellValue('O'.$i, 'Nov')
                                  ->setCellValue('P'.$i, 'Dec')
                                  ->setCellValue('Q'.$i, 'Q4')
                                  ->setCellValue('R'.$i, 'Total');

    $objPHPExcel->getActiveSheet()->getStyle('A'.$i.':R'.$i)->applyFromArray($style_total);
    $i++;

    $cat = '';
    $cat_rev = array('jan'=>0, 'janc'=>0, 'feb'=>0, 'febc'=>0, 'mar'=>0, 'marc'=>0, 'q1'=>0, 'q1c'=>0,
                     'apr'=>0, 'aprc'=>0, 'may'=>0, 'mayc'=>0, 'jun'=>0, 'junc'=>0, 'q2'=>0, 'q2c'=>0,
                     'jul'=>0, 'julc'=>0, 'aug'=>0, 'augc'=>0, 'sep'=>0, 'sepc'=>0, 'q3'=>0, 'q3c'=>0,
                     'oct'=>0, 'octc'=>0, 'nov'=>0, 'novc'=>0, 'des'=>0, 'desc'=>0, 'q4'=>0, 'q4c'=>0, 'totc'=>0);
    $tot_rev = array('jan'=>0, 'janc'=>0, 'feb'=>0, 'febc'=>0, 'mar'=>0, 'marc'=>0, 'q1'=>0, 'q1c'=>0,
                     'apr'=>0, 'aprc'=>0, 'may'=>0, 'mayc'=>0, 'jun'=>0, 'junc'=>0, 'q2'=>0, 'q2c'=>0,
                     'jul'=>0, 'julc'=>0, 'aug'=>0, 'augc'=>0, 'sep'=>0, 'sepc'=>0, 'q3'=>0, 'q3c'=>0,
                     'oct'=>0, 'octc'=>0, 'nov'=>0, 'novc'=>0, 'des'=>0, 'desc'=>0, 'q4'=>0, 'q4c'=>0, 'totc'=>0);
    $cat_job = array('jan'=>0, 'janc'=>0, 'feb'=>0, 'febc'=>0, 'mar'=>0, 'marc'=>0, 'q1'=>0, 'q1c'=>0,
                     'apr'=>0, 'aprc'=>0, 'may'=>0, 'mayc'=>0, 'jun'=>0, 'junc'=>0, 'q2'=>0, 'q2c'=>0,
                     'jul'=>0, 'julc'=>0, 'aug'=>0, 'augc'=>0, 'sep'=>0, 'sepc'=>0, 'q3'=>0, 'q3c'=>0,
                     'oct'=>0, 'octc'=>0, 'nov'=>0, 'novc'=>0, 'des'=>0, 'desc'=>0, 'q4'=>0, 'q4c'=>0);
    $tot_job = array('jan'=>0, 'janc'=>0, 'feb'=>0, 'febc'=>0, 'mar'=>0, 'marc'=>0, 'q1'=>0, 'q1c'=>0,
                     'apr'=>0, 'aprc'=>0, 'may'=>0, 'mayc'=>0, 'jun'=>0, 'junc'=>0, 'q2'=>0, 'q2c'=>0,
                     'jul'=>0, 'julc'=>0, 'aug'=>0, 'augc'=>0, 'sep'=>0, 'sepc'=>0, 'q3'=>0, 'q3c'=>0,
                     'oct'=>0, 'octc'=>0, 'nov'=>0, 'novc'=>0, 'des'=>0, 'desc'=>0, 'q4'=>0, 'q4c'=>0);





    $date = getdate();
    $today = $date['mon'];
    $mon_format1 = array('jan' => 1, 'feb' => 2, 'mar' => 3, 'apr' => 4, 'may' => 5, 'jun' => 6, 'jul' => 7, 'aug' => 8, 'sep' => 9,
                            'oct' => 10, 'nov' => 11, 'des' => 12);
    $mon_format2 = array(1 => 'janc', 2 => 'febc', 3 => 'marc', 4 => 'aprc', 5 => 'mayc', 6 => 'junc', 7 => 'julc', 8 => 'augc', 9 => 'sepc',
                            10 => 'octc', 11 => 'novc', 12 => 'desc');
    foreach($xarr as $k => $xval) {

        $tot_div = $xval[$mon_format2[$today]];
        foreach($mon_format2 as $mk => $mval){
            if ($xval[$mval] > 0) {
                $tot_div = $xval[$mval];
            } else {
                continue;
            }
        }
        var_dump($tot_div);
        if ($cat=='') {
            $cat = $xval['defcol_one_id'];
            echo "FIRST ASSET: ".$cat." ";
            //fill cat_rev
            foreach ($cat_rev as $id=>$sval){
                if (($id != 'defcol_one_id') && ($id != 'defcol_two_id') && ($id != 'totc')) {
                    $cat_rev[$id] = $xval[$id];
                    $cat_job[$id] = $yarr[$k][$id];
                }
            }
            $cat_rev['totc'] = $tot_div;
            $cat_rev['defcol_one_id'] = $xval['defcol_one_id'];
            $cat_rev['defcol_two_id'] = $xval['defcol_two_id'];
            echo "TOTAL ASSETS IN COUNTRY".$cat_rev['defcol_two_id'][1];

            fill_section($objPHPExcel, $tot_div, $xval, $yarr[$k], $i, $style_normal, $style_stotal);

        }
        else if ($cat == $xval['defcol_one_id']) {
            foreach ($cat_rev as $id=>$sval){
                if (($id != 'defcol_one_id') && ($id != 'defcol_two_id') && ($id != 'totc')) {
                    $cat_rev[$id] += $xval[$id];
                    $cat_job[$id] += $yarr[$k][$id];
                }
            }
            $cat_rev['totc'] += $tot_div;
            $cat_rev['defcol_one_id'] = $xval['defcol_one_id'];
            $cat_rev['defcol_two_id'] = $xval['defcol_two_id'];
            //xls the line

            fill_section($objPHPExcel, $tot_div, $xval, $yarr[$k], $i, $style_normal, $style_stotal);

        }
        else {
            echo "OLD ASSET: ".$cat." ";
            //sum tot_rev
            foreach ($tot_rev as $id=>$sval){
                if (($id != 'defcol_one_id') && ($id != 'defcol_two_id')) {
                    $tot_rev[$id] += $cat_rev[$id];
                    $tot_job[$id] += $cat_job[$id];
                }
            }
            $tot_rev['defcol_one_id'] = $xval['defcol_one_id'];
            $tot_rev['defcol_two_id'] = 'TGT CORPORATE';
            //xls the cat_rev

            echo "TOTAL ASSETS IN COUNTRY".$cat_rev['defcol_two_id'][1];

            $cat_rev['defcol_two_id'][1] = "TGT Corporate";

            fill_section($objPHPExcel, $cat_rev['totc'], $cat_rev, $cat_job, $i, $style_normal, $style_stotal);

            $objPHPExcel->getActiveSheet()->mergeCells('A'.$i.':R'.$i);
            $i++;
            $objPHPExcel->getActiveSheet()->setCellValue('A'.$i, strtoupper($xval["defcol_one_id"]));
            $objPHPExcel->getActiveSheet()->mergeCells('A'.$i.':R'.$i);
            $objPHPExcel->getActiveSheet()->getStyle('A'.$i.':R'.$i)->applyFromArray($style_total);
            $i++;
            $objPHPExcel->getActiveSheet()->mergeCells('A'.$i.':R'.$i);
            $i++;

    $objPHPExcel->getActiveSheet()->setCellValue('A'.$i, 'Country/Technology')
                                  ->setCellValue('B'.$i, 'Jan')
                                  ->setCellValue('C'.$i, 'Feb')
                                  ->setCellValue('D'.$i, 'Mar')
                                  ->setCellValue('E'.$i, 'Q1')
                                  ->setCellValue('F'.$i, 'Apr')
                                  ->setCellValue('G'.$i, 'May')
                                  ->setCellValue('H'.$i, 'Jun')
                                  ->setCellValue('I'.$i, 'Q2')
                                  ->setCellValue('J'.$i, 'Jul')
                                  ->setCellValue('K'.$i, 'Aug')
                                  ->setCellValue('L'.$i, 'Sep')
                                  ->setCellValue('M'.$i, 'Q3')
                                  ->setCellValue('N'.$i, 'Oct')
                                  ->setCellValue('O'.$i, 'Nov')
                                  ->setCellValue('P'.$i, 'Dec')
                                  ->setCellValue('Q'.$i, 'Q4')
                                  ->setCellValue('R'.$i, 'Total');

    $objPHPExcel->getActiveSheet()->getStyle('A'.$i.':R'.$i)->applyFromArray($style_total);
    $i++;

            fill_section($objPHPExcel, $tot_div, $xval, $yarr[$k], $i, $style_normal, $style_stotal);

            $cat = $xval['defcol_one_id'];
            //new cat_rev
            foreach ($cat_rev as $id=>$sval){
                if (($id != 'defcol_one_id') && ($id != 'defcol_two_id') && ($id != 'totc')) {
                    $cat_rev[$id] = $xval[$id];
                    $cat_job[$id] = $yarr[$k][$id];
                }
            }
            $cat_rev['totc'] = $tot_div;
            $cat_rev['defcol_one_id'] = $xval['defcol_one_id'];
            $cat_rev['defcol_two_id'] = $xval['defcol_two_id'];
        }
    } //foreach $xarr
    $last_elem = end($xarr);
    $last_job = end($yarr);

    $tot_div = $last_elem[$mon_format2[$today]];
        foreach($mon_format2 as $mk => $mval){
            if ($xval[$mval] > 0) {
                $tot_div = $last_elem[$mval];
            } else {
                continue;
            }
        }

    foreach ($cat_rev as $id=>$sval){
        if (($id != 'defcol_one_id') && ($id != 'defcol_two_id') && ($id != 'totc')) {
            $cat_rev[$id] += $last_elem[$id];
            $cat_job[$id] += $last_job[$id];
        }
    }
    $cat_rev['totc'] += $tot_div;

    foreach ($tot_rev as $id=>$sval){
        if (($id != 'defcol_one_id') && ($id != 'defcol_two_id')) {
            $tot_rev[$id] += $cat_rev[$id];
            $tot_job[$id] += $cat_job[$id];
        }
    }
    $tot_rev['defcol_one_id'] = $xval['defcol_one_id'];
    $tot_rev['defcol_two_id'] = 'TGT CORPORATE';
    //xls the cat_rev

    $cat_rev['defcol_two_id'][1] = "TGT Corporate";
    fill_section($objPHPExcel, $cat_rev['totc'], $cat_rev, $cat_job, $i, $style_normal, $style_stotal);


    $tot_rev['defcol_two_id'] = array();
    $tot_rev['defcol_two_id'][1] = 'TGT CORPORATE';
    $tot_rev['defcol_one_id'] = 'ALL TOOLS';

    $objPHPExcel->getActiveSheet()->mergeCells('A'.$i.':R'.$i);
    $i++;
    $objPHPExcel->getActiveSheet()->setCellValue('A'.$i, $tot_rev['defcol_one_id']);
    $objPHPExcel->getActiveSheet()->mergeCells('A'.$i.':R'.$i);
    $objPHPExcel->getActiveSheet()->getStyle('A'.$i.':R'.$i)->applyFromArray($style_total);
    $i++;
    $objPHPExcel->getActiveSheet()->mergeCells('A'.$i.':R'.$i);
    $i++;

            echo "TOTAL ASSETS IN COUNTRY".$tot_rev['defcol_two_id'][1];
            var_dump($tot_rev['totc']);

    fill_section($objPHPExcel, $tot_rev['totc'], $tot_rev, $tot_job, $i, $style_stotal, $style_total);

    //revenue

    $objPHPExcel->getActiveSheet()->getStyle('E1:E'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_stotal);
    $objPHPExcel->getActiveSheet()->getStyle('I1:I'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_stotal);
    $objPHPExcel->getActiveSheet()->getStyle('M1:M'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_stotal);
    $objPHPExcel->getActiveSheet()->getStyle('Q1:Q'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_stotal);
    $objPHPExcel->getActiveSheet()->getStyle('R1:R'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_total);

    $objPHPExcel->getActiveSheet()->getStyle(
        'A1:' .
        $objPHPExcel->getActiveSheet()->getHighestColumn().
        $objPHPExcel->getActiveSheet()->getHighestRow()
    )->applyFromArray($style_borders);

    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, "Excel2007");
    $objWriter->save("/home/webmaster/tgtoil.com/wp-content/themes/MyCuisine/xmlrpc_lib/xls_files/".$fname.".xlsx");
}


function save_au_report_excel($xarr, $fname,$style_header, $style_borders, $style_normal, $style_ctitle, $style_stotal, $style_total) {

        foreach ($xarr as $key => $val) {
            $asset_arr[$key] = $val["defcol_one_id"];
            $cou_arr[$key] = $val["defcol_two_id"][1];
        }

        array_multisort($cou_arr, SORT_ASC, $asset_arr, SORT_ASC, $xarr);

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
                                  ->setCellValue('R1', 'Total');

    $cat = '';
    $cat_rev = array('jan'=>0, 'janc'=>0, 'feb'=>0, 'febc'=>0, 'mar'=>0, 'marc'=>0, 'q1'=>0, 'q1c'=>0,
                     'apr'=>0, 'aprc'=>0, 'may'=>0, 'mayc'=>0, 'jun'=>0, 'junc'=>0, 'q2'=>0, 'q2c'=>0,
                     'jul'=>0, 'julc'=>0, 'aug'=>0, 'augc'=>0, 'sep'=>0, 'sepc'=>0, 'q3'=>0, 'q3c'=>0,
                     'oct'=>0, 'octc'=>0, 'nov'=>0, 'novc'=>0, 'des'=>0, 'desc'=>0, 'q4'=>0, 'q4c'=>0);
    $tot_rev = array('jan'=>0, 'janc'=>0, 'feb'=>0, 'febc'=>0, 'mar'=>0, 'marc'=>0, 'q1'=>0, 'q1c'=>0,
                     'apr'=>0, 'aprc'=>0, 'may'=>0, 'mayc'=>0, 'jun'=>0, 'junc'=>0, 'q2'=>0, 'q2c'=>0,
                     'jul'=>0, 'julc'=>0, 'aug'=>0, 'augc'=>0, 'sep'=>0, 'sepc'=>0, 'q3'=>0, 'q3c'=>0,
                     'oct'=>0, 'octc'=>0, 'nov'=>0, 'novc'=>0, 'des'=>0, 'desc'=>0, 'q4'=>0, 'q4c'=>0);

    $date = getdate();
    $today = $date['mon'];
    $mon_format1 = array('jan' => 1, 'feb' => 2, 'mar' => 3, 'apr' => 4, 'may' => 5, 'jun' => 6, 'jul' => 7, 'aug' => 8, 'sep' => 9,
                            'oct' => 10, 'nov' => 11, 'des' => 12);
    $mon_format2 = array(1 => 'janc', 2 => 'febc', 3 => 'marc', 4 => 'aprc', 5 => 'mayc', 6 => 'junc', 7 => 'julc', 8 => 'augc', 9 => 'sepc',
                            10 => 'octc', 11 => 'novc', 12 => 'desc');
    $i=2;
    foreach($xarr as $k => $xval) {

        if ($cat=='') {
            $cat = $xval['defcol_two_id'][1];
            //fill cat_rev
            foreach ($cat_rev as $id=>$sval){
                $cat_rev[$id] = $xval[$id];
            }
            //xls the line
            $tot_div = $xval[$mon_format2[$today]];
            $total_col = ($xval['q1']+$xval['q2']+$xval['q3']+$xval['q4'])/(($tot_div != 0) ? $tot_div : 1);
            fill_row($objPHPExcel, $xval, $total_col, strtoupper($xval['defcol_one_id']), $i, $style_normal);
            $i++;
        }
        else if ($cat == $xval['defcol_two_id'][1]) {
            //sum cat_rev
            foreach ($cat_rev as $id=>$sval){
                $cat_rev[$id] += $xval[$id];
            }
            //xls the line
            $tot_div = $xval[$mon_format2[$today]];
            $total_col = ($xval['q1']+$xval['q2']+$xval['q3']+$xval['q4'])/(($tot_div != 0) ? $tot_div : 1);
            fill_row($objPHPExcel, $xval, $total_col, strtoupper($xval['defcol_one_id']), $i, $style_normal);
            $i++;
        }
        else {
            //sum tot_rev
            foreach ($tot_rev as $id=>$sval){
                $tot_rev[$id] += $cat_rev[$id];
            }
            //xls the cat_rev
            $tot_div = $xval[$mon_format2[$today]];
            $total_col = ($cat_rev['q1']+$cat_rev['q2']+$cat_rev['q3']+$cat_rev['q4'])/(($tot_div != 0) ? $tot_div : 1);
            fill_row($objPHPExcel, $cat_rev, $total_col, strtoupper($cat), $i, $style_stotal);
            $i++;
            $tot_div = $xval[$mon_format2[$today]];
            $total_col = ($xval['q1']+$xval['q2']+$xval['q3']+$xval['q4'])/(($tot_div != 0) ? $tot_div : 1);
            fill_row($objPHPExcel, $xval, $total_col, strtoupper($xval['defcol_one_id']), $i, $style_normal);
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
        $tot_rev[$id] += $last_elem[$id];
    }
    //xls cat and grand total
                $tot_div = $xval[$mon_format2[$today]];
                $total_col = ($cat_rev['q1']+$cat_rev['q2']+$cat_rev['q3']+$cat_rev['q4'])/(($tot_div != 0) ? $tot_div : 1);
                fill_row($objPHPExcel, $cat_rev, $total_col, strtoupper($cat), $i, $style_stotal);
                $i++;
    $tot_div = $xval[$mon_format2[$today]];
    $total_col = ($tot_rev['q1']+$tot_rev['q2']+$tot_rev['q3']+$tot_rev['q4'])/(($tot_div != 0) ? $tot_div : 1);

    fill_row($objPHPExcel, $tot_rev, $total_col, 'TOTAL', $i, $style_total);

    $objPHPExcel->getActiveSheet()->getStyle('A1:A'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_ctitle);

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
    $objWriter->save("/home/webmaster/tgtoil.com/wp-content/themes/MyCuisine/xmlrpc_lib/xls_files/Asset_Util_by_".$fname.".xlsx");
}

function fill_row($objPHPExcel, $tot_rev, $total_col, $tot_div, $name, $i, $style) {
    $objPHPExcel->getActiveSheet()->setCellValue('A'.$i, $name)
                                          ->setCellValue('B'.$i, $tot_rev['jan']/(($tot_rev['janc'])? $tot_rev['janc'] : 1))
                                          ->setCellValue('C'.$i, $tot_rev['feb']/(($tot_rev['febc'])? $tot_rev['febc'] : 1))
                                          ->setCellValue('D'.$i, $tot_rev['mar']/(($tot_rev['marc'])? $tot_rev['marc'] : 1))
                                          ->setCellValue('E'.$i, $tot_rev['q1']/(($tot_rev['marc'])? $tot_rev['marc'] : 1))
                                          ->setCellValue('F'.$i, $tot_rev['apr']/(($tot_rev['aprc'])? $tot_rev['aprc'] : 1))
                                          ->setCellValue('G'.$i, $tot_rev['may']/(($tot_rev['mayc'])? $tot_rev['mayc'] : 1))
                                          ->setCellValue('H'.$i, $tot_rev['jun']/(($tot_rev['junc'])? $tot_rev['junc'] : 1))
                                          ->setCellValue('I'.$i, $tot_rev['q2']/(($tot_rev['junc'])? $tot_rev['junc'] : 1))
                                          ->setCellValue('J'.$i, $tot_rev['jul']/(($tot_rev['julc'])? $tot_rev['julc'] : 1))
                                          ->setCellValue('K'.$i, $tot_rev['aug']/(($tot_rev['augc'])? $tot_rev['augc'] : 1))
                                          ->setCellValue('L'.$i, $tot_rev['sep']/(($tot_rev['sepc'])? $tot_rev['sepc'] : 1))
                                          ->setCellValue('M'.$i, $tot_rev['q3']/(($tot_rev['sepc'])? $tot_rev['sepc'] : 1))
                                          ->setCellValue('N'.$i, $tot_rev['oct']/(($tot_rev['octc'])? $tot_rev['octc'] : 1))
                                          ->setCellValue('O'.$i, $tot_rev['nov']/(($tot_rev['novc'])? $tot_rev['novc'] : 1))
                                          ->setCellValue('P'.$i, $tot_rev['des']/(($tot_rev['desc'])? $tot_rev['desc'] : 1))
                                          ->setCellValue('Q'.$i, $tot_rev['q4']/(($tot_rev['desc'])? $tot_rev['desc'] : 1))
                                          ->setCellValue('R'.$i, $total_col/($tot_div ? $tot_div : 1));

    if ($style != 0) $objPHPExcel->getActiveSheet()->getStyle('A'.$i.':R'.$i)->applyFromArray($style);
    $objPHPExcel->getActiveSheet()->getStyle(('A'.$i.':R'.$i))->getNumberFormat()->setFormatCode('#,##0.00_-');
}

function fill_row3($objPHPExcel, $tot_rev, $total_col, $tot_div, $name, $i, $style) {
    $objPHPExcel->getActiveSheet()->setCellValue('A'.$i, $name)
                                          ->setCellValue('B'.$i, num_form($tot_rev['jan']/(($tot_rev['janc'])? $tot_rev['janc'] : 1)))
                                          ->setCellValue('C'.$i, num_form($tot_rev['feb']/(($tot_rev['febc'])? $tot_rev['febc'] : 1)))
                                          ->setCellValue('D'.$i, num_form($tot_rev['mar']/(($tot_rev['marc'])? $tot_rev['marc'] : 1)))
                                          ->setCellValue('E'.$i, num_form($tot_rev['q1']/(($tot_rev['marc'])? $tot_rev['marc'] : 1)))
                                          ->setCellValue('F'.$i, num_form($tot_rev['apr']/(($tot_rev['aprc'])? $tot_rev['aprc'] : 1)))
                                          ->setCellValue('G'.$i, num_form($tot_rev['may']/(($tot_rev['mayc'])? $tot_rev['mayc'] : 1)))
                                          ->setCellValue('H'.$i, num_form($tot_rev['jun']/(($tot_rev['junc'])? $tot_rev['junc'] : 1)))
                                          ->setCellValue('I'.$i, num_form($tot_rev['q2']/(($tot_rev['junc'])? $tot_rev['junc'] : 1)))
                                          ->setCellValue('J'.$i, num_form($tot_rev['jul']/(($tot_rev['julc'])? $tot_rev['julc'] : 1)))
                                          ->setCellValue('K'.$i, num_form($tot_rev['aug']/(($tot_rev['augc'])? $tot_rev['augc'] : 1)))
                                          ->setCellValue('L'.$i, num_form($tot_rev['sep']/(($tot_rev['sepc'])? $tot_rev['sepc'] : 1)))
                                          ->setCellValue('M'.$i, num_form($tot_rev['q3']/(($tot_rev['sepc'])? $tot_rev['sepc'] : 1)))
                                          ->setCellValue('N'.$i, num_form($tot_rev['oct']/(($tot_rev['octc'])? $tot_rev['octc'] : 1)))
                                          ->setCellValue('O'.$i, num_form($tot_rev['nov']/(($tot_rev['novc'])? $tot_rev['novc'] : 1)))
                                          ->setCellValue('P'.$i, num_form($tot_rev['des']/(($tot_rev['desc'])? $tot_rev['desc'] : 1)))
                                          ->setCellValue('Q'.$i, num_form($tot_rev['q4']/(($tot_rev['desc'])? $tot_rev['desc'] : 1)))
                                          ->setCellValue('R'.$i, num_form($total_col/($tot_div ? $tot_div : 1)));

    if ($style != 0) $objPHPExcel->getActiveSheet()->getStyle('A'.$i.':R'.$i)->applyFromArray($style);
}

function fill_row1($objPHPExcel, $tot_rev, $total_col, $name, $i, $style) {
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

function fill_row2($objPHPExcel, $tot_rev, $total_col, $name, $i, $style) {
    $objPHPExcel->getActiveSheet()->setCellValue('A'.$i, $name)
                                          ->setCellValue('B'.$i, num_form($tot_rev['janc']))
                                          ->setCellValue('C'.$i, num_form($tot_rev['febc']))
                                          ->setCellValue('D'.$i, num_form($tot_rev['marc']))
                                          ->setCellValue('E'.$i, num_form($tot_rev['marc']))
                                          ->setCellValue('F'.$i, num_form($tot_rev['aprc']))
                                          ->setCellValue('G'.$i, num_form($tot_rev['mayc']))
                                          ->setCellValue('H'.$i, num_form($tot_rev['junc']))
                                          ->setCellValue('I'.$i, num_form($tot_rev['junc']))
                                          ->setCellValue('J'.$i, num_form($tot_rev['julc']))
                                          ->setCellValue('K'.$i, num_form($tot_rev['augc']))
                                          ->setCellValue('L'.$i, num_form($tot_rev['sepc']))
                                          ->setCellValue('M'.$i, num_form($tot_rev['sepc']))
                                          ->setCellValue('N'.$i, num_form($tot_rev['octc']))
                                          ->setCellValue('O'.$i, num_form($tot_rev['novc']))
                                          ->setCellValue('P'.$i, num_form($tot_rev['des']))
                                          ->setCellValue('Q'.$i, num_form($tot_rev['desc']))
                                          ->setCellValue('R'.$i, num_form($total_col));

    if ($style != 0) $objPHPExcel->getActiveSheet()->getStyle('A'.$i.':R'.$i)->applyFromArray($style);
}

function fill_section($objPHPExcel, $tot_div, $xval, $yval, &$i, $style1, $style2) {
            fill_row2($objPHPExcel, $xval, $tot_div, strtoupper($xval["defcol_one_id"])." - QUANTITY", $i, $style1); //for the number of tools
            $i++;
            $total_col1 = $yval['q1']+$yval['q2']+$yval['q3']+$yval['q4'];
            fill_row1($objPHPExcel, $yval, $total_col1, strtoupper($xval["defcol_one_id"])." - JOBS", $i, $style1);
            $i++;
            $total_col2 = $xval['q1']+$xval['q2']+$xval['q3']+$xval['q4'];
            fill_row1($objPHPExcel, $xval, $total_col2, strtoupper($xval["defcol_one_id"])." REVENUE", $i, $style1); //revenue
            $i++;
            fill_row($objPHPExcel, $yval, $total_col1, $tot_div, strtoupper($xval["defcol_two_id"][1])." - JOBS / TOOL", $i, $style2);
            $i++;
            fill_row3($objPHPExcel, $xval, $total_col2, $tot_div, strtoupper($xval["defcol_two_id"][1])." - REVENUE / ASSET, $", $i, $style2); //revenue
            //merge the five cells
            //$objPHPExcel->getActiveSheet()->mergeCells('A'.($i-4).':A'.$i);
            $i++;
            $objPHPExcel->getActiveSheet()->mergeCells('A'.$i.':R'.$i);
            $i++;
}

function num_form($number) {
    if ($number != 0)  {
        return number_format($number,0,'.',',');
    }
    else {
        return '';
    }
}

?>