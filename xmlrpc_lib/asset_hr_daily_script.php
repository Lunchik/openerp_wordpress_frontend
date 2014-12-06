<?php

include 'xmlrpc.inc';
/** Include PHPExcel **/
require_once "/home/webmaster/tgtoil.com/wp-includes/utilites/ExClasses/PHPExcel.php";
require_once 'report_class.php';
require_once 'xml_styles.php';



/**********
HR Reports - connect, get data, generate file
First - Vacation balance
**********/
echo "Hello!";
$VB_Report = new Report;
$relation_vb = array("hr.vacation.report", "hr_tgt.rotation.method");
$required_keys_vb = array(
            array(
                       new xmlrpcval("employee_id", "array"),
                       new xmlrpcval("country_id", "array"),
                       new xmlrpcval("rotation", "array"),
                       new xmlrpcval("is_rotator", "boolean"),
                       new xmlrpcval("land", "int"),
                       new xmlrpcval("sea", "int"),
                       new xmlrpcval("base", "int"),
                       new xmlrpcval("dayoff", "int"),
                       new xmlrpcval("vacation", "int")
            ),
            array(
                       new xmlrpcval("id", "int"),
                       new xmlrpcval("days_off", "int"),
                       new xmlrpcval("days_work", "int"),

            )
    );
$search_keys_vb = array(
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
$VB_Report->constructor('rpc_user', 'rpc_password', 'db_name', 'xmlrpc_interface_url', $relation_vb, $required_keys_vb, $search_keys_vb);
$VB_Report->connect_xmlrpc();
for ($i=0; $i < count($VB_Report->relation_names); $i++){
    $VB_Report->get_xml_data($i);
}

var_dump($VB_Report->data_arr);

$VBalance = save_hr_report_excel1($VB_Report->data_arr, $style_header, $style_borders, $style_normal, $style_ctitle, $style_stotal, $style_total);
function save_hr_report_excel1($varr, $style_header, $style_borders, $style_normal, $style_ctitle, $style_stotal, $style_total) {

    $xarr=$varr[0];

        foreach ($xarr as $key => &$val) {
                if (!$val["country_id"]) {
                    $val["country_id"] = array( 0 => (int)-1, 1 => '');
                    $loc_arr[$key] = $val["country_id"][1];
                    $name_arr[$key] = $val["employee_id"][1];
                }
                else {
                    $loc_arr[$key] = $val["country_id"][1];
                    $name_arr[$key] = $val["employee_id"][1];
                }
        }

        array_multisort($loc_arr, SORT_ASC, $name_arr, SORT_ASC, $xarr);

    // Create new PHPExcel object

    $objPHPExcel = new PHPExcel();

    // Set document properties
    $objPHPExcel->getProperties()->setCreator("OpenERP")
    							 ->setLastModifiedBy("OpenERP")
    							 ->setTitle("Asset Report")
    							 ->setSubject("Office 2007 XLSX Test Document")
    							 ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
    							 ->setKeywords("office 2007 openxml php")
    							 ->setCategory("Test result file");

    // Set default font
    $objPHPExcel->getDefaultStyle()->getFont()->setName('Arial')
                                              ->setSize(10);

    // Add some data, resembling some different data types


    $objPHPExcel->removeSheetByIndex(0);
        $loc_name = '';
        $loc_num = 0;

    $VBalance = array();
    $first_time = 1;

    $j = 0;
    while ($xarr[$j]["country_id"][0] == -1) {
        unset($xarr[$j]);
        $j++;
    }
    $j=0;

    $yarr = array();
    foreach($xarr as $key => $val) {
        $yarr[$j] = $val;
        $j++;
    }

    $xarr = $yarr;

    for($k = 0; $k < count($xarr); $k++) {

        $loc_name = $xarr[$k]["country_id"][1];

        $myWorkSheet = new PHPExcel_Worksheet($objPHPExcel, $loc_name);

        $myWorkSheet->setCellValue('A1', 'Location')
                                            ->setCellValue('B1', $loc_name);

        $myWorkSheet->setCellValue('A3', 'Employee')
                          ->setCellValue('B3', 'Status')
                          ->setCellValue('C3', 'Rotation Method')
                          ->setCellValue('D3', 'Vacation Days')
                          ->setCellValue('E3', 'Vacation Balance');

        $i = 4;

        while (isset($xarr[$k]) && ($loc_name == $xarr[$k]["country_id"][1])) {
            if ($xarr[$k]["rotation"]) {
                $rot_data = $varr[1][$xarr[$k]["rotation"][0]-1];
                $vbalance = ($xarr[$k]["base"] + $xarr[$k]["land"] + $xarr[$k]["sea"] + $xarr[$k]["vacation"])*($rot_data["days_off"]/$rot_data["days_work"]) - $xarr[$k]["dayoff"];
                $VBalance[$xarr[$k]["employee_id"][1]] = ($xarr[$k]["is_rotator"]) ? $vbalance : ($xarr[$k]["an_leaves"] - $xarr[$k]["vacation"]);
                $myWorkSheet->setCellValue('A'.$i, $xarr[$k]["employee_id"][1])
                                  ->setCellValue('B'.$i, ($xarr[$k]["is_rotator"]) ? "Rotator" : "Resident")
                                  ->setCellValue('C'.$i, $xarr[$k]["rotation"][1])
                                  ->setCellValue('D'.$i, ($xarr[$k]["is_rotator"]) ? $xarr[$k]["vacation"] : "--")
                                  ->setCellValue('E'.$i, ($xarr[$k]["rotation"]) ? $vbalance : ($xarr[$k]["an_leaves"] - $xarr[$k]["vacation"]));
            } else {
                if ($xarr[$k]["is_rotator"]) $VBalance[$xarr[$k]["employee_id"][1]] = ($xarr[$k]["an_leaves"] - $xarr[$k]["vacation"]);
                $myWorkSheet->setCellValue('A'.$i, $xarr[$k]["employee_id"][1])
                                  ->setCellValue('B'.$i, ($xarr[$k]["is_rotator"]) ? "Rotator" : "Resident")
                                  ->setCellValue('C'.$i, "--")
                                  ->setCellValue('D'.$i, ($xarr[$k]["is_rotator"]) ? $xarr[$k]["vacation"] : "--")
                                  ->setCellValue('E'.$i, ($xarr[$k]["an_leaves"] - $xarr[$k]["vacation"]));
            }
            $i++;
            $k++;
        }
        $k--;
        $loc_num ++;

        $objPHPExcel->addSheet($myWorkSheet);
        $myWorkSheet->getStyle('A1:'.$myWorkSheet->getHighestColumn().$myWorkSheet->getHighestRow())->applyFromArray($style_normal);
        $myWorkSheet->getStyle('A3:A'.$myWorkSheet->getHighestRow())->applyFromArray($style_ctitle);
        $myWorkSheet->getStyle('A3:'.$myWorkSheet->getHighestColumn().'3')->applyFromArray($style_total);
        $myWorkSheet->getStyle(
            'B4:' .
            $myWorkSheet->getHighestColumn().
            $myWorkSheet->getHighestRow()
        )->getNumberFormat()->setFormatCode('#,##0.00');

        $myWorkSheet->getStyle(
            'A3:' .
            $myWorkSheet->getHighestColumn().
            $myWorkSheet->getHighestRow()
        )->applyFromArray($style_borders);

    } //foreach



    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, "Excel2007");
    $objWriter->save("/home/webmaster/tgtoil.com/wp-content/themes/MyCuisine/xmlrpc_lib/xls_files/hr_vacation_balance.xlsx");

    return $VBalance;

}

$HL_Report = new Report;
$rel_root = "hr.loading.report";
$relation_hl = array();
$search_keys_hl = array();
for ($i=1; $i <=12; $i++) {
    $relation_hl[] = $rel_root.$i;
    $search_keys_hl[] = array( //for hr.vacation.report
                              new xmlrpcval(array(new xmlrpcval("id" , "string"),
                              new xmlrpcval("!=","string"),
                              new xmlrpcval("-1","string")),"array"),
                        );
    $required_keys_hl[] = array(
                                new xmlrpcval("employee_id", "array"),
                                new xmlrpcval("res_country", "array"),
                                new xmlrpcval("land", "int"),
                                new xmlrpcval("sea", "int"),
                                new xmlrpcval("base", "int"),
                                new xmlrpcval("dayoff", "int"),
                          );
}
$required_keys_hl = array(
            array(
                       new xmlrpcval("employee_id", "array"),
                       new xmlrpcval("res_country", "array"),
                       new xmlrpcval("land", "int"),
                       new xmlrpcval("sea", "int"),
                       new xmlrpcval("base", "int"),
                       new xmlrpcval("dayoff", "int"),
            )
    );
$HL_Report->constructor('rpc_user', 'rpc_password', 'db_name', 'xmlrpc_interface_url', $relation_hl, $required_keys_hl, $search_keys_hl);
$HL_Report->connect_xmlrpc();
for ($i=0; $i < count($HL_Report->relation_names); $i++){
    $HL_Report->get_xml_data($i);
}

save_hr_report_excel2($HL_Report->data_arr, $VBalance, $style_header, $style_borders, $style_normal, $style_ctitle, $style_stotal, $style_total);
function save_hr_report_excel2($varr, $VBalance, $style_header, $style_borders, $style_normal, $style_ctitle, $style_stotal, $style_total) {

    // Create new PHPExcel object
    $xarr = $varr[0];

    //echo " Create new PHPExcel object" ;
    $objPHPExcel = new PHPExcel();

    // Set document properties
    $objPHPExcel->getProperties()->setCreator("OpenERP")
    							 ->setLastModifiedBy("OpenERP")
    							 ->setTitle("Asset Report")
    							 ->setSubject("Office 2007 XLSX Test Document")
    							 ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
    							 ->setKeywords("office 2007 openxml php")
    							 ->setCategory("Test result file");

    // Set default font
    $objPHPExcel->getDefaultStyle()->getFont()->setName('Arial')
                                              ->setSize(10);

    $mformat = array( 1 => 'Jan', 2 => 'Feb', 3 => 'Mar', 4 => 'Apr', 5 => 'May', 6 => 'Jun', 7 => 'Jul', 8 => 'Aug',
                      9 => 'Sep', 10 => 'Oct', 11 => 'Nov', 12 => 'Dec' );

    $objPHPExcel->removeSheetByIndex(0);

    foreach($varr as $key => $xarr) {
        //add new worksheet for the month
        $mon = $mformat[$key+1];
        $myWorkSheet = new PHPExcel_Worksheet($objPHPExcel, $mon);
        $objPHPExcel->addSheet($myWorkSheet);

        //sort the xarr
        foreach ($xarr as $key => $val) {
                $loc_arr[$key] = $val["res_country"][1];
                $name_arr[$key] = $val["employee_id"][1];
        }

        array_multisort($loc_arr, SORT_ASC, $name_arr, SORT_ASC, $xarr);

        var_dump($xarr);

        $myWorkSheet->setCellValue('A1', 'Employee Name')
                                          ->setCellValue('B1', $mon)
                                          ->setCellValue('F1', 'Vacation Balance');
        $myWorkSheet->mergeCells('A1:A2')
                                          ->mergeCells('B1:E1')
                                          ->mergeCells('F1:F2');
        $myWorkSheet->setCellValue('B2', 'Base')
                                          ->setCellValue('C2', 'Offshore')
                                          ->setCellValue('D2', 'Onshore')
                                          ->setCellValue('E2', 'Days Off');

        $cou = '';
        $cou_tot = array("base" => 0, "sea" => 0, "land" => 0, "dayoff" => 0, "balance" => 0);
        $hcount = 0;
        $total = array("base" => 0, "sea" => 0, "land" => 0, "dayoff" => 0, "balance" => 0);
        $htotal = 0;
        $i = 4;
        $first_time = 1;
        foreach($xarr as $kk => $xval){

            if ($first_time) {
                $first_time = 0;
                $cou = $xval["res_country"][1];
                //first elem - print the employee
                $myWorkSheet->setCellValue('A'.$i, $xval["employee_id"][1])
                                      ->setCellValue('B'.$i, $xval["base"])
                                      ->setCellValue('C'.$i, $xval["sea"])
                                      ->setCellValue('D'.$i, $xval["land"])
                                      ->setCellValue('E'.$i, $xval["dayoff"])
                                      ->setCellValue('F'.$i, $VBalance[$xval["employee_id"][1]]);
                //init the required values
                foreach ($cou_tot as $id=>$sval){
                    $cou_tot[$id] = (isset($xval[$id])) ? $xval[$id] : 0;
                }
                $cou_tot["balance"] = $VBalance[$xval["employee_id"][1]];
                //$cou_tot = $xval;
                $hcount++;
                $i++;
            } else if ($cou == $xval["res_country"][1]) {
                //print the employee
                $myWorkSheet->setCellValue('A'.$i, $xval["employee_id"][1])
                                      ->setCellValue('B'.$i, $xval["base"])
                                      ->setCellValue('C'.$i, $xval["sea"])
                                      ->setCellValue('D'.$i, $xval["land"])
                                      ->setCellValue('E'.$i, $xval["dayoff"])
                                      ->setCellValue('F'.$i, $VBalance[$xval["employee_id"][1]]);
                //sum up and increment required values
                foreach ($cou_tot as $id=>$sval){
                    $cou_tot[$id] += (isset($xval[$id])) ? $xval[$id] : 0;
                }
                $cou_tot["balance"] += $VBalance[$xval["employee_id"][1]];
                $hcount++;
                $i++;
            } else {
                //new country - print the old country, nullify the $cou value
                $myWorkSheet->setCellValue('A'.$i, 'Total: '.$cou)
                                      ->setCellValue('B'.$i, $cou_tot["base"])
                                      ->setCellValue('C'.$i, $cou_tot["sea"])
                                      ->setCellValue('D'.$i, $cou_tot["land"])
                                      ->setCellValue('E'.$i, $cou_tot["dayoff"])
                                      ->setCellValue('F'.$i, $cou_tot["balance"]);
                $i++;
                $myWorkSheet->setCellValue('A'.$i, 'Headcount: '.$cou) //
                                      ->setCellValue('B'.$i, $hcount);
                $myWorkSheet->mergeCells('B'.$i.':F'.$i);
                $i++;
                $myWorkSheet->setCellValue('A'.$i, 'Per head: '.$cou) //
                                      ->setCellValue('B'.$i, $cou_tot["base"]/$hcount)
                                      ->setCellValue('C'.$i, $cou_tot["sea"]/$hcount)
                                      ->setCellValue('D'.$i, $cou_tot["land"]/$hcount)
                                      ->setCellValue('E'.$i, $cou_tot["dayoff"]/$hcount)
                                      ->setCellValue('F'.$i, $cou_tot["balance"]/$hcount);
                $myWorkSheet->getStyle('A'.($i-2).':'.$myWorkSheet->getHighestColumn().$i)->applyFromArray($style_total);
                foreach ($total as $id=>$sval){
                    $total[$id] += (isset($cou_tot[$id])) ? $cou_tot[$id] : 0;
                }
                $htotal += $hcount;
                $i++;
                //print the employee, increment and sum all required values
                $cou = $xval["res_country"][1];
                $myWorkSheet->setCellValue('A'.$i, $xval["employee_id"][1])
                                      ->setCellValue('B'.$i, $xval["base"])
                                      ->setCellValue('C'.$i, $xval["sea"])
                                      ->setCellValue('D'.$i, $xval["land"])
                                      ->setCellValue('E'.$i, $xval["dayoff"])
                                      ->setCellValue('F'.$i, $VBalance[$xval["employee_id"][1]]);
                foreach ($cou_tot as $id=>$sval){
                    $cou_tot[$id] = (isset($xval[$id])) ? $xval[$id] : 0;
                }
                $cou_tot["balance"] = $VBalance[$xval["employee_id"][1]];
                $hcount = 1;
                $i++;
            }

        }

        //last country
        $xval = end($xarr);
                $myWorkSheet->setCellValue('A'.$i, 'Total: '.$cou)
                                      ->setCellValue('B'.$i, $cou_tot["base"])
                                      ->setCellValue('C'.$i, $cou_tot["sea"])
                                      ->setCellValue('D'.$i, $cou_tot["land"])
                                      ->setCellValue('E'.$i, $cou_tot["dayoff"])
                                      ->setCellValue('F'.$i, $cou_tot["balance"]);
                $i++;
                $myWorkSheet->setCellValue('A'.$i, 'Headcount: '.$cou) //
                                      ->setCellValue('B'.$i, $hcount);
                $myWorkSheet->mergeCells('B'.$i.':F'.$i);
                $i++;
                $myWorkSheet->setCellValue('A'.$i, 'Per head: '.$cou) //
                                      ->setCellValue('B'.$i, $cou_tot["base"]/$hcount)
                                      ->setCellValue('C'.$i, $cou_tot["sea"]/$hcount)
                                      ->setCellValue('D'.$i, $cou_tot["land"]/$hcount)
                                      ->setCellValue('E'.$i, $cou_tot["dayoff"]/$hcount)
                                      ->setCellValue('F'.$i, $cou_tot["balance"]/$hcount);
        $myWorkSheet->getStyle('A'.($i-2).':'.$myWorkSheet->getHighestColumn().$i)->applyFromArray($style_total);
        $i++;
        foreach ($total as $id=>$sval){
            $total[$id] += (isset($cou_tot[$id])) ? $cou_tot[$id] : 0;
        }
        $htotal += $hcount;
        $i++;
        $myWorkSheet->setCellValue('A'.$i, "TOTAL")
                                      ->setCellValue('B'.$i, $total["base"])
                                      ->setCellValue('C'.$i, $total["sea"])
                                      ->setCellValue('D'.$i, $total["land"])
                                      ->setCellValue('E'.$i, $total["dayoff"])
                                      ->setCellValue('F'.$i, $total["balance"]);
        $i++;
        $myWorkSheet->setCellValue('A'.$i, 'Headcount: TOTAL') //
                                      ->setCellValue('B'.$i, $htotal);
        $myWorkSheet->mergeCells('B'.$i.':F'.$i);
        $i++;
        $myWorkSheet->setCellValue('A'.$i, 'Per head: TOTAL') //
                                      ->setCellValue('B'.$i, $total["base"]/$htotal)
                                      ->setCellValue('C'.$i, $total["sea"]/$htotal)
                                      ->setCellValue('D'.$i, $total["land"]/$htotal)
                                      ->setCellValue('E'.$i, $total["dayoff"]/$htotal)
                                      ->setCellValue('F'.$i, $total["balance"]/$htotal);

        $myWorkSheet->getStyle('A1:'.$myWorkSheet->getHighestColumn().$myWorkSheet->getHighestRow())->applyFromArray($style_normal);
        $myWorkSheet->getStyle('A'.($i-2).':'.$myWorkSheet->getHighestColumn().$i)->applyFromArray($style_total);

        $myWorkSheet->getStyle('A3:A'.$myWorkSheet->getHighestRow())->applyFromArray($style_ctitle);
        $myWorkSheet->getStyle('A1:'.$myWorkSheet->getHighestColumn().'2')->applyFromArray($style_total);
        $myWorkSheet->getStyle('B1:E1')->applyFromArray(array( 'alignment' => array(
                                                           'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                                                       )));
        $myWorkSheet->getStyle(
            'B3:' .
            $myWorkSheet->getHighestColumn().
            $myWorkSheet->getHighestRow()
        )->getNumberFormat()->setFormatCode('#,##0.00');

        $myWorkSheet->getStyle(
            'A1:'.
            $myWorkSheet->getHighestColumn().
            $myWorkSheet->getHighestRow()
        )->applyFromArray($style_borders);

    } //end of foreach sheet

    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, "Excel2007");
    $objWriter->save("/home/webmaster/tgtoil.com/wp-content/themes/MyCuisine/xmlrpc_lib/xls_files/hr_loading.xlsx");

}


/**********
HR Reports - connect, get data, generate file
First - Vacation balance
**********/
echo "Hello!";
$HR_Report = new Report;
$relation_vb = array("hr.consolidated.report0", "hr_tgt.rotation.method");
$required_keys_vb = array(
            array(
                       new xmlrpcval("employee_id", "array")/*,
                       new xmlrpcval("country_id", "array"),
                       new xmlrpcval("rotation", "array"),
                       new xmlrpcval("is_rotator", "boolean"),
                       new xmlrpcval("land", "int"),
                       new xmlrpcval("sea", "int"),
                       new xmlrpcval("base", "int"),
                       new xmlrpcval("dayoff", "int"),
                       new xmlrpcval("vacation", "int")*/
            ),
            array(
                       new xmlrpcval("id", "int"),
                       new xmlrpcval("days_off", "int"),
                       new xmlrpcval("days_work", "int"),

            )
    );
$search_keys_vb = array(
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
$HR_Report->constructor('rpc_user', 'rpc_password', 'db_name', 'xmlrpc_interface_url', $relation_vb, $required_keys_vb, $search_keys_vb);
$HR_Report->connect_xmlrpc();
for ($i=0; $i < count($HR_Report->relation_names); $i++){
    $HR_Report->get_xml_data($i);

}

//var_dump($VB_Report->data_arr);

save_hr_report_excel0($HR_Report->data_arr, $style_header, $style_borders, $style_normal, $style_ctitle, $style_stotal, $style_total);
function save_hr_report_excel0($varr, $style_header, $style_borders, $style_normal, $style_ctitle, $style_stotal, $style_total) {

$months = array(1 => 'Jan', 2 => 'Feb', 3 => 'Mar', 4 => 'Apr', 5 => 'May', 6 => 'Jun', 7 => 'Jul', 8 => 'Aug', 9 => 'Sep', 10 => 'Oct', 11 => 'Nov', 12 => 'Dec');

    $xarr = array();

    foreach($varr[0] as $key => $val) {
        if (!(isset($xarr[$val['tar_mon']]))) {
            $xarr[$val['tar_mon']] = array();
        }
        array_push($xarr[$val['tar_mon']], $val);
    }

        foreach ($xarr as $key => &$val) {
            foreach ($val as $key2 => &$val2) {
                if (!$val2["country_id"]) {
                    $val2["country_id"] = array( 0 => (int)-1, 1 => '');
                    $loc_arr[$key2] = $val2["country_id"][1];
                    $name_arr[$key2] = $val2["employee_id"][1];
                }
                else {
                    $loc_arr[$key2] = $val2["country_id"][1];
                    $name_arr[$key2] = $val2["employee_id"][1];
                }
            }
            array_multisort($loc_arr, SORT_ASC, $name_arr, SORT_ASC, $val);
        }

    $objPHPExcel1 = new PHPExcel();
    $objPHPExcel2 = new PHPExcel();

    $objPHPExcel1->getProperties()->setCreator("OpenERP")
    							 ->setLastModifiedBy("OpenERP")
    							 ->setTitle("Asset Report")
    							 ->setSubject("Office 2007 XLSX Test Document")
    							 ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
    							 ->setKeywords("office 2007 openxml php")
    							 ->setCategory("Test result file");
    $objPHPExcel2->getProperties()->setCreator("OpenERP")
    							 ->setLastModifiedBy("OpenERP")
    							 ->setTitle("Asset Report")
    							 ->setSubject("Office 2007 XLSX Test Document")
    							 ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
    							 ->setKeywords("office 2007 openxml php")
    							 ->setCategory("Test result file");

    $objPHPExcel1->getDefaultStyle()->getFont()->setName('Arial')
                                              ->setSize(10);
    $objPHPExcel2->getDefaultStyle()->getFont()->setName('Arial')
                                              ->setSize(10);

    $objPHPExcel1->removeSheetByIndex(0);
    $objPHPExcel2->removeSheetByIndex(0);

        $mon_name = '';
        $loc_num = 0;

    $VBalance = array();
    $first_time = 1;


    foreach($xarr as $key => &$val) {
        $j = 0;
        while ($val[$j]["country_id"][0] == -1) {
                unset($val[$j]);
                $j++;
        }
        $j=0;
        $yarr = array();
        foreach($val as $key2 => &$val2) {
            $yarr[$j] = $val2;
            $j++;
        }

        $xarr[$key] = &$yarr;
        unset($yarr);
    }

    $vBal = array();

    $lmo = array();
    $cur_year = (int)date('Y');
    foreach($xarr as $k => &$val) {

        $date_pieces = explode("-", $k);
        $year = (int)$date_pieces[0];
        var_dump($date_pieces);

        $j = 0;

        if ($year == $cur_year) {
            $mon_name = $months[(int)$date_pieces[1]];
            var_dump($mon_name);
            var_dump((int)$date_pieces[1]);

            $myWorkSheet1 = new PHPExcel_Worksheet($objPHPExcel1, $mon_name);

            $objPHPExcel1->addSheet($myWorkSheet1);

            $myWorkSheet1->setCellValue('A1', 'Month')
                                                ->setCellValue('B1', $mon_name);

            $myWorkSheet1->setCellValue('A3', 'Employee')
                              ->setCellValue('B3', 'Last Month')
                              ->setCellValue('C3', 'Base')
                              ->setCellValue('D3', 'Onshore')
                              ->setCellValue('E3', 'Offshore')
                              ->setCellValue('F3', 'Off')
                              ->setCellValue('G3', 'Vacation Balance');

            $i1 = 4;

            $first = 1;
            while ($first) {
                if ($val[$j]["is_log"])
                    $first = 0;
                $j++;
            }
            $cou = array('name' => $val[$j]['country_id'][1], 'lm' => 0, 'b' => 0, 'l' => 0,
                         's' => 0, 'o' => 0, 'cm' => 0, 'hc' => 0);
            $tot = array('name' => 'TGT Corporate', 'lm' => 0, 'b' => 0, 'l' => 0,
                         's' => 0, 'o' => 0, 'cm' => 0, 'hc' => 0);
            $lm_keep = 0;
            $su_balance = 0;
            $j = 0;
            while (isset($val[$j]) ) {

                $lmo[$val[$j]["employee_id"][0]] = (isset($lmo[$val[$j]["employee_id"][0]])) ? $lmo[$val[$j]["employee_id"][0]] : 0;

                if ($val[$j]["is_log"]) {
                    echo "$j Val";
                    var_dump($val[$j]['country_id'][1]);
                    echo "$j Cou";
                    var_dump($cou['name']);
                    //if ($first = 1) {
                    $local = $val[$j]['is_loc_engineer'] && ($val[$j]["rotation"][0] == 5);
                    if ($local) var_dump($val[$j]['employee_id']);
                    if ($val[$j]['country_id'][1] == $cou['name']) {
                        //sum
                        $lm_keep += (isset($lmo[$val[$j]["employee_id"][0]])) ? $lmo[$val[$j]["employee_id"][0]] : 0;
                        //$cou['lm'] += $lm_keep;
                        $cou['b'] += $val[$j]["base"] + $val[$j]["wend"];
                        $cou['l'] += $val[$j]["land"];
                        $cou['s'] += $val[$j]["sea"];
                        $cou['o'] += ($local) ? ($val[$j]["dayoff"] + $val[$j]["wend"] + $val[$j]['vac_balance']) : $val[$j]["dayoff"];
                        //$cou['cm'] += $lmo[$val[$j]["employee_id"][0]];
                        $cou['hc']++;
                    } else {
                        //display and reinit
                        $myWorkSheet1->setCellValue('A'.$i1, $cou["name"]." p/head")
                                          ->setCellValue('B'.$i1, $lm_keep/$cou['hc'])
                                          ->setCellValue('C'.$i1, $cou['b']/$cou['hc'] )
                                          ->setCellValue('D'.$i1, $cou['l']/$cou['hc'])
                                          ->setCellValue('E'.$i1, $cou['s']/$cou['hc'] )
                                          ->setCellValue('F'.$i1, $cou['o']/$cou['hc'])
                                          ->setCellValue('G'.$i1, $su_balance/$cou['hc']);
                        $myWorkSheet1->getStyle('A'.$i1.':G'.$i1)->applyFromArray($style_total);

                        //sum the total
                        $tot['lm'] += $lm_keep;
                        $tot['cm'] += $su_balance;
                        $tot['b'] += $cou["b"];
                        $tot['l'] += $cou["l"];
                        $tot['s'] += $cou["s"];
                        $tot['o'] += $cou['o'];
                        //$cou['cm'] = $lmo[$val[$j]["employee_id"][0]];
                        $tot['hc'] += $cou['hc'];

                        $i1 += 2;
                        $cou["name"] = $val[$j]['country_id'][1];
                        $lm_keep = 0;
                        $su_balance = 0;
                        $cou['b'] = $val[$j]["base"] + $val[$j]["wend"];
                        $cou['l'] = $val[$j]["land"];
                        $cou['s'] = $val[$j]["sea"];
                        $cou['o'] = ($local) ? ($val[$j]["dayoff"] + $val[$j]['wend'] + $val[$j]['vac_taken']) : $val[$j]["dayoff"];
                        //$cou['cm'] = $lmo[$val[$j]["employee_id"][0]];
                        $cou['hc'] = 1;
                    }
                    if ($val[$j]["rotation"]) {
                        if ($local) {
                            $vbalance =  $lmo[$val[$j]["employee_id"][0]] + $val[$j]['vac_balance'];
                            $myWorkSheet1->setCellValue('A'.$i1, $val[$j]["employee_id"][1])
                                              ->setCellValue('B'.$i1, $lmo[$val[$j]["employee_id"][0]])
                                              ->setCellValue('C'.$i1, ($val[$j]["base"] ) )
                                              ->setCellValue('D'.$i1, $val[$j]["land"])
                                              ->setCellValue('E'.$i1, $val[$j]["sea"] )
                                              ->setCellValue('F'.$i1, ($val[$j]["dayoff"] + $val[$j]["wend"] + $val[$j]["vac_taken"]))
                                              ->setCellValue('G'.$i1, $vbalance );
                            $lmo[$val[$j]["employee_id"][0]] = $vbalance ;
                            $su_balance += $lmo[$val[$j]["employee_id"][0]];
                        } else {
                            $rot_data = $varr[1][$val[$j]["rotation"][0]-1];
                            $earned = ($val[$j]["base"] + $val[$j]["land"] + $val[$j]["sea"] + $val[$j]["wend"])*($rot_data["days_off"]/$rot_data["days_work"]);
                            $vbalance =  $lmo[$val[$j]["employee_id"][0]] + $earned - ($val[$j]["dayoff"]) + $val[$j]["adjust"];

                            $myWorkSheet1->setCellValue('A'.$i1, $val[$j]["employee_id"][1])
                                              ->setCellValue('B'.$i1, $lmo[$val[$j]["employee_id"][0]])
                                              ->setCellValue('C'.$i1, ($val[$j]["base"] + $val[$j]["wend"]) )
                                              ->setCellValue('D'.$i1, $val[$j]["land"])
                                              ->setCellValue('E'.$i1, $val[$j]["sea"] )
                                              ->setCellValue('F'.$i1, ($val[$j]["dayoff"]))
                                              ->setCellValue('G'.$i1, $vbalance );

                            $lmo[$val[$j]["employee_id"][0]] = $vbalance ;
                            $su_balance += $lmo[$val[$j]["employee_id"][0]];
                        }
                    }
                    else {
                        $myWorkSheet1->setCellValue('A'.$i1, $val[$j]["employee_id"][1])
                                          ->setCellValue('B'.$i1, "Error - no RM specified")
                                          ->setCellValue('C'.$i1, ($val[$j]["base"] + $val[$j]["wend"]))
                                          ->setCellValue('D'.$i1, $val[$j]["land"])
                                          ->setCellValue('E'.$i1, $val[$j]["sea"])
                                          ->setCellValue('F'.$i1, ($val[$j]["dayoff"]))
                                          ->setCellValue('G'.$i1, "--" );
                        $myWorkSheet1->getStyle('A'.$i1.':G'.$i1)->applyFromArray(    array(
                                                                                          'fill' => array(
                                                                                              'type' => PHPExcel_Style_Fill::FILL_SOLID,
                                                                                              'color' => array('rgb' => 'EACCCC')
                                                                                          )
                                                                                      ));
                    }
                    $i1++;

                } //end of algorithm for log engineers
                else {
                    //Array keys will be "employee_id" with the following elements inside:
                    //1. balance for previous years
                    //2. Entitlement
                    //3. Used vacation
                    //balance will be counted
                    if (isset($vBal[$val[$j]["employee_id"][0]])) {
                        $vBal[$val[$j]["employee_id"][0]]['ent'] += $val[$j]["vac_balance"];
                        $vBal[$val[$j]["employee_id"][0]]['used'] += $val[$j]["vac_taken"];
                    } else {
                        $vBal[$val[$j]["employee_id"][0]] = array('name' => $val[$j]["employee_id"][0],
                                        'pre' => $lmo[$val[$j]["employee_id"][0]], 'ent' => $val[$j]["vac_balance"],
                                        'used' => $val[$j]["vac_taken"], 'cbal' => 0);
                        //$vBal[$val[$j]["employee_id"][0]]['cbal'] = $vBal[$val[$j]["employee_id"][0]]['pre'] + $vBal[$val[$j]["employee_id"][0]]['ent'] - $vBal[$val[$j]["employee_id"][0]]['used'];
                    }
                    $lmo[$val[$j]["employee_id"][0]] = $lmo[$val[$j]["employee_id"][0]] + $val[$j]["vac_taken"] + $val[$j]["vac_balance"];
                }

                $j++;
            }
            //display last cou
                        $myWorkSheet1->setCellValue('A'.$i1, $cou["name"]." p/head")
                                          ->setCellValue('B'.$i1, $lm_keep/$cou['hc'])
                                          ->setCellValue('C'.$i1, $cou['b']/$cou['hc'] )
                                          ->setCellValue('D'.$i1, $cou['l']/$cou['hc'])
                                          ->setCellValue('E'.$i1, $cou['s']/$cou['hc'] )
                                          ->setCellValue('F'.$i1, $cou['o']/$cou['hc'])
                                          ->setCellValue('G'.$i1, $su_balance/$cou['hc']);
                        $myWorkSheet1->getStyle('A'.$i1.':G'.$i1)->applyFromArray($style_total);
                        $i1++;

                        //sum the total
                        $tot['lm'] += $lm_keep;
                        $tot['cm'] += $su_balance;
                        $tot['b'] += $cou["b"];
                        $tot['l'] += $cou["l"];
                        $tot['s'] += $cou["s"];
                        $tot['o'] += $cou['o'];
                        //$cou['cm'] = $lmo[$val[$j]["employee_id"][0]];
                        $tot['hc'] += $cou['hc'];

            //display total
                        $myWorkSheet1->setCellValue('A'.$i1, $tot["name"]." p/head")
                                          ->setCellValue('B'.$i1, $tot['lm']/$tot['hc'])
                                          ->setCellValue('C'.$i1, $tot['b']/$tot['hc'] )
                                          ->setCellValue('D'.$i1, $tot['l']/$tot['hc'])
                                          ->setCellValue('E'.$i1, $tot['s']/$tot['hc'] )
                                          ->setCellValue('F'.$i1, $tot['o']/$tot['hc'])
                                          ->setCellValue('G'.$i1, $tot['cm']/$tot['hc']);
                        $myWorkSheet1->getStyle('A'.$i1.':G'.$i1)->applyFromArray($style_total);
                        $i1++;
            $j--;
            $loc_num ++;


            $myWorkSheet1->getStyle('A1:'.$myWorkSheet1->getHighestColumn().$myWorkSheet1->getHighestRow())->applyFromArray($style_normal);
            $myWorkSheet1->getStyle('A3:A'.$myWorkSheet1->getHighestRow())->applyFromArray($style_ctitle);
            $myWorkSheet1->getStyle('A3:'.$myWorkSheet1->getHighestColumn().'3')->applyFromArray($style_total);
            $myWorkSheet1->getStyle(
                'B4:' .
                $myWorkSheet1->getHighestColumn().
                $myWorkSheet1->getHighestRow()
            )->getNumberFormat()->setFormatCode('#,##0.00');

            $myWorkSheet1->getStyle(
                'A3:' .
                $myWorkSheet1->getHighestColumn().
                $myWorkSheet1->getHighestRow()
            )->applyFromArray($style_borders);

        } else {//the end of if year == current year
            while (isset($val[$j]) ) {
                //echo "THIS IS VAL $j";

                $lmo[$val[$j]["employee_id"][0]] = (isset($lmo[$val[$j]["employee_id"][0]])) ? $lmo[$val[$j]["employee_id"][0]] : 0;
                if ($val[$j]["is_log"]) {
                    if ($val[$j]["rotation"]) {
                        $rot_data = $varr[1][$val[$j]["rotation"][0]-1];
                        $earned = ($val[$j]["base"] + $val[$j]["land"] + $val[$j]["sea"] + $val[$j]["wend"])*($rot_data["days_off"]/$rot_data["days_work"]);
                        $vbalance =  $lmo[$val[$j]["employee_id"][0]] + $earned - ($val[$j]["dayoff"] + $val[$j]["wend"]) + $val[$j]["adjust"];
                        $lmo[$val[$j]["employee_id"][0]] = $vbalance ;
                    }
                }
                else {
                    //Array keys will be "employee_id" with the following elements inside:
                    //1. balance for previous years
                    //2. Entitlement
                    //3. Used vacation
                    //balance will be counted
                    if (isset($vBal[$val[$j]["employee_id"][0]])) {
                        $vBal[$val[$j]["employee_id"][0]]['ent'] += $val[$j]["vac_balance"];
                        $vBal[$val[$j]["employee_id"][0]]['used'] += $val[$j]["vac_taken"];
                    } else {
                        $vBal[$val[$j]["employee_id"][0]] = array('name' => $val[$j]["employee_id"][1], 'country' => $val[$j]['country_id'][1],
                                        'pre' => $lmo[$val[$j]["employee_id"][0]], 'ent' => $val[$j]["vac_balance"],
                                        'used' => $val[$j]["vac_taken"], 'cbal' => 0);
                        //$vBal[$val[$j]["employee_id"][0]]['cbal'] = $vBal[$val[$j]["employee_id"][0]]['pre'] + $vBal[$val[$j]["employee_id"][0]]['ent'] - $vBal[$val[$j]["employee_id"][0]]['used'];
                    }
                    $lmo[$val[$j]["employee_id"][0]] = $lmo[$val[$j]["employee_id"][0]] + $val[$j]["vac_taken"] + $val[$j]["vac_balance"];
                }
                $j++;
            }
            echo $k;

        }

    } //foreach

    /* The VACATION BALANCE - display*/
    $myWorkSheet2 = new PHPExcel_Worksheet($objPHPExcel2, 'Balance');
    $myWorkSheet2->setCellValue('A1', 'Employee')
                                  ->setCellValue('B1', 'Previous Year')
                                  ->setCellValue('C1', 'Entitlement')
                                  ->setCellValue('D1', 'Used YTD')
                                  ->setCellValue('E1', 'Balance');
    $objPHPExcel2->addSheet($myWorkSheet2);

    $i2 = 2;
    $first = 1;
    $cou = '';
    $tot_cou = array('pre'=>0, 'ent'=>0, 'used'=>0, 'hc'=>0);
    foreach($vBal as $ke => $vale) {
        if ($first) {
            $cou = $vale['country'];
            $first = 0;
        }
        if ($cou == $vale['country']) {
            $tot_cou['pre'] += $vale['pre'];
            $tot_cou['ent'] += $vale['ent'];
            $tot_cou['used'] += $vale['used'];
            $tot_cou['hc'] ++;
            $myWorkSheet2->setCellValue('A'.$i2, $vale['name'])
                                  ->setCellValue('B'.$i2, $vale['pre'])
                                  ->setCellValue('C'.$i2, $vale['ent'])
                                  ->setCellValue('D'.$i2, abs($vale['used']))
                                  ->setCellValue('E'.$i2, ($vale['pre']+$vale['ent']+$vale['used']));
        }
        else {
            $myWorkSheet2->setCellValue('A'.$i2, $cou." Per Head")
                                  ->setCellValue('B'.$i2, $tot_cou['pre']/$tot_cou['hc'])
                                  ->setCellValue('C'.$i2, $tot_cou['ent']/$tot_cou['hc'])
                                  ->setCellValue('D'.$i2, abs($tot_cou['used'])/$tot_cou['hc'])
                                  ->setCellValue('E'.$i2, (($tot_cou['pre']+$tot_cou['ent']+$tot_cou['used'])/$tot_cou['hc']));
            $myWorkSheet2->getStyle('A'.$i2.':'.$myWorkSheet2->getHighestColumn().$i2)->applyFromArray($style_total);
            //$myWorkSheet2->mergeCells('A'.$i2.':'.$myWorkSheet2->getHighestColumn().$i2);
            $cou = $vale['country'];
            $tot_cou['pre'] = $vale['pre'];
            $tot_cou['ent'] = $vale['ent'];
            $tot_cou['used'] = $vale['used'];
            $tot_cou['hc'] =1;
            $i2++;
            $myWorkSheet2->setCellValue('A'.$i2, $vale['name'])
                                  ->setCellValue('B'.$i2, $vale['pre'])
                                  ->setCellValue('C'.$i2, $vale['ent'])
                                  ->setCellValue('D'.$i2, abs($vale['used']))
                                  ->setCellValue('E'.$i2, ($vale['pre']+$vale['ent']+$vale['used']));
        }
        $i2++;
    }
            $myWorkSheet2->setCellValue('A'.$i2, $cou." Per Head")
                                  ->setCellValue('B'.$i2, $tot_cou['pre']/$tot_cou['hc'])
                                  ->setCellValue('C'.$i2, $tot_cou['ent']/$tot_cou['hc'])
                                  ->setCellValue('D'.$i2, abs($tot_cou['used'])/$tot_cou['hc'])
                                  ->setCellValue('E'.$i2, (($tot_cou['pre']+$tot_cou['ent']+$tot_cou['used'])/$tot_cou['hc']));
            $myWorkSheet2->getStyle('A'.$i2.':'.$myWorkSheet2->getHighestColumn().$i2)->applyFromArray($style_total);
            //$myWorkSheet2->mergeCells('A'.$i2.':'.$myWorkSheet2->getHighestColumn().$i2);

    $myWorkSheet2->getStyle('A1:'.$myWorkSheet2->getHighestColumn().$myWorkSheet2->getHighestRow())->applyFromArray($style_normal);
    $myWorkSheet2->getStyle('A1:A'.$myWorkSheet2->getHighestRow())->applyFromArray($style_ctitle);
    $myWorkSheet2->getStyle('A1:'.$myWorkSheet2->getHighestColumn().'1')->applyFromArray($style_total);
    $myWorkSheet2->getStyle(
                'B2:' .
                $myWorkSheet2->getHighestColumn().
                $myWorkSheet2->getHighestRow()
    )->getNumberFormat()->setFormatCode('#,##0.00');

    $myWorkSheet2->getStyle(
                'A1:' .
                $myWorkSheet2->getHighestColumn().
                $myWorkSheet2->getHighestRow()
    )->applyFromArray($style_borders);

    $objWriter2 = PHPExcel_IOFactory::createWriter($objPHPExcel2, "Excel2007");
    $objWriter2->save("/home/webmaster/tgtoil.com/wp-content/themes/MyCuisine/xmlrpc_lib/xls_files/hr_vacation_balance_v2.xlsx");

    $objWriter1 = PHPExcel_IOFactory::createWriter($objPHPExcel1, "Excel2007");
    $objWriter1->save("/home/webmaster/tgtoil.com/wp-content/themes/MyCuisine/xmlrpc_lib/xls_files/hr_loading_balance_v2.xlsx");

}


/**********
Simple Assets reports - connect, get data, generate file
**********/
$Assets_Report = new Report;
$relation_assets = array("account.asset.asset");
$required_keys_assets = array(
            array(
                       new xmlrpcval('name', "string"),
                       new xmlrpcval('category_id', "array"),
                       new xmlrpcval('code', "string"),
                       new xmlrpcval('location', "array"),
                       new xmlrpcval('purchase_date', "string"),
                       new xmlrpcval('partner_id', "string"),
                       new xmlrpcval('purchase_value', "string"),
                       new xmlrpcval('value_residual', "string"),
                       new xmlrpcval('currency_id', "array"),
                       new xmlrpcval('company_id', "array"),
                       new xmlrpcval('state', "string")
            )
    );
$search_keys_assets = array(
            array(
                        new xmlrpcval(array(new xmlrpcval("id" , "string"),
                        new xmlrpcval("!=","string"),
                        new xmlrpcval("-1","string")),"array"),
            )
    );

$Assets_Report->constructor('rpc_user', 'rpc_password', 'db_name', 'xmlrpc_interface_url', $relation_assets, $required_keys_assets, $search_keys_assets);
$Assets_Report->connect_xmlrpc();

for ($i = 0; $i<count($Assets_Report->relation_names); $i++) {

    $Assets_Report->get_xml_data($i);

    save_asset_report_excel1($Assets_Report->data_arr[$i], $style_header, $style_borders, $style_normal, $style_ctitle, $style_stotal, $style_total);
    $cat_list = array();
    $cat_list = save_asset_report_excel2($Assets_Report->data_arr[$i], $style_header, $style_borders, $style_normal, $style_ctitle, $style_stotal, $style_total);

    save_asset_report_excel3($Assets_Report->data_arr[$i], $cat_list, $style_header, $style_borders, $style_normal, $style_ctitle, $style_stotal, $style_total);
}


/**********
Possibly temp function for generating the assets reports
**********/
function save_asset_report_excel1($varr, $style_header, $style_borders, $style_normal, $style_ctitle, $style_stotal, $style_total) {

    // Create new PHPExcel object
    //echo " Create new PHPExcel object" ;
    $objPHPExcel = new PHPExcel();
    //echo gettype($objPHPExcel);

    usort($varr, function ($a, $b) {
        if ($a["code"][1] == $b["code"][1]) return 0;
        return ($a["code"][1] < $b["code"][1]) ? -1 : 1;
    });

    // Set document properties
    $objPHPExcel->getProperties()->setCreator("OpenERP")
    							 ->setLastModifiedBy("OpenERP")
    							 ->setTitle("Asset Report")
    							 ->setSubject("Office 2007 XLSX Test Document")
    							 ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
    							 ->setKeywords("office 2007 openxml php")
    							 ->setCategory("Test result file");

    // Set default font
    $objPHPExcel->getDefaultStyle()->getFont()->setName('Arial')
                                              ->setSize(10);

    // Add some data, resembling some different data types
    $objPHPExcel->getActiveSheet()->setCellValue('A1', 'Asset Name')
                                  ->setCellValue('B1', 'Asset Category')
                                  ->setCellValue('C1', 'Reference')
                                  ->setCellValue('D1', 'Asset Current Location')
                                  ->setCellValue('E1', 'Purchase Date')
                                  ->setCellValue('F1', 'Partner')
                                  ->setCellValue('G1', 'Gross Value')
                                  ->setCellValue('H1', 'Residual Value')
                                  ->setCellValue('I1', 'Currency')
                                  ->setCellValue('J1', 'Company')
                                  ->setCellValue('K1', 'Status');

    foreach($varr as $k => $xval) {
        $i = $k+2;
        if ($xval["name"] != '') {
            $objPHPExcel->getActiveSheet()->setCellValue('A'.$i, $xval["name"])
                                  ->setCellValue('B'.$i, $xval["category_id"][1])
                                  ->setCellValue('C'.$i, $xval["code"])
                                  ->setCellValue('D'.$i, $xval["location"][1])
                                  ->setCellValue('E'.$i, $xval["purchase_date"])
                                  ->setCellValue('F'.$i, ($xval["partner_id"]) ? $xval["partner_id"][1] : " ")
                                  ->setCellValue('G'.$i, num_form($xval["purchase_value"]))
                                  ->setCellValue('H'.$i, num_form($xval["value_residual"]))
                                  ->setCellValue('I'.$i, $xval["currency_id"][1])
                                  ->setCellValue('J'.$i, $xval["company_id"][1])
                                  ->setCellValue('K'.$i, $xval["state"]);
        }
    }

    $objPHPExcel->getActiveSheet()->getStyle('A1:'.$objPHPExcel->getActiveSheet()->getHighestColumn().$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_normal);
    $objPHPExcel->getActiveSheet()->getStyle('A1:A'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_ctitle);
    $objPHPExcel->getActiveSheet()->getStyle('A1:'.$objPHPExcel->getActiveSheet()->getHighestColumn().'1')->applyFromArray($style_total);

    $objPHPExcel->getActiveSheet()->getStyle(
        'E1:E' .
        $objPHPExcel->getActiveSheet()->getHighestRow()
    )->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_DMYSLASH);

    $objPHPExcel->getActiveSheet()->getStyle(
        'A1:' .
        $objPHPExcel->getActiveSheet()->getHighestColumn().
        $objPHPExcel->getActiveSheet()->getHighestRow()
    )->applyFromArray($style_borders);


    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, "Excel2007");
    $objWriter->save("/home/webmaster/tgtoil.com/wp-content/themes/MyCuisine/xmlrpc_lib/xls_files/assets1.xlsx");

}

function save_asset_report_excel2($xarr, $style_header, $style_borders, $style_normal, $style_ctitle, $style_stotal, $style_total) {

    //prepare the data :
    // $xarr[0] - main data, sort by categories as it goes in $xarr[1], then by locations

    $c_list = array();

    foreach ($xarr as $key => $val) {
        $loc_arr[$key] = $val["location"][1];
        $cat_arr[$key] = $val["child_category"][1];
    }

    array_multisort($loc_arr, SORT_ASC, $cat_arr, SORT_ASC, $xarr);

    // Create new PHPExcel object
    $objPHPExcel = new PHPExcel();

    // Set document properties
    $objPHPExcel->getProperties()->setCreator("OpenERP")
    							 ->setLastModifiedBy("OpenERP")
    							 ->setTitle("Asset Report by Location")
    							 ->setSubject("Office 2007 XLSX Test Document")
    							 ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
    							 ->setKeywords("office 2007 openxml php")
    							 ->setCategory("Test result file");

    // Set default font
    $objPHPExcel->getDefaultStyle()->getFont()->setName('Arial')
                                              ->setSize(10);

    $objPHPExcel->removeSheetByIndex(0);

    $loc_name = '';
    $loc_num = 0;
    $corp = array();
    // Add some data: each location new sheet
    for($e=0; $e < count($xarr); $e++) {

        $loc_name = explode(" (", $xarr[$e]['location'][1]);

        // Create a new worksheet called “My Data”
        $myWorkSheet = new PHPExcel_Worksheet($objPHPExcel, $loc_name[0]);

            //echo $loc_num." ".$loc_name." ";
        $myWorkSheet->setCellValue('A1', 'Location')
                                        ->setCellValue('B1', $loc_name[0]);
        $myWorkSheet->setCellValue('A3', 'Asset Category')
                                                  ->setCellValue('B3', 'Quantity')
                                                  ->setCellValue('C3', 'Gross Value')
                                                  ->setCellValue('D3', 'Residual Value')
                                                  ->setCellValue('E3', 'Currency');

        $k = 0;
        $i = 4;
        $temp_loc = explode(" (",$xarr[$e+$k]['location'][1]);
        $total_gross = 0;
        $total_res = 0;
        $total_num = 0;
        while ($loc_name[0] == $temp_loc[0]) {
            $ca_t = explode(" (", $xarr[$e+$k]["child_category"][1]);
            $cat_id = $ca_t[0];
            $res_val = 0;
            $gros_val = 0;
            if ($cat_id) {
                $ca_t = explode(" (", $xarr[$e+$k]["child_category"][1]);
                $temp_loc = explode(" (",$xarr[$e+$k]['location'][1]);
                $num = 0;
                while (($cat_id == $ca_t[0]) && ($loc_name[0] == $temp_loc[0])) {
                    $res_val += $xarr[$e+$k]["value_residual"];
                    $gros_val += $xarr[$e+$k]["purchase_value"];
                    $k++;
                    $num++;
                    if (isset($xarr[$e+$k])) {
                        $ca_t = explode(" (", $xarr[$e+$k]["child_category"][1]);
                        $temp_loc = explode(" (",$xarr[$e+$k]['location'][1]);
                    }
                    else break;
                }
                if (!isset($corp[$cat_id])) {
                    $corp[$cat_id] = array('num' => 0, 'residual' => 0, 'gross' => 0, 'currency' => $xarr[$e+$k]["currency_id"][1]);
                }
                $corp[$cat_id]['num'] += $num;
                $corp[$cat_id]['residual'] += $res_val;
                $corp[$cat_id]['gross'] += $gros_val;

                $k--;
                $total_num += $num;

                $ca_t = explode(" (", $xarr[$e+$k]["child_category"][1]);
                $c_list[] = $ca_t[0];

                $myWorkSheet->setCellValue('A'.$i, $ca_t[0])
                                              ->setCellValue('B'.$i, $num)
                                              ->setCellValue('C'.$i, num_form($gros_val))
                                              ->setCellValue('D'.$i, num_form($res_val))
                                              ->setCellValue('E'.$i, $xarr[$e+$k]["currency_id"][1]);
                        echo "Index ".$i;
                        var_dump($ca_t[0]);
                        var_dump($num);
                        var_dump($gros_val);
                        var_dump($res_val);
                        var_dump($xarr[$e+$k]["currency_id"][1]);
                $total_gross += $gros_val;
                $total_res += $res_val;
                $i++;
            }
            else echo "cat_id = undefined, array = ".$xarr[$e+$k]["child_category"][1];
            $k++;
            if (!isset($xarr[$e+$k])) break;
        }//end of while location loop
        //add total
        $myWorkSheet->setCellValue('A'.$i, "TOTAL")
                                              ->setCellValue('B'.$i, $total_num)
                                              ->setCellValue('C'.$i, num_form($total_gross))
                                              ->setCellValue('D'.$i, num_form($total_res))
                                              ->setCellValue('E'.$i, $xarr[$e+$k-1]["currency_id"][1]);
        $objPHPExcel->addSheet($myWorkSheet);

        $myWorkSheet->getStyle('A3:A'.$myWorkSheet->getHighestRow())->applyFromArray($style_ctitle);
        $myWorkSheet->getStyle('A3:'.$myWorkSheet->getHighestColumn().'3')->applyFromArray($style_total);
        $myWorkSheet->getStyle('B1')->applyFromArray($style_total);

        $myWorkSheet->getStyle(
            'B4:' .
            $myWorkSheet->getHighestColumn().
            $myWorkSheet->getHighestRow()
        )->applyFromArray($style_normal);

        $myWorkSheet->getStyle(
            'A3:' .
            $myWorkSheet->getHighestColumn().
            $myWorkSheet->getHighestRow()
        )->applyFromArray($style_borders);

        $e += $k-1;
        $loc_num++;

    } //end of for location loop

        // Create a new worksheet called “My Data”
    $myWorkSheet = new PHPExcel_Worksheet($objPHPExcel, 'TGT Global');
        // Attach the “My Data” worksheet as the first worksheet in the PHPExcel object

    $myWorkSheet->setCellValue('A1', 'Location')
                                        ->setCellValue('B1', 'TGT Global');

    $myWorkSheet->setCellValue('A3', 'Asset Category')
                                                  ->setCellValue('B3', 'Quantity')
                                                  ->setCellValue('C3', 'Gross Value')
                                                  ->setCellValue('D3', 'Residual Value')
                                                  ->setCellValue('E3', 'Currency');
    $i=4;
    $num = 0;
    $tot_gross = 0;
    $tot_res = 0;
    $currency = ' ';
    foreach($corp as $type => $value){
        $myWorkSheet->setCellValue('A'.$i, $type)
                    ->setCellValue('B'.$i, $value['num'])
                    ->setCellValue('C'.$i, num_form($value['gross']))
                    ->setCellValue('D'.$i, num_form($value['residual']))
                    ->setCellValue('E'.$i, $value['currency']);
        $num += $value['num'];
        $tot_gross += $value['gross'];
        $tot_res += $value['residual'];
        $currency = $value['currency'];
        echo "Index ".$i;
        var_dump($type);
        var_dump($value);
        $i++;
    }
    $myWorkSheet->setCellValue('A'.$i, 'TGT Global')
                    ->setCellValue('B'.$i, $num)
                    ->setCellValue('C'.$i, num_form($tot_gross))
                    ->setCellValue('D'.$i, num_form($tot_res))
                    ->setCellValue('E'.$i, $currency);

    $objPHPExcel->addSheet($myWorkSheet);
    $myWorkSheet->getStyle('A3:A'.$myWorkSheet->getHighestRow())->applyFromArray($style_ctitle);
    $myWorkSheet->getStyle('A3:'.$myWorkSheet->getHighestColumn().'3')->applyFromArray($style_total);
    $myWorkSheet->getStyle('B1')->applyFromArray($style_total);

        $myWorkSheet->getStyle(
            'B4:' .
            $myWorkSheet->getHighestColumn().
            $myWorkSheet->getHighestRow()
        )->applyFromArray($style_normal);

    $myWorkSheet->getStyle(
        'A3:' .
        $myWorkSheet->getHighestColumn().
        $myWorkSheet->getHighestRow()
    )->applyFromArray($style_borders);

    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, "Excel2007");
    $objWriter->save("/home/webmaster/tgtoil.com/wp-content/themes/MyCuisine/xmlrpc_lib/xls_files/assets_summary.xlsx");

    $result = array_unique($c_list, SORT_STRING);
    sort($result, SORT_STRING);

    return $result;
}



function save_asset_report_excel3($xarr, $cat_list, $style_header, $style_borders, $style_normal, $style_ctitle, $style_stotal, $style_total) {
    //empty
        foreach ($xarr as $key => $val) {
            $loc_arr[$key] = $val["location"][1];
            $cat_arr[$key] = $val["child_category"][1];
        }

        array_multisort($loc_arr, SORT_ASC, $cat_arr, SORT_ASC, $xarr);

        $sh_arr = array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z");


        // Create new PHPExcel object
        $objPHPExcel = new PHPExcel();

        // Set document properties
        $objPHPExcel->getProperties()->setCreator("OpenERP")
        							 ->setLastModifiedBy("OpenERP")
        							 ->setTitle("Asset Report by Location")
        							 ->setSubject("Office 2007 XLSX Test Document")
        							 ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
        							 ->setKeywords("office 2007 openxml php")
        							 ->setCategory("Test result file");

        // Set default font
        $objPHPExcel->getDefaultStyle()->getFont()->setName('Arial')
                                                  ->setSize(10);

        $objPHPExcel->getActiveSheet()->setCellValue('A1', 'Asset Category');
        foreach($cat_list as $ky => $val) {
            $objPHPExcel->getActiveSheet()->setCellValue('A'.($ky +2), $val);
        }

        $loc_name = '';
        $loc_num = 0;
        // Add some data: each location new sheet
        for($e=0; $e < count($xarr); $e++) {
            $cat_loc = 0;
            $loc_name = explode(" (",$xarr[$e]['location'][1]);

            $k = 0;
            $objPHPExcel->getActiveSheet()->setCellValue($sh_arr[$loc_num+1]."1", $loc_name[0]);
            $temp_loc = explode(" (",$xarr[$e+$k]['location'][1]);
            $ass_tot = 0;
            while ($loc_name[0] == $temp_loc[0]) {
                $ca_t = explode(" (", $xarr[$e+$k]["child_category"][1]);
                $cat_id = $ca_t[0];
                $ass_num = 0;
                if ($cat_id) {
                    $temp_loc = explode(" (",$xarr[$e+$k]['location'][1]);
                    while (($cat_id == $ca_t[0] ) && ($loc_name[0] == $temp_loc[0])) {
                        $ass_num++;
                        $k++;
                        if (isset($xarr[$e+$k])) {
                            $ca_t = explode(" (", $xarr[$e+$k]["child_category"][1]);
                            $temp_loc = explode(" (",$xarr[$e+$k]['location'][1]);
                        } else break;
                    }
                    $ass_tot += $ass_num;
                    $k--;

                    for ($ci = $cat_loc; $ci < count($cat_list); $ci++) {
                        $ca_t = explode(" (", $xarr[$e+$k]["child_category"][1]);
                        if ($cat_list[$ci]!= $ca_t[0]) {
                            $objPHPExcel->getActiveSheet()->setCellValue($sh_arr[$loc_num+1].($ci+2), 0);
                        }
                        else {
                            $cat_loc = $ci+1;
                            $objPHPExcel->getActiveSheet()->setCellValue($sh_arr[$loc_num+1].($cat_loc+1), num_form($ass_num));
                            break;
                        }
                    }

                }
                else {
                    //echo "cat_id = undefined, array = ".$xarr[$e+$k]["child_category"][1];
                }
                $k++;
                if (!isset($xarr[$e+$k])) break;
            }//end of while location loop
            //add total

            for ($ci = $cat_loc; $ci < count($cat_list); $ci++) {
                $objPHPExcel->getActiveSheet()->setCellValue($sh_arr[$loc_num+1].($ci+2), 0);
            }
            echo "; REFLECT TOTAL: COLUMN = ".$sh_arr[$loc_num+1].", ROW = ".($cat_loc+1);
            $objPHPExcel->getActiveSheet()->setCellValue($sh_arr[$loc_num+1].(count($cat_list)+1), num_form($ass_tot));
            $e += $k-1;
            $loc_num++;

        } //end of for location loop
        $objPHPExcel->getActiveSheet()->setCellValue('A'.(count($cat_list)+1), 'TOTAL');


    $objPHPExcel->getActiveSheet()->getStyle('A1:A'.$objPHPExcel->getActiveSheet()->getHighestRow())->applyFromArray($style_ctitle);
    $objPHPExcel->getActiveSheet()->getStyle('A1:'.$objPHPExcel->getActiveSheet()->getHighestColumn().'1')->applyFromArray($style_total);

    $objPHPExcel->getActiveSheet()->getStyle(
        'A1:' .
        $objPHPExcel->getActiveSheet()->getHighestColumn().
        $objPHPExcel->getActiveSheet()->getHighestRow()
    )->applyFromArray($style_borders);

    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, "Excel2007");
    $objWriter->save("/home/webmaster/tgtoil.com/wp-content/themes/MyCuisine/xmlrpc_lib/xls_files/assets_quant_summary.xlsx");

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