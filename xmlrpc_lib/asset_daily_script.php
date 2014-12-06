<?php

include 'xmlrpc.inc';
/** Include PHPExcel **/
require_once $_SERVER['DOCUMENT_ROOT']."/wp-includes/utilites/ExClasses/PHPExcel.php";
require_once 'report_class.php';

/*****Connection config
var $user = ' ';
var $password = ' ';
var $dbname = ' ';
var $server_url = ' ';
**/


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
save_hr_report_excel1($VB_Report->data_arr);
function save_hr_report_excel1($varr) {

    $xarr=$varr[0];

    // Create new PHPExcel object
    echo " Create new PHPExcel object" ;
    $objPHPExcel = new PHPExcel();

    // Set document properties
    echo date('H:i:s') , " Set document properties" , EOL;
    $objPHPExcel->getProperties()->setCreator("OpenERP")
    							 ->setLastModifiedBy("OpenERP")
    							 ->setTitle("Asset Report")
    							 ->setSubject("Office 2007 XLSX Test Document")
    							 ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
    							 ->setKeywords("office 2007 openxml php")
    							 ->setCategory("Test result file");

    // Set default font
    echo date('H:i:s') , " Set default font" , EOL;
    $objPHPExcel->getDefaultStyle()->getFont()->setName('Arial')
                                              ->setSize(10);

    // Add some data, resembling some different data types
    echo date('H:i:s') , " Add some data" , EOL;
    $objPHPExcel->getActiveSheet()->setCellValue('A1', 'Employee')
                                  ->setCellValue('B1', 'Status')
                                  ->setCellValue('C1', 'Rotation Method')
                                  ->setCellValue('D1', 'Vacation Days')
                                  ->setCellValue('E1', 'Vacation Balance');


    foreach($xarr as $k => $xval) {
        $i = $k+2;
        $rot_data = $varr[1][$xval["rotation"][0]-1];
        $vbalance = ($xval["base"] + $xval["land"] + $xval["sea"] + $xval["vacation"])*($rot_data["days_off"]/$rot_data["days_work"]) - $xval["dayoff"];
        $objPHPExcel->getActiveSheet()->setCellValue('A'.$i, $xval["employee_id"][1])
                                  ->setCellValue('B'.$i, ($xval["is_rotator"]) ? "Rotator" : "Resident")
                                  ->setCellValue('C'.$i, ($xval["rotation"]) ? $xval["rotation"][1] : "--")
                                  ->setCellValue('D'.$i, ($xval["is_rotator"]) ? $xval["vacation"] : "--")
                                  ->setCellValue('E'.$i, ($xval["is_rotator"]) ? $vbalance : ($xval["an_leaves"] - $xval["vacation"]));

    }



    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, "Excel2007");
    $objWriter->save("hr_vacation_balance.xlsx");

}

$HL_Report = new Report;
$relation_hl = array("hr.loading.month.report");
$required_keys_hl = array(
            array(
                       new xmlrpcval("employee_id", "array"),
                       new xmlrpcval("res_country", "array"),
                       new xmlrpcval("land1", "int"),
                       new xmlrpcval("sea1", "int"),
                       new xmlrpcval("base1", "int"),
                       new xmlrpcval("dayoff1", "int"),
                       new xmlrpcval("vacation1", "int"),
                       new xmlrpcval("land1", "int"),
                       new xmlrpcval("sea1", "int"),
                       new xmlrpcval("base1", "int"),
                       new xmlrpcval("dayoff1", "int"),
                       new xmlrpcval("vacation1", "int"),
            )
    );
$search_keys_hl = array(
            array( //for hr.vacation.report
                        new xmlrpcval(array(new xmlrpcval("id" , "string"),
                        new xmlrpcval("!=","string"),
                        new xmlrpcval("-1","string")),"array"),
            )
    );
$HL_Report->constructor('rpc_user', 'rpc_password', 'db_name', 'xmlrpc_interface_url', $relation_hl, $required_keys_hl, $search_keys_hl);
$HL_Report->connect_xmlrpc();
for ($i=0; $i < count($HL_Report->relation_names); $i++){
    $HL_Report->get_xml_data($i);
}

save_hr_report_excel2($HL_Report->data_arr);
function save_hr_report_excel2($varr) {

    // Create new PHPExcel object
    $xarr = $varr[0];
    echo " Create new PHPExcel object" ;
    $objPHPExcel = new PHPExcel();

    // Set document properties
    echo date('H:i:s') , " Set document properties" , EOL;
    $objPHPExcel->getProperties()->setCreator("OpenERP")
    							 ->setLastModifiedBy("OpenERP")
    							 ->setTitle("Asset Report")
    							 ->setSubject("Office 2007 XLSX Test Document")
    							 ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
    							 ->setKeywords("office 2007 openxml php")
    							 ->setCategory("Test result file");

    // Set default font
    echo date('H:i:s') , " Set default font" , EOL;
    $objPHPExcel->getDefaultStyle()->getFont()->setName('Arial')
                                              ->setSize(10);

    // Add some data, resembling some different data types
    echo date('H:i:s') , " Add some data" , EOL;
    $objPHPExcel->getActiveSheet()->setCellValue('A1', 'Employee Name')
                                  ->setCellValue('B1', 'Previous Month')
                                  ->setCellValue('G1', 'Current Month')
                                  ->setCellValue('L1', 'Vacation Balance');
    $objPHPExcel->getActiveSheet()->mergeCells('A1:A2')
                                  ->mergeCells('B1:F1')
                                  ->mergeCells('G1:K1')
                                  ->mergeCells('L1:L2');
    $objPHPExcel->getActiveSheet()->setCellValue('B2', 'Base')
                                  ->setCellValue('C2', 'Offshore')
                                  ->setCellValue('D2', 'Onshore')
                                  ->setCellValue('E2', 'Days Off')
                                  ->setCellValue('F2', 'Vacation Days');
    $objPHPExcel->getActiveSheet()->setCellValue('G2', 'Base')
                                  ->setCellValue('H2', 'Offshore')
                                  ->setCellValue('I2', 'Onshore')
                                  ->setCellValue('J2', 'Days Off')
                                  ->setCellValue('K2', 'Vacation Days');

    foreach($xarr as $k => $xval) {
        $i = $k+3;

        $objPHPExcel->getActiveSheet()->setCellValue('A'.$i, $xval["employee_id"][1])
                                  ->setCellValue('B'.$i, $xval["base1"])
                                  ->setCellValue('C'.$i, $xval["sea1"])
                                  ->setCellValue('D'.$i, $xval["land1"])
                                  ->setCellValue('E'.$i, $xval["dayoff1"])
                                  ->setCellValue('F'.$i, $xval["vacation1"])
                                  ->setCellValue('G'.$i, $xval["base2"])
                                  ->setCellValue('H'.$i, $xval["sea2"])
                                  ->setCellValue('I'.$i, $xval["land2"])
                                  ->setCellValue('J'.$i, $xval["dayoff2"])
                                  ->setCellValue('K'.$i, $xval["vacation2"])
                                  ->setCellValue('L'.$i, "coming soon");

    }



    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, "Excel2007");
    $objWriter->save("hr_loading.xlsx");

}

/**********
Simple Assets reports - connect, get data, generate file
**********/

/*$Assets_Report = new Report;
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

    save_asset_report_excel1($Assets_Report->data_arr[$i]);
    $cat_list = array();
    $cat_list = save_asset_report_excel2($Assets_Report->data_arr[$i]);

    save_asset_report_excel3($Assets_Report->data_arr[$i], $cat_list);
}
*/

/**********
Possibly temp function for generating the assets reports
**********/
function save_asset_report_excel1($varr) {

    // Create new PHPExcel object
    echo " Create new PHPExcel object" ;
    $objPHPExcel = new PHPExcel();

    // Set document properties
    echo date('H:i:s') , " Set document properties" , EOL;
    $objPHPExcel->getProperties()->setCreator("OpenERP")
    							 ->setLastModifiedBy("OpenERP")
    							 ->setTitle("Asset Report")
    							 ->setSubject("Office 2007 XLSX Test Document")
    							 ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
    							 ->setKeywords("office 2007 openxml php")
    							 ->setCategory("Test result file");

    // Set default font
    echo date('H:i:s') , " Set default font" , EOL;
    $objPHPExcel->getDefaultStyle()->getFont()->setName('Arial')
                                              ->setSize(10);

    // Add some data, resembling some different data types
    echo date('H:i:s') , " Add some data" , EOL;
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
        $objPHPExcel->getActiveSheet()->setCellValue('A'.$i, $xval["name"])
                                  ->setCellValue('B'.$i, $xval["category_id"][1])
                                  ->setCellValue('C'.$i, $xval["code"])
                                  ->setCellValue('D'.$i, $xval["location"][1])
                                  ->setCellValue('E'.$i, $xval["purchase_date"])
                                  ->setCellValue('F'.$i, ($xval["partner_id"]) ? $xval["partner_id"][1] : " ")
                                  ->setCellValue('G'.$i, $xval["purchase_value"])
                                  ->setCellValue('H'.$i, $xval["value_residual"])
                                  ->setCellValue('I'.$i, $xval["currency_id"][1])
                                  ->setCellValue('J'.$i, $xval["company_id"][1])
                                  ->setCellValue('K'.$i, $xval["state"]);

    }



    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, "Excel2007");
    $objWriter->save("assets1.xlsx");

}

function save_asset_report_excel2($xarr) {

    //prepare the data :
    // $xarr[0] - main data, sort by categories as it goes in $xarr[1], then by locations

    /*usort($xarr, function ($a, $b) {
        if ($a["child_category"][1] == $b["child_category"][1]) return 0;
        return ($a["child_category"][1] < $b["child_category"][1]) ? -1 : 1;
    });

        for ($el =0; $el <count($xarr); $el++) {
            //
            echo $xarr[$el]["child_category"][0]." ".$xarr[$el]["child_category"][1]."\n";
            echo $xarr[$el]["location"][0]." ".$xarr[$el]["location"][1]."\n";
        }

    usort($xarr, function ($a, $b) {
        if ($a["location"][1] == $b["location"][1]) return 0;
        return ($a["location"][1] < $b["location"][1]) ? -1 : 1;
    });

    for ($el =0; $el <count($xarr); $el++) {
        //
        echo $xarr[$el]["child_category"][0]." ".$xarr[$el]["child_category"][1]."\n";
        echo $xarr[$el]["location"][0]." ".$xarr[$el]["location"][1]."\n";
    } */

    $c_list = array();

    foreach ($xarr as $key => $val) {
        $loc_arr[$key] = $val["location"][1];
        $cat_arr[$key] = $val["child_category"][1];
    }

    array_multisort($loc_arr, SORT_ASC, $cat_arr, SORT_ASC, $xarr);

        /*for ($el =0; $el <count($xarr); $el++) {
            //
            echo $xarr[$el]["child_category"][0]." ".$xarr[$el]["child_category"][1]."\n";
            echo $xarr[$el]["location"][0]." ".$xarr[$el]["location"][1]."\n";
        }*/

    // Create new PHPExcel object
    echo " Create new PHPExcel object" ;
    $objPHPExcel = new PHPExcel();

    // Set document properties
    echo date('H:i:s') , " Set document properties" , EOL;
    $objPHPExcel->getProperties()->setCreator("OpenERP")
    							 ->setLastModifiedBy("OpenERP")
    							 ->setTitle("Asset Report by Location")
    							 ->setSubject("Office 2007 XLSX Test Document")
    							 ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
    							 ->setKeywords("office 2007 openxml php")
    							 ->setCategory("Test result file");

    // Set default font
    echo date('H:i:s') , " Set default font" , EOL;
    $objPHPExcel->getDefaultStyle()->getFont()->setName('Arial')
                                              ->setSize(10);


    $loc_name = '';
    $loc_num = 0;
    // Add some data: each location new sheet
    for($e=0; $e < count($xarr); $e++) {

        $loc_name = $xarr[$e]['location'][1];

        // Create a new worksheet called “My Data”
        $myWorkSheet = new PHPExcel_Worksheet($objPHPExcel, $loc_name);
        // Attach the “My Data” worksheet as the first worksheet in the PHPExcel object

        $myWorkSheet->setCellValue('A1', 'Location')
                                        ->setCellValue('B1', $loc_name);

        $myWorkSheet->setCellValue('A3', 'Asset Category')
                                                  ->setCellValue('B3', 'Gross Value')
                                                  ->setCellValue('C3', 'Residual Value')
                                                  ->setCellValue('D3', 'Currency');

        $k = 0;
        $i = 4;

        while ($loc_name == $xarr[$e+$k]['location'][1]) {
            echo date('H:i:s') , " Add some data" , EOL;

            $cat_id = $xarr[$e+$k]["child_category"][0];

            $res_val = 0;
            $gros_val = 0;
            if ($cat_id) {
                while (($cat_id == $xarr[$e+$k]["child_category"][0]) && ($loc_name == $xarr[$e+$k]['location'][1])) {
                    $res_val += $xarr[$e+$k]["value_residual"];
                    $gros_val += $xarr[$e+$k]["purchase_value"];
                    $k++;
                }
                $k--;

                //echo $xarr[$e+$k]["child_category"][1];
                $c_list[] = $xarr[$e+$k]["child_category"][1];

                $myWorkSheet->setCellValue('A'.$i, $xarr[$e+$k]["child_category"][1])
                                              ->setCellValue('B'.$i, $gros_val)
                                              ->setCellValue('C'.$i, $res_val)
                                              ->setCellValue('D'.$i, $xarr[$e+$k]["currency_id"][1]);
                $i++;
            }
            else echo "cat_id = undefined";
            $k++;
        }//end of while location loop
        $objPHPExcel->addSheet($myWorkSheet);
        $e += $k-1;
        $loc_num++;

    } //end of for location loop


    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, "Excel2007");
    $objWriter->save("assets_summary.xlsx");

    $result = array_unique($c_list, SORT_STRING);
    sort($result, SORT_STRING);

    return $result;

}



function save_asset_report_excel3($xarr, $cat_list) {
    //empty
    var_dump($xarr);
        foreach ($xarr as $key => $val) {
            $loc_arr[$key] = $val["location"][1];
            $cat_arr[$key] = $val["child_category"][1];
        }

        array_multisort($loc_arr, SORT_ASC, $cat_arr, SORT_ASC, $xarr);

        $sh_arr = array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z");

        // Create new PHPExcel object
        echo " Create new PHPExcel object" ;
        $objPHPExcel = new PHPExcel();
        //echo gettype($objPHPExcel);

        // Set document properties
        echo date('H:i:s') , " Set document properties" , EOL;
        $objPHPExcel->getProperties()->setCreator("OpenERP")
        							 ->setLastModifiedBy("OpenERP")
        							 ->setTitle("Asset Report by Location")
        							 ->setSubject("Office 2007 XLSX Test Document")
        							 ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
        							 ->setKeywords("office 2007 openxml php")
        							 ->setCategory("Test result file");

        // Set default font
        echo date('H:i:s') , " Set default font" , EOL;
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
            $loc_name = $xarr[$e]['location'][1];

            $k = 0;

            $objPHPExcel->getActiveSheet()->setCellValue($sh_arr[$loc_num+1]."1", $loc_name);
            while ($loc_name == $xarr[$e+$k]['location'][1]) {
                echo date('H:i:s') , " Add some data for ".$loc_name , EOL;

                $cat_id = $xarr[$e+$k]["child_category"][0];

                $ass_num = 0;
                if ($cat_id) {
                    while (($cat_id == $xarr[$e+$k]["child_category"][0]) && ($loc_name == $xarr[$e+$k]['location'][1])) {
                        $ass_num++;
                        $k++;
                    }
                    $k--;

                    for ($ci = $cat_loc; $ci < count($cat_list); $ci++) {
                        if (($cat_list[$ci]!= $xarr[$e+$k]["child_category"][1])) {
                            $objPHPExcel->getActiveSheet()->setCellValue($sh_arr[$loc_num+1].($ci+2), 0);
                        }
                        else {
                            $cat_loc = $ci+1;
                            $objPHPExcel->getActiveSheet()->setCellValue($sh_arr[$loc_num+1].($cat_loc+1), $ass_num);
                            break;
                        }
                    }

                }
                else {
                    echo "cat_id = undefined";
                }
                $k++;
            }//end of while location loop
            for ($ci = $cat_loc; $ci < count($cat_list); $ci++) {
                $objPHPExcel->getActiveSheet()->setCellValue($sh_arr[$loc_num+1].($ci+2), 0);
            }
            $e += $k-1;
            $loc_num++;

        } //end of for location loop


        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, "Excel2007");
        $objWriter->save("assets_quant_summary.xlsx");
}

?>