<?php

include 'xmlrpc.inc';



$report_id = @intval($_GET['report_id']);

//connection config
$user = 'admin';
$password = 'TGTadmin1';
$dbname = 'alpha';
$server_url = 'https://dxb.tgtoil.com/xmlrpc/';

$uid = connect_xmlrpc($server_url, $user, $password, $dbname);

//print $uid;

if ($report_id == 1) { //Sales by Client/Cost Centre
    $relation = "sale.by.cc.client.report";
    $cc_cl = 1;
    //$name = "Client/Cost Centre";
    //print "entered 1";
    $result = print_cc_client($server_url, $relation, $uid, $password, $dbname, $cc_cl);
    }
else if ($report_id == 2) { //Sale by Cost Centre/Country
            $relation = "sale.by.cc.report";
            $cc_cl = 1;
            //$name = "Client/Cost Centre";
            //print "entered 1";
            $result = print_cc_client($server_url, $relation, $uid, $password, $dbname, $cc_cl);
            }
else if ($report_id == 3) { //Sale by Service/Client
            $relation = "sale.by.service.client.report";
            $cc_cl = 1;
            //$name = "Client/Cost Centre";
            //print "entered 1";
            $result = print_cc_client($server_url, $relation, $uid, $password, $dbname, $cc_cl);
            }
else if ($report_id == 4) { //Sales by CService/Country
            $relation = "sale.by.service.c.report";
            $cc_cl = 1;
            //$name = "Client/Cost Centre";
            //print "entered 1";
            $result = print_cc_client($server_url, $relation, $uid, $password, $dbname, $cc_cl);
            }
else if ($report_id == 5) { //Sales by Service/Category
            $relation = "sale.by.service.report";
            $cc_cl = 1;
            //$name = "Client/Cost Centre";
            //print "entered 1";
            $result = print_cc_client($server_url, $relation, $uid, $password, $dbname, $cc_cl);
            }
else if ($report_id == 6) { //Sales by Client/Cost Centre
            $relation = "sale.by.service.cat2.report";
            $cc_cl = 1;
            //$name = "Client/Cost Centre";
            //print "entered 1";
            $result = print_cc_client($server_url, $relation, $uid, $password, $dbname, $cc_cl);
            //$result = -1;
            }
else if ($report_id == 7) { //Sales by Client/Country
            $relation = "sale.by.client.report";
            $cc_cl = 1;
            //$name = "Client/Cost Centre";
            //print "entered 1";
            $result = print_cc_client($server_url, $relation, $uid, $password, $dbname, $cc_cl);
            }
else if ($report_id == 8) { //Sales by Service/Operator
            $relation = "sale.by.service.op.report";
            $cc_cl = 1;
            //$name = "Client/Cost Centre";
            //print "entered 1";
            $result = print_cc_client($server_url, $relation, $uid, $password, $dbname, $cc_cl);
            }
else if ($report_id == 9) { //Sales by Operator/Country
            $relation = "sale.by.operator.report";
            $cc_cl = 1;
            //$name = "Client/Cost Centre";
            //print "entered 1";
            $result = print_cc_client($server_url, $relation, $uid, $password, $dbname, $cc_cl);
            }
else if ($report_id == 10) { //Revenue per job per Country //There's a problem with this report
            $relation1 = "sale.web.avg.cou.report";
            //$relation2 = "sale.web.avg.cou.report2";
            //$result = array();
            $result = print_cc_client($server_url, $relation1, $uid, $password, $dbname, 1);
            //$result[] = print_avg_report2($server_url, $relation2, $uid, $password, $dbname);
            //$result = -1;
            }
else if ($report_id == 11) { //Sales by Client/Cost Centre
            $relation1 = "sale.web.avg.pa.report1";
            $relation2 = "sale.web.avg.pa.report2";
            $result = array();
            $result[] = print_avg_report1($server_url, $relation1, $uid, $password, $dbname);
            $result[] = print_avg_report2($server_url, $relation2, $uid, $password, $dbname);
            }
else if ($report_id == 12) { //Sales by Client/Cost Centre
            $relation1 = "sale.web.avg.opr.reportq1";
            $relation2 = "sale.web.avg.opr.report2";
            $result1 = array();
            $result1[] = print_avg_report1($server_url, $relation1, $uid, $password, $dbname);
            //$result1[] = print_avg_report2($server_url, $relation2, $uid, $password, $dbname);
            $result = json_encode($result1);
            }
else if ($report_id == 13) { //Sales by Client/Cost Centre //Problem here too
            $relation1 = "sale.web.avg.de.report1";
            $relation2 = "sale.web.avg.de.report2";
            $result = array();
            $result[] = print_avg_report1($server_url, $relation1, $uid, $password, $dbname);
            $result[] = print_avg_report2($server_url, $relation2, $uid, $password, $dbname);
            }
else if ($report_id == 14) { //Revenue vs. Target YTD
            $relation = "location.rev";
            $cc_cl = 1;
            //$name = "Client/Cost Centre";
            //print "entered 1";
            $result = get_target_rev($server_url, $relation, $uid, $password, $dbname);
            //$result = -1;
            }
else { //wrong value
    $result = -1;
    }
print $result;



function connect_xmlrpc($server_url, $user, $password, $dbname) {

   if(isset($_COOKIE["user_id"]) == true)  {
       if($_COOKIE["user_id"]>0) {
       //echo $_COOKIE["user_id"];
       }
   }

   $sock = new xmlrpc_client($server_url.'common');

   $msg = new xmlrpcmsg('login');
   $msg->addParam(new xmlrpcval($dbname, "string"));
   $msg->addParam(new xmlrpcval($user, "string"));
   $msg->addParam(new xmlrpcval($password, "string"));
   $resp =  $sock->send($msg);

   $val = $resp->value();
   $id = $val->scalarval();

   setcookie("user_id",$id,time()+3600);
   if($id > 0) {
       return $id;
   }else{
       return -1;
   }

}

function mysort($arr) {
       $sortOrder= array(
            'defcol_one_id',
            'defcol_two_id',
            'jan',
            'feb',
            'mar',
            'q1',
            'apr',
            'may',
            'jun',
            'q2',
            'jul',
            'aug',
            'sep',
            'q3',
            'oct',
            'nov',
            'des',
            'q4',
            'total',
            'id'
       );
       $newarr = array();

       foreach ($sortOrder as $key => $value) {
            $newarr[$value] = $arr[$value];
       }

       return $newarr;
}



function read_rpc_val($server_url, $relation, $user_id, $password, $dbname, $cc_cl) {


    if ($cc_cl == 0) { //
        $defone = 'defcol_two_id';
        $deftwo = 'defcol_one_id';
    }
    else {
        $deftwo = 'defcol_two_id';
        $defone = 'defcol_one_id';
    }
    $sock = new xmlrpc_client($server_url.'object');
//$client->return_type = 'phpvals';// or 'xml';
    $sock->return_type = 'xml';

       $key = array(new xmlrpcval(array(new xmlrpcval("defcol_two_id" , "string"),
                     new xmlrpcval("!=","string"),
                     new xmlrpcval("-1","string")),"array"),
               );
       //$key = array(1,2);

       //echo gettype($key);

       $msg = new xmlrpcmsg('execute');
          $msg->addParam(new xmlrpcval($dbname, "string"));
          $msg->addParam(new xmlrpcval($user_id, "int"));
          $msg->addParam(new xmlrpcval($password, "string"));
          $msg->addParam(new xmlrpcval($relation, "string"));
          $msg->addParam(new xmlrpcval("search", "string"));
          $msg->addParam(new xmlrpcval($key, "array"));
          $resp =  $sock->send($msg);

          if ($resp->faultCode())
            echo "Got error on ->send()".$resp->faultCode();

          $val = $resp->value();
          $v = xmlrpc_decode($val);

          //return gettype($v[0]); int, v[0]=1

          //if ($v === NULL)
          //     echo 'Got invalid response';
          //else
             // check if server sent a fault response
          //if (xmlrpc_is_fault($v))
          //     echo 'Got xmlrpc fault '.$v['faultCode'];


          //print_r ((array)$val);
          //echo gettype($resp);

          $ids_read = array();
          foreach($v AS $index => $value) {
                $ids_read[]=new xmlrpcval($value, "int");
          }

           $data_arr = array();
           $key1 = array(
                       new xmlrpcval('defcol_one_id', "string"),
                       new xmlrpcval('defcol_two_id', "string"),
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
                       new xmlrpcval('total', "string"),
                   );

          //return $val;
          //for ($i=0; i< count($v); i++)
          //foreach ($v as $elem) {


              $msg_read = new xmlrpcmsg('execute');
              $msg_read->addParam(new xmlrpcval($dbname, "string"));
              $msg_read->addParam(new xmlrpcval($user_id, "int"));
              $msg_read->addParam(new xmlrpcval($password, "string"));
              $msg_read->addParam(new xmlrpcval($relation, "string"));
              $msg_read->addParam(new xmlrpcval("read", "string"));
              $msg_read->addParam(new xmlrpcval($ids_read, "array"));
              $msg_read->addParam(new xmlrpcval($key1, "array"));

                 $resp =  $sock->send($msg_read);
              //return $resp;
              if ($resp->faultCode())
                          echo "Got error on ->send(read)".$resp->faultCode();
              $str_val = $resp->value();
              $val = xmlrpc_decode($str_val);

              //return $str_val;
              //$new_val = mysort($val);
              //print_r ($new_val);
              //$data_arr[] = $val;

          //}
return json_encode($val);
}


function get_target_rev($server_url, $relation, $user_id, $password, $dbname) {

    $sock = new xmlrpc_client($server_url.'object');
//$client->return_type = 'phpvals';// or 'xml';
    $sock->return_type = 'xml';

    $key = array(new xmlrpcval(array(new xmlrpcval("year" , "string"),
                         new xmlrpcval("=","string"),
                         new xmlrpcval("2014","string")),"array"),
                );
       //$key = array(1,2);

       //echo gettype($key);

       $msg = new xmlrpcmsg('execute');
          $msg->addParam(new xmlrpcval($dbname, "string"));
          $msg->addParam(new xmlrpcval($user_id, "int"));
          $msg->addParam(new xmlrpcval($password, "string"));
          $msg->addParam(new xmlrpcval($relation, "string"));
          $msg->addParam(new xmlrpcval("search", "string"));
          $msg->addParam(new xmlrpcval($key, "array"));
          $resp =  $sock->send($msg);

          if ($resp->faultCode())
            echo "Got error on ->send()".$resp->faultCode();

          $val = $resp->value();
          //return $val;
          $v = xmlrpc_decode($val);

          //return gettype($v[0]); int, v[0]=1

          //if ($v === NULL)
          //     echo 'Got invalid response';
          //else
             // check if server sent a fault response
          //if (xmlrpc_is_fault($v))
          //     echo 'Got xmlrpc fault '.$v['faultCode'];


          //print_r ((array)$val);
          //echo gettype($resp);

          $ids_read = array();
          foreach($v AS $index => $value) {
                $ids_read[]=new xmlrpcval($value, "int");
          }

           $data_arr = array();
           $key1 = array(
                       new xmlrpcval('country_id', "string"),
                       new xmlrpcval('tar_mo', "string"),
                       new xmlrpcval('mon', "string")
                   );

          //return $val;
          //for ($i=0; i< count($v); i++)
          //foreach ($v as $elem) {


              $msg_read = new xmlrpcmsg('execute');
              $msg_read->addParam(new xmlrpcval($dbname, "string"));
              $msg_read->addParam(new xmlrpcval($user_id, "int"));
              $msg_read->addParam(new xmlrpcval($password, "string"));
              $msg_read->addParam(new xmlrpcval($relation, "string"));
              $msg_read->addParam(new xmlrpcval("read", "string"));
              $msg_read->addParam(new xmlrpcval($ids_read, "array"));
              $msg_read->addParam(new xmlrpcval($key1, "array"));

                 $resp =  $sock->send($msg_read);
              //return $resp;
              if ($resp->faultCode())
                          echo "Got error on ->send(read)".$resp->faultCode();
              $str_val = $resp->value();
              $val = xmlrpc_decode($str_val);

              //return $str_val;
              //$new_val = mysort($val);
              //print_r ($new_val);
              //$data_arr[] = $val;

          //}
return json_encode($val);
}

function print_cc_client($server_url, $relation, $uid, $password, $dbname, $cc_cl, $name) {
    //$retval = '';
    //$retval .= "<h1>Client/Cost Centre Revenue report</h1><br>";

    $retval .= read_rpc_val($server_url, $relation, $uid, $password, $dbname, $cc_cl, $name);
    return $retval;
}


function print_avg_report1($server_url, $relation, $uid, $password, $dbname) {

    $sock = new xmlrpc_client($server_url.'object');
//$client->return_type = 'phpvals';// or 'xml';
    $sock->return_type = 'xml';

    print $uid;
       $key = array(new xmlrpcval(array(new xmlrpcval("defcol_two_id" , "string"),
                     new xmlrpcval("!=","string"),
                     new xmlrpcval("-1","string")),"array"),
               );

       $msg = new xmlrpcmsg('execute');
          $msg->addParam(new xmlrpcval($dbname, "string"));
          $msg->addParam(new xmlrpcval($uid, "int"));
          $msg->addParam(new xmlrpcval($password, "string"));
          $msg->addParam(new xmlrpcval($relation, "string"));
          $msg->addParam(new xmlrpcval("search", "string"));
          $msg->addParam(new xmlrpcval($key, "array"));
          $resp =  $sock->send($msg);

          if ($resp->faultCode())
            echo "Got error on ->send()".$resp->faultCode();

          $val = $resp->value();
          $v = xmlrpc_decode($val);

          $ids_read = array();
          foreach($v AS $index => $value) {
                $ids_read[]=new xmlrpcval($value, "int");
          }

           $data_arr = array();
           $key1 = array(
                       new xmlrpcval('defcol_one_id', "string"),
                       new xmlrpcval('defcol_two_id', "string"),
                       new xmlrpcval('jan', "string"),
                       new xmlrpcval('jjan', "string"),
                       new xmlrpcval('feb', "string"),
                       new xmlrpcval('jfeb', "string"),
                       new xmlrpcval('mar', "string"),
                       new xmlrpcval('jmar', "string"),
                       new xmlrpcval('q1', "string"),
                       new xmlrpcval('jq1', "string"),
                       /*new xmlrpcval('apr', "string"),
                       new xmlrpcval('japr', "string"),
                       new xmlrpcval('may', "string"),
                       new xmlrpcval('jmay', "string"),
                       new xmlrpcval('jun', "string"),
                       new xmlrpcval('jjun', "string"),
                       new xmlrpcval('q2', "string"),
                       new xmlrpcval('jq2', "string"),
                       new xmlrpcval('total', "string"),
                       new xmlrpcval('jtotal', "string"),*/
                   );

          //return $val;
          //for ($i=0; i< count($v); i++)
          //foreach ($v as $elem) {


              $msg_read = new xmlrpcmsg('execute');
              $msg_read->addParam(new xmlrpcval($dbname, "string"));
              $msg_read->addParam(new xmlrpcval($uid, "int"));
              $msg_read->addParam(new xmlrpcval($password, "string"));
              $msg_read->addParam(new xmlrpcval($relation, "string"));
              $msg_read->addParam(new xmlrpcval("read", "string"));
              $msg_read->addParam(new xmlrpcval($ids_read, "array"));
              $msg_read->addParam(new xmlrpcval($key1, "array"));

                 $resp =  $sock->send($msg_read);
              //return $resp;
              if ($resp->faultCode())
                          echo "Got error on ->send(read)".$resp->faultCode();
              $str_val = $resp->value();
              $val = xmlrpc_decode($str_val);

              //return $str_val;
              //$new_val = mysort($val);
              //print_r ($new_val);
              //$data_arr[] = $val;

          //}
return $val;
}

function print_avg_report2($server_url, $relation, $uid, $password, $dbname) {

    $sock = new xmlrpc_client($server_url.'object');
//$client->return_type = 'phpvals';// or 'xml';
    $sock->return_type = 'xml';

       $key = array(new xmlrpcval(array(new xmlrpcval("defcol_two_id" , "string"),
                     new xmlrpcval("!=","string"),
                     new xmlrpcval("-1","string")),"array"),
               );

       $msg = new xmlrpcmsg('execute');
          $msg->addParam(new xmlrpcval($dbname, "string"));
          $msg->addParam(new xmlrpcval($uid, "int"));
          $msg->addParam(new xmlrpcval($password, "string"));
          $msg->addParam(new xmlrpcval($relation, "string"));
          $msg->addParam(new xmlrpcval("search", "string"));
          $msg->addParam(new xmlrpcval($key, "array"));
          $resp =  $sock->send($msg);

          if ($resp->faultCode())
            echo "Got error on ->send()".$resp->faultCode();

          $val = $resp->value();
          $v = xmlrpc_decode($val);

          $ids_read = array();
          foreach($v AS $index => $value) {
                $ids_read[]=new xmlrpcval($value, "int");
          }

           $data_arr = array();
           $key1 = array(
                       new xmlrpcval('defcol_one_id', "string"),
                       new xmlrpcval('defcol_two_id', "string"),
                       new xmlrpcval('jul', "string"),
                       new xmlrpcval('jjul', "string"),
                       new xmlrpcval('aug', "string"),
                       new xmlrpcval('jaug', "string"),
                       new xmlrpcval('sep', "string"),
                       new xmlrpcval('jsep', "string"),
                       new xmlrpcval('q3', "string"),
                       new xmlrpcval('jq3', "string"),
                       new xmlrpcval('oct', "string"),
                       new xmlrpcval('joct', "string"),
                       new xmlrpcval('nov', "string"),
                       new xmlrpcval('jnov', "string"),
                       new xmlrpcval('des', "string"),
                       new xmlrpcval('jdes', "string"),
                       new xmlrpcval('q4', "string"),
                       new xmlrpcval('jq4', "string"),
                       new xmlrpcval('total', "string"),
                       new xmlrpcval('jtotal', "string"),
                   );

              $msg_read = new xmlrpcmsg('execute');
              $msg_read->addParam(new xmlrpcval($dbname, "string"));
              $msg_read->addParam(new xmlrpcval($uid, "int"));
              $msg_read->addParam(new xmlrpcval($password, "string"));
              $msg_read->addParam(new xmlrpcval($relation, "string"));
              $msg_read->addParam(new xmlrpcval("read", "string"));
              $msg_read->addParam(new xmlrpcval($ids_read, "array"));
              $msg_read->addParam(new xmlrpcval($key1, "array"));

                 $resp =  $sock->send($msg_read);
              //return $resp;
              if ($resp->faultCode())
                          echo "Got error on ->send(read)".$resp->faultCode();
              $str_val = $resp->value();
              $val = xmlrpc_decode($str_val);

return $val;
}


?>