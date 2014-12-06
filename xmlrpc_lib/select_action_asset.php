<?php

include 'xmlrpc.inc';



$report_id = @intval($_GET['report_id']);

//connection config
$user = 'admin';
$password = 'TGTadmin1';
$dbname = 'alpha';
$server_url = 'https://dxb.tgtoil.com/xmlrpc/';

$relation = array("account.asset.asset", "account.asset.category", "tgt.location");

$required_keys = array(
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
            ),
            array(
                       new xmlrpcval('id', "string"),
                       new xmlrpcval('name', "string")
            ),
            array(
                       new xmlrpcval('id', "string"),
                       new xmlrpcval('name', "string")
            )
    );

$uid = connect_xmlrpc($server_url, $user, $password, $dbname);

//print $uid;

if ($report_id == 1) { //Sales by Client/Cost Centre
    //$relation = "sale.by.cc.client.report";
    $result = print_cc_client($server_url, $relation, $uid, $password, $dbname, $report_id);
    //$result = -1;
    }
else if ($report_id == 2) { //Sales by Client/Cost Centre
         //$relation = "sale.by.cc.client.report";
         $result = print_cc_client($server_url, $relation, $uid, $password, $dbname, $report_id);
         //$result = -1;
         }
else if ($report_id == 3) { //Sales by Client/Cost Centre
         //$relation = "sale.by.cc.client.report";
         $result = print_cc_client($server_url, $relation, $uid, $password, $dbname, $report_id);
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


function get_target_rev($server_url, $relation, $user_id, $password, $dbname, $rep_id) {

    $sock = new xmlrpc_client($server_url.'object');
//$client->return_type = 'phpvals';// or 'xml';
    $sock->return_type = 'xml';

    //return "user ".$user_id." db ".$dbname. " rel ".$relation[0];

    $key1 = array(new xmlrpcval(array(new xmlrpcval("id" , "string"),
                         new xmlrpcval("!=","string"),
                         new xmlrpcval("-1","string")),"array"),
                );
       //$key = array(1,2);
       //$data_arr = array();
       //echo gettype($key);

          $msg = new xmlrpcmsg('execute');
          $msg->addParam(new xmlrpcval($dbname, "string"));
          $msg->addParam(new xmlrpcval($user_id, "int"));
          $msg->addParam(new xmlrpcval($password, "string"));
          $msg->addParam(new xmlrpcval($relation[$rep_id], "string"));
          $msg->addParam(new xmlrpcval("search", "string"));
          $msg->addParam(new xmlrpcval($key1, "array"));
          $resp =  $sock->send($msg);

          if ($resp->faultCode())
            echo "Got error on ->send()".$resp->faultCode();

          $val = $resp->value();
          //return $val;
          $v = xmlrpc_decode($val);
          //return $v;
          $ids_read = array();
          foreach($v AS $index => $value) {
                $ids_read[]=new xmlrpcval($value, "int");
          }

              $msg_read = new xmlrpcmsg('execute');
              $msg_read->addParam(new xmlrpcval($dbname, "string"));
              $msg_read->addParam(new xmlrpcval($user_id, "int"));
              $msg_read->addParam(new xmlrpcval($password, "string"));
              $msg_read->addParam(new xmlrpcval($relation[$rep_id], "string"));
              $msg_read->addParam(new xmlrpcval("read", "string"));
              $msg_read->addParam(new xmlrpcval($ids_read, "array"));
              $msg_read->addParam(new xmlrpcval($required_keys[$rep_id], "array"));

                 $resp =  $sock->send($msg_read);
              //return $resp;
              if ($resp->faultCode())
                          echo "Got error on ->send(read)".$resp->faultCode();
              $str_val = $resp->value();
              $val = xmlrpc_decode($str_val);

              //$data_arr[] = $val;

return json_encode($val);
}

function print_cc_client($server_url, $relation, $user_id, $password, $dbname, $report_id) {
    //$retval = '';
    //$retval .= "<h1>Client/Cost Centre Revenue report</h1><br>";

    $retval = get_target_rev($server_url, $relation, $user_id, $password, $dbname, $report_id-1);
    return $retval;
}





?>