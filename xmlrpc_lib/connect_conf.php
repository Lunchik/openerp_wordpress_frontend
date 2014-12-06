<?php
include 'xmlrpc.inc';

$user = 'rpc_user';
$password = 'rpc_password';
$dbname = 'db_name';
$server_url = 'xmlrpc_interface_url';

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

?>